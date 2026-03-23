use sysinfo::System;
use serde::Serialize;
use tokio::net::TcpListener;
use tokio_tungstenite::accept_async;
use tokio_tungstenite::tungstenite::Message;
use futures_util::SinkExt;
use tokio::sync::{broadcast, mpsc};
use tauri::Manager;
use std::collections::{HashMap, VecDeque};
use std::sync::{Arc, Mutex};

#[cfg(target_os = "windows")]
use std::sync::{
    mpsc as std_mpsc,
    Mutex as StdMutex,
};

#[cfg(target_os = "windows")]
use std::time::Duration;

#[cfg(target_os = "windows")]
use windows::{
    core::{implement, ComInterface, Error as WinError, HSTRING, HRESULT, IUnknown},
    Win32::{
        Media::{
            Audio::{
                ActivateAudioInterfaceAsync,
                IActivateAudioInterfaceAsyncOperation,
                IActivateAudioInterfaceCompletionHandler,
                IActivateAudioInterfaceCompletionHandler_Impl,
                IAudioCaptureClient,
                IAudioClient,
                IMMDevice,
                IMMDeviceEnumerator,
                MMDeviceEnumerator,
                WAVEFORMATEX,
                AUDIOCLIENT_ACTIVATION_PARAMS,
                AUDIOCLIENT_ACTIVATION_PARAMS_0,
                AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK,
                AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS,
                AUDCLNT_BUFFERFLAGS_SILENT,
                AUDCLNT_E_DEVICE_INVALIDATED,
                AUDCLNT_E_WRONG_ENDPOINT_TYPE,
                AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK,
                PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
                VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
                eMultimedia,
                eRender,
            },
        },
        System::Com::{
            CoCreateInstance,
            CoInitializeEx,
            CoTaskMemFree,
            CoUninitialize,
            CLSCTX_ALL,
            BLOB,
            StructuredStorage::{PROPVARIANT, PROPVARIANT_0, PROPVARIANT_0_0, PROPVARIANT_0_0_0},
            COINIT_MULTITHREADED,
        },
        System::Variant::VT_BLOB,
    },
};

struct AppState {
    capture_tx: broadcast::Sender<Vec<u32>>,
    latest_capture_pids: Arc<Mutex<Vec<u32>>>,
}

// 定义我们要传给前端的数据格式
#[derive(Serialize)]
struct ProcessInfo {
    pid: u32,
    name: String,
    memory_mb: u64,
}

// 暴漏给前端 JS 调用的命令：获取活跃进程雷达
#[tauri::command]
fn get_active_processes() -> Vec<ProcessInfo> {
    let mut sys = System::new_all();
    sys.refresh_all();

    let mut grouped: HashMap<String, (u32, u64)> = HashMap::new();

    for (pid, process) in sys.processes() {
        let mut root_pid = *pid;
        let mut cursor_pid = *pid;

        loop {
            let Some(current_proc) = sys.processes().get(&cursor_pid) else {
                break;
            };

            let Some(parent_pid) = current_proc.parent() else {
                break;
            };

            let Some(parent_proc) = sys.processes().get(&parent_pid) else {
                break;
            };

            if parent_proc.name() == current_proc.name() {
                root_pid = parent_pid;
                cursor_pid = parent_pid;
                continue;
            }

            break;
        }

        let name = process.name().to_string();
        let mem_bytes = process.memory();

        let entry = grouped.entry(name).or_insert((root_pid.as_u32(), 0));
        entry.1 = entry.1.saturating_add(mem_bytes);
    }

    let mut process_list: Vec<ProcessInfo> = grouped
        .into_iter()
        .map(|(name, (pid, memory_bytes))| ProcessInfo {
            pid,
            name,
            memory_mb: memory_bytes / 1024 / 1024,
        })
        .filter(|item| item.memory_mb >= 30)
        .collect();

    process_list.sort_by(|a, b| b.memory_mb.cmp(&a.memory_mb));
    process_list
}

#[tauri::command]
async fn start_capture(pid: u32, state: tauri::State<'_, AppState>) -> Result<u32, String> {
    start_capture_multi(vec![pid], state).await
}

#[tauri::command]
async fn start_capture_multi(pids: Vec<u32>, state: tauri::State<'_, AppState>) -> Result<u32, String> {
    let mut normalized = Vec::new();
    for pid in pids {
        if pid == 0 {
            continue;
        }
        if !normalized.contains(&pid) {
            normalized.push(pid);
        }
    }

    if normalized.is_empty() {
        return Err("没有可用的 PID，无法启动采集".into());
    }

    if let Ok(mut guard) = state.latest_capture_pids.lock() {
        *guard = normalized.clone();
    }

    let _ = state.capture_tx.send(normalized);

    query_mix_sample_rate()
}

#[cfg(target_os = "windows")]
fn query_mix_sample_rate() -> Result<u32, String> {
    let mut should_uninit = false;
    match unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) } {
        Ok(()) => {
            should_uninit = true;
        }
        Err(e) if e.code() == HRESULT(0x80010106u32 as i32) => {}
        Err(e) => return Err(hr_msg("查询采样率时 CoInitializeEx 失败", e.code())),
    }

    let result = (|| -> Result<u32, String> {
        let device = get_default_render_device()?;
        let client: IAudioClient = match unsafe { device.Activate(CLSCTX_ALL, None) } {
            Ok(v) => v,
            Err(e) => return Err(hr_msg("查询采样率时 Activate(dummy_client) 失败", e.code())),
        };

        let mix_format_ptr = match unsafe { client.GetMixFormat() } {
            Ok(v) => v,
            Err(e) => return Err(hr_msg("查询采样率时 GetMixFormat 失败", e.code())),
        };

        let sample_rate = unsafe { (*mix_format_ptr).nSamplesPerSec };
        unsafe { CoTaskMemFree(Some(mix_format_ptr as *const std::ffi::c_void)) };
        Ok(sample_rate)
    })();

    if should_uninit {
        unsafe { CoUninitialize() };
    }

    result
}

#[cfg(not(target_os = "windows"))]
fn query_mix_sample_rate() -> Result<u32, String> {
    Err("当前平台不支持采样率查询，仅 Windows 可用".into())
}

async fn start_audio_pump(capture_tx: broadcast::Sender<Vec<u32>>, latest_capture_pids: Arc<Mutex<Vec<u32>>>) {
    let addr = "127.0.0.1:9001";
    // 建立本地服务
    let listener = match TcpListener::bind(&addr).await {
        Ok(v) => v,
        Err(e) => {
            println!("❌ 无法绑定 9001 端口: {e}");
            return;
        }
    };
    println!("🎧 音频专属高铁已发车，监听端口: {}", addr);

    // 死循环：等待前端大厅来连接
    while let Ok((stream, _)) = listener.accept().await {
        let mut pid_rx = capture_tx.subscribe();
        let latest_capture_pids_ref = latest_capture_pids.clone();
        tokio::spawn(async move {
            let mut ws_stream = match accept_async(stream).await {
                Ok(v) => v,
                Err(e) => {
                    println!("❌ WebSocket 握手失败: {e}");
                    return;
                }
            };
            println!("✅ 前端 JS 已连接音频 WebSocket，等待 start_capture/start_capture_multi 指令...");

            let mut pending_pids: Option<Vec<u32>> = match latest_capture_pids_ref.lock() {
                Ok(guard) if !guard.is_empty() => Some(guard.clone()),
                _ => None,
            };

            loop {
                let pids = if let Some(v) = pending_pids.take() {
                    v
                } else {
                    match pid_rx.recv().await {
                        Ok(v) => v,
                        Err(e) => {
                            println!("⚠️ 接收采集 PID 失败: {e}");
                            break;
                        }
                    }
                };

                if pids.is_empty() {
                    continue;
                }

                println!("🚀 准备启动 WASAPI 多进程捕获, pids={:?}", pids);

                let mut pcm_rxs = Vec::new();
                let mut stop_txs = Vec::new();
                let mut capture_handles = Vec::new();

                for pid in &pids {
                    let (pcm_tx, pcm_rx) = mpsc::channel::<Vec<u8>>(64);
                    let (stop_tx, stop_rx) = std_mpsc::channel::<()>();
                    let capture_pid = *pid;
                    let capture_handle = tokio::task::spawn_blocking(move || {
                        run_capture_for_pid(capture_pid, pcm_tx, stop_rx)
                    });

                    pcm_rxs.push(pcm_rx);
                    stop_txs.push(stop_tx);
                    capture_handles.push(capture_handle);
                }

                let mut closed_flags = vec![false; pcm_rxs.len()];
                let mut pending_samples = vec![VecDeque::<f32>::new(); pcm_rxs.len()];
                const MIX_BLOCK_SAMPLES: usize = 480;

                let mut should_restart = false;
                let mut should_exit = false;

                loop {
                    let mut has_new_data = false;

                    for (index, rx) in pcm_rxs.iter_mut().enumerate() {
                        if closed_flags[index] {
                            continue;
                        }

                        loop {
                            match rx.try_recv() {
                                Ok(chunk) => {
                                    append_f32_samples_from_bytes(&mut pending_samples[index], &chunk);
                                    has_new_data = true;
                                }
                                Err(mpsc::error::TryRecvError::Empty) => break,
                                Err(mpsc::error::TryRecvError::Disconnected) => {
                                    closed_flags[index] = true;
                                    break;
                                }
                            }
                        }
                    }

                    let active_indices: Vec<usize> = closed_flags
                        .iter()
                        .enumerate()
                        .filter_map(|(idx, closed)| if !*closed { Some(idx) } else { None })
                        .collect();

                    if active_indices.is_empty() {
                        break;
                    }

                    let min_available = active_indices
                        .iter()
                        .map(|idx| pending_samples[*idx].len())
                        .min()
                        .unwrap_or(0);

                    if min_available >= MIX_BLOCK_SAMPLES {
                        let mut mixed_chunk = Vec::with_capacity(MIX_BLOCK_SAMPLES * 4);

                        for _ in 0..MIX_BLOCK_SAMPLES {
                            let mut sum = 0.0f32;
                            let mut count = 0usize;

                            for idx in &active_indices {
                                if let Some(sample) = pending_samples[*idx].pop_front() {
                                    sum += sample;
                                    count += 1;
                                }
                            }

                            let mixed_sample = if count == 0 {
                                0.0
                            } else {
                                (sum / count as f32).clamp(-1.0, 1.0)
                            };
                            mixed_chunk.extend_from_slice(&mixed_sample.to_le_bytes());
                        }

                        if ws_stream.send(Message::Binary(mixed_chunk.into())).await.is_err() {
                            println!("⚠️ 前端断开连接，停止音频推流");
                            should_exit = true;
                            break;
                        }
                    }

                    if closed_flags.iter().all(|v| *v) {
                        break;
                    }

                    // 非阻塞检查：如果用户再次调用 start_capture/start_capture_multi，立即切换目标 PID 列表。
                    match pid_rx.try_recv() {
                        Ok(new_pids) => {
                            println!("🔁 收到新的 PID 列表指令，切换采集目标: {:?}", new_pids);
                            pending_pids = Some(new_pids);
                            should_restart = true;
                            break;
                        }
                        Err(broadcast::error::TryRecvError::Empty) => {}
                        Err(broadcast::error::TryRecvError::Closed) => {
                            println!("⚠️ PID 控制通道已关闭，停止捕获");
                            should_exit = true;
                            break;
                        }
                        Err(broadcast::error::TryRecvError::Lagged(skipped)) => {
                            println!("⚠️ PID 控制消息积压，跳过 {skipped} 条，仅使用最新 PID 列表");
                        }
                    }

                    if min_available < MIX_BLOCK_SAMPLES && !has_new_data {
                        tokio::time::sleep(std::time::Duration::from_millis(2)).await;
                    }
                }

                for stop_tx in &stop_txs {
                    let _ = stop_tx.send(());
                }

                for handle in capture_handles {
                    match handle.await {
                        Ok(Ok(())) => {}
                        Ok(Err(e)) => println!("⚠️ WASAPI 捕获线程结束: {e}"),
                        Err(e) => println!("⚠️ WASAPI 捕获任务 Join 失败: {e}"),
                    }
                }

                if should_exit {
                    return;
                }

                if should_restart {
                    continue;
                }

                // 当前捕获自然结束后，等待下一次 start_capture(pid)
            }
        });
    }
}

#[cfg(target_os = "windows")]
#[implement(IActivateAudioInterfaceCompletionHandler)]
struct ActivateAudioInterfaceHandler {
    tx: StdMutex<Option<std_mpsc::Sender<Result<IAudioClient, WinError>>>>,
}

#[cfg(target_os = "windows")]
impl ActivateAudioInterfaceHandler {
    fn new(tx: std_mpsc::Sender<Result<IAudioClient, WinError>>) -> Self {
        Self {
            tx: StdMutex::new(Some(tx)),
        }
    }

    fn send_once(&self, value: Result<IAudioClient, WinError>) {
        if let Ok(mut guard) = self.tx.lock() {
            if let Some(tx) = guard.take() {
                let _ = tx.send(value);
            }
        }
    }
}

#[cfg(target_os = "windows")]
impl IActivateAudioInterfaceCompletionHandler_Impl for ActivateAudioInterfaceHandler {
    fn ActivateCompleted(
        &self,
        activateoperation: Option<&IActivateAudioInterfaceAsyncOperation>,
    ) -> windows::core::Result<()> {
        let result = (|| -> Result<IAudioClient, WinError> {
            let op = activateoperation.ok_or_else(|| {
                WinError::new(HRESULT(0x80004005u32 as i32), HSTRING::from("activateoperation 为空"))
            })?;

            let mut activate_hr = HRESULT(0);
            let mut activated: Option<IUnknown> = None;
            unsafe { op.GetActivateResult(&mut activate_hr, &mut activated)? };
            activate_hr.ok()?;

            let unknown = activated.ok_or_else(|| {
                WinError::new(HRESULT(0x80004005u32 as i32), HSTRING::from("未返回激活接口"))
            })?;

            unknown.cast::<IAudioClient>()
        })();

        self.send_once(result);
        Ok(())
    }
}

#[cfg(target_os = "windows")]
fn activate_process_loopback_audio_client(pid: u32) -> Result<IAudioClient, String> {
    let activation_params = AUDIOCLIENT_ACTIVATION_PARAMS {
        ActivationType: AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK,
        Anonymous: AUDIOCLIENT_ACTIVATION_PARAMS_0 {
            ProcessLoopbackParams: AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS {
                TargetProcessId: pid,
                ProcessLoopbackMode: PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
            },
        },
    };

    let blob = BLOB {
        cbSize: std::mem::size_of::<AUDIOCLIENT_ACTIVATION_PARAMS>() as u32,
        pBlobData: (&activation_params as *const AUDIOCLIENT_ACTIVATION_PARAMS) as *mut u8,
    };

    let prop = PROPVARIANT {
        Anonymous: PROPVARIANT_0 {
            Anonymous: std::mem::ManuallyDrop::new(PROPVARIANT_0_0 {
                vt: VT_BLOB,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: PROPVARIANT_0_0_0 { blob },
            }),
        },
    };

    let (tx, rx) = std_mpsc::channel::<Result<IAudioClient, WinError>>();
    let handler: IActivateAudioInterfaceCompletionHandler =
        ActivateAudioInterfaceHandler::new(tx).into();

    let op = unsafe {
        ActivateAudioInterfaceAsync(
            VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
            &IAudioClient::IID,
            Some(&prop as *const _),
            &handler,
        )
    };

    match op {
        Ok(_) => {}
        Err(e) => return Err(hr_msg("ActivateAudioInterfaceAsync 调用失败", e.code())),
    }

    match rx.recv_timeout(Duration::from_secs(5)) {
        Ok(Ok(client)) => Ok(client),
        Ok(Err(e)) => Err(hr_msg("进程回环接口激活回调失败", e.code())),
        Err(e) => Err(format!("等待进程回环激活回调超时/失败: {e}")),
    }
}

#[cfg(target_os = "windows")]
fn get_default_render_device() -> Result<IMMDevice, String> {
    // 第一步：先拿默认渲染设备 IMMDevice。
    let enumerator: IMMDeviceEnumerator = match unsafe {
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
    } {
        Ok(v) => v,
        Err(e) => return Err(hr_msg("创建 IMMDeviceEnumerator 失败", e.code())),
    };

    match unsafe { enumerator.GetDefaultAudioEndpoint(eRender, eMultimedia) } {
        Ok(v) => Ok(v),
        Err(e) => Err(hr_msg("GetDefaultAudioEndpoint 失败", e.code())),
    }
}

#[cfg(target_os = "windows")]
fn activate_process_loopback_client(pid: u32) -> Result<(IAudioClient, *mut WAVEFORMATEX), String> {
    let device = get_default_render_device()?;

    // 第二步：普通模式激活 dummy_client，专门用于获取系统混音格式。
    let dummy_client: IAudioClient = match unsafe {
        device.Activate(CLSCTX_ALL, None)
    } {
        Ok(v) => v,
        Err(e) => return Err(hr_msg("普通模式 Activate(dummy_client) 失败", e.code())),
    };

    let mix_format_ptr = match unsafe { dummy_client.GetMixFormat() } {
        Ok(v) => v,
        Err(e) => return Err(hr_msg("dummy_client.GetMixFormat 失败", e.code())),
    };

    // 第三/四步：用官方进程回环虚拟设备路径 + PID 参数激活真正 loopback client。
    let loopback_client = activate_process_loopback_audio_client(pid)?;

    Ok((loopback_client, mix_format_ptr))
}

#[cfg(target_os = "windows")]
fn hr_hex(code: windows::core::HRESULT) -> String {
    format!("0x{:08X}", code.0 as u32)
}

#[cfg(target_os = "windows")]
fn hr_msg(context: &str, code: windows::core::HRESULT) -> String {
    format!("{context}, HRESULT={}", hr_hex(code))
}

#[cfg(target_os = "windows")]
const DEVICE_INVALIDATED_TAG: &str = "__AUDCLNT_E_DEVICE_INVALIDATED__";

#[cfg(target_os = "windows")]
fn device_invalidated_err(context: &str) -> String {
    format!("{DEVICE_INVALIDATED_TAG} {context}")
}

#[cfg(target_os = "windows")]
fn is_device_invalidated_err(err: &str) -> bool {
    err.contains(DEVICE_INVALIDATED_TAG)
}

fn append_f32_samples_from_bytes(queue: &mut VecDeque<f32>, chunk: &[u8]) {
    for bytes in chunk.chunks_exact(4) {
        let sample = f32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
        queue.push_back(sample);
    }
}

#[cfg(target_os = "windows")]
fn run_capture_for_pid(
    pid: u32,
    pcm_tx: mpsc::Sender<Vec<u8>>,
    stop_rx: std_mpsc::Receiver<()>,
) -> Result<(), String> {
    // WASAPI/COM 必须在线程内初始化；spawn_blocking 正好给我们一个稳定线程。
    match unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) } {
        Ok(()) => {}
        Err(e) => return Err(hr_msg("CoInitializeEx 失败", e.code())),
    }

    let mut recover_attempt = 0u32;
    let result = loop {
        match stop_rx.try_recv() {
            Ok(_) | Err(std_mpsc::TryRecvError::Disconnected) => break Ok(()),
            Err(std_mpsc::TryRecvError::Empty) => {}
        }

        match run_capture_for_pid_inner(pid, &pcm_tx, &stop_rx) {
            Ok(()) => break Ok(()),
            Err(e) if is_device_invalidated_err(&e) => {
                recover_attempt = recover_attempt.saturating_add(1);
                let backoff_ms = (50u64.saturating_mul(1u64 << recover_attempt.min(5))).min(1000);
                println!(
                    "⚠️ 检测到音频设备失效，准备自动重建捕获会话 (attempt={}, backoff={}ms)",
                    recover_attempt,
                    backoff_ms
                );

                let _ = pcm_tx.blocking_send(vec![0u8; 480 * std::mem::size_of::<f32>()]);

                std::thread::sleep(Duration::from_millis(backoff_ms));
                continue;
            }
            Err(e) => break Err(e),
        }
    };

    unsafe { CoUninitialize() };
    result
}

#[cfg(target_os = "windows")]
fn run_capture_for_pid_inner(
    pid: u32,
    pcm_tx: &mpsc::Sender<Vec<u8>>,
    stop_rx: &std_mpsc::Receiver<()>,
) -> Result<(), String> {
    let (audio_client, mix_format_ptr) = activate_process_loopback_client(pid)?;

    let sample_rate = unsafe { (*mix_format_ptr).nSamplesPerSec };
    let channels = unsafe { (*mix_format_ptr).nChannels };
    let bits_per_sample = unsafe { (*mix_format_ptr).wBitsPerSample };
    let format_tag = unsafe { (*mix_format_ptr).wFormatTag };
    println!(
        "🎛️ WASAPI MixFormat: sampleRate={}Hz, channels={}, bitsPerSample={}, formatTag=0x{:04X}",
        sample_rate, channels, bits_per_sample, format_tag
    );

    // 进程回环激活的客户端初始化时使用 LOOPBACK 标志。
    let src_channels = if channels == 0 { 1usize } else { channels as usize };

    let init_result = unsafe {
        audio_client.Initialize(
            AUDCLNT_SHAREMODE_SHARED,
            AUDCLNT_STREAMFLAGS_LOOPBACK,
            0,
            0,
            mix_format_ptr as *const WAVEFORMATEX,
            None,
        )
    };

    // GetMixFormat 返回的内存由 COM 分配，必须释放。
    unsafe { CoTaskMemFree(Some(mix_format_ptr as *const std::ffi::c_void)) };

    match init_result {
        Ok(()) => {}
        Err(e) => return Err(hr_msg("IAudioClient::Initialize 失败", e.code())),
    }

    let capture_client: IAudioCaptureClient = match unsafe { audio_client.GetService() } {
        Ok(v) => v,
        Err(e) if e.code() == AUDCLNT_E_WRONG_ENDPOINT_TYPE => {
            return Err(format!(
                "GetService<IAudioCaptureClient> 失败: WRONG_ENDPOINT_TYPE ({}). 通常表示拿到的不是捕获端 IAudioClient，请检查进程回环激活链路",
                hr_hex(e.code())
            ));
        }
        Err(e) => return Err(hr_msg("GetService<IAudioCaptureClient> 失败", e.code())),
    };

    match unsafe { audio_client.Start() } {
        Ok(()) => {}
        Err(e) => return Err(hr_msg("IAudioClient::Start 失败", e.code())),
    }

    loop {
        // 来自前端的新 PID 会通过异步层发送 stop 信号进来。
        match stop_rx.try_recv() {
            Ok(_) | Err(std_mpsc::TryRecvError::Disconnected) => {
                break;
            }
            Err(std_mpsc::TryRecvError::Empty) => {}
        }

        let mut packet_frames = match unsafe { capture_client.GetNextPacketSize() } {
            Ok(v) => v,
            Err(e) if e.code() == AUDCLNT_E_DEVICE_INVALIDATED => {
                return Err(device_invalidated_err("GetNextPacketSize: 音频设备失效"));
            }
            Err(e) => return Err(format!("GetNextPacketSize 失败, HRESULT={}", hr_hex(e.code()))),
        };

        if packet_frames == 0 {
            std::thread::sleep(Duration::from_millis(2));
            continue;
        }

        while packet_frames > 0 {
            let mut data_ptr: *mut u8 = std::ptr::null_mut();
            let mut frames_to_read: u32 = 0;
            let mut flags: u32 = 0;

            match unsafe {
                capture_client.GetBuffer(
                    &mut data_ptr,
                    &mut frames_to_read,
                    &mut flags,
                    None,
                    None,
                )
            } {
                Ok(()) => {}
                Err(e) if e.code() == AUDCLNT_E_DEVICE_INVALIDATED => {
                    return Err(device_invalidated_err("GetBuffer: 音频设备失效"));
                }
                Err(e) => {
                    return Err(format!(
                        "IAudioCaptureClient::GetBuffer 失败, HRESULT={}",
                        hr_hex(e.code())
                    ))
                }
            }

            // 防御 1：frames==0 时不做任何切片读取，按规范归还空缓冲后继续。
            if frames_to_read == 0 {
                if let Err(e) = unsafe { capture_client.ReleaseBuffer(0) } {
                    return Err(hr_msg("IAudioCaptureClient::ReleaseBuffer(0) 失败", e.code()));
                }

                packet_frames = match unsafe { capture_client.GetNextPacketSize() } {
                    Ok(v) => v,
                    Err(e) if e.code() == AUDCLNT_E_DEVICE_INVALIDATED => {
                        return Err(device_invalidated_err("GetNextPacketSize(循环): 音频设备失效"));
                    }
                    Err(e) => {
                        return Err(format!(
                            "GetNextPacketSize(循环) 失败, HRESULT={}",
                            hr_hex(e.code())
                        ))
                    }
                };
                continue;
            }

            let output_bytes = (frames_to_read as usize).saturating_mul(std::mem::size_of::<f32>());

            // 防御 2：静音包或空指针包严禁 from_raw_parts，直接填零维持时间轴。
            let payload = if data_ptr.is_null() || (flags & AUDCLNT_BUFFERFLAGS_SILENT.0 as u32) != 0 {
                vec![0u8; output_bytes]
            } else {
                let expected_samples = match (frames_to_read as usize).checked_mul(src_channels) {
                    Some(v) => v,
                    None => {
                        if let Err(e) = unsafe { capture_client.ReleaseBuffer(frames_to_read) } {
                            return Err(format!(
                                "samples 计算溢出时 ReleaseBuffer 失败, HRESULT={}",
                                hr_hex(e.code())
                            ));
                        }
                        return Err(format!(
                            "检测到异常 samples 大小: frames_to_read={}, channels={}",
                            frames_to_read, src_channels
                        ));
                    }
                };

                let f32_slice = unsafe { std::slice::from_raw_parts(data_ptr as *const f32, expected_samples) };

                let mut mono_data = Vec::<u8>::with_capacity(output_bytes);

                for i in 0..frames_to_read as usize {
                    let frame_start = i * src_channels;
                    let left = f32_slice[frame_start];
                    let right = if src_channels > 1 {
                        f32_slice[frame_start + 1]
                    } else {
                        left
                    };
                    let mono = (left + right) / 2.0;
                    mono_data.extend_from_slice(&mono.to_le_bytes());
                }

                mono_data
            };

            // 无论数据分支如何，都必须归还缓冲。
            if let Err(e) = unsafe { capture_client.ReleaseBuffer(frames_to_read) } {
                return Err(format!(
                    "IAudioCaptureClient::ReleaseBuffer 失败, HRESULT={}",
                    hr_hex(e.code())
                ));
            }

            if pcm_tx.blocking_send(payload).is_err() {
                break;
            }

            packet_frames = match unsafe { capture_client.GetNextPacketSize() } {
                Ok(v) => v,
                Err(e) if e.code() == AUDCLNT_E_DEVICE_INVALIDATED => {
                    return Err(device_invalidated_err("GetNextPacketSize(循环): 音频设备失效"));
                }
                Err(e) => {
                    return Err(format!(
                        "GetNextPacketSize(循环) 失败, HRESULT={}",
                        hr_hex(e.code())
                    ))
                }
            };
        }
    }

    let _ = unsafe { audio_client.Stop() };
    Ok(())
}

#[cfg(not(target_os = "windows"))]
fn run_capture_for_pid(
    _pid: u32,
    _pcm_tx: mpsc::Sender<Vec<u8>>,
    _stop_rx: std::sync::mpsc::Receiver<()>,
) -> Result<(), String> {
    Err("当前平台不支持 WASAPI 进程回环捕获，仅 Windows 可用".into())
}


#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let (capture_tx, _) = broadcast::channel::<Vec<u32>>(32);
    let latest_capture_pids = Arc::new(Mutex::new(Vec::<u32>::new()));

    if let Err(e) = tauri::Builder::default()
        .manage(AppState {
            capture_tx: capture_tx.clone(),
            latest_capture_pids: latest_capture_pids.clone(),
        })
        .setup(|app| {
            let ws_capture_tx = app.state::<AppState>().capture_tx.clone();
            let ws_latest_pids = app.state::<AppState>().latest_capture_pids.clone();

            tauri::async_runtime::spawn(async {
                start_audio_pump(ws_capture_tx, ws_latest_pids).await;
            });

            if cfg!(debug_assertions) {
                app.handle().plugin(
                    tauri_plugin_log::Builder::default()
                        .level(log::LevelFilter::Info)
                        .build(),
                )?;
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![get_active_processes, start_capture, start_capture_multi])
        .run(tauri::generate_context!())
    {
        println!("❌ tauri 运行失败: {e}");
    }
}
