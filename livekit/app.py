from flask import Flask, request, jsonify, render_template_string
from livekit import api
from flask_cors import CORS  # 👈 新加这一行导入
import uuid  # 新增：引入生成随机唯一ID的库

app = Flask(__name__)
CORS(app)  # 👈 新加这一行，让保安直接放行所有来源！

API_KEY = "devkey"
API_SECRET = "secret"
ROOM_NAME = "team-meeting-room"

def build_token(user_name: str, room_name: str):
    unique_identity = f"{user_name}-{uuid.uuid4().hex[:8]}"
    token = api.AccessToken(API_KEY, API_SECRET) \
        .with_identity(unique_identity) \
        .with_name(user_name) \
        .with_grants(api.VideoGrants(
            room_join=True,
            room=room_name,
        ))
    return token.to_jwt()

@app.route('/')
def index():
    with open('index.html', 'r', encoding='utf-8') as f:
        return render_template_string(f.read())

@app.route('/api/get_token')
@app.route('/token')
def get_token():
    user_name = request.args.get('user', '访客')
    room_name = request.args.get('room', ROOM_NAME)

    token_jwt = build_token(user_name, room_name)
    return jsonify({"token": token_jwt, "room": room_name})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)