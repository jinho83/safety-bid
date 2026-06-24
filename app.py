# app.py - 안전 종합평가제 시뮬레이터 Flask 백엔드
from flask import Flask, send_file, jsonify, send_from_directory
from flask_cors import CORS
import os

from data_processor import build_contracts, clear_cache

app = Flask(__name__)
CORS(app)  # 개발 편의를 위한 CORS 허용


# ─── 라우트 ──────────────────────────────────────────────
@app.route('/')
def index():
    """index.html 서빙"""
    return send_file(os.path.join(os.path.dirname(__file__), 'index.html'))

@app.route('/lib/<path:filename>')
def serve_lib(filename):
    """lib/ 폴더의 정적 JS 파일 서빙"""
    lib_dir = os.path.join(os.path.dirname(__file__), 'lib')
    return send_from_directory(lib_dir, filename)

@app.route('/contracts.js')
def serve_contracts_js():
    """contracts.js 파일 서빙"""
    return send_file(os.path.join(os.path.dirname(__file__), 'contracts.js'))

@app.route('/get_excel_data')
def get_excel_data():
    """엑셀 데이터를 JSON으로 반환"""
    try:
        contracts = build_contracts()
        return jsonify(contracts)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/reload')
def reload_cache():
    """캐시 초기화 (엑셀 파일 변경 후 재로딩)"""
    clear_cache()
    return jsonify({'message': '캐시가 초기화되었습니다. 다음 요청 시 엑셀 파일을 다시 읽습니다.'})


if __name__ == '__main__':
    print("=" * 50)
    print("안전 종합평가제 시뮬레이터 서버 시작")
    print("접속 주소: http://localhost:5000")
    print("=" * 50)
    app.run(debug=True, port=5000)
