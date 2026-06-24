# build_data.py
import os
import json
from data_processor import build_contracts

def main():
    print("엑셀 데이터를 읽어서 contracts.js 및 contracts.json 파일로 변환 중...")
    contracts = build_contracts()
    
    output_js_path = os.path.join(os.path.dirname(__file__), 'contracts.js')
    with open(output_js_path, 'w', encoding='utf-8') as f:
        f.write("window.INITIAL_CONTRACTS = ")
        json.dump(contracts, f, ensure_ascii=False, indent=2)
        f.write(";\n")
        
    output_json_path = os.path.join(os.path.dirname(__file__), 'contracts.json')
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(contracts, f, ensure_ascii=False, indent=2)
        
    print(f"변환 완료! 생성된 파일: {output_js_path}, {output_json_path} (계약 수: {len(contracts)}개)")

if __name__ == '__main__':
    main()

