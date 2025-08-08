
import json
import pandas as pd
from tkinter import Tk, filedialog
import os

def convert_json_to_excel():
    """
    Reads a JSON file, extracts specific keys, and saves the data to an Excel file.
    """
    print("변환할 JSON 파일을 선택해주세요.")
    # --- Get Input JSON file ---
    root = Tk()
    root.withdraw()  # Hide the main window
    json_path = filedialog.askopenfilename(
        title="Select a JSON file",
        filetypes=[("JSON files", "*.json")]
    )

    if not json_path:
        print("파일이 선택되지 않았습니다. 프로그램을 종료합니다.")
        return

    # --- Read JSON file ---
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            # Handle JSON files with multiple JSON objects per line
            data = [json.loads(line) for line in f if line.strip()]
    except json.JSONDecodeError:
        # If reading line-by-line fails, try reading the whole file as a single JSON array/object
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except json.JSONDecodeError:
            print(f"오류: {json_path} 파일의 JSON 형식이 올바르지 않습니다.")
            return
    except Exception as e:
        print(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        return

    # --- Process JSON data ---
    # User specified that the data is within a "rows" key.
    
    # Ensure data is a list to handle both single JSON objects and line-by-line JSON
    if not isinstance(data, list):
        data = [data]

    records_to_process = []
    for payload in data:
        if isinstance(payload, dict) and 'rows' in payload:
            row_data = payload.get('rows')
            if isinstance(row_data, list):
                records_to_process.extend(row_data)
            else:
                print("경고: 'rows' 키의 값이 리스트가 아닙니다. 해당 항목을 건너뜁니다.")
        else:
            # Fallback for JSONs that don't have the 'rows' key
            if isinstance(payload, dict):
                records_to_process.append(payload)

    if not records_to_process:
        print("오류: JSON 데이터에서 'rows' 키를 찾을 수 없거나 처리할 레코드가 없습니다.")
        return

    keys_to_extract = ['id', 'message_sending_datetime', 'send_type', 'cust_name', 'cust_phone_number']
    
    extracted_data = []
    for item in records_to_process:
        if isinstance(item, dict):
            record = {key: item.get(key) for key in keys_to_extract}
            extracted_data.append(record)

    if not extracted_data:
        print("처리할 데이터가 없거나 JSON 객체에서 키를 찾을 수 없습니다.")
        return

    # --- Create DataFrame and save to Excel ---
    df = pd.DataFrame(extracted_data)
    
    print("저장할 Excel 파일의 경로와 이름을 지정해주세요.")
    # --- Get Output Excel file path ---
    excel_path = filedialog.asksaveasfilename(
        title="Save Excel file as...",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not excel_path:
        print("저장이 취소되었습니다. 프로그램을 종료합니다.")
        return

    try:
        df.to_excel(excel_path, index=False)
        print(f"성공적으로 변환하여 {excel_path} 파일에 저장했습니다.")
    except Exception as e:
        print(f"Excel 파일 저장 중 오류가 발생했습니다: {e}")

if __name__ == "__main__":
    convert_json_to_excel()
