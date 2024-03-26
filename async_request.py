import sys
import requests
import win32com.client

def async_get_chat_completion_content(prompt, cell_address, result_cell_address):
    url = "http://localhost:11434/v1/chat/completions"
    headers = {"Content-Type": "application/json"}
    payload = {
        "model": "gemma:7b",
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": prompt}
        ]
   
    }
    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 200:
        content = response.json()["choices"][0]["message"]["content"]
    else:
        content = "Error: " + str(response.status_code)

    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(excel.ActiveWorkbook.FullName)
    worksheet = workbook.ActiveSheet
    result_cell = worksheet.Range(result_cell_address)
    result_cell.Value = content
    workbook.Save()

if __name__ == "__main__":
    prompt = sys.argv[1]
    cell_address = sys.argv[2]
    result_cell_address = sys.argv[3]
    async_get_chat_completion_content(prompt, cell_address, result_cell_address)
