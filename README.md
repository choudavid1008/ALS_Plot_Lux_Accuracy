# Lux Accuracy 分析工具 (PySide6)

這是一個使用 PySide6 開發的 GUI 工具，用於從 Excel 檔案中提取 Lux 數據並產生成準確度分析圖表。

## 解決 ModuleNotFoundError (環境問題排除)

如果您在執行 `pip install -r requirements.txt` 後依然出現 `ModuleNotFoundError`，通常是因為您的 `python` 與 `pip` 指令版本不一致（例如 python 是 3.11，但 pip 把套件裝到了 3.8）。

### 快速修復命令
請嘗試使用以下指令來安裝，這會確保套件裝在目前的 Python 環境中：
```bash
python -m pip install -r requirements.txt
```

---

## 標準安裝說明 (建議)

建議使用虛擬環境 (Virtual Environment) 以避免套件衝突。

### 1. 執行自動設定腳本 (Windows)
在資料夾中找到 **`setup_windows.bat`** 並雙擊執行。它會自動建立虛擬環境並安裝所有套件。

### 2. VS Code 設定 (重要)
安裝完後，請務必在 VS Code 中切換到正確的環境：
1. 按下 `Ctrl+Shift+P`。
2. 輸入並選擇 **"Python: Select Interpreter"**。
3. 選擇路徑中包含 **`venv\Scripts\python.exe`** 的那一項。

## 功能特點
1. **讀取 Excel (讀檔)**：載入 Excel 活頁簿，並在界面上直接檢視內容。
2. **動態配置**：可於 UI 上自由設定範圍、係數儲存格。
3. **自動化繪圖 (執行)**：一鍵生成 2x2 的準確度分析圖。
4. **自動存檔**：圖表儲存至 `output_file/lux_accuracy_for_all.png`。

## 檔案結構
- `main.py`: 主程式。
- `setup_windows.bat`: 自動安裝腳本。
- `requirements.txt`: 套件清單。
- `input_file/`: 輸入目錄。
- `output_file/`: 輸出目錄。
