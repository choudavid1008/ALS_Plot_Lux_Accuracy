# Lux Accuracy 分析工具 (PySide6)

這是一個使用 PySide6 開發的 GUI 工具，用於從 Excel 檔案中提取 Lux 數據並產生成準確度分析圖表。

## 解決 ModuleNotFoundError: No module named 'PySide6'

如果您在執行時遇到此錯誤，表示您的 Python 環境尚未安裝必要的套件。請按照以下「安裝說明」進行操作。

## 安裝說明

請確保您的電腦已安裝 Python 3.8+。建議使用虛擬環境 (Virtual Environment) 以避免套件衝突。

### 1. 建立並啟動虛擬環境 (建議)

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**macOS/Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 2. 安裝必要套件

在啟動虛擬環境後，執行以下命令：

```bash
pip install -r requirements.txt
```

### 3. VS Code 設定 (重要)

如果您使用 VS Code：
1. 按下 `Ctrl+Shift+P` (或 `Cmd+Shift+P`)。
2. 輸入並選擇 **"Python: Select Interpreter"**。
3. 選擇剛才建立的虛擬環境 (路徑通常包含 `./venv/Scripts/python.exe` 或 `./venv/bin/python`)。

## 功能特點

1.  **讀取 Excel (讀檔)**：載入 Excel 活頁簿，並在界面上直接檢視內容。
2.  **動態配置**：可於 UI 上自由設定 CCT 模式範圍、係數 (Cr, Cg, Cb, Cc, Cwb) 的儲存格位置，以及數據範圍。
3.  **自動化繪圖 (執行)**：一鍵生成 2x2 的 Lux 準確度分析圖。
4.  **自動存檔**：圖表會自動儲存至 `output_file/lux_accuracy_for_all.png`。

## 使用方法

1.  執行主程式：
    ```bash
    python main.py
    ```
2.  點擊 **"Load Excel File (讀檔)"** 選擇數據檔案。
3.  點擊 **"Run & Generate Plot (執行)"** 產出結果。

## 檔案結構

-   `main.py`: 主程式邏輯與介面。
-   `requirements.txt`: 必要的 Python 套件清單。
-   `input_file/`: 建議存放輸入檔案目錄。
-   `output_file/`: 產生的分析圖表儲存目錄。
