# Lux Accuracy 分析工具 (PySide6)

這是一個使用 PySide6 開發的 GUI 工具，用於從 Excel 檔案中提取 Lux 數據並產生成準確度分析圖表。

## 功能特點

1.  **讀取 Excel (讀檔)**：載入 Excel 活頁簿，並在界面上直接檢視內容。
2.  **動態配置**：可於 UI 上自由設定 CCT 模式範圍、係數 (Cr, Cg, Cb, Cc, Cwb) 的儲存格位置，以及數據範圍。
3.  **自動化繪圖 (執行)**：一鍵生成 2x2 的 Lux 準確度分析圖，包含：
    -   四種指定的 Sheet: `Negroni (red)`, `Pine (green)`, `Haze midnight (black)`, `Silver`。
    -   自動對齊 CCT 模式與顏色。
    -   繪製理想線與 ±10% 的誤差區間。
4.  **自動存檔**：產生的圖表會自動儲存至 `output_file/lux_accuracy_for_all.png`。

## 安裝說明

請確保您的電腦已安裝 Python 3.8+，然後在終端機執行以下命令安裝所需套件：

```bash
pip install -r requirements.txt
```

## 使用方法

1.  執行主程式：
    ```bash
    python main.py
    ```
2.  點擊 **"Load Excel File (讀檔)"** 選擇您的數據檔案。
3.  在 **"Data View (表格與設定)"** 頁籤中確認您的參數設定（預設值已根據您的需求填入）。
4.  點擊 **"Run & Generate Plot (執行)"**。
5.  程式會自動切換到 **"Plot View (圖表)"** 顯示分析結果，並將圖檔存入 `output_file` 資料夾。

## 檔案結構

-   `main.py`: 主程式邏輯與介面。
-   `input_file/`: 建議存放輸入 Excel 檔案的目錄。
-   `output_file/`: 產生的分析圖表儲存目錄。
-   `requirements.txt`: 必要的 Python 套件清單。
