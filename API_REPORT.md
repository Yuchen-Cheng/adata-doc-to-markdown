# Document to Markdown API — 技術報告

## 目錄
1. [系統概覽](#1-系統概覽)
2. [環境設定](#2-環境設定)
3. [API 呼叫方式](#3-api-呼叫方式)
4. [程式運作流程](#4-程式運作流程)
5. [各 Converter 腳本流程](#5-各-converter-腳本流程)
6. [Dify Vision 圖片描述流程](#6-dify-vision-圖片描述流程)
7. [錯誤處理](#7-錯誤處理)

---

## 1. 系統概覽

本系統是一個 FastAPI 服務，接收上傳的辦公室文件（DOCX / DOC / PDF / PPTX / XLS / XLSX），將內容轉換成 Markdown，並透過 Dify Vision Workflow 為文件內的圖片產生文字描述。

### 檔案結構

```
converter_api/
├── main.py                        # FastAPI 主程式，所有 API endpoint 與後處理邏輯
├── markitdown_to_md.py            # DOCX / DOC 轉換腳本
├── pdf_to_markdown.py             # PDF 轉換腳本
├── ppt_to_markdown.py             # PPTX 轉換腳本
├── excel_sop_kb_to_markdown.py    # XLS / XLSX 通用轉換腳本
├── excel_ISO_to_markdown.py       # XLS / XLSX TW- 開頭特規腳本（含 Mermaid 流程圖）
├── .env                           # 機敏設定（API key、Workflow ID 等）
└── requirements.txt
```

---

## 2. 環境設定

### `.env` 檔案

```env
DIFY_API_KEY=app-Qn1KiCDdStLSdeCboYmfHHXx
DIFY_WORKFLOW_ID=301f6809-f0db-4d57-b500-cc1edcc15dcd
DIFY_BASE_URL=https://ai-dev.adata.com
DIFY_USER_ID=converter_api
```

| 變數 | 用途 | 必填 |
|---|---|---|
| `DIFY_API_KEY` | Dify API Bearer Token | ✅ |
| `DIFY_WORKFLOW_ID` | Vision Workflow 的 ID | ✅ |
| `DIFY_BASE_URL` | Dify 服務基底網址 | 選填（預設 `https://ai-dev.adata.com`） |
| `DIFY_USER_ID` | 呼叫 Dify API 時帶的 user 識別 | 選填（預設 `converter_api`） |
| `CONVERTERS_DIR` | 轉換腳本所在目錄 | 選填（預設與 `main.py` 同目錄） |

### 啟動服務

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

服務預設監聽 `http://127.0.0.1:8000`，Swagger UI 位於 `http://127.0.0.1:8000/docs`。

---

## 3. API 呼叫方式

所有 endpoint 均為 `POST`，接受 `multipart/form-data`，唯一必填欄位為 `file`。

### 共用回應格式

所有 endpoint 回傳單一 **JSON 物件**。多頁 / 多工作表的內容會以 `\n\n` 串接成一個字串：

```json
{
  "filename": "report.xlsx",
  "markdown": "# Sheet1\n\n表格內容...\n\n# Sheet2\n\n...",
  "markdown_with_descriptions": "# Sheet1\n\n表格內容...\n\n圖片描述文字...\n\n# Sheet2\n\n..."
}
```

| 欄位 | 說明 |
|---|---|
| `filename` | 上傳的原始檔名 |
| `markdown` | 完整文件內容 Markdown，圖片已移除，多頁 / 多工作表以 `\n\n` 串接 |
| `markdown_with_descriptions` | 完整文件內容 Markdown，圖片已替換為 Dify Vision 產生的文字描述，多頁 / 多工作表以 `\n\n` 串接 |

> `/convert/auto` 額外包含 `"detected_type"` 欄位（如 `"pdf"`）。

---

### 3.1 `POST /convert/docx`

轉換 `.docx` 或 `.doc` 檔案。

**cURL**
```bash
curl -X POST http://127.0.0.1:8000/convert/docx \
  -F "file=@report.docx"
```

**Python**
```python
import requests

with open("report.docx", "rb") as f:
    resp = requests.post(
        "http://127.0.0.1:8000/convert/docx",
        files={"file": ("report.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
    )
results = resp.json()          # list with 1 element, sheet = ""
print(results[0]["markdown"])
```

**回應範例**
```json
{
  "filename": "report.docx",
  "markdown": "# 標題\n\n段落文字...",
  "markdown_with_descriptions": "# 標題\n\n段落文字..."
}
```

| 接受副檔名 | `.docx` `.doc` |
|---|---|

---

### 3.2 `POST /convert/pdf`

轉換 `.pdf` 檔案。

**cURL**
```bash
curl -X POST http://127.0.0.1:8000/convert/pdf \
  -F "file=@manual.pdf"
```

**回應範例**
```json
{
  "filename": "manual.pdf",
  "markdown": "頁面文字...",
  "markdown_with_descriptions": "頁面文字...\n\n圖片A描述..."
}
```

| 接受副檔名 | `.pdf` |
|---|---|

---

### 3.3 `POST /convert/pptx`

轉換 `.pptx` 檔案。

**cURL**
```bash
curl -X POST http://127.0.0.1:8000/convert/pptx \
  -F "file=@slides.pptx"
```

| 接受副檔名 | `.pptx` |
|---|---|

---

### 3.4 `POST /convert/excel`

轉換 `.xls` 或 `.xlsx` 檔案。  
檔名以 `TW-` 開頭時自動使用 ISO 特規腳本（含 Mermaid 流程圖）。

**cURL**
```bash
curl -X POST http://127.0.0.1:8000/convert/excel \
  -F "file=@data.xlsx"
```

**多工作表回應範例**
```json
{
  "filename": "data.xlsx",
  "markdown": "# Sales\n\n...\n\n# Inventory\n\n...",
  "markdown_with_descriptions": "# Sales\n\n...\n\n# Inventory\n\n..."
}
```

**單一工作表回應範例**
```json
{
  "filename": "data.xlsx",
  "markdown": "# Sheet1\n\n...",
  "markdown_with_descriptions": "# Sheet1\n\n..."
}
```

| 接受副檔名 | `.xls` `.xlsx` |
|---|---|
| `TW-*.xlsx` | 使用 `excel_ISO_to_markdown.py` |
| 其他 | 使用 `excel_sop_kb_to_markdown.py` |

---

### 3.5 `POST /convert/auto`

自動偵測副檔名並選擇對應轉換器。

**cURL**
```bash
curl -X POST http://127.0.0.1:8000/convert/auto \
  -F "file=@unknown_file.xlsx"
```

**額外欄位**：多一個 `"detected_type"` 欄位

```json
{
  "filename": "unknown_file.xlsx",
  "markdown": "...",
  "markdown_with_descriptions": "...",
  "detected_type": "xlsx"
}
```

| 接受副檔名 | `.doc` `.docx` `.xls` `.xlsx` `.pdf` `.pptx` |
|---|---|

---

## 4. 程式運作流程

### 4.1 整體流程圖

```
HTTP Request (multipart/form-data, file)
        │
        ▼
[Endpoint 函式]  ──副檔名驗證──▶ HTTP 400 (不符)
        │ 通過
        ▼
[tempfile.TemporaryDirectory()]   ← 建立全程暫存目錄
        │
        ├─ 將上傳檔案寫入暫存目錄
        │
        ▼
[choose_excel_converter() / CONVERTER_MAP]
        │  選出對應的 .py 腳本
        ▼
[_run_script()]
        │  subprocess 執行腳本
        │  腳本輸出 <SheetName>/index.md（每個工作表一份）至暫存目錄
        │  returncode != 0 ──▶ HTTP 500
        ▼
[_collect_output()]
        │
        ├─ 找到多個 */index.md（Excel 多工作表）
        │     多工作表：每份獨立處理，sheet = 資料夾名稱
        │     單工作表：sheet = ""
        │
        └─ 找到單一 index.md（docx/pdf/pptx）
              sheet = ""
        │
        ▼  對每份 index.md 呼叫 _split_markdown()
        │
        ├─▶ markdown               (圖片全部移除)
        │
        └─▶ markdown_with_descriptions
                │
                └─ 每張圖片呼叫 image_to_description()
                        │  上傳至 Dify → 執行 Vision Workflow
                        └─▶ 文字描述取代原圖片位置
        │
        ▼
[_build_response()]
        │  合併結果包裝成 {filename, markdown, markdown_with_descriptions}
        ▼
[TemporaryDirectory 自動清除所有暫存檔]
        │
        ▼
HTTP 200  JSON Object { filename, markdown, markdown_with_descriptions }
```

---

### 4.2 `main.py` 函式一覽

| 函式 | 輸入 | 輸出 | 說明 |
|---|---|---|---|
| `get_file_type(filename)` | 檔名字串 | `"docx"` / `"pdf"` / … / `"unknown"` | 由副檔名判斷檔案類型 |
| `choose_excel_converter(filename)` | 檔名字串 | 腳本檔名字串 | `TW-` 開頭選 ISO 腳本，否則選通用腳本 |
| `_run_script(script_name, input_file, output_dir)` | 腳本名稱、輸入路徑、輸出路徑 | `(True, "")` 或 `(False, 錯誤訊息)` | 以 subprocess 執行轉換腳本，最長等待 300 秒 |
| `_collect_output(output_dir)` | 暫存輸出目錄 | `(markdown: str, markdown_with_descriptions: str)` | 讀取所有 `*/index.md` 或單一 `index.md`，多份內容以 `\n\n` 串接後呼叫 `_split_markdown()` |
| `_split_markdown(raw)` | 含 data URI 圖片的 Markdown 字串 | `(no_images: str, with_descriptions: str)` | 找出所有 `![](data:…)` 圖片分兩路處理 |
| `image_to_description(image_bytes, mime_type)` | 圖片 bytes、MIME 類型 | 文字描述字串 | 上傳至 Dify → 執行 Vision Workflow → 回傳 `outputs.result` |
| `_build_response(filename, markdown, markdown_with_descriptions, extra)` | 檔名、兩份 markdown 字串、額外欄位 | `JSONResponse`（單一物件） | 組裝最終 API 回應 |
| `_convert(script, upload)` | 腳本名稱、UploadFile | `JSONResponse` | docx / pdf / pptx 的共用驅動函式 |

---

### 4.3 Excel 多工作表處理

兩個 Excel 腳本都採用**每個工作表一個子目錄**的輸出格式：

```
output_dir/
├── Sales/
│   └── index.md      ← 第 1 張工作表
├── Inventory/
│   └── index.md      ← 第 2 張工作表
└── Summary/
    └── index.md      ← 第 3 張工作表
```

`_collect_output()` 以 `output_dir.glob("*/index.md")` 掃描全部子目錄，依資料夾名稱排序後，將每張工作表的 markdown 以 `\n\n` 串接成單一字串後再處理。最終回應只有一個 JSON 物件，多張工作表的內容都在同一個 `markdown` 欄位中。

---

## 5. 各 Converter 腳本流程

轉換腳本由 `_run_script()` 以 **獨立 subprocess** 呼叫。  
所有圖片均以 **base64 data URI** 直接嵌入 `index.md`，不寫入任何圖片檔案。

---

### 5.1 `markitdown_to_md.py`（DOCX / DOC）

```
Input: input.docx / input.doc
       output_dir/
        │
        ├─ [.doc 輸入] convert_doc_to_docx()
        │      Windows COM (Word) 轉換為暫存 .docx（系統 temp，轉換後刪除）
        │
        ▼
   MarkItDown.convert(keep_data_uris=True)
        │  輸出：Markdown，圖片已是 data:image/xxx;base64,... 格式
        │
        ▼
   re.sub() 逐一處理每個 data URI
        ├─ EMF / WMF：convert_emf_to_png_data_uri()  (記憶體內轉 PNG)
        └─ 其他格式：直接保留原 data URI
        │
        ▼
Output: output_dir/<stem>/index.md
```

---

### 5.2 `pdf_to_markdown.py`（PDF）

```
Input: input.pdf  /  output_dir/
        │
        ▼
   Pass 1：掃描全頁，計算圖片 MD5，找出重複圖片集合
        │
        ▼
   Pass 2：逐頁 extract_page()
        ├─ 圖片：PNG bytes → MD5 dedup → base64 data URI
        ├─ 文字：移除頁尾，過濾已在表格的行
        └─ 表格：table_to_markdown()
        │
        ▼
   各頁以 --- 分隔合併
        │
Output: output_dir/index.md
```

---

### 5.3 `ppt_to_markdown.py`（PPTX）

```
Input: input.pptx  /  output_dir/
        │
        ▼
   逐張 Slide → process_slide()
        ├─ Shape 依座標排序
        ├─ GROUP：遞迴 process_group_shape()
        ├─ TABLE：table_to_markdown()
        ├─ PICTURE：image bytes → MD5 dedup → base64 data URI
        └─ TEXT：段落 / bullet 文字
        │
        ▼
   各 Slide 以 --- 分隔合併
        │
Output: output_dir/index.md
```

---

### 5.4 `excel_sop_kb_to_markdown.py`（XLS / XLSX 通用）

```
Input: input.xlsx  /  output_dir/
        │
        ├─ [.xls] convert_xls_to_xlsx()  COM 轉暫存 .xlsx
        │
        ▼
   openpyxl load_workbook()
        │
        ▼
   逐張 Worksheet → worksheet_to_markdown()
        ├─ extract_text_from_sheet()    → [(row, col, text), ...]
        └─ extract_images_from_sheet()  → image bytes → base64 data URI
        │  依 (row, col) 排序後合併
        │
        ▼  每張工作表寫入獨立子目錄
        │
Output: output_dir/<SheetName>/index.md  × N 張工作表
```

---

### 5.5 `excel_ISO_to_markdown.py`（TW- 開頭 Excel，含流程圖）

```
Input: TW-*.xlsx  /  output_dir/
        │
        ▼
   win32com Excel.Application 開啟
        │
        ▼
   逐張 Worksheet → convert_sheet()
        ├─ Phase 1-3：偵測 Connector，收集 Node，抽取 Edge
        ├─ Phase 4：Yes/No 標籤匹配 Decision Edge
        ├─ Phase 5：Swim-lane 偵測（部門標頭 → Mermaid subgraph）
        ├─ Phase 6：讀取 UsedRange 表格資料
        ├─ Phase 7：組裝 Markdown（Mermaid + 表格）
        └─ Phase 8：extract_images_from_sheet()
                    .xlsx zip 內 xl/media/ → base64 data URI
        │
Output: output_dir/<SheetName>/index.md  × N 張工作表
```

---

## 6. Dify Vision 圖片描述流程

`image_to_description()` 在 `_split_markdown()` 內對每張圖片呼叫一次。

```
image_bytes (raw bytes)  /  mime_type (e.g. "image/png")
        │
        ▼
Step 1：POST {DIFY_BASE_URL}/v1/files/upload
        Headers: Authorization: Bearer {DIFY_API_KEY}
        Body (multipart):
          file = (image.png, image_bytes, mime_type)
          user = converter_api
        Response: { "id": "<file_id>", ... }
        │
        ▼
Step 2：POST {DIFY_BASE_URL}/v1/workflows/{DIFY_WORKFLOW_ID}/run
        Headers: Authorization: Bearer {DIFY_API_KEY}
                 Content-Type: application/json
        Body:
          {
            "inputs": {
              "img": {
                "transfer_method": "local_file",
                "upload_file_id": "<file_id>",
                "type": "image"
              }
            },
            "response_mode": "blocking",
            "user": "converter_api"
          }
        Response: { "data": { "outputs": { "result": "圖片描述..." } } }
        │
        ▼
Output: "圖片描述文字"  ← 取代 markdown_with_descriptions 中的圖片位置
```

| 步驟 | Timeout |
|---|---|
| 檔案上傳 | 30 秒 |
| Workflow 執行 | 120 秒 |

失敗時回傳 `[Image description unavailable: ...]`，不中斷整體轉換。

---

## 7. 錯誤處理

| 情境 | HTTP 狀態 | 回應內容 |
|---|---|---|
| 上傳不支援的副檔名 | `400` | `"Only .xxx files are accepted"` |
| `/convert/auto` 不支援的副檔名 | `400` | `"Unsupported extension '.xxx'"` |
| 轉換腳本找不到 | `500` | `"Converter script not found: ..."` |
| 轉換腳本回傳非零 exit code | `500` | `"Conversion failed: <stderr>"` |
| 轉換超過 300 秒 | `500` | `"Conversion timed out (300 s)"` |
| 腳本未產生任何 `.md` 檔 | `500` | `"Converter produced no Markdown output"` |
| Dify 上傳失敗 | *(不中斷)* | 圖片位置替換為 `[Image description unavailable: upload failed — ...]` |
| Dify Workflow 失敗 | *(不中斷)* | 圖片位置替換為 `[Image description unavailable: workflow failed — ...]` |

所有暫存目錄（`tempfile.TemporaryDirectory`）不論成功或失敗均自動清除，**不在伺服器留下任何檔案**。
