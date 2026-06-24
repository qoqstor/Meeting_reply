# Live Translation + Reply Assistant v2

**即時翻譯 + AI 回覆助理** — 與 [jt-live-whisper](https://github.com/KoljaB/RealtimeSTT) 整合，提供即時語音翻譯監控與 AI 雙向回覆建議。

---

## 功能總覽

| 功能 | 說明 |
|------|------|
| 📡 **即時翻譯顯示** | 監控 jt-live-whisper log，日文原文 / 中文翻譯左右並列，自動捲動 |
| 🤖 **AI 回覆建議** | 取最近 7 句 / 10 分鐘語音，生成日文回覆建議 |
| 📝 **中文草稿** | 同步輸出中文草稿，方便中文使用者理解並說出對應回覆 |
| ⚡ **自動觸發** | 偵測句尾標點（。！？），靜默 3 秒後自動生成回覆 |
| 🧠 **對話歷史記憶** | 多輪對話送進 AI，最多保留 10 輪歷史 |
| 🔊 **TTS 播放** | 生成後可朗讀日文回覆或中文草稿（edge-tts） |
| 📊 **匯出逐字稿** | 一鍵匯出 Excel / Word / TXT，含原文、翻譯、AI 回覆 |
| 🎙 **jt-live-whisper 控制** | 直接從 UI 啟動 / 停止翻譯程式，含模式選單 |
| 🔀 **多 AI Provider** | Ollama（本地免費）、Anthropic Claude、Google Gemini |
| 📂 **可設定情境 Profile** | `profiles/*.json` 儲存每個專案的 system prompt |

---

## 快速開始

### 1. 安裝依賴

```bash
pip install anthropic google-genai pyperclip websockets edge-tts python-docx openpyxl
```

> **TTS / 匯出為選用功能**，不安裝對應套件仍可正常使用核心功能。

### 2. 啟動

```bash
python sei_reply_assistant.py
```

首次啟動會詢問 jt-live-whisper 安裝目錄（包含 `start.ps1` 的資料夾），選擇後自動記憶。

### 3. 使用流程

1. 點擊 **🎙 開啟即時翻譯** 啟動 jt-live-whisper
2. 對方說話 → 即時翻譯區自動更新（左=日文，右=中文）
3. 點擊綠色 **▶ 生成回覆建議** 或啟用「自動觸發」
4. AI 輸出日文回覆（左）+ 中文草稿（右）
5. 點「▶ TTS」朗讀，或「複製」貼至對話視窗
6. 點「匯出逐字稿」存成 Excel / Word

---

## AI Provider 設定

| Provider | API Key | 費用 | 推薦模型 |
|----------|---------|------|---------|
| **Ollama** | 不需要 | 免費（本地） | `qwen2.5:7b` |
| **Anthropic** | `sk-ant-api03-...` | 付費 | `claude-haiku-4-5-20251001` |
| **Gemini** | `AIza...` | 依帳號 | `gemini-2.0-flash` |

Ollama 使用前需先安裝並拉取模型：
```bash
ollama pull qwen2.5:7b
```

---

## Profile 情境設定

`profiles/*.json` 每個檔案是一個情境。可用 UI 的「新增情境」按鈕建立，或直接複製 `General_JA.json` 修改：

```json
{
  "name": "MyProject",
  "description": "我的專案情境說明",
  "reply_lang": "ja",
  "system_prompt": "あなたは会議の回答アシスタントです。...\n{\"ja_reply\": \"日本語の返答...\", \"zh_reply\": \"中文草稿...\"}"
}
```

> **v2 注意**：system_prompt 需包含 JSON 格式指示（`ja_reply` / `zh_reply`），才能正確生成雙向回覆。

---

## v2 新功能說明

### 自動觸發模式
- 句尾偵測：`。！？!?`
- 偵測到句尾後靜默 3 秒才觸發（避免邊說邊觸發）
- 可與手動按鈕同時使用

### 雙向回覆
- AI 同步生成「日文回覆」與「中文草稿」
- 使用 JSON 格式輸出（`ja_reply` + `zh_reply`）
- 若 AI 未返回 JSON 格式，自動 fallback 至全文顯示

### 對話歷史記憶
- 每次生成將對話加入 `_conversation_history`
- 下次生成時帶入完整歷史（Ollama / Anthropic 使用 messages array，Gemini 使用文字串接）
- 點擊「清除對話歷史」重置
- 最多保留 10 輪（20 則訊息）

### TTS 播放
- 依賴 `edge-tts`（`pip install edge-tts`）
- 日文語音：`ja-JP-NanamiNeural`
- 中文語音：`zh-TW-HsiaoChenNeural`
- 音訊存為暫存 MP3，播放後 90 秒自動刪除

### 匯出逐字稿
- Excel（需 `openpyxl`）：含藍色標頭、自動欄寬
- Word（需 `python-docx`）：表格格式，含匯出時間
- TXT：純文字，含時間戳與各欄位

---

## 目錄結構

```
Meeting_reply/
├── sei_reply_assistant.py   # 主程式（tkinter GUI）
├── profiles/
│   ├── General_JA.json      # 通用日文情境（範例）
│   └── General_EN.json      # 通用英文情境（範例）
├── requirements.txt
└── README.md
```

執行後自動建立：
- `settings.json` — 記住 provider / model / profile / jt-live-whisper 路徑
- `logs/` — jt-live-whisper 逐字稿輸出（開發模式備用）

---

## 系統需求

- **OS**：Windows 10/11 x64
- **Python**：3.10+
- **依賴**：`jt-live-whisper`（需另行安裝，提供即時 ASR 翻譯）

---

## License

MIT
