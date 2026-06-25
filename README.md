# Live Translation + Reply Assistant v2.1

**即時翻譯 + AI 回覆助理** — 與 [jt-live-whisper](https://github.com/jasoncheng7115/jt-live-whisper) 整合，提供即時語音翻譯監控與 AI 雙向回覆建議。

---

## 功能總覽

| 功能 | 說明 |
|------|------|
| 📡 **即時翻譯顯示** | 監控 jt-live-whisper log，支援所有翻譯模式（ja2zh / ja_zh 等），日文 / 中文左右並列 |
| 🤖 **AI 回覆建議** | 取最近 30 句 / 10 分鐘語音，生成日文回覆建議 |
| 📝 **中文草稿** | 同步輸出中文草稿，方便中文使用者理解並說出對應回覆 |
| ⚡ **自動觸發** | 偵測句尾標點（。！？），靜默 8 秒後自動生成回覆 |
| 🧠 **對話歷史記憶** | 多輪對話送進 AI，最多保留 10 輪歷史 |
| 🔊 **TTS 播放** | 生成後可朗讀日文回覆或中文草稿（edge-tts） |
| 📊 **匯出逐字稿** | 一鍵匯出 Excel / Word / TXT，含原文、翻譯、AI 回覆 |
| ↕️ **可調整版面** | 上下兩區 PanedWindow，拖曳 sash 自由調整翻譯區與回覆區比例 |
| 🔄 **自動更新** | 啟動時背景檢查 GitHub releases，有新版自動下載並重啟 |
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

> **注意**：system_prompt 需包含 JSON 格式指示（`ja_reply` / `zh_reply`），才能正確生成雙向回覆。自訂 profile 請使用獨特檔名，避免自動更新覆蓋。

---

## 版本說明

### v2.1（2026-06-25）

- **修正**：所有翻譯模式（ja2zh、ko2zh 等）均可正確監控，不再只認 ja_zh
- **修正**：單向模式翻譯內容正常顯示（LINE_RE 修正）
- **改善**：PanedWindow 可調整版面，拖曳 sash 改變上下比例
- **調整**：MAX_SENTENCES 7 → 30，AUTO_TRIGGER_DEBOUNCE 3s → 8s
- **新增**：自動更新機制（GitHub releases 背景檢查 + 一鍵更新）

### v2 新功能說明

#### 自動觸發模式
- 句尾偵測：`。！？!?`
- 偵測到句尾後靜默 **8 秒**才觸發（避免邊說邊觸發，並給 Ollama 足夠時間）
- 可與手動按鈕同時使用

#### 雙向回覆
- AI 同步生成「日文回覆」與「中文草稿」
- 使用 JSON 格式輸出（`ja_reply` + `zh_reply`）
- 若 AI 未返回 JSON 格式，自動 fallback 至全文顯示

#### 對話歷史記憶
- 每次生成將對話加入 `_conversation_history`
- 下次生成時帶入完整歷史（Ollama / Anthropic 使用 messages array，Gemini 使用文字串接）
- 點擊「清除對話歷史」重置
- 最多保留 10 輪（20 則訊息）

#### TTS 播放
- 依賴 `edge-tts`（`pip install edge-tts`）
- 日文語音：`ja-JP-NanamiNeural`
- 中文語音：`zh-TW-HsiaoChenNeural`

#### 匯出逐字稿
- Excel（需 `openpyxl`）：含藍色標頭、自動欄寬
- Word（需 `python-docx`）：表格格式，含匯出時間
- TXT：純文字，含時間戳與各欄位

---

## 自動更新

程式啟動時會在背景檢查 GitHub releases。若偵測到新版本：

1. 彈出對話框詢問是否更新
2. 確認後自動下載新版 zip（進度顯示於狀態列）
3. 解壓縮並寫入更新腳本
4. 程式關閉 → 腳本覆蓋檔案 → 自動重啟

> **jt-live-whisper 與 Whisper 模型不受影響**，只更新 SEI exe 本身。

---

## 目錄結構

```
Meeting_reply/
├── sei_reply_assistant.py      # 主程式（tkinter GUI）
├── profiles/
│   ├── General_JA.json         # 通用日文情境（範例）
│   └── General_EN.json         # 通用英文情境（範例）
├── .github/
│   └── workflows/
│       └── release.yml         # push v* tag 自動打包並發布 Release
├── CHANGELOG.md
├── requirements.txt
└── README.md
```

執行後自動建立：
- `settings.json` — 記住 provider / model / profile / jt-live-whisper 路徑
- `logs/` — jt-live-whisper 逐字稿輸出（開發模式備用）

---

## 安裝 jt-live-whisper（即時翻譯引擎）

本工具的即時翻譯功能依賴 [jt-live-whisper](https://github.com/jasoncheng7115/jt-live-whisper)。

### 方法一：下載 SEI_Bundle（建議，完整離線版）

包含 SEI Reply Assistant + jt-live-whisper + Whisper 模型（small / medium）+ NLLB 翻譯模型，完全離線即用。

1. [📥 下載 SEI_Bundle（Google Drive，~3.6 GB）](https://drive.google.com/file/d/1bSdbysjpoS4SY2S5EBkGdSgtHEugRPPw/view?usp=sharing)
2. 解壓縮後直接執行 `SEI_Reply_Assistant.exe`

> Whisper 模型選擇：啟動時可選 `small`（快，~2–3s）或 `medium`（準確，~8–12s）

### 方法二：從原始碼執行

```bash
git clone https://github.com/jasoncheng7115/jt-live-whisper.git
cd jt-live-whisper
pip install -r requirements.txt
```

### 設定路徑

啟動 `sei_reply_assistant.py` 後，點擊 **🎙 開啟即時翻譯** 旁的資料夾圖示，選擇 jt-live-whisper 的安裝目錄（含 `start.ps1` 或 `jt-live-whisper.exe` 的資料夾）。路徑會自動記憶於 `settings.json`。

---

## 發布新版本

```bash
git tag v2.x.x
git push origin v2.x.x
```

GitHub Actions 自動建置 `sei_reply_assistant_v2.x.x.zip` 並建立 Release。已安裝的客戶端啟動時自動偵測並提示更新。

---

## 系統需求

- **OS**：Windows 10/11 x64
- **Python**：3.10+（若使用 SEI_Bundle 則不需要）

---

## License

MIT
