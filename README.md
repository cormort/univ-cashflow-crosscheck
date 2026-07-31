---
title: 大學現金流量表勾稽底稿產生器
emoji: 📊
colorFrom: blue
colorTo: indigo
sdk: static
app_file: index.html
pinned: false
---

# 大學現金流量表勾稽底稿產生器

上傳 `12.平衡xxxx.xlsx`（預計平衡表）與 `3.現流xxxx.xlsx`（現金流量預計表），
產生「總表 + 47 校明細」的勾稽底稿 xlsx。

## 檔案不會上傳

整支程式在瀏覽器裡用 [Pyodide](https://pyodide.org)（Python 編譯成 WebAssembly）執行。
來源檔只進到瀏覽器分頁的記憶體，**不會送到 Hugging Face 或任何伺服器**；
產出的 xlsx 也是在本機組出來後用 blob 下載。這個 Space 是純靜態網頁，沒有後端。

第一次開啟要下載約 10MB 的 Python 執行環境，之後由瀏覽器快取。實測產生一次約 6 秒。

## 產出內容

- 年度取自來源檔標題（`中華民國NNN年…`），可在介面覆寫
- 47 張明細表與總表**逐列對齊**，總表跨表加總一律指同一列，整欄公式可直接往下拉
- 調整欄 D/F 的公式保留，金額仍由人工填；填完 K 欄與檢查欄自動算
- 五份原始資料只寫一份在總表底下，明細表用 VLOOKUP／OFFSET 指過去

## 檔案

| 檔案 | 說明 |
|---|---|
| `index.html` | 網頁介面（Pyodide，無後端） |
| `build_gouji.py` | 產生邏輯；瀏覽器與命令列共用 |
| `template.xlsx` + `.mirror.json` | 骨架範本，**已清掉所有年度金額**，只留科目名稱 |
| `test_build.py` | 自我檢查：與參考檔逐格比對 B/J/L、驗證逐列對齊 |

## 部署

純靜態網站，兩邊都可以，檔案內容完全一樣。

**GitHub Pages**：把 `index.html`、`build_gouji.py`、`template.xlsx`、
`template.xlsx.mirror.json`、`.nojekyll` 推到 repo，
Settings → Pages → Source 選該分支的根目錄即可。
`.nojekyll` 一定要有，否則 Jekyll 會插手處理靜態檔。
子路徑（`https://<帳號>.github.io/<repo>/`）已實測可用。

**Hugging Face Space**：建 Space、SDK 選 **Static**，把上表檔案（含 `README.md`）推上去。

`template.xlsx` 約 1.1MB，兩邊都不需要 git-lfs。

> 注意：GitHub Pages 一律是公開的（即使 repo 是 private）。
> 這裡沒有敏感內容——範本已清空金額，來源檔也從不離開瀏覽器——但網址等於公開。
> 想限制存取就用 HF Space 的 Private。

## 命令列用法

```bash
pip install openpyxl
python build_gouji.py 12.平衡0814.xlsx 3.現流0814.xlsx
python test_build.py 大學115現金流量表勾稽檔-47所大學.xlsx 大學116現金流量表勾稽檔-47所大學.xlsx
```

從新的歷年勾稽檔重做範本：

```bash
python build_gouji.py x y -t 大學NNN現金流量表勾稽檔-47所大學.xlsx --strip-template template.xlsx
```

## 本機預覽網頁版

```bash
python -m http.server 8765   # 然後開 http://127.0.0.1:8765
```
