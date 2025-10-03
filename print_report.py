import os, time, subprocess, argparse, sys
from weasyprint import HTML
from bs4 import BeautifulSoup
import json, tempfile
from pathlib import Path
import re

# Windows の場合だけ win32print を使う
if sys.platform.startswith("win"):
    import win32print

# 必要な場合のみ（WeasyPrint用のDLLパス）
DLL_DIR = r"C:\msys64\mingw64\bin"
if os.path.isdir(DLL_DIR):
    os.add_dll_directory(DLL_DIR)

DEFAULT_SUMATRA = os.path.join(os.path.dirname(__file__), "bin", "SumatraPDF.exe")


def update_report_meta(soup, data):
    meta_div = soup.find("div", class_="report-meta")
    if not meta_div:
        return
    meta_div.clear()

    # 日付
    new_date = soup.new_tag("div")
    new_date.string = data["header"]["date"]
    meta_div.append(new_date)

    # 患者ID
    new_id = soup.new_tag("div")
    new_id.string = f"【患者ID 　　　　】"
    meta_div.append(new_id)


def update_exam_summary(soup, data):
    """
    検査サマリー（表、biopsy、開始/終了時刻）の更新
    """
    # --- 表 (tbody) ---
    tbody = soup.select_one(".exam-summary__table tbody")
    if tbody:
        tbody.clear()
        for row in data["checks"]["rows"]:
            tr = soup.new_tag("tr")
            td1 = soup.new_tag("td", **{"class": "label"}); td1.string = row["label"]
            td2 = soup.new_tag("td", **{"class": "mark"});  td2.string = row["mark"]
            td3 = soup.new_tag("td", **{"class": "time", "lang": "en"});  td3.string = row["time"]
            tr.extend([td1, td2, td3])
            tbody.append(tr)

    # --- 生検情報 ---
    biopsy_div = soup.find("div", class_="exam-summary__biopsy-info")
    if biopsy_div:
        biopsy_div.clear()
        label_span = soup.new_tag("span", **{"class": "biopsy-label"})
        label_span.string = data["checks"]["biopsy"].get("method", "")
        target_span = soup.new_tag("span", **{"class": "biopsy-target"})
        target_span.string = data["checks"]["biopsy"].get("target", "")
        biopsy_div.extend([label_span, target_span])

    # --- 開始/終了時刻 + 体位図 ---
    aside = soup.find("aside", class_="exam-summary__times")
    if aside:
        aside.clear()
        start_div = soup.new_tag("div", **{"lang": "en"})
        start_div.string = f"検査開始時刻　{data['checks']['times']['start']}"
        end_div = soup.new_tag("div", **{"lang": "en"})
        end_div.string = f"終了時刻　　{data['checks']['times']['end']}"
        img_div = soup.new_tag("div", **{"class": "exam-summary__position-image"})
        img_tag = soup.new_tag("img", src="position.png", alt="検査体位図")
        img_div.append(img_tag)
        aside.extend([start_div, end_div, img_div])

def update_exam_timeline(soup, data):
    """
    タイムライン部分を JSON データから更新する
    """
    section = soup.find("section", class_="exam-timeline")
    if not section:
        return

    # ヘッダー行を残して他を削除
    header = section.find("div", class_="exam-timeline__header")
    section.clear()
    if header:
        section.append(header)

    # JSONデータから行を生成
    for tl in data.get("timeline", []):
        row_div = soup.new_tag("div", **{"class": "exam-timeline__row"})

        # 左ラベル
        caption_div = soup.new_tag("div", **{"class": "exam-timeline__caption"})
        caption_div.string = tl["caption"]
        row_div.append(caption_div)

        # 右側（トラック）
        track_div = soup.new_tag("div", **{"class": "exam-timeline__track"})

        # 時間マーカー
        if tl.get("time_markers"):
            tm_container = soup.new_tag("div", **{"class": "exam-timeline__time-markers"})
            for m in tl["time_markers"]:
                marker = soup.new_tag("div", **{
                    "class": "exam-timeline__time-marker",
                    "style": f"left: {m['x']};"
                })
                span = soup.new_tag("span", **{"lang": "en"})
                span.string = m["label"]
                marker.append(span)
                tm_container.append(marker)
            track_div.append(tm_container)

        # イベントマーカー
        if tl.get("event_markers"):
            em_container = soup.new_tag("div", **{"class": "exam-timeline__markers"})
            for m in tl["event_markers"]:
                marker = soup.new_tag("div", **{
                    "class": "exam-timeline__marker",
                    "style": f"left: {m['x']};"
                })
                span = soup.new_tag("span", **{"lang": "en"}, **{"data-char": m.get("char", f"{m['label']}")})
                span.string = m["label"]
                marker.append(span)
                em_container.append(marker)
            track_div.append(em_container)

        # 画像
        img = soup.new_tag("img", src=tl["img"], alt=f"{tl['caption']}タイムライン")
        track_div.append(img)

        row_div.append(track_div)
        section.append(row_div)


def build_static_html_from_json(template_html: str, json_path: str) -> str:

    print("TEMPLATE_ABS:", Path(template_html).resolve())
    print("JSON_ABS    :", Path(json_path).resolve())

    html_text = Path(template_html).read_text(encoding="utf-8")
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    soup = BeautifulSoup(html_text, "lxml")

    # 更新処理を呼び出す
    update_report_meta(soup, data)
    update_exam_summary(soup, data)
    update_exam_timeline(soup, data)

    # --- 一時HTMLを書き出し ---
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    Path(tmp.name).write_text(str(soup), encoding="utf-8")

    # デバッグ用にも保存
    debug_html = Path("debug_output.html")
    debug_html.write_text(soup.prettify(), encoding="utf-8")
    print(f"DEBUG: also copied to {debug_html.resolve()}")

    return tmp.name

def html_to_pdf(html_path: str, pdf_path: str) -> str:
    html_abs = os.path.abspath(html_path)
    pdf_abs  = os.path.abspath(pdf_path)
    if not os.path.exists(html_abs):
        raise FileNotFoundError(html_abs)
    if os.path.exists(pdf_abs):
        try: os.remove(pdf_abs)
        except PermissionError: raise RuntimeError(f"PDF使用中: {pdf_abs}")
    
    base = Path(pdf_abs).parent.resolve()   # ← 元の grok.html があるディレクトリ
    print(f"base: {base}")
    HTML(filename=html_path, base_url=base).write_pdf(pdf_path)
    for _ in range(10):
        if os.path.exists(pdf_abs) and os.path.getsize(pdf_abs) > 0:
            return pdf_abs
        time.sleep(0.1)
    raise RuntimeError("PDF生成に失敗（サイズ0バイト）")

def print_with_sumatra(pdf_abs: str, printer: str, sumatra_exe: str):
    exe = sumatra_exe or DEFAULT_SUMATRA
    if not os.path.exists(exe):
        raise FileNotFoundError(f"SumatraPDF.exe が見つかりません: {exe}")
    subprocess.run([exe, "-print-to", printer, "-exit-on-print", pdf_abs], check=True)

def main():
    p = argparse.ArgumentParser(description="HTML→PDF→印刷（SumatraPDF利用）")
    p.add_argument("html", help="入力HTMLファイル")
    p.add_argument("pdf", help="出力PDFファイル")
    p.add_argument("--mode", choices=["device", "pdf"], default="device",
                   help='device=実機プリンタ / pdf=Microsoft Print to PDF')
    p.add_argument("--printer", default="", help="実機プリンタ名（mode=device時）")
    p.add_argument("--sumatra", default="", help="SumatraPDF.exe のパス（未指定で ./bin/SumatraPDF.exe）")
    args = p.parse_args()

    template_html = args.html
    json_path = os.path.join(os.path.dirname(args.html), "report.json")

    # HTML + JSON → 静的HTML
    static_html = build_static_html_from_json(template_html, json_path)

    # 静的HTML → PDF
    pdf_abs = html_to_pdf(static_html, args.pdf)


    # if args.mode == "pdf":
    #     # Microsoft Print to PDF へ出力（保存ダイアログが出ます）
    #     printer = r"Microsoft Print to PDF"
    # else:
    #     printer = args.printer or win32print.GetDefaultPrinter()

    # # PDF → プリンタへ出力
    # print_with_sumatra(pdf_abs, printer, args.sumatra)
    # print(f"送信: {pdf_abs} → {printer}")

if __name__ == "__main__":
    main()
