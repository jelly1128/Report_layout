import os, time, subprocess, argparse, sys
from weasyprint import HTML
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

def build_static_html_from_json(template_html: str, json_path: str) -> str:

    print("TEMPLATE_ABS:", Path(template_html).resolve())
    print("JSON_ABS    :", Path(json_path).resolve())

    html_text = Path(template_html).read_text(encoding="utf-8")
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))

    # 古い患者IDプレースホルダを削除
    html_text = re.sub(r"<div>【患者ID.+?】</div>", "", html_text)

    # --- report-meta の中身を置換 ---
    meta_html = f"""
      <div class="report-meta">
        <div>{data['header']['date']}</div>
        <div>【患者ID {data['header']['patient_id']}】</div>
      </div>
    """
    html_text = re.sub(
        r"<div class=\"report-meta\">.*?</div>",
        meta_html,
        html_text,
        flags=re.DOTALL
    )

    # --- 検査サマリー（表 tbody）を JSON rows から構築 ---
    rows_html = []
    for row in data["checks"]["rows"]:
        rows_html.append(f"""
            <tr>
              <td class="label">{row['label']}</td>
              <td class="mark">{row['mark']}</td>
              <td class="time">{row['time']}</td>
            </tr>
        """)
    new_tbody = "<tbody>" + "".join(rows_html) + "</tbody>"
    html_text = re.sub(
        r"<tbody>.*?</tbody>",
        new_tbody,
        html_text,
        flags=re.DOTALL
    )

    # --- 生検情報 ---
    biopsy = data["checks"].get("biopsy", {})
    biopsy_html = f"""
      <div class="exam-summary__biopsy-info">
        <span class="biopsy-label">{biopsy.get("method","")}</span>
        <span class="biopsy-target">{biopsy.get("target","")}</span>
      </div>
    """
    html_text = re.sub(
        r"<div class=\"exam-summary__biopsy-info\">.*?</div>",
        biopsy_html,
        html_text,
        flags=re.DOTALL
    )

    # --- 開始/終了時刻 + 体位図 ---
    times_html = f"""
      <aside class="exam-summary__times">
        <div>検査開始時刻　{data["checks"]['times']['start']}</div>
        <div>終了時刻　　　{data["checks"]['times']['end']}</div>
        <div class="exam-summary__position-image">
          <img src="position.png" alt="検査体位図" />
        </div>
      </aside>
    """
    html_text = re.sub(
        r"<aside class=\"exam-summary__times\">.*?</aside>",
        times_html,
        html_text,
        flags=re.DOTALL
    )

    # --- タイムライン ---
    rows_html = []
    for tl in data["timeline"]:
        markers_html = ""
        if tl.get("time_markers"):
            markers_html += '<div class="exam-timeline__time-markers">' + "".join(
                f'<div class="exam-timeline__time-marker" style="left:{m["x"]};"><span>{m["label"]}</span></div>'
                for m in tl["time_markers"]
            ) + "</div>"

        if tl.get("event_markers"):
            markers_html += '<div class="exam-timeline__markers">' + "".join(
                f'<div class="exam-timeline__marker" style="left:{m["x"]};"><span>{m["label"]}</span></div>'
                for m in tl["event_markers"]
            ) + "</div>"

        rows_html.append(f"""
          <div class="exam-timeline__row">
            <div class="exam-timeline__caption">{tl['caption']}</div>
            <div class="exam-timeline__track">
              {markers_html}
              <img src="{tl['img']}" alt="{tl['caption']}タイムライン">
            </div>
          </div>
        """)

    new_timeline = f"""
    <section class="exam-timeline">
      <div class="exam-timeline__header">
        <div class="exam-timeline__caption">経過時間</div>
      </div>
      {''.join(rows_html)}
    </section>
    """
    html_text = re.sub(
        r"<section class=\"exam-timeline\">.*?</section>",
        new_timeline,
        html_text,
        flags=re.DOTALL
    )

    # --- ギャラリー ---
    gallery_html = ""
    for g in data.get("gallery", []):
        thumbs_html = "".join(
            f"""
            <div class="exam-gallery__thumb">
              <span class="exam-gallery__thumb-label">{g['label']}<small>{img['index']}</small> <span>{img['time']}</span></span>
              <img src="{img['src']}" alt="">
            </div>
            """
            for img in g["images"]
        )
        gallery_html += f"""
        <div class="exam-gallery__block">
          <div class="exam-gallery__caption"><strong>{g["label"]}</strong><br>{g["caption"]}</div>
          <div class="exam-gallery__thumbnails">
            {thumbs_html}
          </div>
        </div>
        """

    new_gallery = f"<section class=\"exam-gallery\">{gallery_html}</section>"
    html_text = re.sub(
        r"<section class=\"exam-gallery\">.*?</section>",
        new_gallery,
        html_text,
        flags=re.DOTALL
    )

    # --- 一時HTMLを書き出し ---
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    Path(tmp.name).write_text(html_text, encoding="utf-8")

    # デバッグ用にも保存
    debug_html = Path("debug_output.html")
    debug_html.write_text(html_text, encoding="utf-8")
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
