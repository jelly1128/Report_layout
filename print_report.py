import os, time, subprocess, argparse, win32print
from weasyprint import HTML
import json, tempfile
from pathlib import Path
import re

# 必要な場合のみ（WeasyPrint用のDLLパス）
DLL_DIR = r"C:\msys64\mingw64\bin"
if os.path.isdir(DLL_DIR):
    os.add_dll_directory(DLL_DIR)

DEFAULT_SUMATRA = os.path.join(os.path.dirname(__file__), "bin", "SumatraPDF.exe")

def build_static_html_from_json(template_html: str, json_path: str) -> str:

    print("TEMPLATE_ABS:", Path(template_html).resolve())
    print("JSON_ABS    :", Path(json_path).resolve())
    p = Path(json_path)
    print("EXISTS      :", p.exists(), "SIZE:", p.stat().st_size if p.exists() else -1)
    with open(p, "rb") as f:
        head = f.read(16)
    print("HEAD_BYTES  :", head)

    
    """grok.html の <section id="gallery"> を JSONの images で置換した静的HTMLを作る"""
    html_text = Path(template_html).read_text(encoding="utf-8")
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))

    # --- report-meta の中身を置換 ---
    meta_html = f"""
      <div class="report-meta">
        <div>{data['header']['date']}　【患者ID {data['header']['patient_id']}】</div>
      </div>
    """
    html_text = re.sub(
        r"<div class=\"report-meta\">.*?</div>",
        meta_html,
        html_text,
        flags=re.DOTALL
    )

    # --- tbody を JSON rows から構築 ---
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

    # JSONから aside 内のHTMLを構築
    times_html = f"""
    <aside class="checks__times">
      <div>検査開始時刻　{data["checks"]['times']['start']}</div>
      <div>終了時刻　　　{data["checks"]['times']['end']}</div>

      <div class="position-image">
          <img src="position.png" alt="検査体位図" />
      </div>
    </aside>
    """

    # 元の <aside class="checks__times">…</aside> を置換
    html_text = re.sub(
        r"<aside class=\"checks__times\">.*?</aside>",
        times_html,
        html_text,
        flags=re.DOTALL
    )

    # timeline の各 row を構築
    rows_html = []
    for tl in data["timeline"]:
        # markers があれば div を作る
        markers_html = ""
        if tl.get("markers"):
            markers_html = '<div class="markers">' + "".join(
                f'<div class="marker" style="left: {m["x"]};">'
                f'<span>{m["label"]}</span></div>'
                for m in tl["markers"]
            ) + "</div>"

        rows_html.append(f"""
        <div class="timeline-row">
            <div class="timeline-caption">{tl['caption']}</div>
            <div class="timeline">
            {markers_html}
            <img src="{tl['img']}">
            </div>
        </div>
        """)

    # <section class="timeline-block">…</section> を置換
    new_section = f"<section class=\"timeline-block\">{''.join(rows_html)}\n    </section>"
    html_text = re.sub(
        r"<section class=\"timeline-block\">.*?</section>",
        new_section,
        html_text,
        flags=re.DOTALL
    )

    # --- ギャラリー置換 ---
    gallery_html = ""
    for g in data.get("gallery", []):
        imgs_html = "".join(
            f'<img src="{img["src"]}" alt="{img.get("alt","")}">'
            for img in g["images"]
        )
        gallery_html += f"""
        <div class="gallery-block">
        <div class="gallery-caption">{g["caption"]}</div>
        <div class="gallery-thumbnails">
            {imgs_html}
        </div>
        </div>
        """

    new_gallery = f"<section class='gallery'>{gallery_html}</section>"
    html_text = re.sub(
        r"<section class=\"gallery\">.*?</section>",
        new_gallery,
        html_text,
        flags=re.DOTALL
    )

    # imgs = []
    # for img in data.get("images", []):
    #     src = img.get("src", "")
    #     alt = img.get("alt", "")
    #     imgs.append(f'<img src="{src}" alt="{alt}">')
    # gallery_html = "\n        ".join(imgs) if imgs else ""

    # # ギャラリー置換（JSは不要なので削除）
    # html_text = html_text.replace(
    #     '<section class="gallery" id="gallery">',
    #     '<section class="gallery" id="gallery">\n        ' + gallery_html
    # )
    # 一時HTMLを書き出し
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    Path(tmp.name).write_text(html_text, encoding="utf-8")

    # デバッグ用にも保存
    debug_html = Path("report/debug_output.html")
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

    if args.mode == "pdf":
        # Microsoft Print to PDF へ出力（保存ダイアログが出ます）
        printer = r"Microsoft Print to PDF"
    else:
        printer = args.printer or win32print.GetDefaultPrinter()

    # PDF → プリンタへ出力
    print_with_sumatra(pdf_abs, printer, args.sumatra)
    print(f"送信: {pdf_abs} → {printer}")

if __name__ == "__main__":
    main()
