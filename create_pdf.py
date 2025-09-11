from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pypdf import PdfWriter, PdfReader  # インポートは同じ
from io import BytesIO

# 新しいテキスト付きPDF作成（同じ）
packet = BytesIO()
can = canvas.Canvas(packet, pagesize=letter)
can.drawString(100, 100, "add text")
can.save()
packet.seek(0)
new_pdf = PdfReader(packet)

# 既存PDF読み込みと結合
existing_pdf = PdfReader("Ideal_Report.pdf")  # open不要、ファイルパス直接可
output = PdfWriter()

page = existing_pdf.pages[0]
new_page = new_pdf.pages[0]
page.merge_page(new_page)
output.add_page(page)

# 保存
with open("output.pdf", "wb") as output_file:
    output.write(output_file)