from pypdf import PdfReader
import sys

def extract_pages(pdf_path, start_page, end_page, output_txt):
    reader = PdfReader(pdf_path)
    with open(output_txt, "w", encoding="utf-8") as f:
        for i in range(start_page - 1, min(end_page, len(reader.pages))):
            try:
                text = reader.pages[i].extract_text()
                f.write(f"--- PDF Page {i+1} ---\n")
                f.write(text + "\n")
            except Exception as e:
                f.write(f"--- PDF Page {i+1} Error: {e} ---\n")

if __name__ == "__main__":
    pdf_path = r"e:\教学管理\信息系统项目管理师教程-第四版.pdf"
    extract_pages(pdf_path, 300, 450, r"e:\教学管理\ch11_search.txt")
    print("Extraction complete: e:\教学管理\ch11_search.txt")
