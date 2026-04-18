import pypdf
import os

def extract_chapter_11_perfectly(pdf_path, output_path):
    # Using pypdf to extract text page by page
    # Chapter 11 starts at book page 338, which is relative page 353 in the PDF
    with open(pdf_path, 'rb') as f:
        reader = pypdf.PdfReader(f)
        text = ""
        # 353 to 370 (indices)
        for i in range(352, 370):
            if i < len(reader.pages):
                page_text = reader.pages[i].extract_text()
                text += f"\n--- PDF Page {i+1} ---\n"
                text += page_text
        
        with open(output_path, 'w', encoding='utf-8') as out:
            out.write(text)

if __name__ == "__main__":
    pdf = r"e:\教学管理\信息系统项目管理师教程-第四版.pdf"
    out = r"e:\教学管理\ch11_prof_text_final.txt"
    extract_chapter_11_perfectly(pdf, out)
