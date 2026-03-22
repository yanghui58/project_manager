import os
from pypdf import PdfReader

def search_pdf_content(pdf_path):
    reader = PdfReader(pdf_path)
    print(f"Total pages: {len(reader.pages)}")
    
    # Check page 1-20 for TOC
    print("--- TOC Search ---")
    for i in range(min(50, len(reader.pages))):
        text = reader.pages[i].extract_text()
        if text:
            if "第11章" in text:
                print(f"Page {i+1}: Found '第11章'")
            if "项目成本管理" in text:
                print(f"Page {i+1}: Found '项目成本管理'")
        else:
            print(f"Page {i+1}: No text extracted (likely image/scan)")

if __name__ == "__main__":
    pdf_path = r"e:\教学管理\信息系统项目管理师教程-第四版.pdf"
    search_pdf_content(pdf_path)
