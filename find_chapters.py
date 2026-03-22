import os
from pypdf import PdfReader

def search_chapters(pdf_path):
    reader = PdfReader(pdf_path)
    found_chapters = {}
    
    # Check table of contents or search beginning of pages
    # Usually TOC is in the first 20 pages
    for i in range(min(30, len(reader.pages))):
        text = reader.pages[i].extract_text()
        if "第10章" in text and "项目进度管理" in text:
            print(f"Potential TOC entry for Ch10 at page {i+1}")
        if "第11章" in text and "项目成本管理" in text:
            print(f"Potential TOC entry for Ch11 at page {i+1}")

    # Broad search for chapter start pages
    for i in range(len(reader.pages)):
        # Check first 200 characters of each page for chapter title
        text = reader.pages[i].extract_text()[:200]
        if "第10章" in text and "项目进度管理" in text:
            found_chapters["Ch10"] = i
            print(f"Found Chapter 10 at page {i+1}")
        if "第11章" in text and "项目成本管理" in text:
            found_chapters["Ch11"] = i
            print(f"Found Chapter 11 at page {i+1}")
        if len(found_chapters) == 2:
            break
            
    return found_chapters

if __name__ == "__main__":
    pdf_path = r"e:\教学管理\信息系统项目管理师教程-第四版.pdf"
    chapters = search_chapters(pdf_path)
    print(f"Chapters found: {chapters}")
