from pypdf import PdfReader

def find_chapter_range(pdf_path):
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    
    ch11_start = -1
    ch12_start = -1
    
    for i in range(total_pages):
        # Extract first 500 chars to find chapter headers
        try:
            text = reader.pages[i].extract_text()
            if not text: continue
            
            if "第11章" in text and "项目成本管理" in text:
                print(f"Chapter 11 starts at page {i+1}")
                ch11_start = i
            if "第12章" in text and "项目质量管理" in text:
                print(f"Chapter 12 starts at page {i+1}")
                ch12_start = i
                break
        except:
            continue
            
    return ch11_start, ch12_start

if __name__ == "__main__":
    pdf_path = r"e:\教学管理\信息系统项目管理师教程-第四版.pdf"
    start, end = find_chapter_range(pdf_path)
    print(f"Range: {start+1} to {end+1}")
    
    if start != -1:
        reader = PdfReader(pdf_path)
        content = ""
        actual_end = end if end != -1 else start + 50 # Fallback
        for i in range(start, actual_end):
            content += f"--- Page {i+1} ---\n"
            content += reader.pages[i].extract_text() + "\n"
        
        with open(r"e:\教学管理\chapter11_raw.txt", "w", encoding="utf-8") as f:
            f.write(content)
        print("Chapter 11 content saved to chapter11_raw.txt")
