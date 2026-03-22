import pdfplumber

def check_page_349():
    with pdfplumber.open(r'e:\教学管理\信息系统项目管理师教程-第四版.pdf') as pdf:
        # User said page 349. Note pdfplumber pages are 0-indexed. 
        # So page 349 in PDF viewer is index 348.
        page = pdf.pages[348]
        text = page.extract_text()
        print("--- Viewer Page 349 (Index 348) ---")
        print(text[:200])

        # Let's also check index 352 (Viewer page 353) which is where Book Page 338 
        # (Chapter 11 start) was previously found.
        page2 = pdf.pages[352]
        text2 = page2.extract_text()
        print("--- Viewer Page 353 (Index 352) ---")
        print(text2[:200])
        
if __name__ == '__main__':
    check_page_349()
