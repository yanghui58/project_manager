import pypdfium2 as pdfium
from rapidocr_onnxruntime import RapidOCR
import os

def extract_text_via_ocr():
    pdf_path = r"e:\教学管理\信息系统项目管理师教程-第四版.pdf"
    out_path = r"e:\教学管理\ch11_ocr_text.txt"
    
    ocr = RapidOCR()
    pdf = pdfium.PdfDocument(pdf_path)
    
    all_text = ""
    
    # Chapter 11 spans from viewer index 352 to 369 (inclusive)
    # We add 1 more page just in case.
    for page_idx in range(352, 372):
        if page_idx >= len(pdf):
            break
        print(f"Processing Page {page_idx + 1}...")
        all_text += f"\n--- PDF Page {page_idx + 1} ---\n"
        
        page = pdf[page_idx]
        # Render page to PIL image. scale=2 yields ~144 DPI (good enough for OCR).
        # We can use scale=3 for better accuracy.
        bitmap = page.render(scale=3)
        pil_image = bitmap.to_pil()
        
        # Save temp image for OCR
        tmp_img = "temp_page.png"
        pil_image.save(tmp_img)
        
        # OCR extraction
        result, _ = ocr(tmp_img)
        
        if result:
            for line in result:
                # result is [ [[x1,y1],[x2,y2],[x3,y3],[x4,y4]], text, confidence ]
                text = line[1]
                all_text += text + "\n"
                
        # Clean up
        if os.path.exists(tmp_img):
            os.remove(tmp_img)
            
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(all_text)
        
    print(f"OCR Extraction Complete! Saved to {out_path}")

if __name__ == "__main__":
    extract_text_via_ocr()
