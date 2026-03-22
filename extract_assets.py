import fitz
import os

def extract_images_and_tables(pdf_path, output_dir, start_page, end_page):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    doc = fitz.open(pdf_path)
    for page_num in range(start_page - 1, end_page):
        page = doc.load_page(page_num)
        
        # Extract images
        images = page.get_images(full=True)
        print(f"Page {page_num + 1}: Found {len(images)} images")
        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            ext = base_image["ext"]
            img_name = f"page_{page_num + 1}_img_{img_index + 1}.{ext}"
            with open(os.path.join(output_dir, img_name), "wb") as f:
                f.write(image_bytes)
        
        # Pixmap of the whole page (for tables or complex diagrams)
        # We can also take a screenshot of specific areas if we find them, 
        # but for now let's just save the page as an image to look at it.
        pix = page.get_pixmap()
        pix.save(os.path.join(output_dir, f"page_{page_num + 1}_full.png"))

if __name__ == "__main__":
    pdf = r"e:\教学管理\信息系统项目管理师教程-第四版.pdf"
    out = r"e:\教学管理\ch11_images"
    # Chapter 11 is roughly 354-370
    extract_images_and_tables(pdf, out, 354, 370)
