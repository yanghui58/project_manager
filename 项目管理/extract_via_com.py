import win32com.client
import os

def extract_via_com():
    pdf_path = os.path.abspath(r'e:\教学管理\信息系统项目管理师教程-第四版.pdf')
    out_path = os.path.abspath(r'e:\教学管理\ch11_wps_text.txt')
    
    print("Dispatching Word/WPS...")
    try:
        word = win32com.client.Dispatch('KWPS.Application')
        print("Using KWPS.Application")
    except Exception:
        try:
            word = win32com.client.Dispatch('Word.Application')
            print("Using Word.Application")
        except Exception as e:
            print("Failed to dispatch COM:", e)
            return

    word.Visible = True  # As requested, open it so it can be seen
    print(f"Opening PDF: {pdf_path}")
    
    try:
        # Format 2 is wdFormatText
        doc = word.Documents.Open(pdf_path, ReadOnly=True, ConfirmConversions=False)
        print("PDF opened, saving as text...")
        doc.SaveAs2(out_path, FileFormat=2, Encoding=65001) # UTF-8 text
        doc.Close(False)
        print(f"Successfully saved to {out_path}")
    except Exception as e:
        print("Error during document operations:", e)
    finally:
        word.Quit()

if __name__ == "__main__":
    extract_via_com()
