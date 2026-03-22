import os
import re

def extract_content_regex(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Define sections
    sections = [
        "11.3.1", "11.3.2", "11.3.3",
        "11.4.1", "11.4.2", "11.4.3",
        "11.5.1", "11.5.2", "11.5.3",
        "11.6.1", "11.6.2", "11.6.3",
        "12.1" # End marker
    ]
    
    results = {}
    for i in range(len(sections)-1):
        start = sections[i]
        end = sections[i+1]
        
        # Look for the section title and text following it
        # The section title might be split by newlines or formatting
        pattern = re.escape(start) + r"(.*?)" + re.escape(end)
        match = re.search(pattern, content, re.DOTALL)
        if match:
            results[start] = match.group(1).strip()
        else:
            # Try a looser match if literal fail
            results[start] = "Not found"
            
    return results

if __name__ == "__main__":
    search_file = r"e:\教学管理\ch11_search.txt"
    outfile = r"e:\教学管理\ch11_outline_v2.txt"
    
    data = extract_content_regex(search_file)
    with open(outfile, "w", encoding="utf-8") as out:
        for k, v in data.items():
            out.write(f"=== {k} ===\n{v[:2000]}\n\n") # Limit length
    
    print(f"Outline saved to {outfile}")
