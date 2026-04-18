import os

def extract_section(file_path, start_pattern, next_pattern=None):
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    
    extracted = []
    found = False
    for line in lines:
        if start_pattern in line:
            found = True
        if found:
            if next_pattern and next_pattern in line:
                break
            extracted.append(line.strip())
            if len(extracted) > 100: # Safety break
                break
    return "\n".join(extracted)

if __name__ == "__main__":
    search_file = r"e:\教学管理\ch11_search.txt"
    outfile = r"e:\教学管理\ch11_outline.txt"
    
    with open(outfile, "w", encoding="utf-8") as out:
        for section in ["11.3", "11.4", "11.5", "11.6"]:
            out.write(f"=== Process {section} ===\n")
            for sub in [".1", ".2", ".3"]:
                pattern = f"{section}{sub}"
                text = extract_section(search_file, pattern, f"{section}.{int(sub[1])+1}" if sub != ".3" else f"11.{int(section[3])+1}.1")
                out.write(f"\n--- {pattern} ---\n{text}\n")
    
    print(f"Outline saved to {outfile}")
