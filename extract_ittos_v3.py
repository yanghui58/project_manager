import re

def clean_text(text):
    # Remove excessive newlines and spaces
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_ittos(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Simple markers for sections
    # Processes: 11.3 (Plan), 11.4 (Estimate), 11.5 (Budget), 11.6 (Control)
    process_names = {
        "11.3": "规划成本管理",
        "11.4": "估算成本",
        "11.5": "制定预算",
        "11.6": "控制成本"
    }
    
    results = {}
    for code, name in process_names.items():
        results[name] = {"输入": [], "工具与技术": [], "输出": []}
        for i, sub in enumerate(["输入", "工具与技术", "输出"], 1):
            pattern = f"{code}.{i}"
            # Find the section and take the next few paragraphs
            # We look for bullet points or numbered lists
            start_idx = content.find(pattern)
            if start_idx != -1:
                chunk = content[start_idx:start_idx+1500]
                # Extract numbered items (1. 2. 3. ...)
                items = re.findall(r'\d\s*\.\s*([^\d\n]+)', chunk)
                results[name][sub] = [item.strip() for item in items if len(item.strip()) > 1][:10]
    
    return results

if __name__ == "__main__":
    search_file = r"e:\教学管理\ch11_search.txt"
    data = extract_ittos(search_file)
    for p, ittos in data.items():
        print(f"Process: {p}")
        for k, v in ittos.items():
            print(f"  {k}: {v}")
