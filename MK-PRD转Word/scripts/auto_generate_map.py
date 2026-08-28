import os
import yaml
import glob
import re

def extract_page_info(md_path):
    """
    从 Markdown 文件中提取所有的页面名称。
    支持多种格式：
    1. ### 1.1 页面名称：车辆档案列表
    2. ## 页面1：保养记录列表
    3. | 页面名称 | XXX | （表格行中的页面名称）
    """
    page_names = []
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 格式1: ### 1.1 页面名称：XXX
    matches = re.findall(r'###\s*\d+\.\d+\s*页面名称[:：]\s*(.*?)(?=\n|$)', content)
    # 格式2: ## 页面N：XXX
    matches += re.findall(r'##\s*页面\d+[:：]\s*(.*?)(?=\n|$)', content)
    # 格式3: 表格行 | 页面名称 | XXX |
    table_matches = re.findall(r'\|\s*页面名称\s*\|\s*(.*?)\s*\|', content)

    seen = set()
    for m in matches + table_matches:
        name = m.strip()
        if name and name not in seen:
            seen.add(name)
            page_names.append(name)

    return page_names

def find_best_html_match(page_name, html_files):
    """
    智能匹配：在 html_files 列表中寻找与 page_name 最匹配的 html 文件
    """
    best_match = None
    max_score = 0
    
    # 清理页面名称：去掉特殊字符，保留核心中文和英文
    clean_page_name = re.sub(r'[^\w\u4e00-\u9fa5]', '', page_name).lower()
    
    for html_file in html_files:
        score = 0
        html_base = os.path.basename(html_file).lower()
        
        # 清理 html 文件名：去掉 .html 后缀和特殊字符
        clean_html_base = re.sub(r'\.html$', '', html_base)
        clean_html_base = re.sub(r'[^\w\u4e00-\u9fa5]', '', clean_html_base)
        
        # 1. 完全包含 (名字完全在 HTML 文件名中)
        if clean_page_name in clean_html_base:
            score += 20
            
        # 2. 核心词交集计算 (按两个字的分词来算交集)
        page_grams = set([clean_page_name[i:i+2] for i in range(len(clean_page_name)-1)]) if len(clean_page_name) > 1 else set([clean_page_name])
        html_grams = set([clean_html_base[i:i+2] for i in range(len(clean_html_base)-1)]) if len(clean_html_base) > 1 else set([clean_html_base])
        
        if page_grams and html_grams:
            intersection = page_grams.intersection(html_grams)
            score += len(intersection) * 2
            
        # 3. 页面类型后缀特征匹配
        if "列表" in clean_page_name and ("list" in html_base or "列表" in html_base):
            score += 5
        if ("表单" in clean_page_name or "新建" in clean_page_name or "编辑" in clean_page_name) and ("form" in html_base or "edit" in html_base or "新建" in html_base or "编辑" in html_base):
            score += 5
        if ("详情" in clean_page_name or "查看" in clean_page_name) and ("detail" in html_base or "详情" in html_base):
            score += 5
        if "弹窗" in clean_page_name and ("modal" in html_base or "dialog" in html_base or "弹窗" in html_base):
            score += 5
            
        if score > max_score:
            max_score = score
            best_match = html_file
            
    # 只要得分大于阈值，就认为找到了
    if max_score >= 2:
        return best_match
    return None

def auto_generate_map(prd_dir, prototype_dir, output_map_file="export-map.yaml"):
    """
    自动扫描目录，推断映射关系，并生成 export-map.yaml (1对多页面结构)
    """
    print(f"🔍 开始自动扫描 PRD 目录: {prd_dir}")
    print(f"🔍 开始自动扫描 原型 目录: {prototype_dir}")
    
    stories_dir = os.path.join(prd_dir, "stories")
    if not os.path.exists(stories_dir):
        print(f"⚠️ 在 {prd_dir} 未找到 stories 目录。请检查路径。")
        return
        
    md_files = glob.glob(os.path.join(stories_dir, "*.md"))
    md_files.sort() # 按 01, 02 排序
    
    html_files = glob.glob(os.path.join(prototype_dir, "pages", "*.html"))
    # 为了跨平台，转换为正斜杠表示的统一相对路径
    html_basenames = [os.path.relpath(f, prototype_dir).replace('\\', '/') for f in html_files]
    
    mappings = []
    
    for md_file in md_files:
        page_names = extract_page_info(md_file)
        
        # 强制转换为正斜杠，确保 yaml 在多端可用
        rel_md_path = os.path.relpath(md_file, prd_dir).replace('\\', '/')
        
        if not page_names:
            print(f"❓ 未能从 {os.path.basename(md_file)} 中提取到规范的页面名称。")
            continue
            
        pages_mapping = []
        for page_name in page_names:
            matched_html = find_best_html_match(page_name, html_basenames)
            
            if matched_html:
                print(f"✅ 找到匹配: {os.path.basename(md_file)} [{page_name}] -> {matched_html}")
                pages_mapping.append({
                    "page_name": page_name,
                    "html": matched_html
                })
            else:
                print(f"❓ 未找到匹配: {os.path.basename(md_file)} [{page_name}]")
                pages_mapping.append({
                    "page_name": page_name,
                    "html": "手动填写.html"
                })
                
        if pages_mapping:
            mappings.append({
                "markdown": rel_md_path,
                "pages": pages_mapping
            })

    # 从 prd_dir 提取模块名称 (例如 "PRD文档/车辆管理-v2" -> "车辆管理")
    # 当 prd_dir 是相对路径 "." 时，basename 会得到 "."，因此用 abspath 兜底
    abs_prd = os.path.abspath(prd_dir)
    dir_name = os.path.basename(os.path.normpath(abs_prd))
    module_name = re.sub(r'-v\d+.*$', '', dir_name)
    if not module_name or module_name == '.':
         module_name = "自动生成"

    # 生成 YAML
    config_data = {
        "config": {
            "project_name": module_name,
            "prd_dir": prd_dir,
            "prototype_dir": prototype_dir,
            "output_dir": "./output",
            "screenshot_dir": "./assets/screenshots",
            "output_filename": f"{module_name}产品需求说明书.docx",
            "reference_doc": "可选：这里可以填公司模板路径"
        },
        "mapping": mappings
    }
    
    with open(output_map_file, 'w', encoding='utf-8') as f:
        yaml.dump(config_data, f, allow_unicode=True, sort_keys=False, default_flow_style=False)
        
    print(f"\n🎉 自动生成映射配置完成！已保存至: {output_map_file}")
    print(f"⚠️ 请打开文件检查自动匹配的结果，对于 '手动填写.html' 的项需要您人工补全。")

if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 4:
        prd_dir = sys.argv[1]
        prototype_dir = sys.argv[2]
        output_yaml = sys.argv[3]
        auto_generate_map(prd_dir, prototype_dir, output_yaml)
    else:
        print("用法: python auto_generate_map.py <prd_dir> <prototype_dir> <output_yaml>")
        print("请提供完整的目录路径。")
