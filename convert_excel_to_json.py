
import pandas as pd
import json

def convert_chapters(chapter_text, section_text="", lesson_text=""):
    # 创建基本的章节结构
    result = {
        "text": "",
        "type": "chapter",
        "sub": {}
    }
    # 处理一级目录
    if isinstance(chapter_text, str) and chapter_text.strip():
        result["text"] = chapter_text.strip()
    
    # 处理二级目录
    if isinstance(section_text, str) and section_text.strip():
        result["sub"] = {
            "text": section_text.strip(),
            "type": "section",
            "sub": []
        }
        
        # 处理三级目录
        if isinstance(lesson_text, str) and lesson_text.strip():
            result["sub"]["sub"] = [
                {
                    "text": lesson_text.strip(),
                    "type": "lesson"
                }
            ]
    
    return result

def process_field_value(key, value):
    """根据字段类型处理值"""
    if pd.isna(value):
        if key == "choice" or key == "knowledge":
            return []
        elif key == "chapters":
            return {"text": "", "type": "chapter", "sub": {}}
        elif key == "difficulty":
            return 0.0
        else:
            return ""
            
    if key == "choice":
        # 按回车符分割选项
        choices = str(value).split('\n')
        # 去除空选项并清理每个选项
        return [choice.strip() for choice in choices if choice.strip()]
    elif key == "knowledge":
        # 按回车符分割知识点
        knowledge_points = str(value).split('\n')
        # 去除空知识点并清理每个知识点
        return [k.strip() for k in knowledge_points if k.strip()]
    elif key == "chapters":
        return convert_chapters(str(value))
    elif key == "difficulty":
        # 确保difficulty是浮点数
        try:
            return float(value)
        except:
            return 0.0
    else:
        return str(value).strip()

def convert_excel_to_json(excel_file_path, field_mapping=None):
    # 读取指定的工作表 最终结果
    try:
        df = pd.read_excel(excel_file_path, sheet_name="最终结果")
    except Exception as e:
        print(f"错误：未找到'最终结果'工作表")
        return None, None
    
    # 工作簿名称
    workbook_name = excel_file_path.split('/')[-1].split('.')[0]
    
    # 如果没有提供字段映射，则使用默认映射
    if not field_mapping:
        field_mapping = {}
    
    # 创建结果列表
    result = []
    
    # 反向映射 - 只处理有映射关系的字段
    valid_columns = set(field_mapping.keys())
    
    # 特殊字段 - 目录相关
    chapter_field = None
    section_field = None
    lesson_field = None
    
    
    # 查找章节相关字段的映射
    for col, field in field_mapping.items():
        if field == "chapters":
            chapter_field = col
        elif field == "section":
            section_field = col
        elif field == "lesson":
            lesson_field = col
    
    # 遍历每一行数据
    for index, row in df.iterrows():
        item = {}
        
        # 预先获取章节相关数据
        chapter_value = row.get(chapter_field, "") if chapter_field else ""
        section_value = row.get(section_field, "") if section_field else ""
        lesson_value = row.get(lesson_field, "") if lesson_field else ""
        
        # 处理其他字段
        for column in df.columns:
            if column in valid_columns:
                target_field = field_mapping.get(column)
                
                # 特殊处理章节字段
                if target_field == "chapters":
                    item[target_field] = convert_chapters(chapter_value, section_value, lesson_value)
                # 跳过已经在chapters中处理的section和lesson字段
                elif target_field in ["section", "lesson"]:
                    continue
                else:
                    item[target_field] = process_field_value(target_field, row[column])
        
        result.append(item)
    
    return result, workbook_name

def save_json(data, output_file):
    # 将数据保存为JSON文件，使用中文编码
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    # 转写路径
    excel_file = "./标注结果-9787574804616新编基础训练·化学 人教版 九年级 下册化学2024.pdf-NEW-20250325111357.xlsx"
    
    # 映射名
    field_mapping = {
        "【必填】question（题干）": "question",
        "【必填】choice（选项，选择题必填，其他非必填）": "choice",
        "【必填】answer（答案）": "answer",
        "analysis（解析）": "analysis",
        "comment（点评）": "comment",
        "tishi（展现形式）": "tishi",
        "tilei（题类）": "tilei",
        "difficulty（难度）": "difficulty",
        "subject（学科）": "subject",
        "step（学段）": "step",
        "grade（年级）": "grade",
        "【必填】bookIsbn（书本isbn）": "bookname",
        "【必填】bookIsbn（书本isbn）": "bookIsbn",
        "【必填】chapters（一级目录）": "chapters",
        "section（二级目录）": "section",
        "lesson（三级目录）": "lesson",
        "knowledge（知识点）": "knowledge"
    }
    
    # 转换并保存
    json_data, workbook_name = convert_excel_to_json(excel_file, field_mapping)
    if json_data:
        output_json = f"{workbook_name}.json"
        save_json(json_data, output_json)
        print(f"转换完成！数据已保存到 {output_json}") 