'''
Author: 曲洪利 quhongli999@163.com
Date: 2025-03-21 10:55:48
LastEditors: 曲洪利 quhongli999@163.com
LastEditTime: 2025-03-26 15:02:38
FilePath: /py/convert_excel_to_json_catalogue.py
Description: 

Copyright (c) 2025 by ${git_name_email}, All Rights Reserved. 
'''
import pandas as pd
import json
import os

def convert_chapters(chapter_text, section_text="", lesson_text=""):
    """将章节文本转换为多级层级结构"""
    result = {
        "text": "",
        "type": "chapter",
        "sub": []
    }
    
    # 处理一级目录
    if isinstance(chapter_text, str) and chapter_text.strip():
        result["text"] = chapter_text.strip()
    
    # 处理二级目录
    if isinstance(section_text, str) and section_text.strip():
        section_obj = {
            "text": section_text.strip(),
            "type": "section",
            "sub": []
        }
        
        # 处理三级目录
        if isinstance(lesson_text, str) and lesson_text.strip():
            # 按回车符分割多个课时
            lessons = lesson_text.split('\n')
            for lesson in lessons:
                if lesson.strip():
                    section_obj["sub"].append({
                        "text": lesson.strip(),
                        "type": "lesson",
                        "sub": []
                    })
        
        result["sub"].append(section_obj)
    
    return result

def process_field_value(key, value):
    """根据字段类型处理值"""
    if pd.isna(value):
        if key == "choice" or key == "knowledge":
            return []
        elif key == "chapters":
            return {"text": "", "type": "chapter", "sub": []}
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

def merge_chapters(chapters_list):
    """合并相同的章节内容，每一层级都去重"""
    # 使用字典存储已存在的章节，以text为键
    chapters_dict = {}
    
    for chapter in chapters_list:
        chapter_text = chapter["text"]
        if chapter_text not in chapters_dict:
            chapters_dict[chapter_text] = chapter
        else:
            # 合并二级目录
            for section in chapter["sub"]:
                section_text = section["text"]
                if section_text not in [s["text"] for s in chapters_dict[chapter_text]["sub"]]:
                    chapters_dict[chapter_text]["sub"].append(section)
                else:
                    # 合并三级目录
                    existing_section = next(s for s in chapters_dict[chapter_text]["sub"] if s["text"] == section_text)
                    for lesson in section["sub"]:
                        if lesson["text"] not in [l["text"] for l in existing_section["sub"]]:
                            existing_section["sub"].append(lesson)
    
    return list(chapters_dict.values())

def convert_excel_to_json(excel_file_path, field_mapping=None):
    # 读取指定的工作表 最终结果
    try:
        df = pd.read_excel(excel_file_path, sheet_name="最终结果")
    except Exception as e:
        print(f"错误：未找到'最终结果'工作表")
        return None, None
    
    # 工作簿名称 - 保留完整文件名（包含时间戳）
    workbook_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    
    # 如果没有提供字段映射，则使用默认映射
    if not field_mapping:
        field_mapping = {}
    
    # 创建结果列表
    chapters_list = []
    
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
        # 预先获取章节相关数据
        chapter_value = row.get(chapter_field, "") if chapter_field else ""
        section_value = row.get(section_field, "") if section_field else ""
        lesson_value = row.get(lesson_field, "") if lesson_field else ""
        
        # 创建章节对象
        chapter_obj = convert_chapters(chapter_value, section_value, lesson_value)
        if chapter_obj["text"]:  # 只添加非空的章节
            chapters_list.append(chapter_obj)
    
    # 合并相同的章节
    merged_chapters = merge_chapters(chapters_list)
    
    return merged_chapters, workbook_name

def save_json(data, output_file):
    # 将数据保存为JSON文件，使用中文编码
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    # 转写路径
    excel_file = "./标注结果-9787574804616新编基础训练·化学 人教版 九年级 下册化学2024.pdf-NEW-20250325111357.xlsx"
    
    # 映射名
    field_mapping = {
        "【必填】chapters（一级目录）": "chapters",
        "section（二级目录）": "section",
        "lesson（三级目录）": "lesson",
    }
    
    # 转换并保存
    json_data, workbook_name = convert_excel_to_json(excel_file, field_mapping)
    if json_data:
        output_json = f"{workbook_name}.json"
        save_json(json_data, output_json)
        print(f"转换完成！数据已保存到 {output_json}") 