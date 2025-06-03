import re
import pandas as pd
from docx import Document
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

def parse_audit_doc(doc_path):
    """解析审计报告Word文档的核心函数"""
    doc = Document(doc_path)
    data = []
    
    # 状态控制变量
    start_processing = False
    in_table = False
    current_section = None

    # 数据存储容器
    current_record = {
        'level1': '', 'level2': '', 'level3': '',
        'pre_risk': [], 'risk': [], 'suggestion': [],
        'response': {'confirm': [], 'plan': [], 'responsible': [], 'time': [], 'other': []},
        'current_response': None
    }

    def save_record():
        """保存当前记录并重置容器"""
        if current_record['level3']:
            data.append([
                current_record['level1'],
                current_record['level2'],
                current_record['level3'],
                '\n'.join(current_record['pre_risk']).strip(),
                '\n'.join(current_record['risk']).strip(),
                '\n'.join(current_record['suggestion']).strip(),
                '\n'.join(current_record['response']['confirm']).strip(),
                '\n'.join(current_record['response']['plan']).strip(),
                '\n'.join(current_record['response']['responsible']).strip(),
                '\n'.join(current_record['response']['time']).strip(),
                '\n'.join(current_record['response']['other']).strip()
            ])
        # 重置记录（保留一二级标题）
        current_record.update({
            'level3': '',
            'pre_risk': [],
            'risk': [],
            'suggestion': [],
            'response': {'confirm': [], 'plan': [], 'responsible': [], 'time': [], 'other': []},
            'current_response': None
        })

    for para in doc.paragraphs:
        raw_text = para.text.strip()
        
        # 1. 定位审计正文开始
        if not start_processing:
            if re.match(r'^三[、.]\s*审计正文', raw_text):
                start_processing = True
            continue
            
        # 2. 处理表格内容
        if in_table:
            if not raw_text:
                in_table = False
            continue
        if raw_text.startswith('表'):
            in_table = True
            continue
            
        # 3. 标题识别
        # 一级标题（示例："（一）项目管理"）
        if re.match(r'^（[一二三四五六七八九十]+）', raw_text):
            if current_record['level3']:
                save_record()
            current_record['level1'] = re.sub(
                r'^（[一二三四五六七八九十]+）', 
                '', 
                raw_text
            ).strip()
        
        # 二级标题（示例："（X风险）1.1.需求至立项管理"）
        elif re.match(r'^（.*?）\s*\d+\.\d+', raw_text):
            match = re.search(r'（.*?）\s*(\d+\.\d+)(\.?)\s*([^_]+)', raw_text)
            if match:
                if current_record['level3']:
                    save_record()
                current_record['level2'] = match.group(3).strip(' ._')
        
        # 三级标题（示例："1.1.1 OA系统..." 或 "1.1.1(示例)标题"）
        elif re.match(r'^\d+\.\d+\.\d+', raw_text):
            if current_record['level3']:
                save_record()
            # 使用正则表达式精确提取标题内容
            match = re.match(r'^(\d+\.\d+\.\d+)\s*(.*)', raw_text)
            if match:
                current_record['level3'] = match.group(2).strip()  # 提取数字编号后的全部内容
                current_section = 'pre_risk'  # 开始收集风险前内容
        
        # 4. 内容解析
        else:
            if not raw_text or in_table:
                continue

            # 处理回复部分的内容
            if current_section == 'response':
                # 匹配回复子标题
                if re.match(r'^1\.\s*确认意见', raw_text):
                    current_record['current_response'] = 'confirm'
                elif re.match(r'^2\.\s*改进计划', raw_text):
                    current_record['current_response'] = 'plan'
                elif re.match(r'^3\.\s*整改部门及负责人', raw_text):
                    current_record['current_response'] = 'responsible'
                elif re.match(r'^4\.\s*整改完成时间', raw_text):
                    current_record['current_response'] = 'time'
                else:
                    # 将内容添加到当前回复的子部分
                    target = current_record['response'][current_record['current_response']] \
                             if current_record['current_response'] \
                             else current_record['response']['other']
                    target.append(raw_text)
            else:
                # 段落类型判断（非回复部分）
                if '相关风险' in raw_text:
                    current_section = 'risk'
                    continue
                elif '改进建议' in raw_text:
                    current_section = 'suggestion'
                    continue
                elif '公司管理层回复' in raw_text:
                    current_section = 'response'
                    continue

                # 收集内容到当前section
                if current_section == 'pre_risk':
                    current_record['pre_risk'].append(raw_text)
                elif current_section == 'risk':
                    current_record['risk'].append(raw_text)
                elif current_section == 'suggestion':
                    current_record['suggestion'].append(raw_text)

    save_record()
    return data

def process_files(word_files):
    """批量处理文件的主逻辑"""
    success_count = 0
    error_count = 0
    error_list = []

    for idx, word_file in enumerate(word_files, 1):
        try:
            print(f"\n正在处理文件 ({idx}/{len(word_files)}): {word_file}")
            data = parse_audit_doc(word_file)
            
            if not data:
                raise ValueError("未提取到有效数据，请检查文档格式")
                
            df = pd.DataFrame(data, columns=[
                "一级标题", "二级标题", "三级标题",
                "相关风险前内容", "相关风险", "改进建议",
                "确认意见", "改进计划", "整改部门及负责人",
                "完成时间", "备注"
            ])
            
            excel_file = word_file.replace(".docx", "_审计报告.xlsx")
            df.to_excel(excel_file, index=False, engine='openpyxl')
            print(f"✓ 成功生成：{excel_file}")
            success_count += 1
            
        except Exception as e:
            print(f"✕ 处理失败：{str(e)}")
            error_count += 1
            error_list.append((word_file, str(e)))

    # 输出汇总报告
    print("\n" + "="*50)
    print(f"处理完成！共处理 {len(word_files)} 个文件")
    print(f"成功转换：{success_count} 个")
    print(f"转换失败：{error_count} 个")
    
    if error_list:
        print("\n失败文件列表：")
        for file, error in error_list:
            print(f"- {file}\n  错误原因：{error}")

def main():
    """程序入口"""
    Tk().withdraw()
    
    print("请选择要转换的Word文档（可多选）")
    word_files = askopenfilenames(
        title="选择审计报告文档",
        filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")],
        multiple=True
    )
    
    if word_files:
        process_files(word_files)
    else:
        print("未选择任何文件")

if __name__ == "__main__":
    main()