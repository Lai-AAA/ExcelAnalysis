#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据解析与考勤验证脚本
功能：批量解析多个Excel数据源，自动完成考勤验证和黑名单生成
"""

import pandas as pd
import json
import os
import sys
from pathlib import Path
import argparse
from datetime import datetime

# 必需的15个核心字段
REQUIRED_FIELDS = [
    '序号', '学年学期', '是否为工学院举办', '行政班级', '学号',
    '姓名', '活动类型', '活动名称', '加分类型', '加分分数',
    '奖项', '部门', '负责人', '联系电话', '备注'
]

def parse_excel_file(file_path):
    """解析单个Excel文件"""
    try:
        # 读取"数据源"工作表
        df = pd.read_excel(file_path, sheet_name='数据源')
        
        # 验证必需字段
        missing_fields = [field for field in REQUIRED_FIELDS if field not in df.columns]
        if missing_fields:
            raise ValueError(f"缺少以下列：{', '.join(missing_fields)}")
        
        return df
    except Exception as e:
        print(f"解析文件 {file_path} 失败: {str(e)}")
        return None

def process_data(df):
    """处理数据：去重、格式转换、有效性校验"""
    # 去除首尾空格
    df['姓名'] = df['姓名'].astype(str).str.strip()
    df['活动名称'] = df['活动名称'].astype(str).str.strip()
    df['行政班级'] = df['行政班级'].astype(str).str.strip()
    
    # 有效性校验：过滤姓名为空或活动名称为空的数据
    valid_mask = (df['姓名'] != '') & (df['活动名称'] != '') & (df['姓名'] != 'nan') & (df['活动名称'] != 'nan')
    invalid_df = df[~valid_mask].copy()
    df = df[valid_mask].copy()
    
    # 去重：按"班级+姓名+活动名称"
    df = df.drop_duplicates(subset=['行政班级', '姓名', '活动名称'], keep='first')
    
    # 格式转换
    df['加分分数'] = pd.to_numeric(df['加分分数'], errors='coerce').fillna(0)
    df['学号'] = df['学号'].astype(str).str.strip()
    
    return df, invalid_df

def validate_attendance(df, config):
    """验证考勤状态"""
    keywords = config.get('keywords', ['已考勤', '到场', '参与'])
    
    # 检查备注列是否包含考勤关键词
    remark_col = df['备注'].astype(str)
    is_attended = remark_col.str.contains('|'.join(keywords), case=False, na=False)
    
    attended_df = df[is_attended].copy()
    unattended_df = df[~is_attended].copy()
    unattended_df['未考勤原因'] = '备注中未找到考勤关键词'
    
    return attended_df, unattended_df

def generate_report(attended_df, unattended_df, invalid_df, output_dir):
    """生成统计报告"""
    report_lines = []
    report_lines.append("=" * 60)
    report_lines.append("数据解析与考勤验证报告")
    report_lines.append("=" * 60)
    report_lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("")
    
    # 总体统计
    total = len(attended_df) + len(unattended_df) + len(invalid_df)
    report_lines.append(f"总数据量: {total}")
    report_lines.append(f"有效数据: {len(attended_df) + len(unattended_df)}")
    report_lines.append(f"无效数据: {len(invalid_df)}")
    report_lines.append(f"已考勤（可加分）: {len(attended_df)}")
    report_lines.append(f"未考勤（黑名单）: {len(unattended_df)}")
    report_lines.append("")
    
    # 班级统计
    if len(attended_df) > 0:
        class_stats = attended_df.groupby('行政班级').size().sort_values(ascending=False)
        report_lines.append("各班级已考勤人数统计（前10名）:")
        for class_name, count in class_stats.head(10).items():
            report_lines.append(f"  {class_name}: {count}人")
        report_lines.append("")
    
    # 活动类型统计
    if len(attended_df) > 0:
        activity_stats = attended_df.groupby('活动类型').size()
        report_lines.append("活动类型分布:")
        for activity_type, count in activity_stats.items():
            report_lines.append(f"  {activity_type}: {count}人")
        report_lines.append("")
    
    # 考勤率统计
    if len(attended_df) + len(unattended_df) > 0:
        attendance_rate = len(attended_df) / (len(attended_df) + len(unattended_df)) * 100
        report_lines.append(f"考勤率: {attendance_rate:.2f}%")
    
    report_text = "\n".join(report_lines)
    
    # 保存报告
    report_path = os.path.join(output_dir, 'report.txt')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report_text)
    
    print(report_text)
    return report_path

def main():
    parser = argparse.ArgumentParser(description='数据解析与考勤验证脚本')
    parser.add_argument('--input', type=str, required=True, help='输入文件路径或文件夹路径')
    parser.add_argument('--config', type=str, default='config.json', help='配置文件路径')
    parser.add_argument('--output', type=str, default='./output', help='输出目录')
    
    args = parser.parse_args()
    
    # 创建输出目录
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 加载配置
    config = {}
    if os.path.exists(args.config):
        with open(args.config, 'r', encoding='utf-8') as f:
            config = json.load(f)
    else:
        # 使用默认配置
        config = {
            'keywords': ['已考勤', '到场', '参与']
        }
        print(f"配置文件不存在，使用默认配置")
    
    # 收集所有数据
    all_dataframes = []
    input_path = Path(args.input)
    
    if input_path.is_file():
        # 单个文件
        df = parse_excel_file(input_path)
        if df is not None:
            all_dataframes.append(df)
    elif input_path.is_dir():
        # 文件夹批量处理
        excel_files = list(input_path.glob('*.xlsx'))
        for excel_file in excel_files:
            df = parse_excel_file(excel_file)
            if df is not None:
                all_dataframes.append(df)
    else:
        print(f"错误: 输入路径不存在: {args.input}")
        sys.exit(1)
    
    if not all_dataframes:
        print("错误: 没有成功解析任何文件")
        sys.exit(1)
    
    # 合并所有数据
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    print(f"成功解析 {len(combined_df)} 行数据")
    
    # 处理数据
    valid_df, invalid_df = process_data(combined_df)
    print(f"有效数据: {len(valid_df)} 行，无效数据: {len(invalid_df)} 行")
    
    # 考勤验证
    attended_df, unattended_df = validate_attendance(valid_df, config)
    print(f"已考勤: {len(attended_df)} 人，未考勤: {len(unattended_df)} 人")
    
    # 保存结果
    parsed_data_path = os.path.join(output_dir, 'parsed_data.xlsx')
    blacklist_path = os.path.join(output_dir, 'blacklist.xlsx')
    
    attended_df.to_excel(parsed_data_path, index=False, sheet_name='已考勤数据')
    print(f"已保存解析数据到: {parsed_data_path}")
    
    if len(unattended_df) > 0:
        blacklist_df = unattended_df[['行政班级', '姓名', '学号', '未考勤原因']].copy()
        blacklist_df.to_excel(blacklist_path, index=False, sheet_name='黑名单')
        print(f"已保存黑名单到: {blacklist_path}")
    
    # 生成报告
    report_path = generate_report(attended_df, unattended_df, invalid_df, output_dir)
    print(f"已生成报告到: {report_path}")
    
    print("\n处理完成！")

if __name__ == '__main__':
    main()

