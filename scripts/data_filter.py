#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多维度筛选脚本
功能：按指定维度筛选数据，支持复杂条件组合
"""

import pandas as pd
import json
import os
import sys
import argparse
import re

def load_data(input_file):
    """加载解析后的数据"""
    try:
        df = pd.read_excel(input_file, sheet_name='已考勤数据')
        return df
    except Exception as e:
        print(f"加载数据失败: {str(e)}")
        sys.exit(1)

def filter_data(df, config):
    """按配置筛选数据"""
    filtered_df = df.copy()
    
    # 学年学期筛选
    if 'semester' in config and config['semester']:
        filtered_df = filtered_df[filtered_df['学年学期'] == config['semester']]
    
    # 活动类型筛选
    if 'activityTypes' in config and config['activityTypes']:
        filtered_df = filtered_df[filtered_df['活动类型'].isin(config['activityTypes'])]
    
    # 活动名称筛选（模糊搜索）
    if 'activityName' in config and config['activityName']:
        keyword = config['activityName'].lower()
        filtered_df = filtered_df[
            filtered_df['活动名称'].astype(str).str.lower().str.contains(keyword, na=False)
        ]
    
    # 加分类型筛选
    if 'scoreType' in config and config['scoreType']:
        filtered_df = filtered_df[filtered_df['加分类型'] == config['scoreType']]
    
    # 加分分数范围筛选
    if 'minScore' in config and config['minScore'] is not None:
        filtered_df = filtered_df[filtered_df['加分分数'] >= config['minScore']]
    if 'maxScore' in config and config['maxScore'] is not None:
        filtered_df = filtered_df[filtered_df['加分分数'] <= config['maxScore']]
    
    # 班级筛选（支持正则匹配）
    if 'classes' in config and config['classes']:
        if config.get('useRegex', False):
            # 正则匹配
            pattern = '|'.join(config['classes'])
            filtered_df = filtered_df[
                filtered_df['行政班级'].astype(str).str.contains(pattern, na=False, regex=True)
            ]
        else:
            # 精确匹配
            filtered_df = filtered_df[filtered_df['行政班级'].isin(config['classes'])]
    
    # 是否包含其他学院
    if 'includeOtherCollege' in config and not config['includeOtherCollege']:
        filtered_df = filtered_df[filtered_df['是否为工学院举办'] == '是']
    
    # 排序
    if config.get('sortByClass', True):
        filtered_df = filtered_df.sort_values(['行政班级', '姓名'], ascending=[True, True])
    
    return filtered_df

def generate_log(filtered_df, original_count, config, output_dir):
    """生成筛选日志"""
    log_lines = []
    log_lines.append("=" * 60)
    log_lines.append("数据筛选日志")
    log_lines.append("=" * 60)
    log_lines.append(f"原始数据量: {original_count}")
    log_lines.append(f"筛选后数据量: {len(filtered_df)}")
    log_lines.append(f"筛选条件:")
    
    for key, value in config.items():
        if value and key not in ['useRegex', 'sortByClass']:
            log_lines.append(f"  {key}: {value}")
    
    log_lines.append("")
    log_lines.append("筛选结果统计:")
    
    if len(filtered_df) > 0:
        # 按班级统计
        class_stats = filtered_df.groupby('行政班级').size().sort_values(ascending=False)
        log_lines.append("各班级人数（前10名）:")
        for class_name, count in class_stats.head(10).items():
            log_lines.append(f"  {class_name}: {count}人")
        
        log_lines.append("")
        # 按活动类型统计
        activity_stats = filtered_df.groupby('活动类型').size()
        log_lines.append("活动类型分布:")
        for activity_type, count in activity_stats.items():
            log_lines.append(f"  {activity_type}: {count}人")
    
    log_text = "\n".join(log_lines)
    
    # 保存日志
    log_path = os.path.join(output_dir, 'filter_log.txt')
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write(log_text)
    
    print(log_text)
    return log_path

def main():
    parser = argparse.ArgumentParser(description='多维度数据筛选脚本')
    parser.add_argument('--input', type=str, required=True, help='输入文件路径（parsed_data.xlsx）')
    parser.add_argument('--config', type=str, required=True, help='筛选条件配置文件路径（JSON格式）')
    parser.add_argument('--output', type=str, default='./output', help='输出目录')
    
    args = parser.parse_args()
    
    # 创建输出目录
    os.makedirs(args.output, exist_ok=True)
    
    # 加载数据
    df = load_data(args.input)
    original_count = len(df)
    print(f"加载数据: {original_count} 行")
    
    # 加载筛选配置
    if not os.path.exists(args.config):
        print(f"错误: 配置文件不存在: {args.config}")
        sys.exit(1)
    
    with open(args.config, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    # 执行筛选
    filtered_df = filter_data(df, config)
    print(f"筛选完成: {len(filtered_df)} 行")
    
    # 保存结果
    output_path = os.path.join(args.output, 'filtered_data.xlsx')
    filtered_df.to_excel(output_path, index=False, sheet_name='筛选结果')
    print(f"已保存筛选结果到: {output_path}")
    
    # 生成日志
    log_path = generate_log(filtered_df, original_count, config, args.output)
    print(f"已生成筛选日志到: {log_path}")
    
    print("\n筛选完成！")

if __name__ == '__main__':
    main()

