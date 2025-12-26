#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档合并助手 - 将多个MD文件合并为一个手册
"""

import os
import re


def read_file(filepath):
    """读取文件内容"""
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read()


def adjust_heading_levels(content, add_levels=1):
    """
    调整标题级别
    例如: # 标题 -> ## 标题
    """
    lines = content.split('\n')
    adjusted_lines = []

    for line in lines:
        # 匹配标题行
        match = re.match(r'^(#+)\s+(.+)$', line)
        if match:
            hashes = match.group(1)
            title_text = match.group(2)
            # 增加标题级别
            new_hashes = '#' * (len(hashes) + add_levels)
            adjusted_lines.append(f"{new_hashes} {title_text}")
        else:
            adjusted_lines.append(line)

    return '\n'.join(adjusted_lines)


def create_运营手册():
    """创建门店运营手册"""

    docs_dir = 'docs'
    output_file = 'newmd/03_门店运营手册.md'

    # 文件映射: (文件编号, 新章节编号, 章节名称)
    files = [
        ('10', '10', '西班牙受欢迎的冰淇淋酸奶冰连锁对标拆解表'),
        ('11', '11', '门店模型与动线翻台'),
        ('12', '12', '项目总设计文件'),
        ('13', '13', '预拌粉规格与代工厂交付包'),
        ('14', '14', '门店SOP'),
    ]

    # 创建手册头部
    header = """# 03_门店运营手册

> **用途**:本手册整合了软冰门店运营的全部实施指南,涵盖市场调研、门店设计、项目总体规划、预拌粉交付和门店SOP。适用于门店经理、运营总监和执行人员。

> **使用指南**:第10章提供市场对标参考,第11章提供门店设计方案,第12章是总体规划,第13-14章提供具体执行标准。建议先理解市场定位(第10章),再规划门店模型(第11章),最后执行标准化流程(第12-14章)。

---

"""

    content_parts = [header]

    # 遍历文件并合并
    for orig_num, new_num, chapter_name in files:
        # 找到匹配的文件
        filename_pattern = f"{orig_num}_*.md"
        matching_files = []
        for f in os.listdir(docs_dir):
            if f.startswith(f"{orig_num}_") and f.endswith('.md'):
                matching_files.append(f)

        if not matching_files:
            print(f"警告: 未找到编号 {orig_num} 的文件")
            continue

        filepath = os.path.join(docs_dir, matching_files[0])
        print(f"正在处理: {filepath}")

        # 读取文件
        file_content = read_file(filepath)

        # 移除第一行的标题(# 开头的)
        lines = file_content.split('\n')
        content_without_title = '\n'.join(lines[1:])

        # 添加章节头部
        chapter_header = f"## 第{new_num}章 {chapter_name}\n\n"

        # 调整所有标题级别(原来的# -> ###, ## -> ####, 等等)
        adjusted_content = adjust_heading_levels(content_without_title, add_levels=2)

        content_parts.append(chapter_header)
        content_parts.append(adjusted_content)
        content_parts.append("\n\n---\n\n")

    # 添加附录
    footer = """## 附录：手册内部交叉引用说明

本门店运营手册整合自原文档 10-14 章节。文中涉及的其他手册引用对应关系如下：

- **手册第1卷**：技术培训手册
  - 第1章：软冰核心原理与参数窗口
  - 第2章：原料体系与糖体系手册
  - 第3章：配方框架与工艺模板库
  - 第4章：抹茶专项
  - 第5章：巧克力专项
  - 第6章：调参旋钮与排障树总表
  - 第7章：研发记录与训练计划

- **手册第2卷**：设备采购手册
  - 第8章：软冰机选型规格书
  - 第9章：Makro扫货与最小采购清单

- **附录**：
  - 第0章：总索引与交付清单
  - 第99章：口径差异与对立观点汇总
"""

    content_parts.append(footer)

    # 合并所有内容
    final_content = ''.join(content_parts)

    # 替换文档内部的交叉引用
    # 例如: "见 12" -> "见手册第3卷第12章" (不替换,保持原样)
    # 例如: "见 06" -> "见手册第1卷第6章"

    final_content = re.sub(r'见\s+06([^\d])', r'见手册第1卷第6章\1', final_content)
    final_content = re.sub(r'见\s+07([^\d])', r'见手册第1卷第7章\1', final_content)
    final_content = re.sub(r'见\s+13/14([^\d])', r'见本卷第13/14章\1', final_content)
    final_content = re.sub(r'文件\s+12', '本卷第12章', final_content)
    final_content = re.sub(r'文件\s+13', '本卷第13章', final_content)
    final_content = re.sub(r'文件\s+14', '本卷第14章', final_content)

    # 写入文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_content)

    print(f"✓ 门店运营手册已创建: {output_file}")


if __name__ == '__main__':
    create_运营手册()
