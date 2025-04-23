#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
脚本功能: 增强版docx转换为markdown，支持更多格式和样式
依赖库: python-docx, mammoth, html2markdown
特点: 使用mammoth库中间转换为HTML，能更准确处理样式和格式
使用方法: python docx_to_markdown_enhanced.py 输入文件.docx [输出文件.md]
"""

import os
import sys
import tempfile
import shutil
import re
import json
from html import escape as html_escape
try:
    import mammoth
    import html2markdown
    from docx import Document
except ImportError:
    print("请安装所需依赖：pip install python-docx mammoth html2markdown")
    sys.exit(1)

def clean_markdown_text(text):
    """
    清理和修复Markdown文本中的常见问题
    
    Args:
        text (str): 原始Markdown文本
    
    Returns:
        str: 修复后的Markdown文本
    """
    # 修复连续多个空行变成一个空行
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    # 修复表格格式问题
    lines = text.split('\n')
    for i in range(len(lines) - 2):
        # 检查是否为表格头部和分隔行
        if (lines[i].startswith('|') and 
            lines[i].endswith('|') and 
            '---' in lines[i+1]):
            # 确保分隔行格式正确
            header_cells = lines[i].count('|') - 1
            separator = '|' + '|'.join([' --- ' for _ in range(header_cells)]) + '|'
            lines[i+1] = separator
    
    # 修复标题前后的空行
    fixed_lines = []
    for i, line in enumerate(lines):
        if line.startswith('#') and i > 0 and fixed_lines[-1] != '':
            fixed_lines.append('')  # 标题前添加空行
            fixed_lines.append(line)
            if i < len(lines) - 1 and lines[i+1] != '':
                fixed_lines.append('')  # 标题后添加空行
        else:
            fixed_lines.append(line)
    
    # 修复列表项目间的空行问题
    text = '\n'.join(fixed_lines)
    
    # 修复图片链接
    text = re.sub(r'!\[(.*?)\]\((?!http)(.*?)\)', r'![\1](images/\2)', text)
    
    # 移除HTML转义
    text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
    
    return text

def convert_images_to_markdown_links(html_content, image_dir, root_path):
    """
    将HTML中的img标签转换为Markdown图片链接，并保存图片
    
    Args:
        html_content (str): HTML内容
        image_dir (str): 图片保存目录
        root_path (str): 文档根目录
    
    Returns:
        str: 处理后的HTML内容
    """
    if not os.path.exists(image_dir):
        os.makedirs(image_dir)
    
    # 处理base64编码的图片
    img_pattern = re.compile(r'<img[^>]*src="data:image/([^;]+);base64,([^"]+)"[^>]*>')
    img_count = 0
    
    def replace_image(match):
        nonlocal img_count
        img_count += 1
        img_format = match.group(1)
        img_data = match.group(2)
        
        # 保存图片
        filename = f"image_{img_count}.{img_format}"
        img_path = os.path.join(image_dir, filename)
        rel_path = os.path.join('images', filename)
        
        import base64
        with open(img_path, 'wb') as f:
            f.write(base64.b64decode(img_data))
        
        # 替换为Markdown图片语法
        return f'![图片]({rel_path})'
    
    html_with_images = img_pattern.sub(replace_image, html_content)
    return html_with_images

def create_custom_style_map():
    """创建自定义的样式映射"""
    return {
        "p[style-name='Heading 1']": "h1",
        "p[style-name='Heading 2']": "h2",
        "p[style-name='Heading 3']": "h3",
        "p[style-name='Heading 4']": "h4",
        "p[style-name='Heading 5']": "h5",
        "p[style-name='Title']": "h1.title",
        "p[style-name='Subtitle']": "h2.subtitle",
        "p[style-name='Quote']": "blockquote",
        "p[style-name='Intense Quote']": "blockquote.intense",
        "r[style-name='Strong']": "strong",
        "r[style-name='Emphasis']": "em",
        "p[style-name='List Paragraph']": "p.list",
        "table": "table.docx-table",
        "p[style-name='Caption']": "p.caption",
        "p[style-name='TOC Heading']": "h1.toc-heading",
        "p[style-name='TOC 1']": "p.toc-1",
        "p[style-name='TOC 2']": "p.toc-2",
        "p[style-name='TOC 3']": "p.toc-3",
        "r[style-name='Hyperlink']": "a",
        "p[style-name='Code']": "pre.code",
        "r[style-name='Code Char']": "code"
    }

def docx_to_markdown(input_path, output_path=None):
    """
    将docx文档转换为markdown格式，支持更多格式和样式
    
    Args:
        input_path (str): 输入docx文件路径
        output_path (str, optional): 输出markdown文件路径。如果未提供，将生成同名.md文件
    
    Returns:
        str: 输出文件的路径
    """
    if not os.path.exists(input_path):
        print(f"错误: 文件 '{input_path}' 不存在!")
        return None
        
    # 如果没有指定输出路径，则生成默认输出路径
    if output_path is None:
        file_dir = os.path.dirname(input_path) or '.'
        file_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(file_dir, file_name + ".md")
    
    # 创建图片目录
    img_dir = os.path.join(os.path.dirname(output_path) or '.', 'images')
    
    try:
        # 创建自定义样式映射
        style_map = create_custom_style_map()
        style_map_str = "\n".join([f"{key} => {value}" for key, value in style_map.items()])
        
        # 使用mammoth进行转换
        with open(input_path, "rb") as docx_file:
            # 转换为HTML
            result = mammoth.convert_to_html(
                docx_file,
                style_map=style_map_str,
                include_embedded_style_map=True,
                ignore_empty_paragraphs=True,
                convert_image=mammoth.images.img_element(lambda image, _: {
                    "src": f"data:image/{image.content_type.split('/')[-1]};base64,{image.base64_bytes()}"
                })
            )
            
            html = result.value
            
            # 处理转换过程中的警告
            if result.messages:
                print("\n转换过程中的警告：")
                for message in result.messages:
                    print(f"- {message}")
        
        # 处理HTML中的图片
        html = convert_images_to_markdown_links(html, img_dir, os.path.dirname(output_path) or '.')
        
        # 将处理后的HTML转换为Markdown
        markdown_text = html2markdown.convert(html)
        
        # 清理和修复Markdown文本
        markdown_text = clean_markdown_text(markdown_text)
        
        # 写入文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_text)
        
        print(f"已成功将docx文档转换为markdown格式")
        print(f"输出文件: {output_path}")
        
        # 检查是否转换了图片
        if os.path.exists(img_dir) and os.listdir(img_dir):
            print(f"已提取图片到 {img_dir} 目录")
        
        return output_path
        
    except Exception as e:
        print(f"转换过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def main():
    """主函数，处理命令行参数并执行转换操作"""
    if len(sys.argv) < 2:
        print("用法: python docx_to_markdown_enhanced.py 输入文件.docx [输出文件.md]")
        return
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    docx_to_markdown(input_path, output_path)

if __name__ == "__main__":
    main() 