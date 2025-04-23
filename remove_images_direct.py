#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
脚本功能: 通过直接修改DOCX内部结构来删除所有图片
原理: DOCX文件本质上是一个ZIP压缩包，包含多个XML文件和媒体文件夹
使用方法: python remove_images_direct.py 输入文件.docx [输出文件.docx]
"""

import os
import sys
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET

# 定义DOCX中XML命名空间
namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

def register_namespaces():
    """注册XML命名空间"""
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

def remove_images_from_docx(input_path, output_path=None):
    """
    通过直接操作DOCX内部结构来删除所有图片
    
    Args:
        input_path (str): 输入docx文件路径
        output_path (str, optional): 输出docx文件路径。如果未提供，将生成'无图片_' + 原文件名
    
    Returns:
        str: 输出文件的路径
    """
    if not os.path.exists(input_path):
        print(f"错误: 文件 '{input_path}' 不存在!")
        return None
    
    # 如果没有指定输出路径，则生成默认输出路径
    if output_path is None:
        file_dir = os.path.dirname(input_path)
        file_name = os.path.basename(input_path)
        output_path = os.path.join(file_dir, "无图片_" + file_name)
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    try:
        # 解压DOCX文件到临时目录
        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # 注册命名空间
        register_namespaces()
        
        # 处理document.xml文件(主文档内容)
        doc_path = os.path.join(temp_dir, 'word', 'document.xml')
        if os.path.exists(doc_path):
            tree = ET.parse(doc_path)
            root = tree.getroot()
            
            # 查找并删除所有drawing元素（包含图片）
            drawings_removed = 0
            for drawing in root.findall('.//w:drawing', namespaces):
                drawing.getparent().remove(drawing)
                drawings_removed += 1
            
            # 查找并删除所有pict元素（包含较旧格式的图片）
            picts_removed = 0
            for pict in root.findall('.//w:pict', namespaces):
                pict.getparent().remove(pict)
                picts_removed += 1
            
            # 保存修改后的XML
            tree.write(doc_path, encoding='UTF-8', xml_declaration=True)
            
            print(f"已从文档中移除 {drawings_removed} 个drawing元素和 {picts_removed} 个pict元素")
        
        # 删除media文件夹（包含所有图片文件）
        media_dir = os.path.join(temp_dir, 'word', 'media')
        if os.path.exists(media_dir):
            media_files = len(os.listdir(media_dir))
            shutil.rmtree(media_dir)
            print(f"已删除media文件夹中的 {media_files} 个文件")
        
        # 重新打包DOCX文件
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root_dir, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    zipf.write(
                        file_path, 
                        os.path.relpath(file_path, temp_dir)
                    )
        
        print(f"新文档已保存至: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"处理文档时出错: {str(e)}")
        return None
    
    finally:
        # 清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def main():
    """主函数，处理命令行参数并执行图片移除操作"""
    if len(sys.argv) < 2:
        print("用法: python remove_images_direct.py 输入文件.docx [输出文件.docx]")
        return
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    remove_images_from_docx(input_path, output_path)

if __name__ == "__main__":
    main() 