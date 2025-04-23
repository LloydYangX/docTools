#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
脚本功能: 使用lxml库高效地解析和修改DOCX文件，移除所有图片
优点: 使用lxml处理XML更加高效且完整支持命名空间
依赖库: lxml
使用方法: python remove_images_lxml.py 输入文件.docx [输出文件.docx]
"""

import os
import sys
import zipfile
import shutil
import tempfile
from lxml import etree

def remove_images_from_docx(input_path, output_path=None):
    """
    使用lxml解析DOCX并删除所有图片
    
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
        
        # 定义命名空间
        nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
        
        # 处理document.xml文件(主文档内容)
        doc_path = os.path.join(temp_dir, 'word', 'document.xml')
        if os.path.exists(doc_path):
            # 解析XML
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(doc_path, parser)
            root = tree.getroot()
            
            # 查找并删除所有图片元素
            img_count = 0
            
            # 删除内联图片 - drawing元素
            drawings = root.xpath('//w:drawing', namespaces=nsmap)
            for drawing in drawings:
                drawing.getparent().remove(drawing)
                img_count += 1
            
            # 删除旧式图片 - pict元素
            picts = root.xpath('//w:pict', namespaces=nsmap)
            for pict in picts:
                pict.getparent().remove(pict)
                img_count += 1
            
            # 删除其他可能包含图片的元素
            shapes = root.xpath('//w:object', namespaces=nsmap)
            for shape in shapes:
                shape.getparent().remove(shape)
                img_count += 1
            
            # 保存修改后的XML
            tree.write(doc_path, encoding='UTF-8', xml_declaration=True, pretty_print=True)
            
            print(f"已从文档中移除 {img_count} 个图片元素")
        
        # 删除media文件夹（包含所有图片文件）
        media_dir = os.path.join(temp_dir, 'word', 'media')
        if os.path.exists(media_dir):
            media_files = len(os.listdir(media_dir))
            shutil.rmtree(media_dir)
            print(f"已删除media文件夹中的 {media_files} 个文件")
        
        # 清理文档关系文件中的图片引用
        rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        if os.path.exists(rels_path):
            try:
                rels_tree = etree.parse(rels_path, parser)
                rels_root = rels_tree.getroot()
                
                # 删除指向图片的关系
                img_rels = rels_root.xpath('//r:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]', 
                                        namespaces=nsmap)
                for rel in img_rels:
                    rel.getparent().remove(rel)
                
                # 保存修改后的关系文件
                rels_tree.write(rels_path, encoding='UTF-8', xml_declaration=True, pretty_print=True)
                print(f"已从关系文件中删除 {len(img_rels)} 个图片引用")
            except Exception as e:
                print(f"清理关系文件时出错: {str(e)}")
        
        # 重新打包DOCX文件
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)
        
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
        print("用法: python remove_images_lxml.py 输入文件.docx [输出文件.docx]")
        return
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    remove_images_from_docx(input_path, output_path)

if __name__ == "__main__":
    main() 