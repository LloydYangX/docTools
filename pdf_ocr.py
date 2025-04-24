#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PDF文档图片OCR识别示例
从PDF文档中提取图片，然后使用PaddleOCR进行文字识别
"""

import os
import io
import fitz  # PyMuPDF
import argparse
import tempfile
from PIL import Image
from pathlib import Path
from paddleocr_recognition import OCRProcessor

def extract_images_from_pdf(pdf_path, output_dir=None, min_width=100, min_height=100):
    """
    从PDF文档中提取图片
    
    参数:
        pdf_path: PDF文件路径
        output_dir: 输出目录，默认为None(不保存图片文件)
        min_width: 最小图片宽度，小于此宽度的图片会被忽略
        min_height: 最小图片高度，小于此高度的图片会被忽略
        
    返回:
        images: 图片对象列表（PIL.Image）
        image_paths: 保存的图片文件路径列表（如果output_dir为None则为空列表）
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
    
    # 打开PDF文件
    pdf_document = fitz.open(pdf_path)
    images = []
    image_paths = []
    
    # 如果指定了输出目录且不存在，则创建
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    pdf_name = os.path.basename(pdf_path).rsplit('.', 1)[0]
    
    # 遍历每一页
    for page_index, page in enumerate(pdf_document):
        # 获取页面上的图片
        image_list = page.get_images(full=True)
        
        # 处理页面上的每张图片
        for img_index, img_info in enumerate(image_list):
            img_index_in_doc = img_info[0]  # 图片在文档中的索引
            base_img = pdf_document.extract_image(img_index_in_doc)
            image_bytes = base_img["image"]
            
            # 使用PIL打开图片
            try:
                img = Image.open(io.BytesIO(image_bytes))
                width, height = img.size
                
                # 忽略过小的图片
                if width < min_width or height < min_height:
                    continue
                
                # 保存图片
                images.append(img)
                
                if output_dir:
                    img_path = os.path.join(output_dir, f"{pdf_name}_p{page_index+1}_img{img_index+1}.png")
                    img.save(img_path)
                    image_paths.append(img_path)
                    
            except Exception as e:
                print(f"处理图片时出错: {e}")
    
    return images, image_paths

def main():
    parser = argparse.ArgumentParser(description='PDF文档图片OCR识别示例')
    parser.add_argument('pdf_path', help='PDF文件路径')
    parser.add_argument('-o', '--output_dir', help='输出目录路径')
    parser.add_argument('-l', '--lang', default='ch', help='识别语言, 默认为中文(ch)')
    parser.add_argument('--gpu', action='store_true', help='是否使用GPU加速')
    parser.add_argument('--min-width', type=int, default=100, help='最小图片宽度，小于此宽度的图片会被忽略')
    parser.add_argument('--min-height', type=int, default=100, help='最小图片高度，小于此高度的图片会被忽略')
    parser.add_argument('--save-images', action='store_true', help='是否保存提取的图片')
    
    args = parser.parse_args()
    
    # 创建输出目录(如果指定的话)
    if args.output_dir:
        if not os.path.exists(args.output_dir):
            os.makedirs(args.output_dir)
        images_dir = os.path.join(args.output_dir, "images")
        results_dir = os.path.join(args.output_dir, "results")
    else:
        # 使用临时目录
        temp_dir = tempfile.mkdtemp()
        images_dir = os.path.join(temp_dir, "images")
        results_dir = os.path.join(temp_dir, "results")
    
    # 创建图片目录(如果需要保存图片)
    if args.save_images:
        if not os.path.exists(images_dir):
            os.makedirs(images_dir)
    
    # 创建结果目录
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)
    
    # 提取PDF中的图片
    print(f"正在从PDF文件提取图片: {args.pdf_path}")
    output_images_dir = images_dir if args.save_images else None
    images, image_paths = extract_images_from_pdf(
        args.pdf_path, 
        output_images_dir,
        min_width=args.min_width,
        min_height=args.min_height
    )
    
    print(f"从PDF中提取了 {len(images)} 张图片")
    
    if len(images) == 0:
        print("未找到符合条件的图片，退出程序")
        return
    
    # 初始化OCR处理器
    print("初始化OCR处理器...")
    processor = OCRProcessor(
        use_gpu=args.gpu,
        lang=args.lang,
        use_angle_cls=True
    )
    
    # 处理提取的图片
    all_text = []
    
    if args.save_images:
        # 如果保存了图片，直接处理图片文件
        print("处理提取的图片文件...")
        processor.process_directory(images_dir, results_dir)
        
        # 读取所有OCR结果文本
        for img_path in image_paths:
            img_name = os.path.basename(img_path).rsplit('.', 1)[0]
            txt_path = os.path.join(results_dir, f"{img_name}_ocr.txt")
            if os.path.exists(txt_path):
                with open(txt_path, 'r', encoding='utf-8') as f:
                    text = f.read().strip()
                    if text:
                        all_text.append(f"--- 图片: {img_name} ---")
                        all_text.append(text)
                        all_text.append("")
    else:
        # 内存中处理图片对象
        print("处理提取的图片...")
        
        # 创建临时图片文件
        temp_img_paths = []
        with tempfile.TemporaryDirectory() as temp_dir:
            for i, img in enumerate(images):
                temp_img_path = os.path.join(temp_dir, f"temp_img_{i+1}.png")
                img.save(temp_img_path)
                temp_img_paths.append(temp_img_path)
            
            # 处理临时图片文件
            results = processor.process_directory(temp_dir, results_dir)
            
            # 读取所有OCR结果文本
            for i, temp_img_path in enumerate(temp_img_paths):
                img_name = os.path.basename(temp_img_path).rsplit('.', 1)[0]
                txt_path = os.path.join(results_dir, f"{img_name}_ocr.txt")
                if os.path.exists(txt_path):
                    with open(txt_path, 'r', encoding='utf-8') as f:
                        text = f.read().strip()
                        if text:
                            all_text.append(f"--- 图片 {i+1} ---")
                            all_text.append(text)
                            all_text.append("")
    
    # 合并所有OCR结果
    if all_text:
        pdf_name = os.path.basename(args.pdf_path).rsplit('.', 1)[0]
        combined_txt_path = os.path.join(results_dir, f"{pdf_name}_all_ocr.txt")
        with open(combined_txt_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(all_text))
        print(f"合并OCR结果已保存至: {combined_txt_path}")
    
    print("\n处理完成!")
    print(f"处理图片数量: {len(images)}")
    print(f"识别结果保存在: {results_dir}")

if __name__ == "__main__":
    main() 