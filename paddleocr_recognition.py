#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
使用PaddleOCR进行图片文字识别
支持单张图片识别和批量图片识别
支持多语言识别
"""

import os
import cv2
import argparse
import numpy as np
from PIL import Image
from pathlib import Path
from paddleocr import PaddleOCR, draw_ocr

class OCRProcessor:
    """PaddleOCR图像文字识别处理器"""
    
    def __init__(self, use_gpu=False, lang='ch', use_angle_cls=True):
        """
        初始化OCR处理器
        
        参数:
            use_gpu: 是否使用GPU加速
            lang: 识别语言，如'ch'(中文),'en'(英文),'jp'(日语)等
            use_angle_cls: 是否使用方向分类器
        """
        self.ocr = PaddleOCR(use_angle_cls=use_angle_cls, lang=lang, use_gpu=use_gpu)
        self.lang = lang
        self.use_gpu = use_gpu
    
    def recognize_image(self, img_path):
        """
        识别单张图片中的文字
        
        参数:
            img_path: 图片路径
            
        返回:
            result: OCR识别结果列表
        """
        if not os.path.exists(img_path):
            raise FileNotFoundError(f"图片文件不存在: {img_path}")
            
        result = self.ocr.ocr(img_path, cls=True)
        return result
    
    def visualize_result(self, img_path, result, output_path=None):
        """
        可视化OCR识别结果
        
        参数:
            img_path: 原始图片路径
            result: OCR识别结果
            output_path: 输出图片路径，默认为None（在当前目录生成带有'_result'后缀的图片）
        
        返回:
            output_path: 输出图片路径
        """
        if not result or len(result) == 0:
            print(f"警告: 图片 {img_path} 没有识别到任何文字")
            return None
        
        # 读取图片
        img = cv2.imread(img_path)
        if img is None:
            raise ValueError(f"无法读取图片: {img_path}")
        
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        boxes = [line[0] for line in result[0]]
        txts = [line[1][0] for line in result[0]]
        scores = [line[1][1] for line in result[0]]
        
        # 绘制结果 - 使用系统字体路径
        font_path = '/System/Library/Fonts/PingFang.ttc'  # macOS系统中文字体
        if not os.path.exists(font_path):
            # 尝试使用其他常见系统字体
            alternative_fonts = [
                '/System/Library/Fonts/STHeiti Light.ttc',  # macOS另一个中文字体
                '/System/Library/Fonts/Hiragino Sans GB.ttc',  # macOS另一个中文字体
                '/Library/Fonts/Arial Unicode.ttf',  # 通用Unicode字体
                './fonts/simfang.ttf'  # 回退到原始路径
            ]
            for alt_font in alternative_fonts:
                if os.path.exists(alt_font):
                    font_path = alt_font
                    break
        
        im_show = draw_ocr(img, boxes, txts, scores, font_path=font_path)
        im_show = Image.fromarray(im_show)
        
        # 生成输出路径
        if output_path is None:
            img_name = os.path.basename(img_path)
            name, ext = os.path.splitext(img_name)
            output_path = os.path.join(os.path.dirname(img_path), f"{name}_result{ext}")
        
        im_show.save(output_path)
        print(f"结果图片已保存至: {output_path}")
        return output_path
    
    def extract_text(self, result):
        """
        提取OCR识别结果中的文本内容
        
        参数:
            result: OCR识别结果
            
        返回:
            text: 提取的文本内容
        """
        if not result or len(result) == 0:
            return ""
        
        text = "\n".join([line[1][0] for line in result[0]])
        return text
    
    def save_to_txt(self, text, output_path=None, img_path=None):
        """
        将识别的文本保存为TXT文件
        
        参数:
            text: 文本内容
            output_path: 输出TXT文件路径，默认为None
            img_path: 原始图片路径，用于生成默认输出路径
            
        返回:
            output_path: 输出文件路径
        """
        if output_path is None and img_path is not None:
            img_name = os.path.basename(img_path)
            name, _ = os.path.splitext(img_name)
            output_path = os.path.join(os.path.dirname(img_path), f"{name}_ocr.txt")
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        print(f"文本已保存至: {output_path}")
        return output_path
    
    def process_directory(self, dir_path, output_dir=None, extensions=('.jpg', '.jpeg', '.png', '.bmp')):
        """
        批量处理目录中的图片
        
        参数:
            dir_path: 图片目录路径
            output_dir: 输出目录路径，默认为None（在原目录生成结果）
            extensions: 支持的图片扩展名元组
            
        返回:
            results: 处理结果字典，键为图片路径，值为识别结果
        """
        if not os.path.isdir(dir_path):
            raise NotADirectoryError(f"目录不存在: {dir_path}")
        
        # 如果指定了输出目录且不存在，则创建
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 获取所有图片文件
        image_files = []
        for ext in extensions:
            image_files.extend(list(Path(dir_path).glob(f"*{ext}")))
            image_files.extend(list(Path(dir_path).glob(f"*{ext.upper()}")))
        
        if not image_files:
            print(f"警告: 目录 {dir_path} 中没有找到支持的图片文件")
            return {}
        
        results = {}
        total = len(image_files)
        
        for i, img_path in enumerate(image_files):
            img_path_str = str(img_path)
            print(f"正在处理 [{i+1}/{total}]: {img_path_str}")
            
            try:
                # 识别图片
                result = self.recognize_image(img_path_str)
                results[img_path_str] = result
                
                # 提取文本
                text = self.extract_text(result)
                
                # 确定输出路径
                if output_dir:
                    img_name = os.path.basename(img_path_str)
                    name, ext = os.path.splitext(img_name)
                    out_img_path = os.path.join(output_dir, f"{name}_result{ext}")
                    out_txt_path = os.path.join(output_dir, f"{name}_ocr.txt")
                else:
                    out_img_path = None  # 使用默认路径
                    out_txt_path = None  # 使用默认路径
                
                # 可视化结果并保存文本
                if result and len(result[0]) > 0:  # 只有识别到内容才保存结果
                    self.visualize_result(img_path_str, result, out_img_path)
                    self.save_to_txt(text, out_txt_path, img_path_str)
                else:
                    print(f"图片 {img_path_str} 没有识别到任何文字")
            
            except Exception as e:
                print(f"处理图片 {img_path_str} 时出错: {e}")
        
        return results

def main():
    parser = argparse.ArgumentParser(description='使用PaddleOCR进行图片文字识别')
    parser.add_argument('input', help='输入图片路径或目录路径')
    parser.add_argument('-o', '--output', help='输出目录路径（批量处理时）或输出文件路径（单张图片）')
    parser.add_argument('-l', '--lang', default='ch', help='识别语言, 默认为中文(ch)', 
                        choices=['ch', 'en', 'fr', 'german', 'korean', 'japan'])
    parser.add_argument('--gpu', action='store_true', help='是否使用GPU加速')
    parser.add_argument('--no-angle-cls', action='store_true', help='不使用方向分类器')
    parser.add_argument('--txt-only', action='store_true', help='只输出文本文件，不生成可视化图片')
    
    args = parser.parse_args()
    
    # 初始化OCR处理器
    processor = OCRProcessor(
        use_gpu=args.gpu,
        lang=args.lang,
        use_angle_cls=not args.no_angle_cls
    )
    
    # 处理输入
    input_path = args.input
    if os.path.isdir(input_path):
        # 批量处理目录
        processor.process_directory(input_path, args.output)
    elif os.path.isfile(input_path):
        # 处理单张图片
        result = processor.recognize_image(input_path)
        text = processor.extract_text(result)
        
        print("\n识别结果:")
        print("-" * 30)
        print(text)
        print("-" * 30)
        
        # 保存文本结果
        if args.output:
            # 检查输出路径是目录还是文件
            if os.path.isdir(args.output) or args.output.endswith('/'):
                if not os.path.exists(args.output):
                    os.makedirs(args.output)
                output_txt = None  # 使用默认文件名
                output_img = None  # 使用默认文件名
            else:
                # 指定了具体文件名
                output_txt = args.output if args.output.endswith('.txt') else f"{args.output}.txt"
                output_dir = os.path.dirname(output_txt)
                if output_dir and not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                output_img = output_txt.rsplit('.', 1)[0] + '_result.png'
        else:
            output_txt = None  # 使用默认文件名
            output_img = None  # 使用默认文件名
        
        processor.save_to_txt(text, output_txt, input_path)
        
        # 可视化结果
        if not args.txt_only and result and len(result[0]) > 0:
            processor.visualize_result(input_path, result, output_img)
    else:
        print(f"错误: 输入路径 {input_path} 不存在")

if __name__ == "__main__":
    main() 