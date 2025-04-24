#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
批量图片OCR处理示例脚本
使用paddleocr_recognition.py中的OCRProcessor类进行批量图片处理
"""

import os
import argparse
from paddleocr_recognition import OCRProcessor

def main():
    parser = argparse.ArgumentParser(description='批量图片OCR处理示例')
    parser.add_argument('input_dir', help='输入图片目录路径')
    parser.add_argument('-o', '--output_dir', help='输出结果目录路径')
    parser.add_argument('-l', '--lang', default='ch', help='识别语言, 默认为中文(ch)')
    parser.add_argument('--gpu', action='store_true', help='是否使用GPU加速')
    
    args = parser.parse_args()
    
    # 创建输出目录(如果指定的话)
    if args.output_dir and not os.path.exists(args.output_dir):
        os.makedirs(args.output_dir)
    
    # 初始化OCR处理器
    print("初始化OCR处理器...")
    processor = OCRProcessor(
        use_gpu=args.gpu,
        lang=args.lang,
        use_angle_cls=True
    )
    
    # 处理图片目录
    print(f"开始处理目录: {args.input_dir}")
    results = processor.process_directory(args.input_dir, args.output_dir)
    
    # 统计结果
    total_files = len(results)
    files_with_text = sum(1 for result in results.values() if result and len(result[0]) > 0)
    
    print("\n处理完成!")
    print(f"总共处理: {total_files} 个文件")
    print(f"识别到文字: {files_with_text} 个文件")
    print(f"无文字内容: {total_files - files_with_text} 个文件")
    
    if args.output_dir:
        print(f"结果已保存到: {args.output_dir}")

if __name__ == "__main__":
    main() 