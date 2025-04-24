# PaddleOCR 图像文字识别脚本

这是一个基于PaddleOCR的图像文字识别工具，支持单张图片识别和批量目录处理，可识别多种语言。

## 功能特点

- 支持单张图片OCR识别
- 支持批量处理整个目录的图片
- 支持多语言识别（中文、英文、日语、韩语等）
- 支持识别结果可视化（带标注框的图片）
- 支持将识别结果保存为文本文件
- 支持GPU加速（如果系统有CUDA环境）

## 环境要求

- Python 3.6+
- PaddlePaddle 2.5.0+
- PaddleOCR 2.6.0+
- OpenCV 4.5.0+
- Pillow 8.0.0+

## 安装依赖

使用pip安装所需依赖：

```bash
pip install -r requirements.txt
```

如果您希望使用GPU加速，请安装GPU版本的PaddlePaddle：

```bash
pip install paddlepaddle-gpu
```

## 使用方法

### 单张图片识别

```bash
python paddleocr_recognition.py image.jpg
```

指定输出路径：

```bash
python paddleocr_recognition.py image.jpg -o output/result.txt
```

### 批量处理目录

```bash
python paddleocr_recognition.py images_directory/
```

指定输出目录：

```bash
python paddleocr_recognition.py images_directory/ -o output_directory/
```

### 其他参数

- `-l, --lang`: 设置识别语言，默认为中文(`ch`)，可选项包括：`ch`(中文)、`en`(英文)、`fr`(法语)、`german`(德语)、`korean`(韩语)、`japan`(日语)
- `--gpu`: 使用GPU加速
- `--no-angle-cls`: 不使用方向分类器
- `--txt-only`: 只输出文本文件，不生成可视化图片

示例：

```bash
# 使用英文识别模型
python paddleocr_recognition.py image.jpg -l en

# 使用GPU加速
python paddleocr_recognition.py image.jpg --gpu

# 不使用方向分类器（提高速度，但可能影响识别精度）
python paddleocr_recognition.py image.jpg --no-angle-cls

# 只输出文本文件，不生成可视化图片
python paddleocr_recognition.py image.jpg --txt-only
```

## 输出文件

脚本会生成两种文件：

1. 文本文件（`*_ocr.txt`）：包含识别的纯文本内容
2. 可视化图片（`*_result.jpg/png`等）：原始图片上标注了识别区域和文字

## 示例

输入图片：
![输入图片](example_input.jpg)

识别结果图片：
![识别结果](example_input_result.jpg)

识别文本（example_input_ocr.txt）：
```
识别出的文本内容
多行文本示例
```

## 注意事项

1. 首次运行时，脚本会自动下载相应的OCR模型文件（约300MB），请确保网络连接正常
2. 对于大尺寸图片，处理速度可能较慢，建议使用GPU加速
3. 使用可视化功能需要确保系统中有中文字体文件

## 常见问题解决

### 字体问题

如果遇到字体相关错误（`OSError: unknown file format`），请确保系统中有中文字体文件。对于 macOS 用户，系统默认包含多种中文字体如 PingFang.ttc、STHeiti Light.ttc 等。脚本已配置为自动检测这些字体。

### 模块依赖问题

如果遇到 `ModuleNotFoundError`，请确保已安装所有必需的依赖：

```bash
pip install setuptools paddlepaddle paddleocr opencv-python Pillow
```

对于PDF处理功能，还需要安装：

```bash
pip install PyMuPDF
```

### 虚拟环境问题

为避免依赖冲突，建议使用虚拟环境：

```bash
# 创建虚拟环境
python -m venv .venv

# 激活虚拟环境
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 使用示例

识别单张图片并查看结果：

```bash
source .venv/bin/activate  # 激活虚拟环境
python paddleocr_recognition.py your_image.jpg
```

批量处理图片：

```bash
source .venv/bin/activate  # 激活虚拟环境
python batch_ocr.py your_images_directory/ -o output_directory/
``` 

批量PDF文档中的图片：

```bash
source .venv/bin/activate  # 激活虚拟环境
python pdf_ocr.py your_document.pdf -o output_directory/ --save-images
``` 