# DOCX图片移除工具

这个项目提供了三个不同的Python脚本，用于从Word文档(.docx)中移除所有图片，同时保留文档的其他内容。

## 脚本说明

### 1. remove_images_from_docx.py

- **实现方式**：使用python-docx库解析和修改Word文档
- **优点**：API简洁清晰，对文档结构的操作更加直观
- **缺点**：依赖较多，运行速度相对较慢
- **适用场景**：希望使用更高级API处理文档的场景

### 2. remove_images_direct.py

- **实现方式**：直接操作DOCX文件的内部ZIP结构和XML内容
- **优点**：依赖较少，只需要Python标准库
- **缺点**：XML处理能力有限，可能存在命名空间处理问题
- **适用场景**：在环境受限无法安装第三方库的情况下使用

### 3. remove_images_lxml.py

- **实现方式**：使用lxml库高效解析和修改XML
- **优点**：性能最佳，XML处理能力强大，支持完整的命名空间
- **缺点**：需要安装lxml库
- **适用场景**：处理大型文档或批量处理多个文档时首选

## 安装依赖

根据选择的脚本，安装相应的依赖：

```bash
# 对于remove_images_from_docx.py
pip install python-docx

# 对于remove_images_lxml.py
pip install lxml

# 对于remove_images_direct.py
# 不需要额外依赖，使用Python标准库即可
```

## 使用方法

所有脚本的使用方法相同：

```bash
# 基本用法
python 脚本名.py 输入文件.docx [输出文件.docx]

# 例如:
python remove_images_lxml.py 我的文档.docx 无图片文档.docx
```

如果不指定输出文件名，脚本会自动在原文件名前添加"无图片_"前缀。

## 工作原理

这些脚本的工作原理基于DOCX文件的结构：

1. DOCX文件本质上是一个ZIP压缩包
2. 压缩包内包含多个XML文件和资源文件夹
3. 文档的主要内容存储在word/document.xml中
4. 图片等媒体文件存储在word/media/文件夹中
5. 文档与图片之间的关系通过XML元素和关系文件(word/_rels/document.xml.rels)建立

脚本通过以下步骤移除图片：

1. 解压DOCX文件到临时目录
2. 修改XML文件，删除所有图片相关元素
3. 删除media文件夹中的图片文件
4. 清理关系文件中的图片引用
5. 重新打包成DOCX文件

## 注意事项

- 处理前请备份原始文档
- 这些脚本会移除所有图片，没有选择性保留的功能
- 处理后的文档可能在图片原位置留下空白区域 

## Word文档转Markdown脚本

本仓库包含两个将Word文档(.docx)转换为Markdown格式的脚本:

### 1. docx_to_markdown.py

使用python-docx库直接解析.docx文件并转换为Markdown。

**特点:**
- 直接使用python-docx解析文档结构
- 支持基本的文本样式、标题、列表和表格
- 能够提取并保存文档中的图片

**依赖:**
```
python-docx>=0.8.11
```

**使用方法:**
```
python docx_to_markdown.py 输入文件.docx [输出文件.md]
```

### 2. docx_to_markdown_enhanced.py

使用mammoth库进行中间HTML转换，支持更丰富的格式和样式。

**特点:**
- 通过HTML中间转换，更准确地处理复杂样式和格式
- 支持更多的Word样式元素（引用、目录、代码块等）
- 更好的表格和列表格式支持
- 优化的图片处理方式

**依赖:**
```
python-docx>=0.8.11
mammoth>=1.5.0
html2markdown>=0.1.7
```

**使用方法:**
```
python docx_to_markdown_enhanced.py 输入文件.docx [输出文件.md]
```

## 两个转换脚本的比较

| 功能 | docx_to_markdown.py | docx_to_markdown_enhanced.py |
| --- | --- | --- |
| 基本文本转换 | ✓ | ✓ |
| 图片提取 | ✓ | ✓ |
| 表格支持 | 基本支持 | 更好的格式处理 |
| 复杂样式 | 有限支持 | 更全面支持 |
| 转换原理 | 直接解析 | HTML中间转换 |
| 依赖库 | 较少 | 较多 |
| 速度 | 较快 | 可能较慢 |

## 移除Word文档中的图片脚本

本仓库还包含以下几个用于移除Word文档中图片的脚本:

1. **remove_images_from_docx.py**: 使用python-docx库实现
2. **remove_images_direct.py**: 使用Python标准库实现
3. **remove_images_lxml.py**: 使用lxml库实现

详细说明请参考本文档上方的"DOCX图片移除工具"部分。 