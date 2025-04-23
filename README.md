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