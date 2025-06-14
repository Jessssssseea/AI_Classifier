# AI文件分类器

## 简介
AI文件分类器是一个基于Python开发的工具，用于自动对指定文件夹中的文件进行分类。该工具支持多种文件格式，包括DOCX、PDF、PPTX等，通过文件名和文件内容进行分类判断，并将文件移动到相应的学科文件夹中。同时，工具还提供了用户反馈机制，允许用户对分类结果进行确认。

## 功能特点
1. **多格式支持**：支持DOCX、PDF、PPTX、PPT、DOC、WPS、MP4、WBD等多种文件格式。
2. **智能分类**：先通过文件名进行初步分类，若无法确定则提取文件内容，使用AI模型进行分类。
3. **用户反馈**：对分类结果发送通知，用户可确认分类是否正确，错误时文件会移回原位置。
4. **配置灵活**：可通过配置文件设置监控文件夹、输出文件夹、延迟时间等参数。
5. **日志记录**：记录文件处理过程中的信息，方便查看和排查问题。

## 安装与配置
### 1. 安装依赖
确保你已经安装了Python 3.x，并安装以下依赖库：
```bash
pip install -r requirements.txt
```

### 2. 配置文件
编辑`config.json`文件，设置监控文件夹、输出文件夹、延迟时间等参数：
```json
{
    "WATCH_FOLDER": "D:/定期练手/自动归纳/monitor",
    "OUTPUT_BASE_FOLDER": "D:/定期练手/自动归纳/classified",
    "DELAY_SECONDS": 3,
    "SUPPORTED_EXTS": [
        ".docx",
        ".pdf",
        ".pptx",
        ".ppt",
        ".doc",
        ".wps",
        ".mp4",
        ".wbd"
    ],
    "AUTO_START": true
}
```

### 3. 模型文件
确保`subject_classifier.pkl`和`tfidf_vectorizer.pkl`模型文件存在，若不存在，请先运行训练脚本。

## 使用方法
### 1. 启动文件监控
运行`file_classifier.py`、`file_classifier_beta.py`或`file_classifier_canary.py`脚本，开始监控指定文件夹中的文件创建事件：
```bash
python file_classifier.py
```

### 2. 查看日志
可以通过右键点击系统托盘图标，选择“查看日志”来查看文件处理过程中的信息。

### 3. 清空日志
同样在系统托盘图标菜单中，选择“清空日志”可以清空日志文件。

### 4. 配置编辑器
可以通过右键点击系统托盘图标，选择“打开配置”来打开配置编辑器，修改监控文件夹、输出文件夹、延迟时间等参数。

## 代码结构
- **`file_classifier.py`**：核心文件，实现文件监控、分类和移动功能。
- **`file_classifier_beta.py`**：增加用户反馈处理，通过发送通知让用户确认分类结果。
- **`file_classifier_canary.py`**：与`file_classifier_beta.py`类似，增加了一些日志记录和错误处理。
- **`extract.py`**：提供多种文件格式的内容提取功能，包括DOCX、PPTX、PDF等。
- **`config_editor.py`**：用于编辑配置文件，设置监控文件夹、输出文件夹、延迟时间等参数。
- **`config.json`**：存储配置信息。
- **`file_classifier.log`**：日志文件，记录文件处理过程中的信息。

## 注意事项
1. 确保[Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)已安装，并使用`config_editor`正确设置`Tesseract`路径。
2. 若文件在复制过程中，可能会出现文件无法访问的情况，工具会等待文件复制完成后再进行处理。
3. 对于某些特殊格式的文件（如WBD），可能需要特殊解析器，目前暂不支持内容提取。

## 贡献
如果你有任何建议或发现了问题，请在GitHub上提交[issue](https://github.com/Jessssssseea/AI-courseware-induction/issues)或[pull request](https://github.com/Jessssssseea/AI-courseware-induction/pulls)。

## 许可证
本项目采用[MIT许可证](https://opensource.org/licenses/MIT)。
