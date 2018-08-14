Python Batch Rename 批量重命名
----------------

[2018.8.14] 更新添加中文支持,更新docx,预计添加txt,doc等重命名

## 批量PDF重命名

- Origin from

 * [pdftitle.py-hanjianwei](https://gist.github.com/hanjianwei/6838974)

- Requirements

 * [pyPDF](https://github.com/mstamy2/PyPDF2)
 * [PDFMiner](https://github.com/euske/pdfminer/)


提前安装最新版模块`PDFMiner`和`pyPDF`,安装完将拷贝所有py文件到PYTHONPATH目录下,没有的话拷贝到python环境变量目录,本人安装anaconda所以
`C:\Users\xxxx\Anaconda3\Scripts`(MacOS/Linux按需更改),该目录在环境变量Path里面，因此在cmd里面运行脚本：

Usage:

    复制文件到当前文档,执行脚本
    单文件： pdf-rename-batch.py test.pdf
    多文件： pdf-rename-batch.py *.pdf

## 批量Word文档重命名

灵感来自于重命名邮单和发票

- [x] docx
- [] doc

post_tag 用于取名停止位置例如:"编号""邮编"等

Usage:

    复制文件到当前文档,执行脚本
    docx-rename-batch.py [--dry-run] [-p|--post-tag post_tag] test.docx

That's all,enjoy.


