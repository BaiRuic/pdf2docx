## PDF2DOCX

一个Python脚本，可以将pdf文件**100%复刻**为word, 实现思路为先将pdf转为jpeg图片，然后将图片插入到页边距为0的word中。

### 要求：
Python3.x环境，需要的第三方包：pdf2image, docx, python-docx, tqdm。可通过`pip install pdf2image docx python-docx tqdm`安装。

### 使用方法：
1. 下载该脚本至待转换pdf的文件路径下
    + 例如可通过wget下载该脚本 `wget https://github.com/BaiRuic/pdf2docx/blob/master/pdf2docx.py -O pdf2docx.py`
    + 或者直接下载该仓库，然后手动将.py文件移动至指定路径。
2. 打开命令行，执行 `python pdf2docx.py <pdf file name>`即可。
    
    注：如果想指定转换质量，可以通过可选参数`dpi`来设置图片的dpi；进一步的如果想设置转换后的文件名称，可以通过可选参数`output`来设置文件名。 
    
    例：当待转换pdf文件名为 `main.pdf`, 转换后的`.docx`文件名称希望为 `example.docx`，想要设置的`dpi`为400，则在命令行输入`python pdf2docx.py main.pdf --output=example.docx --dpi=400`

### 注意
+ 使用该方法转换得到的word不可编辑！
+ dpi不适宜调太大，建议400，最多800，不然转换后的.docx文件很大。
