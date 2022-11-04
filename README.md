### 利用ocr将pdf转为docx

#### 背景
该项目首先将pdf文件转为图片形式，再使用百度的paddleocr对这些图片文件分别进行识别，利用PPStructure对识别的内容进行结构化，最终将结构化的内容保存成docx文件，
最后将所有docx文件进行合并，形成一个docx文件。

#### 代码结构
这里将paddleocr源码拿过来进行了相应修改，以支持中文路径
```
├── pdf_recovery_doc.py  # 主程序，用于pdf到docx转换
├── image_process.py    # 用于pdf到image转换
```


#### 主程序
```
def pdf2doc(pdf_file):
    '''
    pdf2doc
    :param pdf_file:
    :return:
    '''
    table_engine = PPStructure(recovery=True, lang='ch', enable_mkldnn=True)
    start_time = time.time()
    if pdf_file.endswith('.pdf') or pdf_file.endswith('.PDF'):
        pdf_file_base_name = os.path.splitext(pdf_file)[0]
        image_path = pdf_file_base_name + '/' + 'images'

        # 1.pdf to image
        pdf_image(pdf_file, image_path)
        image_files = os.listdir(image_path)

        # 2.image to docx
        for img_file in image_files:
            if img_file.endswith('png'):
                img = cv2.imdecode(np.fromfile(os.path.join(image_path, img_file), dtype=np.uint8), -1)
                img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                if img is not None:
                    result = table_engine(img)
                    save_structure_res(result, pdf_file_base_name, img_file.split('.')[0])
                    h, w, _ = img.shape
                    res = sorted_layout_boxes(result, w)
                    convert_info_docx(img, res, pdf_file_base_name, img_file.split('.')[0])

        # 3.merge all docx
        merge_docx_v1(pdf_file_base_name)

    end_time = time.time()
    print('pdf转换doc时间: {}'.format(end_time - start_time))
```
#### 例子
##### pdf文件
![image](https://raw.githubusercontent.com/jiangnanboy/pdf_to_docx/master/imgs/raw.png)

##### docx文件
![image](https://raw.githubusercontent.com/jiangnanboy/pdf_to_docx/master/imgs/result.png)

