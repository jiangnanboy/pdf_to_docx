import os
import fitz
import datetime


def pdf_image(pdf_file, image_path):
    '''
    pdf to image
    :param pdf_path:
    :param image_path:
    :return:
    '''
    start_time = datetime.datetime.now()
    pdf_doc = fitz.open(pdf_file)
    for pg in range(pdf_doc.pageCount):
        page = pdf_doc[pg]
        rotate = int(0)
        zoom_x = 4
        zoom_y = 4
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)
        if not os.path.exists(image_path):
            os.makedirs(image_path)
        pix.writePNG(image_path + '/' + 'images_%s.png' % pg)
    end_time = datetime.datetime.now()
    print('pdf转换图片时间: {}s'.format((end_time - start_time).seconds))
