from pptx import Presentation
from pptx.util import Inches


class PPTXItemPipeline(object):
    def __init__(self, file_name):
        self.__file_name = file_name

    @classmethod
    def from_crawler(cls, crawler):
        return cls(
            file_name=crawler.settings.get('PPTX_FILE_NAME'),
        )

    def open_spider(self, spider):
        self.__presentation = Presentation()

    def close_spider(self, spider):
        self.__presentation.save(self.__file_name)

    def process_item(self, item, spider):
        blank_slide_layout = self.__presentation.slide_layouts[1]
        slide = self.__presentation.slides.add_slide(blank_slide_layout)
        slide.shapes.title.text = ''.join(item['title'])

        #left = top = width = height = Inches(1)
        #txBox = slide.shapes.add_textbox(left, top, width, height)
        #tf = txBox.text_frame
        #tf.text = ''.join(item['title'])
        #p = tf.add_paragraph()
        #p.text = "This is a second paragraph that's bold"

        return item
