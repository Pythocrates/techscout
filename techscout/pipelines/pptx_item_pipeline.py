import os

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches, Pt
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer


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
        blank_slide_layout = self.__presentation.slide_layouts[5]
        slide = self.__presentation.slides.add_slide(blank_slide_layout)
        slide.shapes.title.text = ''.join(item['title'])
        slide.shapes.title.text_frame.paragraphs[0].runs[0].font.size = Pt(20)

        left, top, width, height = [Inches(i) for i in [1, 1.5, 8, 6]]
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame

        text = ' '.join(item['text'])
        parser = PlaintextParser.from_string(text, Tokenizer('english'))
        summarizer = LexRankSummarizer()
        sentences = summarizer(parser.document, 4)
        text = ' '.join(str(s) for s in sentences)
        tf.text = text
        tf.paragraphs[0].runs[0].font.size = Pt(10)
        tf.margin_left = 0
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        left, top, width, height = [Inches(i) for i in [1, 7, 8, 2]]
        tags_box = slide.shapes.add_textbox(left, top, width, height)
        tf = tags_box.text_frame
        tags = item['tags']
        tf.text = ', '.join(tags).replace('  ', ' ')
        tf.paragraphs[0].runs[0].font.size = Pt(8)
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        image_path = item['images'][0]['path']
        slide.shapes.add_picture(
            os.path.join(
                spider.settings.attributes['IMAGES_STORE'].value,
                image_path),
            Inches(1), Inches(4), height=Inches(2))

        return item
