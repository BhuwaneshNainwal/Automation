from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Flowable, Paragraph
from .static import create_border, PAGE_SIZE, PAGE_HEIGHT, PAGE_WIDTH
from typing import Tuple


class PDFPage:
    def __init__(self):
        self.items: tuple[Flowable] = None

    def draw_page(self, canvas):
        create_border(canvas)
        for item in self.items:
            item.draw(canvas)

    def add(self, item):
        self.items = item


class PDFItem:
    def __init__(self, flowable, x, y):
        self.position = x, y
        self.flowable: Flowable = flowable  # To be drawn

    def draw(self, canvas):
        self.flowable.wrap(1000, 400)
        self.flowable.drawOn(canvas, *self.position)


class PDF:
    def __init__(self, dest, source=None, img_loc=None, size=(800, 850)):
        self.source = source
        self.pages = []
        self.size = size
        self.canvas = Canvas(dest, pagesize=self.size)
        print("Generating ", dest)
        pass

    def prepare(self, size: Tuple[int, int]):
        for page in self.pages:
            page.draw_page(self.canvas)
            self.canvas.showPage()
            self.canvas.setPageSize(size)
        self.canvas.save()

    def add_page(self, page: PDFPage):
        self.pages.append(page)
