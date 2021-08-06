from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4

MARGIN_LEFT = MARGIN_TOP = cm / 2
PAGE_SIZE = (850, 950)
PAGE_WIDTH, PAGE_HEIGHT = PAGE_SIZE


def create_border(page):
    page.rect(MARGIN_LEFT, MARGIN_TOP, PAGE_WIDTH - 2 * MARGIN_LEFT, PAGE_HEIGHT - 2 * MARGIN_LEFT, fill=0)
    return
