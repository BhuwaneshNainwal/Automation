from matplotlib import pyplot as plt
from io import BytesIO

FONT_FOR_TITLE = {'fontsize': 16}
BG_COLOR = (0, 0.6, 1)


def pie(data, label, title, auto, explode=None):
    file = BytesIO()
    fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))
    fig.patch.set_facecolor((*BG_COLOR, 0))
    try:
        wedges, texts, autotexts = ax.pie(data, shadow=True, autopct=auto, explode=explode)
    except ValueError:
        wedges, texts, autotexts = ax.pie(data, shadow=True, autopct=auto, explode=None)
    ax.legend(wedges, label, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    ax.set_title(title, FONT_FOR_TITLE)
    plt.savefig(file, format="png")
    plt.clf()
    return file


def time(data, label, title):
    file = BytesIO()
    fig, ax = plt.subplots(figsize=(3, 3))
    fig.patch.set_facecolor((*BG_COLOR, 0))
    ax.bar(label, data, width=0.4)
    ax.set_title(title, FONT_FOR_TITLE)
    plt.savefig(file, format="png")
    plt.clf()
    return file
