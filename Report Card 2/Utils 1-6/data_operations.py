import pandas as pd
from matplotlib import pyplot as plt
from io import BytesIO
from typing import Tuple
from numpy import int64

FONT_FOR_TITLE = {'fontsize': 18}


def get_data(filename: str, sheet: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    head = ['Student No', 'First Name of Candidate',
            'Last Name of Candidate', 'Full Name of Candidate',
            'Registration', 'Grade', 'Name of school', 'Gender',
            'Date of Birth', 'Date and time of test', 'City of Residence',
            'Country of Residence', 'Question No.', 'What you marked',
            'Correct Answer', 'Outcome (Correct/Incorrect/Not Attempted)',
            'Score if correct', 'Your score']
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        data = pd.read_excel(filename, name=head, sheet_name=sheet, dtype={'Your score': int}, converters={'Your score': int})
    elif filename.endswith('.csv'):
        data = pd.read_csv(filename, name=head, sheet_name=sheet, dtype={'Your score': int}, converters={'Your score': int})
    else:
        raise TypeError('Invalid file extension.')
    data = data.dropna(axis=1, how="all")
    data = data.dropna(axis=0, thresh=6)
    data = data.fillna(value=" ")
    data.columns = head
    data = data.iloc[1:]
    # if not data.columns.any() == head:
    #     data.columns = head
    # if (data[0:1] == head).values.any():
    #     data = data.iloc[1:]
    static = data[head[:12]].drop_duplicates()
    cont = data[head[12:]+['Country of Residence', 'Student No']]
    return static, cont


def get_percent_of_attempted_questions(data, take_col, take_col_value, filter_col_1=None, col_1_val=None,
                                       filter_col_2=None, col_2_val=None):
    if not (filter_col_1 and col_1_val):
        return data[take_col].value_counts(take_col_value)[take_col_value] * 100
    elif filter_col_2 and col_2_val:
        temp = data[data[filter_col_1].isin(col_1_val)]
        temp = temp[temp[filter_col_2] == col_2_val][take_col]
    else:
        temp = data[data[filter_col_1].isin(col_1_val)][take_col]
    try:
        percent = temp.value_counts(take_col_value)[take_col_value] * 100
    except KeyError:
        percent = temp.value_counts(take_col_value)
    return percent


def get_median_mode(grouped_data, mode=None):
    if isinstance(grouped_data.sum(), int64) or isinstance(grouped_data.sum(), int):
        temp = grouped_data
    else:
        temp = grouped_data.sum()
    if mode:
        try:
            return temp.mode()[0]
        except:
            print(temp)
            raise ValueError("Data given is empty fix it.")
    return temp.median()


def get_accuracy(data):
    temp = data[data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Correct', 'Incorrect'])].groupby('Question No.')[
        'Outcome (Correct/Incorrect/Not Attempted)']
    if temp:
        return temp.value_counts('Correct') * 100
    else:
        return dict(Correct=0, Incorrect=0)


def get_hist(data, title, xlbl="", ylbl="", label=None, width=3,
             height=3, threshold=0.1, fontsize=17, left=0.3, right=0.7, top=0.9, hspace=0.2):
    def autolabel(rects, xpos='center'):
        ha = {'center': 'center', 'right': 'left', 'left': 'right'}
        offset = {'center': 0.5, 'right': 0.57, 'left': 0.43}  # x_txt = x + w*off
        for rect in rects:
            ht = rect.get_height()
            ax.text(rect.get_x() + rect.get_width() * offset[xpos], 1.01 * ht,
                    f'{ht:.2f}', ha=ha[xpos], va='bottom')

    if not label:
        label = tuple(range(len(xlbl)))
    file = BytesIO()
    plt.clf()
    fig, ax = plt.subplots(figsize=(width, height))
    fig.patch.set_facecolor((1, 1, 1, 0.01))
    rects = ax.bar(label, data, width=len(xlbl) * threshold, align='edge', color="blue")
    plt.xticks(label, xlbl, fontweight='bold', fontsize=fontsize, horizontalalignment='center', rotation=6)
    plt.rcParams['ytick.labelsize'] = 12
    ax.set_title(title, FONT_FOR_TITLE, y=1.1)
    ax.set_ylabel(ylbl, fontsize=16)
    autolabel(rects)
    plt.margins(None, 0.2)
    fig.subplots_adjust(bottom=0.13, left=left, right=right, top=0.85, hspace=hspace)
    plt.savefig(file, format="png", dpi=140)
    plt.plot()
    plt.close()
    return file
