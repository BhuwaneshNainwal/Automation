from scipy.stats import percentileofscore
from time import perf_counter
from Wisd_new.utils import PDF
from Wisd_new.static import PAGE_HEIGHT, PAGE_WIDTH
from Wisd_new.data_operations import get_data, pd
from Wisd_new.page1 import get_page_1
from Wisd_new.page2 import get_page_2, get_page_3, get_page_4, get_page_5, get_page_6
from Wisd_new.data_operations import get_median_mode, get_percent_of_attempted_questions, get_accuracy

STATIC, CONTINUOUS = get_data("../Project/data.xlsx", sheet="1-6")
CONTINUOUS = CONTINUOUS.astype({'Your score': int})
GROUPED = CONTINUOUS.groupby('Student No')

for (_, stu), (_, mark) in zip(STATIC.iterrows(), GROUPED):
    pdf = PDF(dest=f"{stu['Full Name of Candidate']}.pdf")
    pdf.add_page(get_page_1(stu, mark, img_loc="Pics//"))
    all_attempts = []
    all_accuracy = []
    all_scores = []
    total = mark['Your score'].sum()
    MEDIAN = get_median_mode(GROUPED['Your score'])
    MODE = get_median_mode(GROUPED['Your score'], mode=True)
    WHOLE_DATA = CONTINUOUS
    AVERAGE_ACCURACY = get_percent_of_attempted_questions(WHOLE_DATA,
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          'Correct',
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    AVERAGE_ATTEMPTS = get_percent_of_attempted_questions(WHOLE_DATA, 'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    all_accuracy.append(AVERAGE_ACCURACY)
    all_attempts.append(AVERAGE_ATTEMPTS['Incorrect', 'Correct'].sum())
    QUES_ATTEMPTS = (1 - WHOLE_DATA.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Unattempted')) * 100
    rest_accuracy = get_accuracy(WHOLE_DATA)
    rest_score = WHOLE_DATA[['Your score', 'Question No.']].astype({'Your score': int, 'Question No.':str}).groupby('Question No.')['Your score'].sum() / len(WHOLE_DATA.groupby('Student No'))
    all_score = GROUPED['Your score'].sum()
    all_scores.append(all_score.mean())
    MEAN = all_score.mean()
    try:
        m = dict(QUES_ATTEMPTS)
    except:
        m = {'Q1':0}
    if total in all_score:
        percentile = percentileofscore(all_score, total, kind='weak')
    else:
        new_scoreboard = all_score.append(pd.Series([total]))
        percentile = percentileofscore(new_scoreboard, total, kind='weak')
    k = perf_counter()
    pdf.add_page(get_page_2(stu_details=(stu['First Name of Candidate'], stu['Country of Residence']), mean=MEAN,
                            stu_record=mark, median=MEDIAN, mode=MODE, avg_attempts=AVERAGE_ATTEMPTS.sum(),
                            percentile=100 - percentile,
                            avg_accuracy=AVERAGE_ACCURACY, total=total, rest_accuracy=dict(rest_accuracy),
                            rest_score=dict(rest_score), att_group=m))

    print(f"Time taken :{perf_counter() - k}")
    same_country = CONTINUOUS[CONTINUOUS['Country of Residence'] == stu['Country of Residence']]
    WHOLE_DATA = same_country
    MEDIAN = get_median_mode(same_country
                             .groupby('Student No')['Your score'])
    MODE = get_median_mode(same_country
                           .groupby('Student No')['Your score'], mode=True)
    AVERAGE_ACCURACY = get_percent_of_attempted_questions(WHOLE_DATA,
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          'Correct',
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    AVERAGE_ATTEMPTS = get_percent_of_attempted_questions(WHOLE_DATA, 'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    QUES_ATTEMPTS = WHOLE_DATA.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect'])[['Correct','Incorrect']].sum()* 100
    rest_accuracy = get_accuracy(WHOLE_DATA)
    all_score = WHOLE_DATA.groupby('Student No')['Your score'].sum()
    all_accuracy.append(AVERAGE_ACCURACY)
    all_attempts.append(AVERAGE_ATTEMPTS['Incorrect', 'Correct'].sum())
    all_scores.append(WHOLE_DATA['Your score'].sum())
    MEAN = all_score.mean()
    if total in all_score:
        percentile = percentileofscore(all_score, total, kind='weak')
    else:
        new_scoreboard = all_score.append(pd.Series([total]))
        percentile = percentileofscore(new_scoreboard, total, kind='weak')
    rest_score = WHOLE_DATA.groupby('Question No.')['Your score'].sum() / len(WHOLE_DATA.groupby('Student No'))

    pdf.add_page(get_page_3(stu_details=(stu['First Name of Candidate'], stu['Country of Residence']), mean=MEAN,
                            stu_record=mark, median=MEDIAN, mode=MODE, avg_attempts=AVERAGE_ATTEMPTS.sum(),
                            percentile=100 - percentile,
                            avg_accuracy=AVERAGE_ACCURACY, total=total, rest_accuracy=dict(rest_accuracy),
                            rest_score=dict(rest_score), att_group=m))

    print(f"Time taken :{perf_counter() - k}")
    WHOLE_DATA = GROUPED.sum().sort_values(by='Your score', ascending=False).reset_index()
    WHOLE_DATA = WHOLE_DATA[:len(WHOLE_DATA) // 10]
    OTHERS = CONTINUOUS[CONTINUOUS['Student No'].isin(WHOLE_DATA['Student No'])]
    MEDIAN = get_median_mode(WHOLE_DATA['Your score'])
    MODE = get_median_mode(WHOLE_DATA['Your score'], mode=True)
    AVERAGE_ACCURACY = get_percent_of_attempted_questions(OTHERS,
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          'Correct',
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    AVERAGE_ATTEMPTS = get_percent_of_attempted_questions(OTHERS, 'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    QUES_ATTEMPTS = OTHERS.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect'])[['Correct','Incorrect']].sum()* 100
    rest_accuracy = get_accuracy(OTHERS)
    all_score = OTHERS.groupby('Student No')['Your score'].sum()
    MEAN = all_score.mean()
    if total in all_score:
        percentile = percentileofscore(all_score, total, kind='weak')
    else:
        new_scoreboard = all_score.append(pd.Series([total]))
        percentile = percentileofscore(new_scoreboard, total, kind='weak')
    all_accuracy.append(AVERAGE_ACCURACY)
    all_attempts.append(AVERAGE_ATTEMPTS['Incorrect', 'Correct'].sum())
    all_scores.append(all_score.mean())
    rest_score = OTHERS.groupby('Question No.')['Your score'].sum() / len(OTHERS.groupby('Student No'))
    k = perf_counter()
    pdf.add_page(get_page_4(stu_details=(stu['First Name of Candidate'], stu['Country of Residence']), mean=MEAN,
                            stu_record=mark, median=MEDIAN, mode=MODE, avg_accuracy=AVERAGE_ACCURACY,
                            avg_attempts=AVERAGE_ATTEMPTS.sum(), percentile=100 - percentile, total=total,
                            rest_accuracy=dict(rest_accuracy), rest_score=dict(rest_score),
                            att_group=m))
    print(perf_counter() - k)
    WHOLE_DATA = same_country.groupby('Student No').sum().sort_values(by="Your score", ascending=False).reset_index()
    WHOLE_DATA = WHOLE_DATA[:int(len(WHOLE_DATA) // 10) or 10]
    OTHERS = CONTINUOUS[CONTINUOUS['Student No'].isin(WHOLE_DATA['Student No'])]
    MEDIAN = get_median_mode(WHOLE_DATA['Your score'])
    MODE = get_median_mode(WHOLE_DATA['Your score'], mode=True)
    AVERAGE_ACCURACY = get_percent_of_attempted_questions(OTHERS,
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          'Correct',
                                                          'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    AVERAGE_ATTEMPTS = get_percent_of_attempted_questions(OTHERS, 'Outcome (Correct/Incorrect/Not Attempted)',
                                                          ['Incorrect', 'Correct'])
    QUES_ATTEMPTS = OTHERS.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect'])[['Correct','Incorrect']].sum()* 100
    rest_accuracy = get_accuracy(OTHERS)
    all_score = OTHERS.groupby('Student No')['Your score'].sum()
    all_accuracy.append(AVERAGE_ACCURACY)
    all_attempts.append(AVERAGE_ATTEMPTS['Incorrect', 'Correct'].sum())
    all_scores.append(OTHERS['Your score'].sum())
    MEAN = all_score.mean()
    if total in all_score:
        percentile = percentileofscore(all_score, total, kind='weak')
    else:
        new_scoreboard = all_score.append(pd.Series([total]))
        percentile = percentileofscore(new_scoreboard, total, kind='weak')
    rest_score = OTHERS.groupby('Question No.')['Your score'].sum() / len(OTHERS.groupby('Student No'))
    k = perf_counter()
    pdf.add_page(get_page_5(stu_details=(stu['First Name of Candidate'], stu['Country of Residence']), mean=MEAN,
                            stu_record=mark, median=MEDIAN, mode=MODE, avg_accuracy=AVERAGE_ACCURACY,
                            avg_attempts=AVERAGE_ATTEMPTS.sum(), percentile=100 - percentile, total=total,
                            rest_accuracy=dict(rest_accuracy), rest_score=dict(rest_score),
                            att_group=m))

    print(f"Time taken :{perf_counter() - k}")
    all_scores.append(total)
    try:
        accc = mark[mark['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])]['Incorrect', 'Correct'].value_counts().sum()
    except KeyError:
        accc = 0
    print(accc)
    all_attempts.append(accc)
    all_accuracy.append(get_percent_of_attempted_questions(mark, 'Outcome (Correct/Incorrect/Not Attempted)',
                                                           'Correct', 'Outcome (Correct/Incorrect/Not Attempted)',
                                                           ['Incorrect', 'Correct']))
    k = perf_counter()
    print(all_attempts)
    pdf.add_page(
        get_page_6(all_attempts, all_accuracy, all_scores, (stu['First Name of Candidate'], stu['Country of Residence'])))
    pdf.prepare()
    print(f"Time taken :{perf_counter() - k}")
    break
