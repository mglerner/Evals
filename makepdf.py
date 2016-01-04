#!/usr/bin/env python
from __future__ import division, print_function
import argparse,sys,os

from collections import Counter
from matplotlib import pyplot as plt
import wordcloud
from wordcloud import WordCloud
import pandas as pd, numpy as np
from IPython.display import display, HTML 
from weasyprint import HTML as weasyHTML


parser = argparse.ArgumentParser()
parser.add_argument('xlfilename',help='Name of the excel file to use')
parser.add_argument('-V','--verbose',help='Be extra verbose',action='store_true',default=False)
parser.add_argument('-W','--include-word-count',help='Include a table with the word count',action='store_true',default=False)
parser.add_argument('-s','--stop-words',nargs='+',type=str,help='Extra stop words, i.e. words NOT to include in the wordcloud and word count. E.g. -s class course lab')

args = parser.parse_args()

if not args.xlfilename.endswith('xlsx'):
    sys.exit('You must specify a .xlsx file, likely downloaded from Moodle')

stopwords = wordcloud.STOPWORDS
if args.stop_words is not None:
    stopwords = stopwords.union(args.stop_words)

def is_nan(x): 
    try: return np.isnan(x) 
    except: return False #isnan only eats strings

def reformat_answer(answer):
    if is_nan(answer):
        answer = 'No answer given.'
    else:
        if type(answer) in (int, float):
            answer = str(answer)
    return answer
    
css = """
p.large-headline {
    font-family: times, Times New Roman, times-roman, georgia, serif;
    color: #444;
    margin: 0px 0px 100px 0px;
    padding: 40px 40px 40px 40px;
    font-size: 55px;
    line-height: 44px;
    letter-spacing: -1px;
    font-weight: bold;
    text-align: center;
    border-radius: 25px;
    border: 2px solid #111;
    width: 90%;
}

p.medium-headline {
    font-family: times, Times New Roman, times-roman, georgia, serif;
    color: #444;
    margin: 0px -10px 0px 0px;
    padding: 0px 0px 0px 0px;
    font-size: 25px;
    line-height: 24px;
    letter-spacing: -1px;
    font-weight: bold;
    text-align: left;
}

p.name {    
    font-family: times, Times New Roman, times-roman, georgia, serif;
    font-weight: bold;
    font-size: 20px;
    margin-top: 2em;
    margin-bottom: 0em;
}
p.question {
    font-family: times, Times New Roman, times-roman, georgia, serif;
    font-size: 16px;
    color: #111;
    font-weight: bold;
    margin-top: 0em;
    margin-bottom: 0em;
    width: 90%;
}
p.answer {
    font-family: times, Times New Roman, times-roman, georgia, serif;
    font-size: 14px;
    color: #111;
    text-align: justify;
    margin-top: 0em;
    margin-bottom: 0em;
    width: 90%;
}
table
{
  border-collapse: collapse;
}
th
{
  color: #ffffff;
  background-color: #000000;
}
td
{
  background-color: #cccccc;
}
table, th, td
{
  font-family:Arial, Helvetica, sans-serif;
  border: 1px solid black;
  text-align: right;
}
"""

def generatepdf(xl_filename,removeintermediate=False,verbose=False,include_word_count=False):
    pdf_filename = os.path.splitext(xl_filename)[0] + '.pdf'
    html_filename = os.path.splitext(xl_filename)[0] + '.html'
    wc_filename = os.path.splitext(xl_filename)[0] + '-wordcloud.png'

    print("I will write out the following files: {p} {h} {w}".format(p=pdf_filename,
                                                                     h=html_filename,
                                                                     w=wc_filename))
    if removeintermediate:
        print("But I will delete {h} and {w}.".format(h=html_filename,
                                                      w=wc_filename))

    answers = pd.io.excel.read_excel(xl_filename,sheetname='RawData')
    questionmap = pd.io.excel.read_excel(xl_filename,sheetname='QuestionMapper')

    # We want a per-student list of questions and answers. My first
    # thought is to stick everything into a dictionary. We want to
    # make sure to return the results in the correct order, so we
    # could use an ordered dict. I think it's easier just to keep an
    # ordered list of questions.
    questions = questionmap["Question"].values

    # About the below code:
    #
    # When we iterate through the rows, `idx` is the number of the
    # row, and `qd` comes to us as the "question dictionary" where row
    # 1 is expected to name the columns, and we can then look up
    # entries by name. For example, column A happens to be "Column"
    # and column B is "Question", so asking for `qd['Question']` gets
    # the thing in column B.
    #
    # `qm` is then my "question map": it maps something like "Question
    # 1" to "What were the most positive features of this course"
    qm = {}
    for (idx,qd) in questionmap.iterrows():
        qn = qd['Column'].replace(' ','_')
        qt = qd['Question']
        qm[qn] = qt

    if verbose:
        print(qm)

    # Now let's grab the data that should be common to all rows
    path = answers.Path[0]
    course_code = answers.CourseCode[0]
    course_title = answers.CourseTitle[0]
    instructor_name = answers.InstructorName[0]
    enrollments = answers.Enrollments[0]
    # We know we're not extracting the following from each row, so keep quiet about it later.
    knownskips = ['Path','CourseCode','CourseTitle','UniqueID','InstructorName','Enrollments']

    # And now let's slurp up the data per student.
    a = {}
    for (idx,student) in answers.iterrows():
        a[idx] = {}
        for colname in answers.columns:
            col_name = colname.replace(' ','_')
            if col_name in qm:
                #print("Looking up",col_name)
                a[idx][qm[col_name]] = student[colname]
            else:
                if colname not in knownskips:
                    print("Could not find",colname)


    # Now we're ready to stamp out the text, believe it or not. The
    # only cute thing is that `pandas` uses nan ("not a number") to
    # represent missing data. We'll use `numpy` (imported above as
    # `np`) to test for nan, and turn it into "No answer given."

    if verbose:
        print(a[0][questions[0]])

    html = '''
<html>
<head>
<style>
{css}
</style>
</head>
<body>
<div>
<p class="large-headline">{title}</p>
<p class="medium-headline">{code}<br>{instructor}<br>Answers from {a} of {b} enrolled students</p>
<div>'''.format(css=css,
                title=course_title,code=course_code,instructor=instructor_name,
                a=len(a),b=enrollments
    )

    answertext = ''
        
    for idx in sorted(a):
        html += '''<div class="response">
        <p class="name">Student {i} ({n})</p>
        '''.format(
                i=idx+1, n=reformat_answer(a[idx][questions[8]])
            )
        # NOTE: the above line assumes that the name is question 9, index 8.
        
        # NOTE: the below line loops over all of the questions,
        # including the name. That's repeated information, which we
        # could likely remove. We used to go through question[:-1],
        # but the name comes in before custom questions that the
        # instructor can choose to add. So, if we want to remove the
        # name question, we'll need to be smart about matching it
        # above and below here. For now, this is easiest.
        for question in questions:
            question_cor_name = question.replace("[InstructorName]", instructor_name)
            if u'\xa0' in question_cor_name: 
                question_cor_name = question_cor_name.replace(u'\xa0', u' ') #Corrects for unicode encoding error
            answer = reformat_answer(a[idx][question])
            answertext = answertext + ' ' + answer
            html += '''<p class="question">{q}</p>
            <p class="answer">{a}</p>
            '''.format(q=question_cor_name,a=answer)
        html += '</div>\n'

    wordcloud = WordCloud(
                      font_path='Fonts/Raleway-Bold.ttf',
                      stopwords=stopwords,
                      background_color='white',
                      width=1800,
                      height=1400
                     ).generate(answertext)
    plt.imshow(wordcloud)
    plt.axis('off')
    plt.savefig(wc_filename, dpi=300)

    c = Counter([i for i in answertext.lower().split() if i not in stopwords])

    html += '''<div class="content-analysis"><img src="{wc}" style="width:720px;height:560px;"/>'''.format(wc=os.path.split(wc_filename)[-1])

    if include_word_count:
        html += '''<table>
          <caption>Most common words</caption>
          <tr><th>Word</th><th>Count</th></tr>
        '''
        for (w,n) in c.most_common(20):
            html += '<tr><td>{w}</td><td>{n}</td></tr>\n'.format(w=w,n=n)
        html += '''</table>'''
    html += '''
    </div>
    </body>
    </html>
    '''
    f = open(html_filename,'w')
    f.write(html)
    f.close()
    weasyHTML(html_filename).write_pdf(pdf_filename)

        
generatepdf(xl_filename=args.xlfilename, verbose=args.verbose, include_word_count=args.include_word_count)
