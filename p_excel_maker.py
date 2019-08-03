# Install dependencies:
# python37 -m pip install XlsxWriter

from collections import namedtuple
from typing import * 
from abc import ABC
import xlsxwriter
import argparse

class QuestionType(ABC):
    MILULI = 0
    KAMUTI = 1
    ANGLIT = 2

RowData = namedtuple('RowData', 'number ans correct_ans type')

data = [
    RowData(number=1, ans=1, correct_ans=1, type=QuestionType.KAMUTI),
    RowData(number=2, ans=2, correct_ans=1, type=QuestionType.MILULI),
    RowData(number=3, ans=3, correct_ans=1, type=QuestionType.ANGLIT)
]

class Document:
    def __init__(self, document_name: str):
        self._name = document_name
        self._workbook = xlsxwriter.Workbook("{}.xlsx".format(self._name)) 
        self.sheet = self._workbook.add_worksheet()
        self.sheet.right_to_left()

        self._init_formats()

    def format_of(self, name: str):
        """Get a format from a predefined number of formats, based on name"""
        return self._formats.get(name, None)

    def create_format(self, attr):
        """Create a new format, based on custom attributes"""
        return self._workbook.add_format(attr)

    def get_all_formats(self) -> List[str]:
        return self._formats.keys() 

    def __del__(self) -> None:
        self._close()

    def _init_formats(self) -> None:
        self._formats = {
            'top_header': self.create_format({
                'bg_color': 'blue',
                'font_size': 30
            }),
            'correct_answer': self.create_format({
                'bg_color': 'green',
                'font_size': 20
            }),
            'wrong_answer': self.create_format({
                'bg_color': 'red',
                'font_size': 20
            }),
            'answer_numbers': self.create_format({
                'font_size': 20,
                'align': 'center'
            }),
            'blank_gray': self.create_format({
                'bg_color': 'gray'
            }),
            'regular': self.create_format({
                'font_size': 20
            }),
            QuestionType.ANGLIT: self.create_format({
                'font_size': 20,
                'bg_color': '#FFA8BF',
            }),
            QuestionType.MILULI: self.create_format({
                'font_size': 20,
                'bg_color': '#A8D1FF',
            }),
            QuestionType.KAMUTI: self.create_format({
                'font_size': 20,
                'bg_color': '#F1FFA8',
            }),
            
        }

    def _close(self) -> None:
        self._workbook.close()

def init_header(document: Document) -> None:
    header_format = document.format_of('top_header')
    document.sheet.set_row(0, 40, header_format)
    document.sheet.set_column('A:A', 15)
    document.sheet.write('A1', 'מספר')
    document.sheet.set_column('B:B', 15)
    document.sheet.write('B1', 'ת"ש')
    document.sheet.set_column('C:C', 15)
    document.sheet.write('C1', 'ת"נ')
    document.sheet.set_column('D:D', 20)
    document.sheet.write('D1', 'הצלחתי?')
    document.sheet.set_column('E:E', 100)
    document.sheet.write('E1', 'הסקת מסקנות')

def handle_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('input', help="""
    A file containing data in the following structure:
    <Category>
    <Question Number> <Your answer> <Correct Answer>
    where Category needs to be in the following list: כמותי, מילולי, אנגלית
    And the other parameters be integers
    """)

    parser.add_argument('output', help="The name of the target excel file. The program will create a new one each time.")
    args = parser.parse_args()

    return args

def read_file(filename: str):
    from codecs import open

    with open(filename, encoding='utf-8') as input:
        lines = input.readlines()

    data = []
    current_category = None
    for line in lines:
        line = line.rstrip()
        if line in ['כמותי', 'מילולי', 'אנגלית']:
            current_category = {
                'כמותי': QuestionType.KAMUTI,
                'אנגלית': QuestionType.ANGLIT,
                'מילולי': QuestionType.MILULI
            }[line]
        else:
            n, a, ca = (int(x) for x in line.split()) 
            data.append(RowData(number=n, ans=a, correct_ans=ca, type=current_category))

    return data

def write_conclusions(document: Document, data):
    english = data[QuestionType.ANGLIT]
    miluli = data[QuestionType.MILULI]
    kamuti = data[QuestionType.KAMUTI]
    
    e_p = english[True] / (english[True] + english[False]) * 100
    m_p = miluli[True] / (miluli[True] + miluli[False]) * 100 
    k_p = kamuti[True] / (kamuti[True] + kamuti[False]) * 100

    format = document.format_of('regular')
    document.sheet.write('G1', 'E')
    document.sheet.write('G2', str(int(e_p)) + "%", format)
    document.sheet.write('H1', 'K')
    document.sheet.write('H2', str(int(k_p)) + "%", format)
    document.sheet.write('I1', 'M')
    document.sheet.write('I2', str(int(m_p)) + "%", format)

    return e_p, m_p, k_p

def main():
    args = handle_args()
    data = read_file(args.input)

#    data = sorted(data, key=lambda x: x.number)
    document = Document(args.output)
    init_header(document)

    regular_format = document.format_of('regular')
    counter = {
        QuestionType.ANGLIT: {
            True: 0,
            False: 0
        },
        QuestionType.KAMUTI: {
            True: 0,
            False: 0
        },
        QuestionType.MILULI: {
            True: 0,
            False: 0
        },
    }
    for idx, row in enumerate(data):
        question, ans, correct_ans, question_type = row

        got_correct_answer = (ans == correct_ans)
        counter[question_type][got_correct_answer] += 1
        answer_format = document.format_of('correct_answer') if got_correct_answer else document.format_of('wrong_answer')

        idx = idx + 1
        document.sheet.write(idx, 0, str(question), document.format_of(question_type)) 
        document.sheet.write(idx, 1, str(ans), regular_format) 
        document.sheet.write(idx, 2, str(correct_ans), regular_format) 
        document.sheet.write(idx, 3, str(got_correct_answer), answer_format)
        document.sheet.write(idx, 4, '', None if not got_correct_answer else document.format_of('blank_gray'))

    english, miluli, kamuti = write_conclusions(document, counter)
    print("English: {}%, Miluli: {}%, Kamuti: {}%".format(english, miluli, kamuti))

if __name__ == "__main__":
    try:
        main()
    except PermissionError:
        print("Whoopsie! Looks like the output file is already opened!")
    except Exception as e:
        print("Something went wrong and I have no fucking idea why. look at this:")
        print(e)
