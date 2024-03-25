import pdfquery
import pandas as pd
import os


class TradeStationAccountStatement(object):
    pdf = None
    path = None
    account_number = None
    firm_salesman = None
    statement_date = None
    tables = None

    def __init__(self, path, *args, **kwargs):
        self.pdf = pdfquery.PDFQuery(path)
        self.path = path
        self.pdf.load()
        self.summary = {
            'account_number': {'id': 'ACCOUNT NUMBER:', 'value': '', 'type': 'LTTextBoxHorizontal', 'split': False},
            'firm_salesman': {'id': 'FIRM / SALESMAN:', 'value': '', 'type': 'LTTextBoxHorizontal', 'split': False},
            'statement_date': {'id': 'STATEMENT DATE:', 'value': '', 'type': 'LTTextBoxHorizontal', 'split': False},
            'beggining': {'id': 'BEGINNING BALANCE', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'commisions': {'id': 'COMMISSIONS', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'clearing': {'id': 'CLEARING FEES', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'exchange': {'id': 'EXCHANGE FEES', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'nfa': {'id': 'NFA FEES', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'total_fees': {'id': 'TOTAL COMMISSIONS & FEES', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'gross_pl': {'id': 'GROSS P/L', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'net_pl': {'id': 'NET PROFIT/LOSS FROM TRADES', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'end_balance': {'id': 'END BALANCE', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'open_trade_equity': {'id': 'OPEN TRADE EQUITY', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
            'total_equity': {'id': 'TOTAL EQUITY', 'value': '', 'type': 'LTTextLineHorizontal', 'split': True},
        }

    def get_lines(self, table):
        query = f':in_bbox("{table[0]}, {table[1]}, {table[2] + 20}, {table[3]}")'
        lines = self.pdf.pq(query)
        new_list = []

        for line in lines:
            columns = [''] * len(table[4])
            for char in line.layout._objs:
                for idx, sep in enumerate(table[4]):
                    if hasattr(char, 'x0'):
                        if idx < len(table[4]) - 1:
                            if char.x0 > sep and char.x1 < table[4][idx + 1]:
                                columns[idx] = (columns[idx] + char._text)
                        else:
                            if char.x0 > sep and char.x1 < table[2] + 20:
                                columns[idx] = (columns[idx] + char._text)
            columns = [i.strip() if type(i) == str else str(i) for i in columns]
            new_list.append(columns)
        print(new_list)
        self.tables = new_list

    def find_tables(self):
        labels = self.pdf.pq('LTTextLineHorizontal:contains("---")')
        end = self.pdf.pq('LTTextLineHorizontal:contains("*US$-SEGREGATED(F1)*")')
        bboxs = []
        for label in labels:
            table = label.layout
            separators = []
            for obj in table._objs:
                if obj._text == ' ':
                    separators.append((obj.x0 + obj.x1)/2)
            bboxs.append([table.x0, table.y0, table.x1, table.y1, separators])
        tables = []
        for idx, bbox in enumerate(bboxs):
            if idx < len(bboxs) - 1:
                tables.append([bbox[0], bboxs[idx + 1][1], bbox[2], bbox[1], bbox[4]])
            else:
                tables.append([bbox[0], end[0].layout.y1, bbox[2], bbox[1], bbox[4]])

        for table in tables:
            self.get_lines(table)

    def find_summary(self):

        for key, value in self.summary.items():
            line = self.pdf.pq(f'{value["type"]}:contains("{value["id"]}")')
            line.reverse()
            print(type(line))
            print(line.__dict__)
            if len(line) > 0:
                for x in line:
                    text = x.text.strip()
                    if text.startswith(value["id"]):
                        data = text.replace(value["id"], "").strip()
                        if value['split']:
                            self.summary[key]['value'] = data.split(" ")[-1]
                        else:
                            self.summary[key]['value'] = data
        print(self.summary)

    def write_to_excel(self, output_path=None):
        path = output_path if output_path else self.path.replace('.pdf', '.xlsx')
        df = pd.DataFrame(self.tables)
        writer = pd.ExcelWriter(f'{path}', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='welcome', index=False)
        writer.close()


def pdf_to_excel(pdf_file_path, excel_file_path):
    statement = TradeStationAccountStatement(pdf_file_path)
    statement.find_summary()
    statement.write_to_excel(excel_file_path)


if __name__ == "__main__":
    directory = 'files'
    output = 'results'
    for filename in os.listdir(directory):
        if '.pdf' in filename:
            f = os.path.join(directory, filename)
            o = os.path.join(output, filename.replace('.pdf', '.xlsx'))
            # checking if it is a file
            if os.path.isfile(f):
                print(f)
                pdf_to_excel(f, o)
