#!/usr/bin/env python

"""Generate Excel summary from multiple CTP files.

Usage:
    CTP2Excel.py <CTP-file>... [options]

Options:
    -h --help               Show this screen.
    -o --output=<filename>  Specify output filename [default: 'output.xlsx'].
"""
from docopt import docopt


TABLES = ('资金状况', '成交记录', '出入金明细', '平仓明细', '持仓明细', '持仓汇总')

def extract_data(filepath):
    with open(filepath) as f:
        contents = f.readlines()
    # split contents into blocks
    row_id = 0
    block_id = 0
    block_prev = ''
    for i in range(len(contents)):
        for TABLE in TABLES:
            if TABLE in contents[i]:
                block_content = contents[row_id:i]
                if block_id == 0:
                    process_head(block_content)
                elif block_prev == '资金状况':
                    process_summary(block_content)
                elif block_prev == '出入金明细':
                    process_deposit_withdrawal(block_content)
                block_id += 1
                row_id = i
                block_prev = TABLE
                print('found table %s at line %d' % (TABLE, i))


def process_head(content):
    client_id = content[4].split('    ')
    print(''.join(content))
    print(client_id)


def process_summary(content):
    pass

def process_deposit_withdrawal(content):
    pass

if __name__ == "__main__":
    argv = docopt(__doc__)
    CTP_files = argv['<CTP-file>']
    print('Generate summary from %s' % CTP_files)
    for CTP_file in CTP_files:
        print('Processing %s' % CTP_file)
        res = extract_data(CTP_file)