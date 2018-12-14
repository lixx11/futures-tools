#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Wrapper of CTP2Excel script.

Usage:
    CTP2ExcelWrapper.py <RAW-DIR> [options]

Options:
    -h --help               Show this screen.
    -o --output=<folder>    Specify output directory [default: output].
"""


from docopt import docopt
import os
from glob import glob
import subprocess
import pandas as pd


BASE_PATH = os.path.dirname(os.path.abspath(__file__))


if __name__ == "__main__":
    argv = docopt(__doc__)
    raw_dir = argv['<RAW-DIR>']
    output_dir = argv['--output']
    print('Read raw data from %s, and write summary to %s' % (raw_dir, output_dir))

    companies = next(os.walk(raw_dir))[1]
    for company in companies:
        raw_files = glob('%s/%s/*/*.txt' % (raw_dir, company))
        company_dir = os.path.join(output_dir, company)
        res = subprocess.call([
            'python', '%s/CTP2Excel.py' % BASE_PATH,
            '-o', company_dir
            ] + raw_files,
        )
        # print(res.stdout.decode('utf-8'), res.stderr.decode('utf-8'))
    
    # 生成汇总报表
    client_files = glob('%s/*/*.xlsx' % output_dir)
    client_data = []
    bf_data = []
    for client_file in client_files:
        data = pd.read_excel(client_file, sheet_name=None)
        if '结算汇总' in data:
            client_data.append(data['结算汇总'])
        if '银期转账' in data:
            bf_data.append(data['银期转账'])
    
    client_df = pd.concat(client_data)
    bf_df = pd.concat(bf_data)
    final_summary = os.path.join(output_dir, 'summary.xlsx')
    writer = pd.ExcelWriter(final_summary)
    client_df.to_excel(writer, '结算汇总', index=False)
    bf_df.to_excel(writer, '银期转账', index=False)
    writer.save()
    print('=' * 80)
    print('Summary saved to %s' % final_summary)