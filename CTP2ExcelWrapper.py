#!/usr/bin/env python

"""Wrapper of CTP2Excel script.

Usage:
    CTP2ExcelWrapper.py <RAW-DIR> [options]

Options:
    -h --help               Show this screen.
    -e --ext=<extension>    Specify extension of CTP files [default: txt].
    -o --output=<folder>    Specify output directory [default: output].
    --start-date=<DATE>     Specify start date [default: 19900101].
    --end-date=<DATE>       Specify end date [default: NOW].
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
    ext = argv['--ext']
    output_dir = argv['--output']
    start_date = argv['--start-date']
    end_date = argv['--end-date']
    print('从%s获取原始CTP文件并将总结算单写入%s' % (raw_dir, output_dir))
    print('=' * 80)
    companies = next(os.walk(raw_dir))[1]
    for company in companies:
        raw_files = glob('%s/%s/*/*.%s' % (raw_dir, company, ext))
        if len(raw_files) == 0:
            print('WARNING! 未找到%s结算单文件，请检查文件后缀以及文件目录结构是否符合标准！' % company)
            continue
        company_dir = os.path.join(output_dir, company)
        res = subprocess.run([
            'python', '%s/CTP2Excel.py' % BASE_PATH,
            '-o', company_dir,
            '--start-date', start_date,
            '--end-date', end_date,
            ] + raw_files,
            capture_output=True
        )
        print(res.stdout.decode('utf-8'), res.stderr.decode('utf-8'))
        print('-' * 80)
    
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
    print('总结算单已写入%s' % final_summary)