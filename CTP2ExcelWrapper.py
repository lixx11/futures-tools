#!/usr/bin/env python

"""Wrapper of CTP2Excel script.

Usage:
    CTP2ExcelWrapper.py <RAW-DIR> [options]

Options:
    -h --help               Show this screen.
    -e --ext=<extension>    Specify extension of CTP files [default: txt].
    -o --output=<folder>    Specify output directory [default: output].
    --start-date=<DATE>     Specify start date [default: 19990101].
    --end-date=<DATE>       Specify end date [default: NOW].
    --TK=<TOKEN>            Specify tushare-token for trading calendar [default: xxli].
    --CFFEX-return=<NUM>    Specify return factor of CFFEX commission [default: 0.3].
    --INE-return=<NUM>      Specify return factor of INE commission [default: 0.3].
    --SHFE-return=<NUM>     Specify return factor of SHFE commission [default: 0.3].
    --CZCE-return=<NUM>     Specify return factor of CZCE commission [default: 0.3].
    --DCE-IND-return=<NUM>  Specify return factor of DCE industrial products commission [default: 0.3].
    --DCE-AGR-return=<NUM>  Specify return factor of DCE agricultural products [default: 0.3].
"""


from docopt import docopt
import os
import sys
from glob import glob
import subprocess
import pandas as pd
import tushare as ts
from datetime import datetime


BASE_PATH = os.path.dirname(os.path.abspath(__file__))
TD_FILE = 'td.csv'  # trading dates csv file


if __name__ == "__main__":
    argv = docopt(__doc__)
    raw_dir = argv['<RAW-DIR>']
    ext = argv['--ext']
    output_dir = argv['--output']
    start_date = argv['--start-date']
    end_date = argv['--end-date']
    tk = argv['--TK']
    if argv['--end-date'] == 'NOW':
        end_date = datetime.now().strftime('%Y%m%d')
    print('设定起始日期为%s，结束日期为%s' % (start_date, end_date))
    # 获取交易日历
    if tk == 'xxli':
        if os.path.exists('./tk.csv'):
            with open('./tk.csv', 'r') as f:
                contents = f.readlines()
            tk = contents[1].strip()
        else:
            print('无法找到默认的TOKEN文件！')
            sys.exit()
    ts.set_token(tk)
    pro = ts.pro_api()
    cal_df = pro.trade_cal(exchange='', start_date=start_date, end_date=end_date)
    cal_df = cal_df[cal_df['is_open'] == 1]
    cal_df.to_csv(TD_FILE, index='False')
    print('从%s获取原始CTP文件并将总结算单写入%s' % (raw_dir, output_dir))
    print('=' * 80)
    companies = next(os.walk(raw_dir))[1]
    for company in companies:
        print('处理%s期货公司数据' % company)
        raw_files = glob('%s/%s/*/*.%s' % (raw_dir, company, ext))
        if len(raw_files) == 0:
            print('WARNING! 未找到%s内的结算单文件，请检查文件后缀以及文件目录结构是否符合标准！' % company)
            continue
        company_dir = os.path.join(output_dir, company)
        res = subprocess.run([
            'python', '%s/CTP2Excel.py' % BASE_PATH,
            '-o', company_dir,
            '--start-date', start_date,
            '--end-date', end_date,
            '--TD', TD_FILE,
            '--TK', tk,
            '--CFFEX-return', argv['--CFFEX-return'],
            '--INE-return', argv['--INE-return'],
            '--SHFE-return', argv['--SHFE-return'],
            '--CZCE-return', argv['--CZCE-return'],
            '--DCE-IND-return', argv['--DCE-IND-return'],
            '--DCE-AGR-return', argv['--DCE-AGR-return'],
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
    
    client_df = pd.concat(client_data, sort=False)
    bf_df = pd.concat(bf_data)
    final_summary = os.path.join(output_dir, '结算总表_%s_%s.xlsx' % (start_date, end_date))
    writer = pd.ExcelWriter(final_summary)
    client_df.to_excel(writer, '结算汇总', index=False)
    bf_df.to_excel(writer, '银期转账', index=False)
    writer.save()
    print('=' * 80)
    print('总结算单已写入%s' % final_summary)