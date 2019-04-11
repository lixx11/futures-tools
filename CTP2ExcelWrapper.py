#!/usr/bin/env python

"""Wrapper of CTP2Excel script.

Usage:
    CTP2ExcelWrapper.py <RAW-DIR> [options]

Options:
    -h --help               Show this screen.
    -s --start-date=<DATE>  Specify start date [default: 19990101].
    -e --end-date=<DATE>    Specify end date [default: NOW].
    -o --output=<folder>    Specify output directory [default: output].
    --ext=<extension>       Specify extension of CTP files [default: csv].
    --TK=<TOKEN>            Specify tushare-token for trading calendar [default: xxli].
    --rebate-file=<FILE>    Specify rebate configuration file [default: rebate.csv].
    --email-conf=<FILE>     Specify email configuration file [default: email.yml].
"""


from docopt import docopt
import os
import sys
from glob import glob
import subprocess
import pandas as pd
import tushare as ts
from datetime import datetime
import yaml


BASE_PATH = os.path.dirname(os.path.abspath(__file__))
TD_FILE = 'td.csv'  # trading dates csv file
COLUMNS = ('日期', '期初结存', '银期出入金', '手续费返还', '利息返还', '中金所申报费', '出入金合计', '平仓盈亏', '盯市盈亏', 
           '手续费', '交割手续费', '中金所手续费', '上期原油手续费', '上期所手续费', '郑商所手续费', '大商所工业品手续费', '大商所农产品手续费', '期末结存',
           '实际盈亏', '实际份额', '实际净值', '即时手续费返还', '即时期末结存', '即时盈亏', '即时份额', '即时净值') # 结算总表


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
    print(companies)
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
            '--rebate-file', argv['--rebate-file'],
            ] + raw_files,
            capture_output=True
        )
        print(res.stdout.decode('utf-8'), res.stderr.decode('utf-8'))
        print('-' * 80)
    
    # 生成汇总报表
    client_files = glob('%s/*/*_%s_%s.xlsx' % (output_dir, start_date, end_date))
    client_data = []
    bf_data = []
    for client_file in client_files:
        data = pd.read_excel(client_file, sheet_name=None)
        if '结算汇总' in data:
            _data = data['结算汇总'].set_index('日期')
            _data.index = _data.index.astype(str)
            client_data.append(_data)
        if '银期转账' in data:
            _data = data['银期转账'].set_index('日期')
            _data.index = _data.index.astype(str)
            bf_data.append(_data)
    # 结算汇总
    dates = sorted(list(set(
        [date for i in range(len(client_data)) for date in client_data[i].index.values[:-1].tolist()]
        )), key=lambda x: int(x))
    total_data = []
    for date in dates:
        row_dict = {'日期': date}
        for field in COLUMNS:
            if field in ('日期', '实际净值', '即时净值'):
                continue
            row_dict[field] = sum([client_data[i].loc[date, field] for i in range(len(client_data))])
        total_data.append(row_dict)
    client_df = pd.DataFrame(total_data)
    client_df['实际净值'] = client_df['期末结存'] / client_df['实际份额']
    client_df['即时净值'] = client_df['即时期末结存'] / client_df['即时份额']
    # 增加合计行
    last_row = client_df.iloc[-1]
    total_row = {
        '日期': '合计',
        '期末结存': last_row['期末结存'],
        '实际盈亏': last_row['实际盈亏'],
        '实际份额': last_row['实际份额'],
        '实际净值': last_row['实际净值'],
        '即时手续费返还': client_df['即时手续费返还'].sum(),
        '即时期末结存': last_row['即时期末结存'],
        '即时盈亏': last_row['即时盈亏'],
        '即时份额': last_row['即时份额'],
        '即时净值': last_row['即时净值'],
        '银期出入金': client_df['银期出入金'].sum(),
        '手续费返还': client_df['手续费返还'].sum(),
        '利息返还': client_df['利息返还'].sum(),
        '中金所申报费': client_df['中金所申报费'].sum(),
        '出入金合计': client_df['出入金合计'].sum(),
        '平仓盈亏': client_df['平仓盈亏'].sum(),
        '盯市盈亏': client_df['盯市盈亏'].sum(),
        '手续费': client_df['手续费'].sum(),
        '交割手续费': client_df['交割手续费'].sum(),
        '中金所手续费': client_df['中金所手续费'].sum(),
        '上期原油手续费': client_df['上期原油手续费'].sum(),
        '上期所手续费': client_df['上期所手续费'].sum(),
        '郑商所手续费': client_df['郑商所手续费'].sum(),
        '大商所工业品手续费': client_df['大商所工业品手续费'].sum(),
        '大商所农产品手续费': client_df['大商所农产品手续费'].sum(),
    }
    client_df = client_df.append(total_row, ignore_index=True)
    # 银期转账
    total_data = []
    for date in dates:
        row_dict = {
            '日期': date,
            '出金': sum([bf_data[i].loc[date, '出金'] for i in range(len(bf_data)) if date in bf_data[i].index]),
            '入金': sum([bf_data[i].loc[date, '入金'] for i in range(len(bf_data)) if date in bf_data[i].index])
        }
        total_data.append(row_dict)
    bf_df = pd.DataFrame(total_data)
    bf_df['date'] = pd.to_datetime(bf_df['日期'])
    bf_df.sort_values(by='date', ascending=False, inplace=True)
    final_summary = os.path.join(output_dir, '结算总表_%s_%s.xlsx' % (start_date, end_date))
    writer = pd.ExcelWriter(final_summary, engine='xlsxwriter')
    client_df.to_excel(writer, '结算汇总', columns=COLUMNS, index=False, freeze_panes=(1, 1))
    # 调整列宽
    workbook = writer.book
    worksheet = writer.sheets['结算汇总']
    for i, col in enumerate(COLUMNS):
            if 10 <= i <= 15:
                worksheet.set_column(i, i, 4)
            else:
                if client_df[col].dtype == 'float64':
                    max_width = client_df[col].apply(lambda x: len(str('%.2f' % x))).max()
                    max_width = max(max_width, len(col))
                    worksheet.set_column(i, i, max_width+1)
                else:
                    max_width = client_df[col].apply(lambda x: len(str(x))).max()
                    max_width = max(max_width, len(col))
                    worksheet.set_column(i, i, max_width+1)
    bf_df.to_excel(writer, '银期转账', index=False, columns=('日期', '入金', '出金'))
    writer.save()
    print('=' * 80)
    print('总结算单已写入%s' % final_summary)
    # 发送邮件
    email_conf = yaml.load(open(argv['--email-conf'], 'r'))
    if email_conf['send_email']:
        result_files = glob('%s/*/*_%s_%s.xlsx' % (output_dir, start_date, end_date))
        result_files += glob('%s/*_%s_%s.xlsx' % (output_dir, start_date, end_date))
        archive_file = '%s_%s.tar' % (start_date, end_date)
        res = subprocess.run([
            'tar', '-cvf', archive_file,
            ] + result_files,
            capture_output=True
        )

        import smtplib
        import mimetypes
        from email.mime.application import MIMEApplication
        from email.mime.multipart import MIMEMultipart
        from email.mime.base import MIMEBase
        from email.utils import COMMASPACE, formatdate
        from email.header import Header
        from email import encoders

        outer  = MIMEMultipart()
        outer ['From'] = email_conf['sender']['account']
        outer ['To'] = COMMASPACE.join(email_conf['recipients'])
        outer ['Date'] = formatdate(localtime=True)
        outer ['Subject'] = Header('结算表', 'utf-8')

        ctype, encoding = mimetypes.guess_type(archive_file)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        with open(archive_file, 'rb') as fp:
            msg = MIMEBase(maintype, subtype)
            msg.set_payload(fp.read())
            encoders.encode_base64(msg)
        msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(archive_file))
        outer.attach(msg)

        server = smtplib.SMTP(email_conf['server'])

        server.login(email_conf['sender']['account'], email_conf['sender']['passwd'])
        server.sendmail(email_conf['sender']['account'], email_conf['recipients'], outer.as_string())
        server.quit()
