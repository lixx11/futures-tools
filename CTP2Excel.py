#!/usr/bin/env python

"""Generate Excel summary from multiple CTP files.

Usage:
    CTP2Excel.py <CTP-file>... [options]

Options:
    -h --help               Show this screen.
    -o --output=<folder>    Specify output directory [default: output].
    -s --start-date=<DATE>  Specify start date [default: 19900101].
    -e --end-date=<DATE>    Specify end date [default: NOW].
    --TD=<FILE>             Specify trading dates file.
    --TK=<TOKEN>            Specify tushare-token for trading calendar [default: xxli].
    --rebate-file=<FILE>    Specify return configuration file [default: rebate.csv].
"""


from docopt import docopt
import numpy as np
import pandas as pd
import os
import sys
from datetime import datetime, timedelta
import tushare as ts
import matplotlib.pyplot as plt
plt.style.use('ggplot')


TABLES = ('资金状况', '成交记录', '出入金明细', '平仓明细', '持仓明细', '持仓汇总')
COLUMNS = ('账户', '日期', '期初结存', '银期出入金', '手续费返还', '利息返还', '中金所申报费', '出入金合计', '平仓盈亏', '盯市盈亏', 
           '手续费', '交割手续费', '中金所手续费', '上期原油手续费', '上期所手续费', '郑商所手续费', '大商所工业品手续费', '大商所农产品手续费', '期末结存',
           '实际盈亏', '实际份额', '实际净值', '即时手续费返还', '即时期末结存', '即时盈亏', '即时份额', '即时净值') # 结算总表
COMMISSION = (
    ('CFFEX_COMM', ('IF', 'IC', 'IH', 'TS', 'TF', 'T')),  # 中金所
    ('INE_COMM', ('SC',)),  # 上期原油
    ('SHFE_COMM', ('AG', 'AL', 'AU', 'BU', 'CU', 'FU', 'HC', 'NI', 'PB', 'RB', 'RU', 'SN', 'SP', 'WR', 'ZN')),  # 上期所15个品种
    ('CZCE_COMM', ('AP', 'CF', 'CY', 'FG', 'JR', 'LR', 'MA', 'OI', 'PM', 'RI', 'RM', 'RS', 'SF', 'SM', 'SR', 'TA', 'WH', 'ZC')), # 郑商所18个品种
    ('DCE_IND_COMM', ('BB', 'FB', 'I', 'J', 'JM', 'L', 'PP', 'V', 'EG')), # 大商所9个工业品品种
    ('DCE_AGR_COMM', ('A', 'B', 'C', 'CS', 'JD', 'M', 'P', 'Y')) # 大商所8个农产品品种
)
EPSILON = 0.001  # 匹配误差容忍度


def extract_data(filepath):
    with open(filepath) as f:
        contents = f.readlines()
    row_id = 0
    block_prev = ''
    stats = {}
    for i in range(len(contents)):
        for TABLE in TABLES:
            if TABLE in contents[i]:
                block_content = contents[row_id:i]
                if row_id == 0:
                    stats = {**stats, **process_head(block_content)}
                elif block_prev == '资金状况':
                    stats = {**stats, **process_summary(block_content)}
                elif block_prev == '出入金明细':
                    _stats = process_deposit_withdrawal(block_content, stats['company'])
                    if abs(stats['total_deposit_withdrawal'] - _stats['total_deposit_withdrawal']) > EPSILON:
                        print('WARNING! 资金状况出入金(%.2f)与出入金明细不匹配(%.2f)！' 
                              % (stats['total_deposit_withdrawal'], _stats['total_deposit_withdrawal']))
                    stats = {**stats, **_stats}
                elif block_prev == '成交记录':
                    stats = {**stats, **process_transaction(block_content)}
                row_id = i
                block_prev = TABLE
    # process last table
    block_content = contents[row_id:]
    if block_prev == '资金状况':
        stats = {**stats, **process_summary(block_content)}
    elif block_prev == '出入金明细':
        _stats = process_deposit_withdrawal(block_content, stats['company'])
        if abs(stats['total_deposit_withdrawal'] - _stats['total_deposit_withdrawal']) > EPSILON:
            print('WARNING! 资金状况出入金(%.2f)与出入金明细不匹配(%.2f)！' 
                % (stats['total_deposit_withdrawal'], _stats['total_deposit_withdrawal']))
        stats = {**stats, **_stats}
    elif block_prev == '成交记录':
        stats = {**stats, **process_transaction(block_content)}
    return stats


def process_head(content):
    """
    处理结算单表头信息，获取期货公司、用户id、用户姓名以及日期
    """
    for i in range(len(content)):
        if '客户号' in content[i]:
            client_row = i
        if '日期' in content[i]:
            date_row = i
    if '国投安信' in content[0]:
        company = '国投安信'
    elif '兴证期货' in content[0]:
        company = '兴证期货'
    elif '方正中期' in content[0]:
        company = '方正中期'
    else:
        company = ''
    client_id = content[client_row].split('：')[1].strip().split()[0]
    client_name = content[client_row].split('：')[2].strip()
    date = content[date_row].split('：')[-1].strip()
    stats = {
        'company': company,
        'client_id': client_id,
        'client_name': client_name,
        'date': date
    }
    return stats


def process_summary(content):
    """
    处理资金状况，获取期初结存、出入金、平仓盈亏、盯市盈亏、手续费和期末结存等数据。
    """
    for i in range(len(content)):
        if '期初结存' in content[i]:
            balance_bf_row = i
        if '出 入 金' in content[i]:
            total_deposit_withdrawal_row = i
        if '平仓盈亏' in content[i]:
            realized_pl_row = i
        if '盯市盈亏' in content[i]:
            mtm_pl_row = i
        if '手 续 费' in content[i]:
            commission_row = i
        if '交割手续费' in content[i]:
            delivery_row = i
        if '期末结存' in content[i]:
            balance_cf_row = i
    balance_bf = float(content[balance_bf_row].split('：')[1].strip().split()[0])  # 期初结存
    total_deposit_withdrawal = float(content[total_deposit_withdrawal_row].split('：')[1].strip().split()[0])  # 总出入金
    realized_pl = float(content[realized_pl_row].split('：')[1].strip().split()[0])  # 平仓盈亏
    mtm_pl = float(content[mtm_pl_row].split('：')[1].strip().split()[0])  # 盯市盈亏
    commission = -float(content[commission_row].split('：')[1].strip().split()[0])  # 手续费
    delivery_fee = -float(content[delivery_row].split('：')[1].strip().split()[0])  # 交割手续费
    balance_cf = float(content[balance_cf_row].split('：')[-1].strip())  # 期末结存
    total_delta = total_deposit_withdrawal + realized_pl + mtm_pl + commission  # 当期总流水
    if abs(balance_bf + total_delta - balance_cf) > EPSILON:  # 检查期初结存+总流水与期末结存是否匹配
        print('WARNING! 期初结存(%.2f) + 当期总流水(%.2f) = %.2f != 期末结存(%.2f)'
              % (balance_bf, total_delta, balance_bf + total_delta, balance_cf))
    stats = {
        'balance_bf': balance_bf,
        'total_deposit_withdrawal': total_deposit_withdrawal,
        'realized_pl': realized_pl,
        'mtm_pl': mtm_pl,
        'commission': commission,
        'delivery_fee': delivery_fee,
        'balance_cf': balance_cf,
    }
    return stats


def process_deposit_withdrawal(content, company):
    """
    处理出入金，获取各笔出入金类型与金额。
    """
    dw_array = []
    dash_rows = [i for i in range(len(content)) if (len(content[i].strip()) != 0 and len(content[i].strip('-\n')) == 0)]
    for i in range(dash_rows[1] + 1, dash_rows[2]):
        if len(content[i].strip()) == 0:  # skip empty row
            continue
        date, dw_type, deposit, withdrawal, comment  = content[i][1:-2].split('|')
        comment = comment.strip()
        if '中金所申报费' in comment:
            dw_type = '中金所申报费'
        if '手续费减收' in comment:
            dw_type = '手续费返还'
        if '利息' in comment:
            dw_type = '利息返还'
        if len(comment) == 0 and company == '国投安信':
            dw_type = '银期转账'
        if '上海招行' in comment and company == '兴证期货':
            dw_type = '银期转账'
        if '手续费抵免' in comment and company == '方正中期':
            dw_type = '手续费返还'
        dw_array.append({
            'date': date.strip(),
            'dw_type': dw_type.strip(),
            'deposit': float(deposit),
            'withdrawal': float(withdrawal),
        })
    for i in range(dash_rows[2], dash_rows[3]):
        if '|' in content[i]:
            total_row = i
    _, _, total_deposit, total_withdrawal, _ = content[total_row][1:-2].split('|')
    total_deposit = float(total_deposit)
    total_withdrawal = float(total_withdrawal)
    if abs(sum([dw_array[i]['deposit'] for i in range(len(dw_array))]) - total_deposit) > EPSILON:
        print('WARNING! 入金不匹配，请检查出入金表') 
    if abs(sum([dw_array[i]['withdrawal'] for i in range(len(dw_array))]) - total_withdrawal) > EPSILON:
        print('WARNING! 出金不匹配，请检查出入金表')
    total_deposit_withdrawal = total_deposit - total_withdrawal
    stats = {
        'total_deposit_withdrawal': total_deposit_withdrawal,
        'dw_array': dw_array
    }
    return stats


def process_transaction(content):
    """
    处理成交记录，获取手续费金额，包括中金所、上期所、上期原油、郑商所、大商所工业品和大商所农产品共6类。
    """
    dash_rows = [i for i in range(len(content)) if (len(content[i].strip()) != 0 and len(content[i].strip('-\n')) == 0)]
    stats = {}
    for commission in COMMISSION:
        stats[commission[0]] = 0.
    for i in range(dash_rows[1] + 1, dash_rows[2]):
        if len(content[i].strip()) == 0:  # skip empty row
            continue
        _, _, _, instrument, _, _, _, _, _, _, fee, _, _, _ = content[i][1:-2].split('|')
        instrument = instrument.strip().upper()
        fee = float(fee.strip())
        for commission in COMMISSION:
            for symbol in commission[1]:
                if ''.join([c for c in instrument if not c.isdigit()]) == symbol:
                    stats[commission[0]] -= fee
    return stats


def prev_trading_date(all_trading_dates, date):
    """
    获取所给日期前一个交易日
    """
    all_trading_dates.sort(
        key=lambda d: datetime.strptime(d, '%Y%m%d')
    )
    idx = all_trading_dates.index(date)
    return all_trading_dates[idx - 1]


def parse_rebate_conf(path):
    """
    解析手续费返还配置文件
    """
    conf_df = pd.read_csv(path)
    rebate_conf = []
    for i in range(len(conf_df)):
        start_date = str(int(conf_df.iloc[i]['start_date']))
        end_date = str(int(conf_df.iloc[i]['end_date']))
        dates = pd.date_range(start=start_date, end=end_date, closed='left')
        _rebate_conf = pd.DataFrame({
            'CFFEX': conf_df.iloc[i]['CFFEX'],
            'INE': conf_df.iloc[i]['INE'],
            'SHFE': conf_df.iloc[i]['SHFE'],
            'CZCE': conf_df.iloc[i]['CZCE'],
            'DCE_IND': conf_df.iloc[i]['DCE_IND'],
            'DCE_AGR': conf_df.iloc[i]['DCE_AGR']
        }, index=dates)
        rebate_conf.append(_rebate_conf)
    rebate_df = pd.concat(rebate_conf, verify_integrity=True)
    rebate_df.index = rebate_df.index.map(lambda x: x.strftime('%Y%m%d'))
    return rebate_df


if __name__ == "__main__":
    argv = docopt(__doc__)
    CTP_files = argv['<CTP-file>']
    rebate_file = argv['--rebate-file']
    rebate_df = parse_rebate_conf(rebate_file)
    start_date = datetime.strptime(argv['--start-date'], '%Y%m%d')
    if argv['--end-date'] == 'NOW':
        end_date = datetime.now()
    else:
        end_date = datetime.strptime(argv['--end-date'], '%Y%m%d')
    # 获取交易日历：若指定TD文件，直接读取；否则从tushare调取数据
    td = argv['--TD']
    tk = argv['--TK']
    if td is None:
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
        cal_df = pro.trade_cal(
            exchange='', 
            start_date=start_date.strftime('%Y%m%d'), 
            end_date=end_date.strftime('%Y%m%d')
        )
        cal_df = cal_df[cal_df['is_open'] == 1]
    else:
        cal_df = pd.read_csv(td)
    trading_dates = list(map(str, cal_df['cal_date'].values))
    all_stats = []
    all_dates = []
    
    # 移除不在指定日期范围内的CTP文件（根据文件名）
    CTP_files = [CTP_file for CTP_file in CTP_files 
        if start_date <= datetime.strptime(os.path.basename(CTP_file).split('.')[0].split('_')[-1], '%Y%m%d') <= end_date]
    if len(CTP_files) == 0:
        print('指定日期内无CTP文件，跳过。')
        sys.exit()

    CTP_files.sort(key=lambda fpath: datetime.strptime(os.path.basename(fpath).split('.')[0].split('_')[-1], '%Y%m%d'))
    print('正在处理CTP文件（共%d个文件）...' % len(CTP_files))
    for i, CTP_file in enumerate(CTP_files):
        print('处理第%d个CTP文件： %s' % (i, CTP_file))
        stats = extract_data(CTP_file)
        if stats['date'] in all_dates:
            print('跳过重复的CTP文件：%s' % CTP_file)
            continue
        all_dates.append(stats['date'])
        all_stats.append(stats)
    # 处理指定日期范围内完全无交易的情况
    if len(all_stats) == 0:
        all_stats.append(stats)
    client_ids = np.unique([stats['client_id'] for stats in all_stats])
    if len(client_ids) > 1:
        print('WARNING! 发现超过1个账号结算单，仅处理%s账号文件' % client_ids[0])
    client_id = client_ids[0]

    # 数据汇总输出到Excel文件中
    output_dir = argv['--output']
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
    client_stats =[stats for stats in all_stats if stats['client_id'] == client_id]
    client = '%s-%s' % (client_stats[0]['client_id'], client_stats[0]['client_name'])
    client_data = []
    for stats in client_stats:
        row_dict = {}
        row_dict['账户'] = client
        row_dict['日期'] = stats['date']
        row_dict['date'] = datetime.strptime(stats['date'], '%Y%m%d')
        row_dict['期初结存'] = stats['balance_bf']
        row_dict['出入金合计'] = stats['total_deposit_withdrawal']
        row_dict['平仓盈亏'] = stats['realized_pl']
        row_dict['盯市盈亏'] = stats['mtm_pl']
        row_dict['手续费'] = stats['commission']
        row_dict['交割手续费'] = stats['delivery_fee']
        row_dict['期末结存'] = stats['balance_cf']
        # 处理银期出入金
        dw_bf = 0.
        if 'dw_array' in stats:
            dw_bf += sum([dw_item['deposit'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '银期转账'])
            dw_bf -= sum([dw_item['withdrawal'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '银期转账'])
        row_dict['银期出入金'] = dw_bf
        # 处理手续费返还
        dw_rebate_fee = 0.
        if 'dw_array' in stats:
            dw_rebate_fee += sum([dw_item['deposit'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '手续费返还'])
        row_dict['手续费返还'] = dw_rebate_fee
        # 处理利息返还
        dw_rebate_interest = 0.
        if 'dw_array' in stats:
            dw_rebate_interest += sum([dw_item['deposit'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '利息返还'])
        row_dict['利息返还'] = dw_rebate_interest
        # 处理中金所申报费
        dw_cffex_fee = 0. 
        if 'dw_array' in stats:
            dw_cffex_fee -= sum([dw_item['withdrawal'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '中金所申报费'])
        row_dict['中金所申报费'] = dw_cffex_fee
        row_dict['中金所手续费'] = stats['CFFEX_COMM'] if 'CFFEX_COMM' in stats else 0.
        row_dict['上期原油手续费'] = stats['INE_COMM'] if 'INE_COMM' in stats else 0.
        row_dict['上期所手续费'] = stats['SHFE_COMM'] if 'SHFE_COMM' in stats else 0. 
        row_dict['郑商所手续费'] = stats['CZCE_COMM'] if 'CZCE_COMM' in stats else 0. 
        row_dict['大商所工业品手续费'] = stats['DCE_IND_COMM'] if 'DCE_IND_COMM' in stats else 0. 
        row_dict['大商所农产品手续费'] = stats['DCE_AGR_COMM'] if 'DCE_AGR_COMM' in stats else 0.
        client_data.append(row_dict)
    # 总表
    client_df = pd.DataFrame(client_data, columns=COLUMNS)
    client_df['date'] = pd.to_datetime(client_df['日期'])
    client_df.sort_values(by='date', ascending=True, inplace=True)
    client_df.reset_index(inplace=True)
    # 检查出入金
    bug_rows = ((
        client_df['银期出入金'] + client_df['手续费返还'] + client_df['利息返还'] + client_df['中金所申报费'] - client_df['出入金合计']
        ).abs() > EPSILON
    )
    client_df.to_csv('client.csv')
    if bug_rows.sum() > 0:
        print('WARNING! 出入金不匹配，请检查下列日期出入金数据：\n %s' % str(client_df[bug_rows]))
    # 检查期初结存 + 当期总流水 与 期末结存
    bug_rows = ((
        client_df['期初结存'] + client_df['出入金合计'] + client_df['平仓盈亏'] + client_df['盯市盈亏'] + client_df['手续费'] - client_df['期末结存']
        ).abs() > EPSILON
    )
    if bug_rows.sum() > 0:
        print('WARNING! 期初结存+当期流水与期末结存不匹配，请检查下列日期数据：\n %s' % str(client_df[~bug_rows]))
    # 检查期末结存与下期初结存
    bug_rows = np.where(np.abs(client_df['期末结存'].values[:-1] - client_df['期初结存'].values[1:]) > EPSILON)[0].tolist()
    if len(bug_rows) > 0:
        bug_rows.append(max(bug_rows) + 1)
        print('WARNING！期末结存与下期初结存不匹配，请检查下列日期数据：\n %s' % str(client_df.iloc[bug_rows]))
    # 计算盈亏
    total_pl1 = client_df['期末结存'] - client_df['期初结存'] - client_df['银期出入金']  # 当期实际盈亏
    calc_rebate = []
    for i in range(len(client_df)):
        _date = client_df.iloc[i]['日期']
        _calc_rebate = - client_df['中金所手续费'][i] * rebate_df.loc[_date, 'CFFEX'] \
            - client_df['上期原油手续费'][i] * rebate_df.loc[_date, 'INE'] \
            - client_df['上期所手续费'][i] * rebate_df.loc[_date, 'SHFE'] \
            - client_df['郑商所手续费'][i] * rebate_df.loc[_date, 'CZCE'] \
            - client_df['大商所工业品手续费'][i] * rebate_df.loc[_date, 'DCE_IND'] \
            - client_df['大商所农产品手续费'][i] * rebate_df.loc[_date, 'DCE_AGR']
        calc_rebate.append(_calc_rebate)
    client_df['即时手续费返还'] = calc_rebate
    total_pl2 = total_pl1 - client_df['手续费返还'] + client_df['即时手续费返还'] # 当期即时盈亏
    client_df['实际盈亏'] = total_pl1
    client_df['即时盈亏'] = total_pl2
    # 计算份额和净值
    value1, value2 = [1.,], [1.,]
    units1 = [client_df['期末结存'][0] / value1[0], ]
    units2 = [client_df['期末结存'][0] / value2[0], ]
    balance_cf2 = [client_df['期末结存'][0],]  # 即时期末结存
    for i in range(1, len(client_df)):
        _dw_bf = client_df['银期出入金'][i]
        _balance_cf1 = client_df['期末结存'][i]
        _date = client_df['日期'][i]
        _balance_cf2 = balance_cf2[i-1] + total_pl2[i] + _dw_bf  # 当期即时期末结存
        balance_cf2.append(_balance_cf2)
        if abs(_dw_bf) > EPSILON:  # 处理银期出入后份额变化
            if value1[i-1] < EPSILON:  # 净值清零
                value1.append(1.)
                value2.append(1.)
                units1.append(_dw_bf / 1.)
                units2.append(_dw_bf / 1.)
            else:
                units1_delta = _dw_bf / value1[i-1]
                units1.append(units1[i-1] + units1_delta)
                units2_delta = _dw_bf / value2[i-1]
                units2.append(units2[i-1] + units2_delta)
                if units1[i] < EPSILON:  # 份额清零
                    value1.append(0.)
                    value2.append(0.)
                else:
                    value1.append(_balance_cf1 / units1[i])
                    value2.append(_balance_cf2 / units2[i])
        else:  # 保持份额
            units1.append(units1[i-1])
            units2.append(units2[i-1])
            value1.append(_balance_cf1 / units1[i])
            value2.append(_balance_cf2 / units2[i])

    client_df['实际份额'] = units1
    client_df['实际净值'] = value1
    client_df['即时份额'] = units2
    client_df['即时净值'] = value2
    client_df['即时期末结存'] = balance_cf2
    # 填充无交易日数据
    dummy_dates = sorted(list(map(
        lambda x: datetime.strptime(x, '%Y%m%d'),
        set(trading_dates) - set(client_df['日期'].values.tolist()))
    ))
    if len(dummy_dates) > 0:
        first_row_existed = min(dummy_dates) > datetime.strptime(client_df.iloc[0]['日期'], '%Y%m%d')
        for date in dummy_dates:
            date_str = date.strftime('%Y%m%d')
            if not first_row_existed:  # 若第一行不存在则手动初始化
                dummy_row = {
                    '日期': date_str,
                    '账户': client,
                    'date': date,
                }
                first_row_existed = True
            else:
                prev_row = client_df[client_df['date'] < date].sort_values('date').iloc[-1]
                dummy_row = {
                    '日期': date_str,
                    '账户': client,
                    'date': date,
                    '期初结存': prev_row['期末结存'],
                    '期末结存': prev_row['期末结存'],
                    '实际份额': prev_row['实际份额'],
                    '实际净值': prev_row['实际净值'],
                    '即时份额': prev_row['即时份额'],
                    '即时净值': prev_row['即时净值'],
                    '即时期末结存': prev_row['即时期末结存']
                }
            client_df = client_df.append(dummy_row, ignore_index=True)

    client_df.sort_values(by='date', ascending=True, inplace=True)
    client_df.fillna(0, inplace=True)
    client_df = client_df[client_df['日期'].apply(lambda d: datetime.strptime(d, '%Y%m%d')) >= start_date]
    # 增加合计行
    last_row = client_df.iloc[-1]
    total_row = {
        '账户': client_id,
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
    # 写入结算主表
    output_path = os.path.join(
        output_dir, '%s_%s_%s.xlsx' 
        % (client, start_date.strftime('%Y%m%d'), end_date.strftime('%Y%m%d'))
    )
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    client_df.to_excel(
        writer, '结算汇总', index=False, columns=COLUMNS, freeze_panes=(1, 2)
    )
    workbook = writer.book
    worksheet = writer.sheets['结算汇总']
    for i, col in enumerate(COLUMNS):
        if 11 <= i <= 16:
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
    # 作图
    plot_path = os.path.join(
        output_dir, '%s_%s_%s.png'
        % (client, start_date.strftime('%Y%m%d'), end_date.strftime('%Y%m%d'))
    )
    fig, ax = plt.subplots(dpi=120, figsize=(12,6))
    client_df.plot(kind='line', x='date', y='实际净值', label='real value', color='red', ax=ax)
    client_df.plot(kind='line', x='date', y='即时净值', label='instant value', color='green', ax=ax)
    ax.axhline(y=1.0, linestyle='--', color='black')
    ax.set_title('Real profit: %.2fw, value: %.2f; Instant profit: %.2fw, value: %.2f' 
        % (total_row['实际盈亏'], total_row['实际净值'], total_row['即时盈亏'], total_row['即时净值']))
    plt.savefig(plot_path)
    # 银期出入金副表
    bf_array = []
    for stats in client_stats:
        if 'dw_array' not in stats:
            continue
        bf_array += [dw_item for dw_item in stats['dw_array'] if dw_item['dw_type'] == '银期转账']
    if len(bf_array) > 0:
        bf_df = pd.DataFrame(bf_array)
        bf_df.rename(columns={'date': '日期', 'deposit': '入金', 'withdrawal': '出金'}, inplace=True)
        # 检查每日银期出入金
        for date in client_df['日期']:
            if date == '合计':
                continue
            bf_rows = bf_df['日期'] == date
            bug_rows = (bf_df[bf_rows]['入金'].sum() - bf_df[bf_rows]['出金'].sum() - client_df[client_df['日期'] == date]['银期出入金']).abs() > EPSILON
            if bug_rows.sum() > 0:
                print('WARNING! 银期出入金数据不匹配：%s' % date)
        bf_df.to_excel(
            writer, '银期转账', index=False, columns=('日期', '入金', '出金')
        )
    writer.save()
    print('%s --> %s' % (client, output_path))
