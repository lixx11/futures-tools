#!/usr/bin/env python

"""Generate Excel summary from multiple CTP files.

Usage:
    CTP2Excel.py <CTP-file>... [options]

Options:
    -h --help               Show this screen.
    -o --output=<folder>    Specify output directory [default: output].
"""
from docopt import docopt
import numpy as np
import pandas as pd
import os
import sys


TABLES = ('资金状况', '成交记录', '出入金明细', '平仓明细', '持仓明细', '持仓汇总')
COLUMNS = ('账户', '日期', '期初结存', '银期出入金', '手续费返还', '利息返还', '中金所申报费', '出入金合计', '平仓盈亏', '盯市盈亏', '手续费', '期末结存') # 结算总表
EPSILON = 0.01  # 匹配误差容忍度

def extract_data(filepath):
    with open(filepath) as f:
        contents = f.readlines()
    # split contents into blocks
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
                    _stats = process_deposit_withdrawal(block_content)
                    if abs(stats['total_deposit_withdrawal'] - _stats['total_deposit_withdrawal']) > EPSILON:
                        print('WARNING! 资金状况出入金(%.2f)与出入金明细不匹配(%.2f)！' 
                              % (stats['total_deposit_withdrawal'], _stats['total_deposit_withdrawal']))
                    stats = {**stats, **_stats}
                row_id = i
                block_prev = TABLE
    return stats


def process_head(content):
    """
    处理结算单表头信息，获取用户id、用户姓名以及日期
    """

    client_id = content[4].split('：')[1].strip().split()[0]
    client_name = content[4].split('：')[2].strip()
    date = content[5].split('：')[-1].strip()
    stats = {
        'client_id': client_id,
        'client_name': client_name,
        'date': date
    }
    return stats


def process_summary(content):
    """
    处理资金状况，获取期初结存、出入金、平仓盈亏、盯市盈亏、手续费和期末结存等数据。
    """

    balance_bf = float(content[2].split('：')[1].strip().split()[0])  # 期初结存
    total_deposit_withdrawal = float(content[3].split('：')[1].strip().split()[0])  # 总出入金
    realized_pl = float(content[4].split('：')[1].strip().split()[0])  # 平仓盈亏
    mtm_pl = float(content[5].split('：')[1].strip().split()[0])  # 盯市盈亏
    commission = -float(content[7].split('：')[1].strip().split()[0])  # 手续费
    balance_cf = float(content[3].split('：')[-1].strip())  # 期末结存
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
        'balance_cf': balance_cf,
    }
    return stats


def process_deposit_withdrawal(content):
    """
    处理出入金，获取各笔出入金类型与金额。
    """
    dw_array = []
    for i in range(5, len(content)):
        if len(content[i].strip('-\n')) == 0:
            last_row = i
            break
    for i in range(5, last_row):
        date, dw_type, deposit, withdrawal, comment  = content[i][1:-2].split('|')
        if '中金所申报费' in comment:
            dw_type = '中金所申报费'
        dw_array.append({
            'date': date.strip(),
            'dw_type': dw_type.strip(),
            'deposit': float(deposit),
            'withdrawal': float(withdrawal),
        })
    _, _, total_deposit, total_withdrawal, _ = content[last_row + 1][1:-2].split('|')
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


if __name__ == "__main__":
    argv = docopt(__doc__)
    CTP_files = argv['<CTP-file>']
    print('Generate summary from %s' % CTP_files)
    all_stats = []
    for CTP_file in CTP_files:
        print('Processing %s' % CTP_file)
        stats = extract_data(CTP_file)
        all_stats.append(stats)
    client_ids = np.unique([stats['client_id'] for stats in all_stats])
    
    # 数据汇总输出到Excel文件中
    output_dir = argv['--output']
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
    for client_id in client_ids:  # 分别处理每个客户结算数据
        client_stats =[stats for stats in all_stats if stats['client_id'] == client_id]
        client = '%s-%s' % (client_stats[0]['client_id'], client_stats[0]['client_name'])
        client_data = []
        for stats in client_stats:
            row_dict = {}
            row_dict['账户'] = stats['client_id']
            row_dict['日期'] = stats['date']
            row_dict['期初结存'] = stats['balance_bf']
            row_dict['出入金合计'] = stats['total_deposit_withdrawal']
            row_dict['平仓盈亏'] = stats['realized_pl']
            row_dict['盯市盈亏'] = stats['mtm_pl']
            row_dict['手续费'] = stats['commission']
            row_dict['期末结存'] = stats['balance_cf']
            # 处理银期出入金
            dw_bf = 0.
            if 'dw_array' in stats:
                dw_bf += sum([dw_item['deposit'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '银期转账'])
                dw_bf -= sum([dw_item['withdrawal'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '银期转账'])
            row_dict['银期出入金'] = dw_bf
            # 处理手续费返还
            dw_return_fee = 0.
            if 'dw_array' in stats:
                dw_return_fee += sum([dw_item['deposit'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '手续费返还'])
            row_dict['手续费返还'] = dw_return_fee
            # 处理利息返还
            dw_return_interest = 0.
            if 'dw_array' in stats:
                dw_return_interest += sum([dw_item['deposit'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '利息返还'])
            row_dict['利息返还'] = dw_return_interest
            # 处理中金所申报费
            dw_cffex_fee = 0. 
            if 'dw_array' in stats:
                dw_cffex_fee -= sum([dw_item['withdrawal'] for dw_item in stats['dw_array'] if dw_item['dw_type'] == '中金所申报费'])
            row_dict['中金所申报费'] = dw_cffex_fee
            client_data.append(row_dict)
        # 总表
        client_df = pd.DataFrame(client_data, columns=COLUMNS)
        client_df['date'] = pd.to_datetime(client_df['日期'])
        client_df.sort_values(by='date', ascending=True, inplace=True)
        # 检查出入金
        bug_rows = ((
            client_df['银期出入金'] + client_df['手续费返还'] + client_df['利息返还'] + client_df['中金所申报费'] - client_df['出入金合计']
            ).abs() > EPSILON
        )
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
        # 写入结算主表
        writer = pd.ExcelWriter(os.path.join(output_dir, '%s.xlsx' % client))
        client_df.to_excel(
            writer, '结算汇总', index=False, columns=COLUMNS
        )
        # 银期出入金副表
        bf_array = []
        for stats in client_stats:
            if 'dw_array' not in stats:
                continue
            bf_array += [dw_item for dw_item in stats['dw_array'] if dw_item['dw_type'] == '银期转账']
        bf_df = pd.DataFrame(bf_array)
        bf_df.rename(columns={'date': '日期', 'deposit': '入金', 'withdrawal': '出金'}, inplace=True)
        # 检查每日银期出入金
        for date in client_df['日期']:
            bf_rows = bf_df['日期'] == date
            bug_rows = (bf_df[bf_rows]['入金'].sum() - bf_df[bf_rows]['出金'].sum() - client_df[client_df['日期'] == date]['银期出入金']).abs() > EPSILON
            if bug_rows.sum() > 0:
                print('WARNING! 银期出入金数据不匹配：%s' % date)
        bf_df.to_excel(
            writer, '银期转账', index=False, columns=('日期', '入金', '出金')
        )
        writer.save()
        print('-' * 80)
        print('结算单检查完毕，数据写入目录%s' % output_dir)