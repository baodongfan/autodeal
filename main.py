# -*- coding: utf-8 -*-
import pandas as pd
import os


def clear_NA_cols(df):
    """删除空列"""
    # for i in df.columns:
        # if df[i].count() == 0:
        #     print(i)
        #     df.drop(labels=i, axis=1, inplace=True)
    print('done')

def divide_by_name_to_sheets(df, file_name):
    """
    按照姓名把sheet分成独立的sheet
    """
    names = list(df['服务人员'].drop_duplicates())
    with pd.ExcelWriter('{0}_new.xlsx'.format(file_name)) as excel_writer:
        # 循环每一类写入
        for name in names:
            bool_df = df['服务人员'] == name
            my_df = df[bool_df]
            my_df.to_excel(excel_writer, sheet_name=name, index=True)

def divide_by_name_to_excel(df):
    """按照姓名把Excel分成独立的Excel"""
    names = list(df['服务人员'].drop_duplicates())
    for name in names:
        res = df['服务人员'] == name
        df[res].to_excel('{0}.xlsx'.format(name), index=False)

def deal_date(df):
    if '激活日期' in df.columns:
        df['激活日期'] = df['激活日期'].dt.date
    if '末次交易时间' in df.columns:
        df['末次交易时间'] = df['末次交易时间'].dt.date
    if '流转日期' in df.columns:
        df['流转日期'] = pd.to_datetime(df['流转日期'], format='%Y-%m-%d')
        df['流转日期'] = df['流转日期'].dt.date


def combin_files_in_one_file_sheets():
    """
    合并多个Excel到一个Excel的多个sheet里
    :return:
    """
    writer = pd.ExcelWriter(r'result.xlsx')

    for i in os.listdir(r'.')[:-1]:

        if 'py' in i or 'exe' in i or '客户' in i:
            continue
        name = i.split('.')[0]
        print('现在准备合并： ', i)
        df = pd.read_excel(i)
        df.to_excel(writer, sheet_name=name, index=False)
        writer.save()

def deal_huanqiu():
    print('正在处理环球数据，loading#################')
    df = pd.read_excel('客户结果数据明细表-环球.xlsx')
    # 删除不需要的列
    df.drop(['国际户激活日期', '国际户激活金额', '盛宝户激活日期', '盛宝户激活金额', '当日佣转',
             '本月佣转', '当日收入', '当月收入', '累计佣金'], axis=1, inplace=True)
    # 删除空列
    df.dropna(axis=1, how='all', inplace=True)
    # 处理时间
    deal_date(df)
    # 选择组别,筛选二部一组
    df = df[df['服务组别'] == '二部一组']
    df.to_excel('new_环球.xlsx', index=None)
    print('环球处理完成 :) ')
    return df

def deal_shengbao():
    print('正在处理盛宝数据，loading#################')
    df2 = pd.read_excel('客户结果数据明细表-盛宝.xlsx')
    # 删除不需要的列
    df2.drop(['当日佣转', '本月佣转', '当日收入', '当月收入', '累计佣金'], axis=1, inplace=True)
    # 删除空列
    df2.dropna(axis=1, how='all', inplace=True)
    # 资金部分换成万
    df2[['流转金额', '激活金额', '存量','当日入金','当日出金','当日净入金','当月入金','当月出金','当月净入金','当月美股佣金','当月港股佣金','当月期货佣金','当月涡轮牛熊证佣金','当月总佣金']] = df2[['流转金额', '激活金额', '存量','当日入金','当日出金','当日净入金','当月入金','当月出金','当月净入金','当月美股佣金','当月港股佣金','当月期货佣金','当月涡轮牛熊证佣金','当月总佣金']] / 10000
    # 处理时间
    deal_date(df2)
    # 选择组别,筛选二部一组
    df2 = df2[df2['服务组别'] == '二部一组']
    df2.to_excel('new_盛宝.xlsx', index=None)
    print('盛宝处理完成 :) ')
    return df2

def deal_xindan():
    print('正在处理新单操作明细，loading:#################')
    df3 = pd.read_excel('新单操作明细.xlsx')
    df3.dropna(axis=1, how='all', inplace=True)
    df3.drop(['佣金收入'], axis=1, inplace=True)
    deal_date(df3)
    # 选择组别,筛选二部一组
    df3 = df3[df3['服务组别'] == '二部一组']
    df3.to_excel('new_新单.xlsx', index=None)
    print('新单处理完成 :) ')
    return df3

def deal_qihuo():
    print('正在处理期货操作，loading:#################')
    df = pd.read_excel('期货持仓客户明细.xlsx')
    df.dropna(axis=1, how='all', inplace=True)
    df.drop(['交易账号'], axis=1, inplace=True)
    deal_date(df)
    df = df[df['服务组别'] == '二部一组']
    df.to_excel('new_期货.xlsx', index=None)
    print('期货处理完成 :) ')


def combin_three_files_in_one_excel(df1, df2, df3):
    """
    按照姓名把sheet分成独立的sheet
    """
    names = list(df1['服务人员'].drop_duplicates())

if __name__ == '__main__':
    if os.path.exists('客户结果数据明细表-环球.xlsx'):
        deal_huanqiu()
    if os.path.exists('客户结果数据明细表-盛宝.xlsx'):
        deal_shengbao()
    if os.path.exists('新单操作明细.xlsx'):
        deal_xindan()
    if os.path.exists('期货持仓客户明细.xlsx'):
        deal_qihuo()

    combin_files_in_one_file_sheets()
