#!/usr/bin/env python3
"""
案件管理.xlsx → GASアプリ用インポートJSON変換スクリプト

使い方:
  python3 scripts/excel_to_import_json.py [Excelファイルパス] [出力JSONファイルパス]

デフォルト:
  入力: ~/Downloads/案件管理.xlsx
  出力: ./import_data.json
"""
import sys
import json
import math
import datetime
from pathlib import Path

def convert(excel_path, output_path):
    try:
        import pandas as pd
    except ImportError:
        print("pandas が必要です: pip3 install pandas openpyxl")
        sys.exit(1)

    df = pd.read_excel(excel_path, sheet_name='案件一覧')

    def fmtdate(v):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return ''
        if isinstance(v, (datetime.datetime, datetime.date)):
            return v.strftime('%Y-%m-%d')
        return str(v)[:10]

    def safe_str(v):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return ''
        return str(v)

    def safe_float(v):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return 0
        return float(v)

    projects = []
    for _, row in df.iterrows():
        no = row['No.']
        if no is None or (isinstance(no, float) and math.isnan(no)):
            continue
        pid = 'PRJ-' + str(int(no)).zfill(3)
        reward = safe_float(row['報酬'])
        hours  = safe_float(row['稼動時間'])
        hourly = safe_float(row['時給']) or (round(reward / hours) if hours > 0 else 0)
        projects.append({
            'id':             pid,
            'name':           safe_str(row['案件名']),
            'contact':        safe_str(row['担当者']),
            'platform':       safe_str(row['プラットフォーム']),
            'url':            safe_str(row['URL']),
            'contactUrl':     safe_str(row['連絡先']),
            'type':           safe_str(row['種別']),
            'proposalDate':   fmtdate(row['提案日 or 相談日']),
            'status':         safe_str(row['ステータス']),
            'deliveryDate':   fmtdate(row['納品日']),
            'acceptanceDate': fmtdate(row['検収日']),
            'reward':         reward,
            'hours':          hours,
            'hourly':         hourly,
            'tasks':          safe_str(row['タスク']),
            'memo':           safe_str(row['メモ']),
            'docs':           safe_str(row['関連資料']),
            'isInvoice':      False,
            'isDeleted':      False,
            'createdAt':      '',
            'updatedAt':      '',
        })

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)

    print(f'変換完了: {len(projects)}件 → {output_path}')

if __name__ == '__main__':
    excel  = sys.argv[1] if len(sys.argv) > 1 else str(Path.home() / 'Downloads' / '案件管理.xlsx')
    output = sys.argv[2] if len(sys.argv) > 2 else 'import_data.json'
    convert(excel, output)
