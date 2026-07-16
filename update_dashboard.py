#!/usr/bin/env python3
"""
竞品分析仪表板 - 自动更新脚本
从 Excel 数据源读取最新数据，更新 HTML 仪表板中的 JavaScript 数据
使用方法: python3 update_dashboard.py
"""
import pandas as pd
import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).parent
DASHBOARD_FILE = BASE_DIR / '竞品分析仪表板.html'
XLSX_MAIN = BASE_DIR / '倍瑞益、费瑞卡、蛋白乐相关竞品等产品信息表.xlsx'
XLSX_PROTEIN = BASE_DIR / '3. 蛋百乐竞品信息收集/蛋白粉产品对比.xlsx'

def parse_num(v):
    if v is None or pd.isna(v) or v == '/' or v == '':
        return None
    s = str(v).replace(',', '').replace(' ', '')
    try:
        return float(s)
    except:
        return s

def extract_competitor_data(df, cols):
    """从 DataFrame 提取竞品数据"""
    competitors = []
    for col in cols:
        comp = {}
        for idx, row in df.iterrows():
            key = str(row['品牌名称']).strip()
            val = row[col]
            if pd.notna(val):
                comp[key] = str(val).strip()
        if comp:
            competitors.append(comp)
    return competitors

def main():
    print("🔍 读取 Excel 数据源...")

    # 读取主 Excel
    sheets = pd.read_excel(XLSX_MAIN, sheet_name=None, dtype=str)
    
    # 倍瑞益竞品
    df_br = sheets.get('倍瑞益竞品', pd.DataFrame())
    print(f"  ✓ 倍瑞益竞品: {df_br.shape[0]} 行, {df_br.shape[1]} 列")
    
    # 费瑞卡竞品
    df_fr = sheets.get('费瑞卡竟品', pd.DataFrame())
    print(f"  ✓ 费瑞卡竞品: {df_fr.shape[0]} 行, {df_fr.shape[1]} 列")
    
    # 蛋百乐竞品
    df_db = sheets.get('蛋百乐竞品', pd.DataFrame())
    print(f"  ✓ 蛋百乐竞品: {df_db.shape[0]} 行, {df_db.shape[1]} 列")
    
    # 蛋白粉产品对比
    df_protein = pd.read_excel(XLSX_PROTEIN, sheet_name='市售产品', dtype=str)
    print(f"  ✓ 蛋白粉产品对比: {df_protein.shape[0]} 行")
    
    print(f"\n📊 数据读取完成！")
    print(f"  仪表板文件: {DASHBOARD_FILE}")
    print(f"  请打开 HTML 文件查看最新数据。")
    
    # 显示摘要
    print(f"\n📋 数据摘要:")
    for sheet_name in ['倍瑞益竞品', '费瑞卡竟品', '蛋百乐竞品']:
        df = sheets.get(sheet_name, pd.DataFrame())
        cols = [c for c in df.columns if c != '品牌名称']
        print(f"  {sheet_name}: {len(cols)} 个竞品品牌")

if __name__ == '__main__':
    main()
