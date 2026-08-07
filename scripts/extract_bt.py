# -*- coding: utf-8 -*-
"""从 宝淘淘汇总数据.xlsx 提取全量日报+月报数据（紧凑数组格式），输出 JSON

== 口径（用户确认） ==
  净收入    = 订单收入 - 补单收入 - 退款金额 - 其他减免      （补单收入已从净收入中剔除）
  毛利      = 发货收入 - 商品成本 - 快递费                   （发货收入是真实收入，与净收入存在时间差）
  边际贡献  = 毛利 - 平台佣金 - 推广费
  净利润    = 边际贡献 - 总分摊费用                          （补单成本 不参与利润链路）
  补单占比  = 补单收入 / 订单收入                            （补单的关键指标是占比，不是金额）

day.kpi  = [net_rev, order_rev, refund, gross, profit, visitors, buyers, real_buyers,
            promo_fee, promo_rev, rows, net_rate, aov, conv,
            supply_rev, supply_ratio, promo_rate, promo_roi, margin]   (19 项)

mon.kpi  = [net_rev, order_rev, refund, gross, profit, visitors, buyers, real_buyers,
            promo_fee, promo_rev, product_cost, ship_fee, commission,
            supply_cost, other_cost, alloc_cost, rows, net_rate, gross_rate, aov, conv,
            supply_rev, supply_ratio, promo_rate, promo_roi, margin]   (26 项)

day.sh   = [[name, net_rev, profit, visitors, buyers, promo_fee, supply_ratio], ...]
            按 净利润 降序 top12
day.pl   = [[name, net_rev, profit], ...]                            top8
day.st   = [[name, net_rev, profit, rows], ...]                      按 净利润 降序 top8
day.tp   = [[id, shop, net_rev, profit], ...]                        按 净利润 降序 top8
day.wl   = [win, loss]

mon.sh   = [[name, net_rev万, profit万, net_rate%, visitors, buyers, aov,
              promo_fee万, supply_ratio%], ...]                     按 净利润 降序
mon.pl   = [[name, net_rev万, profit万], ...]
mon.st   = [[name, net_rev万, profit万, rows], ...]                  运营维度，按 净利润 降序（日报用）
mon.sup  = [[name, net_rev万, profit万, rows, shop_count], ...]      主管维度，按 净利润 降序（月报用）
mon.tp   = [[id, shop, net_rev万, profit万], ...]                    按 净利润 降序 top15
mon.daily= [[date, net_rev, profit, promo_fee, promo_rev, supply_rev, order_rev], ...]
mon.wl   = [win, loss]
"""
import pandas as pd
import numpy as np
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
FILE = r"Z:\原始文件\欣语\最终结果\宝淘淘汇总数据.xlsx"
OUT = os.path.join(BASE, "_data_bt.json")


def mask_name(s):
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    s = str(s).strip()
    if s in ("", "nan", "NaN", "0", "None"):
        return ""
    if len(s) <= 1:
        return s
    if len(s) == 2:
        return s[0] + "*"
    return s[0] + "***" + s[-1]


def mask_id(v):
    s = str(int(float(v))) if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v)) else str(v).strip()
    if s in ("", "nan", "None"):
        return "-"
    if len(s) <= 8:
        return s
    return s[:4] + "***" + s[-4:]


df = pd.read_excel(FILE, sheet_name='结果', header=0)
df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
df = df.dropna(subset=['日期'])
df['ym'] = df['日期'].dt.to_period('M')
df['yd'] = df['日期'].dt.strftime('%Y-%m-%d')

cost_cols = ['订单收入','补单收入','退款金额','其他减免','净收入','发货收入','商品成本','快递费','毛利',
             '平台佣金','推广费','补单成本','其他直接成本','总分摊费用','净利润','推广收入']
for c in cost_cols:
    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
df['访客数'] = pd.to_numeric(df['访客数'], errors='coerce').fillna(0)
df['买家数'] = pd.to_numeric(df['买家数'], errors='coerce').fillna(0)
df['真实买家数'] = pd.to_numeric(df['真实买家数'], errors='coerce').fillna(0)


def daily_kpi(s):
    net = s['净收入'].sum()
    gross = s['毛利'].sum()
    profit = s['净利润'].sum()
    buyers = s['买家数'].sum()
    visitors = s['访客数'].sum()
    supply_rev = s['补单收入'].sum()
    order_rev = s['订单收入'].sum()
    promo_fee = s['推广费'].sum()
    promo_rev = s['推广收入'].sum()
    commission = s['平台佣金'].sum()
    margin = gross - commission - promo_fee   # 边际贡献
    return [
        round(net, 2), round(order_rev, 2), round(s['退款金额'].sum(), 2), round(gross, 2),
        round(profit, 2), int(visitors), int(buyers), int(s['真实买家数'].sum()),
        round(promo_fee, 2), round(promo_rev, 2), len(s),
        round(profit / net, 6) if net else None,
        round(net / buyers, 2) if buyers else None,
        round(buyers / visitors, 6) if visitors else None,
        round(supply_rev, 2),
        round(supply_rev / order_rev, 6) if order_rev else None,          # 补单占比
        round(promo_fee / net, 6) if net else None,
        round(promo_rev / promo_fee, 6) if promo_fee else None,
        round(margin, 2),
    ]


def monthly_kpi(s):
    net = s['净收入'].sum()
    gross = s['毛利'].sum()
    profit = s['净利润'].sum()
    buyers = s['买家数'].sum()
    visitors = s['访客数'].sum()
    supply_rev = s['补单收入'].sum()
    order_rev = s['订单收入'].sum()
    promo_fee = s['推广费'].sum()
    promo_rev = s['推广收入'].sum()
    commission = s['平台佣金'].sum()
    margin = gross - commission - promo_fee
    return [
        round(net, 2), round(order_rev, 2), round(s['退款金额'].sum(), 2), round(gross, 2),
        round(profit, 2), int(visitors), int(buyers), int(s['真实买家数'].sum()),
        round(promo_fee, 2), round(promo_rev, 2),
        round(s['商品成本'].sum(), 2), round(s['快递费'].sum(), 2), round(commission, 2),
        round(s['补单成本'].sum(), 2), round(s['其他直接成本'].sum(), 2), round(s['总分摊费用'].sum(), 2),
        len(s),
        round(profit / net, 6) if net else None,
        round(gross / net, 6) if net else None,
        round(net / buyers, 2) if buyers else None,
        round(buyers / visitors, 6) if visitors else None,
        round(supply_rev, 2),
        round(supply_rev / order_rev, 6) if order_rev else None,
        round(promo_fee / net, 6) if net else None,
        round(promo_rev / promo_fee, 6) if promo_fee else None,
        round(margin, 2),
    ]


# ============ 按日 ============
days = []
for d, sub in df.groupby('yd'):
    shops = []
    for name, s2 in sub.groupby('店铺'):
        nr = s2['净收入'].sum(); p = s2['净利润'].sum()
        buy = s2['买家数'].sum()
        sr = s2['补单收入'].sum(); orv = s2['订单收入'].sum(); pf = s2['推广费'].sum()
        shops.append([
            str(name), round(nr, 0), round(p, 0),
            int(s2['访客数'].sum()), int(buy),
            round(pf, 0), round(sr / orv, 6) if orv else None
        ])
    shops.sort(key=lambda x: x[2], reverse=True)   # 净利润降序
    plats = []
    for name, s2 in sub.groupby('平台'):
        plats.append([str(name), round(s2['净收入'].sum(), 0), round(s2['净利润'].sum(), 0)])
    plats.sort(key=lambda x: x[2], reverse=True)
    staffs = []
    for name, s2 in sub.groupby('运营'):
        staffs.append([mask_name(name) or "无运营", round(s2['净收入'].sum(), 0), round(s2['净利润'].sum(), 0), len(s2)])
    staffs.sort(key=lambda x: x[2], reverse=True)  # 净利润降序
    top = []
    for (pid, shop), s2 in sub.groupby(['商品ID', '店铺']):
        top.append([mask_id(pid), str(shop), round(s2['净收入'].sum(), 0), round(s2['净利润'].sum(), 0)])
    top.sort(key=lambda x: x[3], reverse=True)    # 净利润降序
    days.append({
        "date": d, "kpi": daily_kpi(sub),
        "sh": shops[:12], "pl": plats[:8], "st": staffs[:8], "tp": top[:8],
        "wl": [int((sub['净利润'] > 0).sum()), int((sub['净利润'] <= 0).sum())],
    })
days.sort(key=lambda x: x["date"])

# ============ 按月 ============
months = []
for ym, sub in df.groupby('ym'):
    daily = []
    for d, s2 in sub.groupby('日期'):
        daily.append([
            d.strftime("%m-%d"),
            round(s2['净收入'].sum(), 0),
            round(s2['净利润'].sum(), 0),
            round(s2['推广费'].sum(), 0),
            round(s2['推广收入'].sum(), 0),
            round(s2['补单收入'].sum(), 0),
            round(s2['订单收入'].sum(), 0),
        ])
    daily.sort(key=lambda x: x[0])
    shops = []
    for name, s2 in sub.groupby('店铺'):
        nr = s2['净收入'].sum(); p = s2['净利润'].sum(); buy = s2['买家数'].sum()
        sr = s2['补单收入'].sum(); orv = s2['订单收入'].sum(); pf = s2['推广费'].sum()
        shops.append([
            str(name),
            round(nr / 10000, 2), round(p / 10000, 2),
            round(p / nr * 100, 1) if nr else None,
            int(s2['访客数'].sum()), int(buy), round(nr / buy, 1) if buy else None,
            round(pf / 10000, 2), round(sr / orv * 100, 1) if orv else None,
        ])
    shops.sort(key=lambda x: x[2], reverse=True)   # 净利润降序
    plats = []
    for name, s2 in sub.groupby('平台'):
        plats.append([str(name), round(s2['净收入'].sum() / 10000, 2), round(s2['净利润'].sum() / 10000, 2)])
    plats.sort(key=lambda x: x[2], reverse=True)
    staffs = []
    for name, s2 in sub.groupby('运营'):
        staffs.append([mask_name(name) or "无运营", round(s2['净收入'].sum() / 10000, 2), round(s2['净利润'].sum() / 10000, 2), len(s2)])
    staffs.sort(key=lambda x: x[2], reverse=True)
    supervisors = []
    for name, s2 in sub.groupby('主管'):
        supervisors.append([mask_name(name) or "无主管",
                            round(s2['净收入'].sum() / 10000, 2),
                            round(s2['净利润'].sum() / 10000, 2),
                            len(s2),
                            s2['店铺'].nunique()])
    supervisors.sort(key=lambda x: x[2], reverse=True)
    top = []
    for (pid, shop), s2 in sub.groupby(['商品ID', '店铺']):
        top.append([mask_id(pid), str(shop), round(s2['净收入'].sum() / 10000, 2), round(s2['净利润'].sum() / 10000, 2)])
    top.sort(key=lambda x: x[3], reverse=True)
    months.append({
        "month": str(ym), "kpi": monthly_kpi(sub),
        "sh": shops, "pl": plats, "st": staffs, "sup": supervisors, "tp": top[:15], "daily": daily,
        "wl": [int((sub['净利润'] > 0).sum()), int((sub['净利润'] <= 0).sum())],
        "complete": bool(len(daily) >= 28),
    })
months.sort(key=lambda x: x["month"])

data = {
    "range": {"min": df['日期'].min().strftime("%Y-%m-%d"), "max": df['日期'].max().strftime("%Y-%m-%d")},
    "days": days, "months": months,
}

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

print("days:", len(days), "months:", len(months), "JSON KB:", round(os.path.getsize(OUT) / 1024, 1))