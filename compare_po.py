# -*- coding: utf-8 -*-
"""Reusable Yakjin Aerie order-vs-PO comparison. Requires: pip install openpyxl"""
import argparse, re
from collections import defaultdict, Counter
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


def excel_rows(path, preferred, cols):
    wb=load_workbook(path, read_only=True, data_only=True)
    sh=wb[preferred] if preferred in wb.sheetnames else wb.worksheets[0]
    rows=[]
    for row in sh.iter_rows(values_only=True):
        values=list(row[:cols])
        rows.append(values+[None]*(cols-len(values)))
    wb.close()
    return rows

def t(x):return str(x or '').strip()
def po7(x):return bool(re.fullmatch(r'[0-9]{7}',t(x)))
def date(x):
    if isinstance(x,(int,float)):
        return (datetime(1899,12,30)+timedelta(days=float(x))).strftime('%Y-%m-%d')
    if isinstance(x,datetime):return x.strftime('%Y-%m-%d')
    s=t(x)
    for fmt in ('%m/%d/%y','%m/%d/%Y','%Y-%m-%d','%d/%m/%Y'):
        try:return datetime.strptime(s,fmt).strftime('%Y-%m-%d')
        except ValueError:pass
    return s

def season(x):
    s=t(x).upper()
    year=re.search(r'20(\d\d)',s)
    yy=year.group(1) if year else (re.search(r'\b(\d\d)\b',s).group(1) if re.search(r'\b(\d\d)\b',s) else '')
    if 'SUM' in s or re.search(r'\bSU\b',s):return 'SU'+yy
    if 'SPR' in s or re.search(r'\bSP\b',s):return 'SP'+yy
    if 'HOL' in s:return 'HOL'+yy
    if 'BTS' in s:return 'BTS'+yy
    return s

def factory(x):
    s=re.sub(r'[^A-Z0-9]+',' ',t(x).upper()).strip()
    if 'VINA KOREA' in s:return 'VINA KOREA'
    if 'YAKJIN JAYA' in s:return 'PT YAKJIN JAYA'
    if 'YAKJIN VIETNAM' in s:return 'YAKJIN VIETNAM'
    if 'YAKJIN CAMBO' in s:return 'YAKJIN CAMBODIA'
    return s

def quantity(x):
    try:return int(round(float(x or 0)))
    except (TypeError,ValueError):return 0

def output_sheet(wb,name,headers,records,widths):
    sh=wb.create_sheet(name)
    sh.append(headers)
    for row in records: sh.append([str(v) if isinstance(v,(list,tuple,set)) else v for v in row])
    for cell in sh[1]:
        cell.fill=PatternFill('solid',fgColor='17365D')
        cell.font=Font(name='Aptos Narrow',size=10,bold=True,color='FFFFFF')
        cell.alignment=Alignment(vertical='center')
    sh.row_dimensions[1].height=29
    for i,w in enumerate(widths,1): sh.column_dimensions[get_column_letter(i)].width=max(12,w/1.3)
    sh.freeze_panes='A2'
    if records:
        table=Table(displayName=re.sub(r'\W','',name)+'Table',ref=f'A1:{get_column_letter(len(headers))}{len(records)+1}')
        table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True)
        sh.add_table(table)
    return sh

def run(order_path,po_path,out_path):
    oo=excel_rows(order_path,'Sheet1',49)
    pp=excel_rows(po_path,'Sheet0',33)
    order=[];excluded=[]
    for r in oo[2:]:
        if not t(r[24]):continue
        item=[season(r[18]),t(r[7]),t(r[15]),t(r[24]),t(r[42]).zfill(3),t(r[40]),date(r[30]),factory(r[3]),quantity(r[44])]
        if po7(r[24]):order.append(item)
        else:excluded.append(item+['PO# not 7 digits'])
    po=[]
    for r in pp[1:]:
        if not t(r[14]) or not t(r[17]):continue
        po.append([season(r[10]),t(r[6]),t(r[17]),t(r[14]),t(r[19]).zfill(3),date(r[23]),factory(r[29]),quantity(r[21])])
    # Main identity uses division/style/PO/color; date and factory are separately diagnosed.
    def ident(r):return (r[2],r[3],r[4])
    og=defaultdict(lambda:defaultdict(int)); pg=defaultdict(lambda:defaultdict(int))
    om=defaultdict(list);pm=defaultdict(list)
    names=defaultdict(set);names_by_code=defaultdict(set)
    for r in order:
        k=ident(r);og[k][(r[6],r[7])]+=r[8];om[k].append(r)
        if r[5]:names[(r[2],r[4])].add(r[5]);names_by_code[r[4]].add(r[5])
    for r in po:
        k=ident(r);pg[k][(r[5],r[6])]+=r[7];pm[k].append(r)
    details=[];stat=Counter();byfactory=defaultdict(lambda:[0,0,0]);byseason=defaultdict(lambda:[0,0,0]); review=[]
    for k in sorted(set(og)|set(pg)):
        a=om.get(k,[]);b=pm.get(k,[])
        oq=sum(r[8] for r in a);pq=sum(r[7] for r in b)
        dates_a=sorted(set(r[6] for r in a));dates_b=sorted(set(r[5] for r in b))
        fa=sorted(set(r[7] for r in a));fb=sorted(set(r[6] for r in b))
        ss=sorted(set(r[0] for r in a));ps=sorted(set(r[0] for r in b))
        n=names.get((k[0],k[2]),set()) or names_by_code.get(k[2],set())
        cname=next(iter(n)) if len(n)==1 else ('COLOR NAME REVIEW REQUIRED' if len(n)>1 else '')
        if len(n)>1:review.append([k[0],k[2],', '.join(sorted(n))])
        if not a:status='ONLY PO'
        elif not b:status='ONLY ORDER'
        elif oq!=pq:status='QTY DIFF'
        elif dates_a!=dates_b:status='DATE CHANGE'
        elif fa!=fb:status='FACTORY CHANGE'
        elif ss!=ps:status='SEASON CHANGE'
        else:status='MATCH'
        stat[status]+=1
        details.append([ss[0] if ss else (ps[0] if ps else ''),a[0][1] if a else (b[0][1] if b else ''),k[0],k[1],k[2],cname,', '.join(dates_a),', '.join(dates_b),', '.join(fa),', '.join(fb),oq,pq,oq-pq,status,', '.join(ps)])
        for r in a:byfactory[r[7]][0]+=r[8];byseason[r[0]][0]+=r[8]
        for r in b:byfactory[r[6]][1]+=r[7];byseason[r[0]][1]+=r[7]
        if status!='MATCH':
            for f in set(fa+fb):byfactory[f][2]+=1
            for s in set(ss+ps):byseason[s][2]+=1
    wb=Workbook()
    wb.remove(wb.active)
    total_o=sum(r[8] for r in order);total_p=sum(r[7] for r in po)
    output_sheet(wb,'Dashboard',['Metric','Value'],[['Order Qty (7-digit PO)',total_o],['PO Detailed Qty',total_p],['Net Difference',total_o-total_p],['Excluded Order Rows',len(excluded)],['Excluded Qty',sum(r[8] for r in excluded)],['Compared Keys',len(details)]]+[[x,stat[x]] for x in ['MATCH','QTY DIFF','DATE CHANGE','FACTORY CHANGE','SEASON CHANGE','ONLY ORDER','ONLY PO']],[34,22])
    output_sheet(wb,'Discrepancy',['Season','Division','Style#','PO#','Color Code','Color Name (Order)','Ex-Fty','H/O','Order Factory','PO Factory','Order Qty','PO Qty','Difference','Status','PO Season'],details,[16,28,13,14,14,34,22,22,23,23,15,15,15,21,16])
    output_sheet(wb,'Factory Summary',['Factory','Order Qty','PO Qty','Difference','Non-match keys'],[[k,v[0],v[1],v[0]-v[1],v[2]] for k,v in sorted(byfactory.items())],[30,18,18,18,19])
    output_sheet(wb,'Season Summary',['Season','Order Qty','PO Qty','Difference','Non-match keys'],[[k,v[0],v[1],v[0]-v[1],v[2]] for k,v in sorted(byseason.items())],[20,18,18,18,19])
    output_sheet(wb,'Color Review',['Style#','Color Code','Conflicting Order Color Names'],review,[15,16,80])
    output_sheet(wb,'Excluded Orders',['Season','Division','Style#','PO#','Color Code','Color Name','Ex-Fty','Factory','Qty','Reason'],excluded,[15,28,14,16,15,32,18,24,15,28])
    output_sheet(wb,'Order Source',['Season','Division','Style#','PO#','Color Code','Color Name','Ex-Fty','Factory','Order Qty'],order,[15,28,14,16,15,32,18,24,18])
    output_sheet(wb,'PO Source',['Season','Division','Style#','PO#','Color Code','H/O','Factory','PO Qty'],po,[15,28,14,16,15,18,24,18])
    wb.save(str(out_path))
    print('Saved:',out_path,'| Order:',total_o,'| PO:',total_p,'| Status:',dict(stat))

if __name__=='__main__':
    p=argparse.ArgumentParser(description='Compare Order Status and PO Detailed View XLSX')
    p.add_argument('order',type=Path);p.add_argument('po',type=Path);p.add_argument('-o','--output',type=Path,default=Path('PO_Discrepancy_Result.xlsx'))
    a=p.parse_args()
    run(a.order,a.po,a.output)
