#!/usr/bin/env python3
"""BU_CF「入力シート」→ Googleスプレッドシート変換用の1タブCSV（数式入り）

usage: python build_gsheet_csv.py <BU_CF.xlsx> <out.csv> [基準日 YYYY-MM-DD]
Drive create_file に textContent / contentMimeType=text/csv で渡すとGoogleスプレッドシートになる。
集計表はA〜J列、物件データはN〜X列（集計式はN:N等の列全体参照）。
"""
import sys, csv, io, os
from datetime import date
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_report as b
SRC, OUT = sys.argv[1], sys.argv[2]
BASE = sys.argv[3] if len(sys.argv) > 3 else date.today().isoformat()
rows=b.load(SRC)
sold=sorted([r for r in rows if r['sell']],key=lambda r:r['sell'])
from collections import Counter
types=[t for t,_ in Counter(r['type'] for r in sold).most_common()]
lo_d, hi_d = min(r['sell'] for r in sold), max(r['sell'] for r in sold)
P=[p for p in b.PERIODS if p[2] >= lo_d and p[1] <= hi_d]
np_=len(P); W=np_+2
out=[]
def add(*cells): out.append(list(cells))
add('ウィローズビルド 期別×種別 販売分析（BU_CF入力シートより／売却決済日で期を判定／基準日'+BASE+'）')
add('金額は万円。粗利＝販売価格−仕入原価−商品化原価（販売諸費用・金利は含めない）。粗利・粗利率は仕入原価入力済の物件のみ。状態=予定は基準日より後の決済（計画値）。12期は3ヶ月変則決算。')
add()
# placeholders; data rows computed later
nsum_rows_est=None
# build summary with symbolic data range, fill later
blocks=[]
PT0=None
lines=[]
def L(i): return b.L(i)
# first determine layout: period table
add('期','開始日','終了日','備考')
PT0=len(out)+1
for n,s,e,note in P: add(n,s.isoformat(),e.isoformat(),note)
PTN=len(out)
add()
summary_start=len(out)
# compute sizes to know data start
T=len(types)
size = 2+T+1+1   # ① title+hdr+types+total+blank
size += 6*(2+T+1+1)
size += 2+5+1 + 2+9+1+1 + 2
DS = summary_start + size + 3  # data header row
D0 = DS+1; DN = D0+len(sold)-1
TY,PER,PR,CO,GP,HO = 'N:N','T:T','R:R','S:S','U:U','V:V'
hdr=lambda first,extra='合計':add(first,*[p[0] for p in P],extra)
def block(title,fn,totfn):
    add(title); hr=len(out)+1; hdr('物件種別')
    top=len(out)+1
    for t in types:
        r=len(out)+1
        add(t,*[fn(f'$A{r}',f'{L(2+j)}${hr}') for j in range(np_)],fn(f'$A{r}',None))
    r=len(out)+1
    add('合計',*[totfn(top,r,L(2+j),f'{L(2+j)}${hr}') for j in range(np_)],totfn(top,r,L(2+np_),None))
    add()
pc=lambda p: f'{PER},{p}' if p else f'{PER},"?*"'
sumup=lambda top,r,c,p: f'=SUM({c}{top}:{c}{r-1})'
block('① 期ごとの物件種別（棟数）', lambda t,p:f'=COUNTIFS({TY},{t},{pc(p)})', sumup)
block('② 売上高（万円）', lambda t,p:f'=ROUND(SUMIFS({PR},{TY},{t},{pc(p)})/10000,0)', sumup)
block('② 粗利（万円）※原価入力済のみ', lambda t,p:f'=ROUND(SUMIFS({GP},{TY},{t},{pc(p)})/10000,0)', sumup)
def gpr(t,p):
    tc=f'{TY},{t},' if t else ''
    return f'=IFERROR(ROUND(SUMIFS({GP},{tc}{pc(p)})/SUMIFS({PR},{tc}{pc(p)},{CO},">0")*100,1),"")'
block('② 粗利率（%）※原価入力済のみ', gpr, lambda top,r,c,p: gpr(None,p))
def hold(t,p):
    tc=f'{TY},{t},' if t else ''
    return f'=IFERROR(ROUND(AVERAGEIFS({HO},{tc}{pc(p)},{HO},">=0"),0),"")'
block('② 平均保有日数（日）', hold, lambda top,r,c,p: hold(None,p))
def miss(t,p): return f'=COUNTIFS({TY},{t},{pc(p)})-COUNTIFS({TY},{t},{pc(p)},{CO},">0")'
block('参考：原価未入力で粗利対象外の棟数', miss, sumup)
add('③ 期ごとの平均成約価格（万円）'); hr=len(out)+1; hdr('指標','全期')
pp=lambda j: f'{L(2+j)}${hr}' if j<np_ else None
add('棟数（価格あり）',*[f'=COUNTIFS({pc(pp(j))},{PR},">0")' for j in range(np_+1)])
add('平均成約価格',*[f'=IFERROR(ROUND(AVERAGEIFS({PR},{pc(pp(j))},{PR},">0")/10000,0),"")' for j in range(np_+1)])
add('最低価格',*[f'=IFERROR(ROUND(MINIFS({PR},{pc(pp(j))},{PR},">0")/10000,0),"")' for j in range(np_+1)])
add('最高価格',*[f'=IFERROR(ROUND(MAXIFS({PR},{pc(pp(j))},{PR},">0")/10000,0),"")' for j in range(np_+1)])
add('売上高合計',*[f'=ROUND(SUMIFS({PR},{pc(pp(j))})/10000,0)' for j in range(np_+1)])
add()
add('③ 価格分布（5,000万円きざみ・棟数）'); hr=len(out)+1; add('価格帯',*[p[0] for p in P],'全期','下限(万円)','上限(万円)')
lo,hi=L(np_+3),L(np_+4)
top=len(out)+1
NB=int(max((r['price'] or 0) for r in sold)//50000000)+1
for k in range(NB):
    r=len(out)+1
    yl=lambda m: f'{m/10000:.1f}億' if m>=10000 else f'{m:,}万'
    lab=('〜5,000万未満' if k==0 else f'{yl(k*5000)}〜{yl((k+1)*5000)}未満')
    lo_c = '">0"' if k==0 else f'">="&${lo}{r}*10000'
    add(lab,*[f'=COUNTIFS({pc(pp(j))},{PR},{lo_c},{PR},"<"&${hi}{r}*10000)' for j in range(np_+1)],k*5000,(k+1)*5000)
r=len(out)+1
add('合計',*[f'=SUM({L(2+j)}{top}:{L(2+j)}{r-1})' for j in range(np_+1)])
add()
DR0=4
def per_of(d):
    for n,a,e,_ in P:
        if a<=d<=e: return n
    return ''
data=[['物件種別','物件名','仕入決済日','売却決済日','販売価格(円)','原価計(円)','期','粗利(円)','保有日数','状態','データ注記']]
for i,x in enumerate(sold):
    r=DR0+1+i
    cost=(x['buy_cost']+x['prod_cost']) if x['buy_cost'] else ''
    data.append([x['type'],x['name'],x['buy'].isoformat() if x['buy'] else '',x['sell'].isoformat(),x['price'] or '',cost,
      per_of(x['sell']),
      (f'=R{r}-S{r}' if (x['price'] and cost!='') else ''), (f'=Q{r}-P{r}' if x['buy'] else ''),
      ('実績' if x['sell'].isoformat()<=BASE else '予定'), x['note']])
while len(out)<DR0-1+len(data): out.append([])
for i,d in enumerate(data):
    row=out[DR0-1+i]
    row+=['']*(12-len(row)); row[:]=row[:12]+['']+d if len(row)<=12 else row
s=io.StringIO(); csv.writer(s,lineterminator='\n').writerows(out)
open(OUT,'w').write(s.getvalue())
print(len(s.getvalue()), 'chars', len(out),'rows')
