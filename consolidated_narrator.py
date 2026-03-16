"""
Consolidated Variance Analysis AI Agent (v5)
=============================================
Reads dept submission workbook with Qtr PL, YTD PL, and BS tabs per department.
Consolidates, identifies material variances, generates AI commentary.

Output: Consolidated Qtr PL, Consolidated YTD PL, Consolidated BS, Dept Summary

Usage:
    python consolidated_narrator.py dept_submissions_input.xlsx
    python consolidated_narrator.py dept_submissions_input.xlsx -o Q2_report.xlsx
    python consolidated_narrator.py dept_submissions_input.xlsx --api-key sk-ant-xxx
"""
import sys,os,json,argparse,getpass
from datetime import datetime
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter
try:
    import anthropic; HAS_API=True
except ImportError:
    HAS_API=False; print("WARNING: pip install anthropic for AI commentary")

MPCT=0.05; MABS=500000; MODEL="claude-sonnet-4-20250514"
B=Font(name='Arial',color='000000',size=10); BB=Font(name='Arial',color='000000',size=10,bold=True)
WB=Font(name='Arial',color='FFFFFF',size=11,bold=True); TF=Font(name='Arial',color='000000',size=14,bold=True)
SF=Font(name='Arial',color='FFFFFF',size=10,bold=True); NF=Font(name='Arial',color='666666',size=9,italic=True)
GF=Font(name='Arial',color='006100',size=10); RF=Font(name='Arial',color='9C0006',size=10)
OF=Font(name='Arial',color='FF6600',size=10,bold=True); IF2=Font(name='Arial',color='999999',size=10,italic=True)
CF=Font(name='Arial',color='1F4E79',size=10); DF=Font(name='Arial',color='003366',size=9)
GB=PatternFill('solid',fgColor='C6EFCE'); RB=PatternFill('solid',fgColor='FFC7CE')
OB=PatternFill('solid',fgColor='FFF2CC'); YB=PatternFill('solid',fgColor='F2F2F2')
DB2=PatternFill('solid',fgColor='4472C4'); DH=PatternFill('solid',fgColor='44546A')
DG=PatternFill('solid',fgColor='548235'); DP=PatternFill('solid',fgColor='7030A0')
TB=Border(left=Side('thin'),right=Side('thin'),top=Side('thin'),bottom=Side('thin'))
AF2='#,##0;(#,##0);"-"'; PF2='0.0%;(0.0%);"-"'

def get_key(ck=None):
    if ck: return ck
    ek=os.environ.get('ANTHROPIC_API_KEY')
    if ek: print("  Using API key from env var."); return ek
    sd=os.path.dirname(os.path.abspath(__file__))
    for fname in ['api_key.txt','API_Key.txt','API_key.txt','Api_Key.txt','apikey.txt','APIKEY.txt']:
        kf=os.path.join(sd,fname)
        if os.path.exists(kf):
            with open(kf) as f: fk=f.read().strip()
            if fk: print(f"  Using API key from {kf}"); return fk
    print("\n  No API key found. Options:")
    print("  1) Enter now  2) Press Enter to skip  3) Put key in API_Key.txt in same folder\n")
    k=getpass.getpass("  Paste key (hidden) or Enter to skip: ").strip()
    return k if k else None

def cv(c,p,ie=False):
    v=-(c-p) if ie else c-p; r=v/abs(p) if p!=0 else None; return v,r
def im(v,p,mp=MPCT,ma=MABS): return abs(v)>=ma or(p is not None and abs(p)>=mp)

def parse(fp):
    wb=load_workbook(fp,data_only=True); ds={}
    for sn in wb.sheetnames:
        ws=wb[sn]; iq=sn.endswith(' Qtr PL'); iy=sn.endswith(' YTD PL'); ib=sn.endswith(' BS')
        if not(iq or iy or ib): continue
        dn=sn.rsplit(' ',2)[0] if(iq or iy) else sn.rsplit(' ',1)[0]
        if dn not in ds: ds[dn]={'qtr':[],'ytd':[],'bs':[],'type':'Unknown'}
        hr=None
        for r in range(1,20):
            v=ws.cell(r,1).value
            if v and str(v).strip()=='Account': hr=r; break
            if v and 'Type:' in str(v):
                if 'Operating' in str(v): ds[dn]['type']='Operating'
                elif 'Back Office' in str(v): ds[dn]['type']='Back Office'
        if hr is None: continue
        cs=None
        for r in range(hr+1,ws.max_row+1):
            a=ws.cell(r,1).value
            if not a or not str(a).strip(): continue
            a=str(a).strip()
            if a.startswith('Total ') or a in('GROSS PROFIT','SEGMENT OPERATING INCOME'): continue
            secs=('Revenue','Cost of Revenue','Segment Operating Expenses','Operating Expenses',
                  'Revenue / Other Income','Segment Assets','Segment Liabilities','Assets','Liabilities')
            if a in secs: cs=a; continue
            if iq:
                try: q2=float(ws.cell(r,2).value or 0); q1=float(ws.cell(r,3).value or 0); q2y=float(ws.cell(r,4).value or 0)
                except: continue
                ds[dn]['qtr'].append({'account':a,'section':cs or'Other','q2_25':q2,'q1_25':q1,'q2_24':q2y,
                    'qoq_expl':str(ws.cell(r,7).value or'').strip(),'yoy_expl':str(ws.cell(r,10).value or'').strip()})
            elif iy:
                try: h1=float(ws.cell(r,2).value or 0); h1y=float(ws.cell(r,3).value or 0)
                except: continue
                ds[dn]['ytd'].append({'account':a,'section':cs or'Other','h1_25':h1,'h1_24':h1y,
                    'ytd_expl':str(ws.cell(r,6).value or'').strip()})
            elif ib:
                try: j=float(ws.cell(r,2).value or 0); d=float(ws.cell(r,3).value or 0)
                except: continue
                ds[dn]['bs'].append({'account':a,'section':cs or'Other','jun':j,'dec':d,
                    'expl':str(ws.cell(r,6).value or'').strip()})
    wb.close(); return ds

def consol(ds,dk,afs,efs):
    c={}
    for dn,dd in ds.items():
        for i in dd[dk]:
            a=i['account']
            if a not in c:
                c[a]={'section':i['section'],'depts':[]}
                for f in afs: c[a][f]=0
                for f in efs: c[a][f]=[]
            for f in afs: c[a][f]+=i.get(f,0)
            for f in efs:
                v=i.get(f,'')
                if v and v!='Immaterial' and v!='Immaterial.' and v!='No variance.' and v!='No variance':
                    c[a][f].append(f'[{dn}] {v}')
            c[a]['depts'].append(dn)
    so=['Revenue','Revenue / Other Income','Cost of Revenue','Segment Operating Expenses',
        'Operating Expenses','Segment Assets','Assets','Segment Liabilities','Liabilities','Other']
    r=[]; s=set()
    for sec in so:
        items=[(k,v) for k,v in c.items() if v['section']==sec and k not in s]
        if items:
            r.append(('__S__',sec))
            for k,v in items: r.append((k,v)); s.add(k)
    rem=[(k,v) for k,v in c.items() if k not in s]
    if rem:
        r.append(('__S__','Other'))
        for k,v in rem: r.append((k,v))
    return r

SYS="""You are a senior corporate controller at a public tech company preparing quarterly variance analysis for 10-Q.
Generating commentary for consolidated {st}.
MATERIALITY: >{mp}%% or >${ma}K

You will return TWO fields per account in JSON: "explanation" and "followup".
- "explanation": The synthesized variance narrative. Clean, professional, no issues flagged here.
- "followup": ONLY populated if there is a problem. Empty string "" if everything checks out.

EXPLANATION RULES:
1. COVER 80%%+ of variance in dollars. State coverage percentage at the end.
2. EVERY driver gets a dollar amount (+$X,XXXK or -$X,XXXK). No driver without dollars.
3. ORDER BY MAGNITUDE. Largest dollar impact first.
4. DEPARTMENT ATTRIBUTION: "[Dept] (+/-$X,XXXK): explanation" — rank by contribution size.
5. USE department explanations. They have real deal names, rates, customer counts. Synthesize into one narrative.
6. Put ONLY the variance explanation in the "explanation" field. No flags, no caveats, no follow-up items.

FOLLOW-UP FLAG RULES (populate "followup" field ONLY when):
1. Department explanations do not cover 80%%+ of the dollar variance — state what %% is explained and what $$ remains.
2. Department explanation is DIRECTIONALLY WRONG — e.g., says "favorable" or "decreased" when the numbers show the opposite. Call out the specific error.
3. Department dollar amounts DON'T ADD UP — the math in the explanation doesn't reconcile to the actual variance. Show the discrepancy.
4. Department explanation contradicts the actual data — e.g., claims balance decreased but it actually increased.
5. Coverage claim is wrong — e.g., says "covers ~80%%" but the actual math shows different coverage.
6. If NONE of these issues exist, set followup to "" (empty string).

RESPONSE FORMAT: ONLY valid JSON. No markdown, no backticks.
{{"account_name": {{"explanation": "...", "followup": "..."}}, ...}}"""

def ai_gen(items,st,mp,ma):
    if not HAS_API: return {}
    ak=os.environ.get('ANTHROPIC_API_KEY')
    if not ak: return {}
    cl=anthropic.Anthropic(api_key=ak)
    try:
        r=cl.messages.create(model=MODEL,max_tokens=8000,
            system=SYS.format(st=st,mp=f"{mp*100:.0f}",ma=f"{ma/1000:.0f}"),
            messages=[{"role":"user","content":f"Generate variance commentary:\n\n{json.dumps(items,indent=2)}"}])
        raw=r.content[0].text.strip()
        if raw.startswith('```'): raw=raw.split('\n',1)[1]
        if raw.endswith('```'): raw=raw[:-3]
        return json.loads(raw.strip())
    except Exception as e: print(f"  AI error: {e}"); return {}

def get_ai(cq,cy,cb,mp,ma):
    cm={}
    def mk_items(data,af1,af2,ef):
        items=[]
        for a,d in data:
            if a=='__S__': continue
            ie=d['section'] in('Cost of Revenue','Segment Operating Expenses','Operating Expenses')
            v,p=cv(d[af1],d[af2],ie)
            if im(v,p,mp,ma):
                items.append({'account':a,'current_period':d[af1],'prior_period':d[af2],
                    'variance':v,'pct':f"{p:.1%}" if p else"N/A",
                    'dept_explanations':d.get(ef,[]),
                    'depts':list(set(d['depts']))})
        return items
    qi=mk_items(cq,'q2_25','q1_25','qoq_expl'); print(f"  QoQ: {len(qi)} material")
    cm['qoq']=ai_gen(qi,"Income Statement QoQ",mp,ma) if qi else {}
    yi=mk_items(cq,'q2_25','q2_24','yoy_expl'); print(f"  YoY: {len(yi)} material")
    cm['yoy']=ai_gen(yi,"Income Statement YoY",mp,ma) if yi else {}
    ti=mk_items(cy,'h1_25','h1_24','ytd_expl'); print(f"  YTD: {len(ti)} material")
    cm['ytd']=ai_gen(ti,"Income Statement YTD",mp,ma) if ti else {}
    bi=[]
    for a,d in cb:
        if a=='__S__': continue
        v=d['jun']-d['dec']; p=v/abs(d['dec']) if d['dec']!=0 else None
        if im(v,p,mp,ma):
            bi.append({'account':a,'jun_2025':d['jun'],'dec_2024':d['dec'],'change':v,
                'pct':f"{p:.1%}" if p else"N/A",
                'dept_explanations':d.get('expl',[]),
                'depts':list(set(d['depts']))})
    print(f"  BS:  {len(bi)} material")
    cm['bs']=ai_gen(bi,"Balance Sheet",mp,ma) if bi else {}
    return cm

def vc(ws,r,c,v,m):
    cl=ws.cell(r,c,v); cl.number_format=AF2; cl.border=TB
    if m: cl.font=GF if v>0 else RF; cl.fill=GB if v>0 else RB
    else: cl.font=B

def ec(ws,r,c,de,ai_resp,m):
    """Write explanation cell. ai_resp is either a string (old) or dict with 'explanation' key."""
    ai_expl = ''
    if isinstance(ai_resp, dict): ai_expl = ai_resp.get('explanation', '')
    elif isinstance(ai_resp, str): ai_expl = ai_resp
    if not m:
        cl=ws.cell(r,c,'Immaterial'); cl.font=IF2; cl.fill=YB
    elif ai_expl:
        cl=ws.cell(r,c,ai_expl); cl.font=CF
    elif de:
        cl=ws.cell(r,c,'\n'.join(de)); cl.font=DF
    else:
        cl=ws.cell(r,c,'No explanation provided'); cl.font=OF; cl.fill=OB
    cl.border=TB; cl.alignment=Alignment(wrap_text=True)

def fu_cell(ws,r,c,ai_resp,m):
    """Write follow-up cell. Only populated by AI when it finds issues."""
    ai_fu = ''
    if isinstance(ai_resp, dict): ai_fu = ai_resp.get('followup', '')
    elif isinstance(ai_resp, str) and ai_resp.startswith('NEEDS_FOLLOWUP'): ai_fu = ai_resp
    if ai_fu:
        cl=ws.cell(r,c,ai_fu); cl.font=OF; cl.fill=OB
    else:
        cl=ws.cell(r,c,''); cl.font=B
    cl.border=TB; cl.alignment=Alignment(wrap_text=True)

def build(cq,cy,cb,cm,ds,op,mp,ma):
    wb=Workbook()
    ts=datetime.now().strftime("%Y-%m-%d %H:%M")
    # -- Qtr PL --
    ws=wb.active; ws.title='Consolidated Qtr PL'; ws.sheet_properties.tabColor='4472C4'
    ws.merge_cells('A1:K1'); ws.cell(1,1,'CONSOLIDATED QUARTERLY P&L — Variance Analysis').font=TF
    ws.merge_cells('A2:K2'); ws.cell(2,1,f'Q2 2025 | Materiality >{mp*100:.0f}% or >${ma/1000:.0f}K | {ts}').font=NF
    r=4
    for i,h in enumerate(['Account','Q2 2025','Q1 2025','Q2 2024','QoQ $ Var','QoQ %','QoQ Explanation','YoY $ Var','YoY %','YoY Explanation','Follow-Up Flags'],1):
        c=ws.cell(r,i,h); c.font=WB; c.fill=DB2; c.border=TB; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[r].height=40; r+=1
    for a,d in cq:
        if a=='__S__':
            ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11)
            ws.cell(r,1,d).font=SF
            for cc in range(1,12): ws.cell(r,cc).fill=DH; ws.cell(r,cc).border=TB
            r+=1; continue
        ie=d['section'] in('Cost of Revenue','Segment Operating Expenses','Operating Expenses')
        qv,qp=cv(d['q2_25'],d['q1_25'],ie); yv,yp=cv(d['q2_25'],d['q2_24'],ie)
        qm=im(qv,qp,mp,ma); ym=im(yv,yp,mp,ma)
        ws.cell(r,1,a).font=B; ws.cell(r,1).border=TB
        for ci,v in[(2,d['q2_25']),(3,d['q1_25']),(4,d['q2_24'])]:
            ws.cell(r,ci,v).font=B; ws.cell(r,ci).number_format=AF2; ws.cell(r,ci).border=TB
        vc(ws,r,5,qv,qm)
        c=ws.cell(r,6,qp if qp else 0); c.number_format=PF2; c.border=TB; c.font=B
        ec(ws,r,7,d.get('qoq_expl',[]),cm.get('qoq',{}).get(a,''),qm)
        vc(ws,r,8,yv,ym)
        c=ws.cell(r,9,yp if yp else 0); c.number_format=PF2; c.border=TB; c.font=B
        ec(ws,r,10,d.get('yoy_expl',[]),cm.get('yoy',{}).get(a,''),ym)
        # Follow-up: AI-generated only — combine QoQ and YoY flags
        qoq_ai = cm.get('qoq',{}).get(a,''); yoy_ai = cm.get('yoy',{}).get(a,'')
        qfu = qoq_ai.get('followup','') if isinstance(qoq_ai,dict) else ''
        yfu = yoy_ai.get('followup','') if isinstance(yoy_ai,dict) else ''
        all_fu = '\n'.join(filter(None, [f'QoQ: {qfu}' if qfu else '', f'YoY: {yfu}' if yfu else '']))
        c=ws.cell(r,11,all_fu)
        if all_fu: c.font=OF; c.fill=OB
        else: c.font=B
        c.border=TB; c.alignment=Alignment(wrap_text=True); r+=1
    ws.column_dimensions['A'].width=38
    for x in range(2,5): ws.column_dimensions[get_column_letter(x)].width=16
    ws.column_dimensions['E'].width=14; ws.column_dimensions['F'].width=10; ws.column_dimensions['G'].width=65
    ws.column_dimensions['H'].width=14; ws.column_dimensions['I'].width=10; ws.column_dimensions['J'].width=65
    ws.column_dimensions['K'].width=55; ws.freeze_panes='B5'

    # -- YTD PL --
    ws2=wb.create_sheet('Consolidated YTD PL'); ws2.sheet_properties.tabColor='7030A0'
    ws2.merge_cells('A1:G1'); ws2.cell(1,1,'CONSOLIDATED YTD P&L — Variance Analysis').font=TF
    ws2.merge_cells('A2:G2'); ws2.cell(2,1,f'H1 2025 vs H1 2024 | Materiality >{mp*100:.0f}% or >{ma/1000:.0f}K | {ts}').font=NF
    r=4
    for i,h in enumerate(['Account','H1 2025','H1 2024','YTD $ Var','YTD %','YTD Explanation','Follow-Up Flags'],1):
        c=ws2.cell(r,i,h); c.font=WB; c.fill=DP; c.border=TB; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws2.row_dimensions[r].height=40; r+=1
    for a,d in cy:
        if a=='__S__':
            ws2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7)
            ws2.cell(r,1,d).font=SF
            for cc in range(1,8): ws2.cell(r,cc).fill=DH; ws2.cell(r,cc).border=TB
            r+=1; continue
        ie=d['section'] in('Cost of Revenue','Segment Operating Expenses','Operating Expenses')
        v,p=cv(d['h1_25'],d['h1_24'],ie); m=im(v,p,mp,ma)
        ws2.cell(r,1,a).font=B; ws2.cell(r,1).border=TB
        ws2.cell(r,2,d['h1_25']).font=B; ws2.cell(r,2).number_format=AF2; ws2.cell(r,2).border=TB
        ws2.cell(r,3,d['h1_24']).font=B; ws2.cell(r,3).number_format=AF2; ws2.cell(r,3).border=TB
        vc(ws2,r,4,v,m)
        c=ws2.cell(r,5,p if p else 0); c.number_format=PF2; c.border=TB; c.font=B
        ec(ws2,r,6,d.get('ytd_expl',[]),cm.get('ytd',{}).get(a,''),m)
        ytd_ai = cm.get('ytd',{}).get(a,'')
        yfu = ytd_ai.get('followup','') if isinstance(ytd_ai,dict) else ''
        c=ws2.cell(r,7,yfu)
        if yfu: c.font=OF; c.fill=OB
        else: c.font=B
        c.border=TB; c.alignment=Alignment(wrap_text=True); r+=1
    ws2.column_dimensions['A'].width=38; ws2.column_dimensions['B'].width=16; ws2.column_dimensions['C'].width=16
    ws2.column_dimensions['D'].width=14; ws2.column_dimensions['E'].width=10; ws2.column_dimensions['F'].width=70
    ws2.column_dimensions['G'].width=55; ws2.freeze_panes='B5'

    # -- BS --
    ws3=wb.create_sheet('Consolidated BS'); ws3.sheet_properties.tabColor='548235'
    ws3.merge_cells('A1:G1'); ws3.cell(1,1,'CONSOLIDATED BALANCE SHEET — Variance Analysis').font=TF
    ws3.merge_cells('A2:G2'); ws3.cell(2,1,f'Jun 30 2025 vs Dec 31 2024 | Materiality >{mp*100:.0f}% or >{ma/1000:.0f}K | {ts}').font=NF
    r=4
    for i,h in enumerate(['Account','Jun 30 2025','Dec 31 2024','Change $','Change %','Explanation','Follow-Up Flags'],1):
        c=ws3.cell(r,i,h); c.font=WB; c.fill=DG; c.border=TB; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws3.row_dimensions[r].height=40; r+=1
    for a,d in cb:
        if a=='__S__':
            ws3.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7)
            ws3.cell(r,1,d).font=SF
            for cc in range(1,8): ws3.cell(r,cc).fill=DH; ws3.cell(r,cc).border=TB
            r+=1; continue
        v=d['jun']-d['dec']; p=v/abs(d['dec']) if d['dec']!=0 else None; m=im(v,p,mp,ma)
        ws3.cell(r,1,a).font=B; ws3.cell(r,1).border=TB
        ws3.cell(r,2,d['jun']).font=B; ws3.cell(r,2).number_format=AF2; ws3.cell(r,2).border=TB
        ws3.cell(r,3,d['dec']).font=B; ws3.cell(r,3).number_format=AF2; ws3.cell(r,3).border=TB
        vc(ws3,r,4,v,m)
        c=ws3.cell(r,5,p if p else 0); c.number_format=PF2; c.border=TB; c.font=B
        ec(ws3,r,6,d.get('expl',[]),cm.get('bs',{}).get(a,''),m)
        bs_ai = cm.get('bs',{}).get(a,'')
        bfu = bs_ai.get('followup','') if isinstance(bs_ai,dict) else ''
        c=ws3.cell(r,7,bfu)
        if bfu: c.font=OF; c.fill=OB
        else: c.font=B
        c.border=TB; c.alignment=Alignment(wrap_text=True); r+=1
    ws3.column_dimensions['A'].width=38; ws3.column_dimensions['B'].width=18; ws3.column_dimensions['C'].width=18
    ws3.column_dimensions['D'].width=16; ws3.column_dimensions['E'].width=10; ws3.column_dimensions['F'].width=70
    ws3.column_dimensions['G'].width=55; ws3.freeze_panes='B5'

    # -- Summary --
    ws4=wb.create_sheet('Dept Summary'); ws4.sheet_properties.tabColor='FF6600'
    ws4.cell(1,1,'DEPARTMENT SUMMARY').font=TF; r=3
    for i,h in enumerate(['Department','Type','Qtr Items','YTD Items','BS Items'],1):
        c=ws4.cell(r,i,h); c.font=WB; c.fill=PatternFill('solid',fgColor='FF6600'); c.border=TB
    r+=1
    for dn,dd in ds.items():
        ws4.cell(r,1,dn).font=BB; ws4.cell(r,1).border=TB
        ws4.cell(r,2,dd['type']).font=B; ws4.cell(r,2).border=TB
        ws4.cell(r,3,len(dd['qtr'])).font=B; ws4.cell(r,3).border=TB
        ws4.cell(r,4,len(dd['ytd'])).font=B; ws4.cell(r,4).border=TB
        ws4.cell(r,5,len(dd['bs'])).font=B; ws4.cell(r,5).border=TB
        r+=1
    ws4.column_dimensions['A'].width=28; ws4.column_dimensions['B'].width=14
    for x in range(3,6): ws4.column_dimensions[get_column_letter(x)].width=14
    wb.save(op); print(f"\nSaved: {op}")

def main():
    pa=argparse.ArgumentParser(description='Consolidated Variance AI Agent v5',formatter_class=argparse.RawDescriptionHelpFormatter)
    pa.add_argument('input_file',help='Dept submissions Excel file')
    pa.add_argument('-o','--output',default='consolidated_variance_report.xlsx',help='Output file')
    pa.add_argument('--materiality-pct',type=float,default=MPCT,help='Materiality pct (default 0.05)')
    pa.add_argument('--materiality-abs',type=float,default=MABS,help='Materiality dollars (default 500000)')
    pa.add_argument('--api-key',default=None,help='Anthropic API key')
    pa.add_argument('--no-ai',action='store_true',help='Skip AI')
    a=pa.parse_args()
    if not os.path.exists(a.input_file): print(f"ERROR: {a.input_file} not found"); sys.exit(1)
    mp=a.materiality_pct; ma=a.materiality_abs
    print("="*60); print("CONSOLIDATED VARIANCE ANALYSIS - AI AGENT v5"); print("="*60)
    print(f"Input:  {a.input_file}\nOutput: {a.output}\nMateriality: >{mp*100:.0f}% OR >${ma/1000:.0f}K\n")
    print("Step 1: Parsing..."); ds=parse(a.input_file)
    for n,d in ds.items(): print(f"  {n} ({d['type']}): {len(d['qtr'])} qtr, {len(d['ytd'])} ytd, {len(d['bs'])} bs")
    print("\nStep 2: Consolidating...")
    cq=consol(ds,'qtr',['q2_25','q1_25','q2_24'],['qoq_expl','yoy_expl'])
    cy=consol(ds,'ytd',['h1_25','h1_24'],['ytd_expl'])
    cb=consol(ds,'bs',['jun','dec'],['expl'])
    print(f"  Qtr: {sum(1 for x,_ in cq if x!='__S__')} | YTD: {sum(1 for x,_ in cy if x!='__S__')} | BS: {sum(1 for x,_ in cb if x!='__S__')}")
    cm={}
    if a.no_ai: print("\nStep 3: Skipped (--no-ai)")
    elif not HAS_API: print("\nStep 3: Skipped (no anthropic)")
    else:
        print("\nStep 3: AI commentary...")
        ak=get_key(a.api_key)
        if ak: os.environ['ANTHROPIC_API_KEY']=ak; cm=get_ai(cq,cy,cb,mp,ma)
        else: print("  No key provided.")
    print("\nStep 4: Building output...")
    build(cq,cy,cb,cm,ds,a.output,mp,ma)
    print(f"\n{'='*60}\nDONE | {len(ds)} depts | Output: {a.output}\n{'='*60}")

if __name__=='__main__': main()
