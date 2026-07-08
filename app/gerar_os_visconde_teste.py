import os, re, sys, shutil, subprocess, unicodedata
from pathlib import Path
from datetime import datetime
import pandas as pd

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.colors import HexColor
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import Table, TableStyle, Paragraph
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.graphics.barcode.code128 import Code128
    from reportlab.lib.utils import ImageReader
except Exception as exc:
    raise SystemExit("A dependência reportlab não está instalada. Execute o instalador da OS Visconde TESTE.\n" + str(exc))

try:
    from pypdf import PdfReader, PdfWriter
except Exception as exc:
    raise SystemExit("A dependência pypdf não está instalada. Execute novamente o instalador principal.\n" + str(exc))

from caminho_base import BASE_DIR, RECURSOS_DIR

DOWNLOADS = BASE_DIR / "downloads" / "roteirizacao"
ROTEIRIZACAO = BASE_DIR / "outputs" / "roteirizacao" / "roteirizacao_atual.xlsx"
OUT = BASE_DIR / "outputs" / "os_visconde_teste"
IND = OUT / "individuais"
POR_TEC = OUT / "por_tecnico"
LOGO = RECURSOS_DIR / "assets" / "logo_visconde_os_mono.png"

W,H=A4; M=24; CW=W-2*M
INK=HexColor('#222222'); LINE=HexColor('#666666'); PALE=HexColor('#F2F2F2'); PALE2=HexColor('#FAFAFA'); GOLD=HexColor('#C79B12')

font_dir = Path('/usr/share/fonts/truetype/dejavu')
if os.name == 'nt':
    font_dir = Path(os.environ.get('WINDIR','C:/Windows')) / 'Fonts'
    regular = font_dir / 'arial.ttf'; bold = font_dir / 'arialbd.ttf'
elif sys.platform == 'darwin':
    regular = Path('/System/Library/Fonts/Supplemental/Arial.ttf')
    bold = Path('/System/Library/Fonts/Supplemental/Arial Bold.ttf')
else:
    regular = font_dir/'DejaVuSans.ttf'; bold = font_dir/'DejaVuSans-Bold.ttf'
if not regular.exists(): regular = font_dir/'DejaVuSans.ttf'
if not bold.exists(): bold = font_dir/'DejaVuSans-Bold.ttf'
pdfmetrics.registerFont(TTFont('DV',str(regular)))
pdfmetrics.registerFont(TTFont('DVB',str(bold)))
styles={
 'section':ParagraphStyle('s',fontName='DVB',fontSize=9.1,leading=10.5,textColor=INK),
 'cell':ParagraphStyle('c',fontName='DV',fontSize=7.0,leading=8.3,textColor=INK),
 'small':ParagraphStyle('sm',fontName='DV',fontSize=6.4,leading=7.5,textColor=INK),
 'center':ParagraphStyle('ce',fontName='DV',fontSize=6.5,leading=7.6,textColor=INK,alignment=TA_CENTER),
 'sign':ParagraphStyle('si',fontName='DV',fontSize=6.4,leading=7.4,textColor=INK,alignment=TA_CENTER),
}

def norm(s):
    s='' if s is None else str(s)
    s=unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode('ascii')
    return re.sub(r'[^A-Z0-9]+',' ',s.upper()).strip()
def val(row,*names):
    idx={norm(c):c for c in row.index}
    for name in names:
        c=idx.get(norm(name))
        if c is not None:
            v=row.get(c)
            if pd.notna(v) and str(v).strip() and str(v).strip().lower()!='nan':
                return str(v).strip().replace('="','').rstrip('"')
    return ''
def digits(v): return re.sub(r'\D','',str(v or ''))
def docfmt(v):
    d=digits(v)
    if len(d)==11: return f'{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}'
    if len(d)==14: return f'{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}'
    return v or ''
def p(t,st='cell'): return Paragraph(str(t or '').replace('&','&amp;').replace('\n','<br/>'),styles[st])
def cell(label,value=''): return p(f'<b>{label}</b><br/>{value if value else "________________________"}')
def no_line_cell(label,value=''):
    return p(f'<b>{label}</b><br/>{value if value else ""}')
def barcode(c,v,x,y,w=145,h=23):
    bc=Code128(str(v),barWidth=.75,barHeight=h,humanReadable=False,quiet=True); s=min(1,w/bc.width)
    c.saveState(); c.translate(x+(w-bc.width*s)/2,y); c.scale(s,1); bc.drawOn(c,0,0); c.restoreState()
def header(c,op,pl,pv,sl='',sv=''):
    y=H-88
    if LOGO.exists():
        c.drawImage(ImageReader(str(LOGO)), M, y+22, 160, 52, mask='auto', preserveAspectRatio=True)

    title_left = M + 175
    title_right = W - M - 190
    title_cx = (title_left + title_right) / 2

    c.setFillColor(INK)
    c.setFont('DVB', 11.2)
    c.drawCentredString(title_cx, y+56, 'ORDEM DE SERVIÇO')

    pill_w = 158
    pill_h = 22
    pill_x = title_cx - pill_w/2
    pill_y = y + 24
    c.setStrokeColor(GOLD)
    c.roundRect(pill_x, pill_y, pill_w, pill_h, 10, stroke=1, fill=0)
    c.setFont('DVB', 7.4)
    c.drawCentredString(title_cx, pill_y + 7.2, op[:34])

    bx = W - M - 178
    by = y + 20
    bw = 178
    bh = 54
    c.roundRect(bx, by, bw, bh, 8, stroke=1, fill=0)
    c.setFont('DVB', 6.6)
    c.drawString(bx + 10, by + bh - 13, pl.upper())
    barcode(c, pv, bx + 12, by + 18, bw - 24, 21)
    c.setFont('DVB', 8)
    c.drawCentredString(bx + bw/2, by + 5.5, str(pv))

    if sv:
        c.setFont('DV', 5.8)
        c.setFillColor(LINE)
        c.drawRightString(W - M, y + 12, f'{sl}: {sv}')
        c.setFillColor(INK)

    c.setStrokeColor(GOLD)
    c.setLineWidth(1)
    c.line(M, y + 6, W - M, y + 6)
    return y
def table_section(c,top,title,rows,widths,spans=None,row_heights=None):
    n=len(widths); data=[[p(title,'section')]+['']*(n-1)]+rows
    t=Table(data,colWidths=widths,rowHeights=[17]+(row_heights or [None]*len(rows)))
    st=[('SPAN',(0,0),(-1,0)),('BACKGROUND',(0,0),(-1,0),PALE),('BOX',(0,0),(-1,-1),.65,LINE),('INNERGRID',(0,1),(-1,-1),.25,HexColor('#BDBDBD')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,1),(-1,-1),4),('BOTTOMPADDING',(0,1),(-1,-1),4)]
    for a,b in spans or []: st.append(('SPAN',a,b))
    t.setStyle(TableStyle(st)); _,h=t.wrap(CW,top-M); t.drawOn(c,M,top-h); return top-h-6
def strip(c,top,items):
    t=Table([[cell(a,b) for a,b in items]],colWidths=[CW/len(items)]*len(items)); t.setStyle(TableStyle([('BOX',(0,0),(-1,-1),.6,LINE),('INNERGRID',(0,0),(-1,-1),.25,HexColor('#BDBDBD')),('BACKGROUND',(0,0),(-1,-1),PALE2),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)])); _,h=t.wrap(CW,100); t.drawOn(c,M,top-h); return top-h-6
def hours(c,top,vals):
    headers=['Horário de funcionamento','Seg.','Ter.','Qua.','Qui.','Sex.','Sáb.','Dom.','Almoço']; row=[p('Funcionamento do EC','center')]+[p(f'{a}<br/>{b}','center') for a,b in vals]
    data=[[p('2. HORÁRIO DE FUNCIONAMENTO','section')]+['']*8,[p(x,'center') for x in headers],row]
    t=Table(data,colWidths=[CW*.21]+[CW*.09875]*8,rowHeights=[17,15,29]); t.setStyle(TableStyle([('SPAN',(0,0),(-1,0)),('BACKGROUND',(0,0),(-1,0),PALE),('BOX',(0,0),(-1,-1),.65,LINE),('INNERGRID',(0,1),(-1,-1),.25,HexColor('#BDBDBD')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(0,1),(-1,-1),'CENTER')])); _,h=t.wrap(CW,100); t.drawOn(c,M,top-h); return top-h-6
def signature_cell(label, blank_lines=4, line='____________________________________'):
    # Mantém a área superior livre para escrita manual e posiciona a linha
    # perto da parte inferior do campo, com a legenda logo abaixo.
    return p(('<br/>' * blank_lines) + f'{line}<br/>{label}','sign')

def term(c,top):
    txt='Declaro que fui informado sobre o serviço executado e que as informações desta Ordem de Serviço são verdadeiras. Concordo com o serviço realizado e com os dados preenchidos neste documento.'
    return table_section(
        c,top,'4. TERMO DE CONCORDÂNCIA',
        [[p(txt,'small'),''],
         [signature_cell('Nome do cliente',2),signature_cell('Documento (RG / CPF)',2)]],
        [CW*.5,CW*.5],spans=[((0,1),(1,1))],row_heights=[26,38]
    )

def closing(c,top):
    return table_section(
        c,top,'5. ENCERRAMENTO',
        [[signature_cell('Nome / documento do cliente',4),signature_cell('Assinatura do cliente',4)],
         [signature_cell('Nome e assinatura do técnico',4),signature_cell('Data, hora e carimbo / observação final',4)]],
        [CW*.5,CW*.5],row_heights=[62,62]
    )
def footer(c):
    c.setStrokeColor(GOLD); c.line(M,18,W-M,18); c.setFont('DV',5.6); c.setFillColor(LINE); c.drawString(M,7,'MODELO OPERACIONAL - TESTE PILOTO'); c.drawRightString(W-M,7,'Visconde Instalação e Manutenção')
def schedule_from_text(text):
    text=str(text or '')
    days=['monday','tuesday','wednesday','thursday','friday','saturday','sunday']
    out=[]
    for d in days:
        m=re.search(d+r'[^0-9]*(\d{1,2}:\d{2})\s*[-a]\s*(\d{1,2}:\d{2})',text,re.I)
        out.append((m.group(1),m.group(2)) if m else ('',''))
    if not any(a for a,b in out):
        m=re.search(r'(\d{1,2})\s*[h:]\s*(\d{0,2})\s*[-a]\s*(\d{1,2})\s*[h:]?\s*(\d{0,2}).{0,20}(SEG\s*[-A]\s*SAB|SEGUNDA\s*A\s*SABADO)',text,re.I)
        if m:
            a=f'{int(m.group(1)):02d}:{int(m.group(2) or 0):02d}'; b=f'{int(m.group(3)):02d}:{int(m.group(4) or 0):02d}'; out=[(a,b)]*6+[('','')]
    return out+[('','')]
def find_reports():
    mob=ogea=None
    for f in sorted(DOWNLOADS.glob('*'),key=lambda x:x.stat().st_mtime,reverse=True):
        if f.suffix.lower() not in {'.csv','.txt','.xlsx','.xls'}: continue
        try:
            if f.suffix.lower() in {'.xlsx','.xls'}:
                cols=pd.read_excel(f,nrows=0).columns
            else:
                enc='latin1' if 'REPORT_SERVICE_ORDER' in f.name.upper() else 'utf-8-sig'
                cols=pd.read_csv(f,sep=';',encoding=enc,nrows=0).columns
            n={norm(c) for c in cols}
            if 'CHAMADO' in n and 'NUMERO REFERENCIA' in n and mob is None: mob=f
            if 'CODIGO' in n and 'CONTRATANTE' in n and ogea is None: ogea=f
        except Exception: pass
    return mob,ogea
def read_report(path,kind):
    if path.suffix.lower() in {'.xlsx','.xls'}: return pd.read_excel(path,dtype=str)
    return pd.read_csv(path,sep=';',dtype=str,encoding='latin1' if kind=='MOBYAN' else 'utf-8-sig')
def write_os(path,kind,row,tech):
    c=canvas.Canvas(str(path),pagesize=A4)
    if kind=='MOBYAN':
        chamado=val(row,'Chamado'); ref=val(row,'Numero Referencia'); op=f"MOBYAN / {val(row,'Contratante') or 'GETNET'}"
        top=header(c,op,'CHAMADO',chamado,'REFERÊNCIA',ref)
        top=strip(c,top,[('Status',val(row,'Status')),('Prazo',val(row,'Data Limite')),('Técnico',tech or val(row,'Técnico')),('Prestador',val(row,'Prestador'))])
        top=table_section(c,top,'1. CLIENTE E LOCAL DE ATENDIMENTO',[[cell('Cliente',val(row,'Nome Cliente')),cell('CNPJ / CPF',docfmt(val(row,'CNPJ / CPF')))],[cell('Código do estabelecimento',val(row,'Cod. Cliente')),cell('Fantasia',val(row,'Nome Cliente'))],[cell('Endereço',val(row,'Endereço')),cell('Bairro / Cidade',f"{val(row,'Bairro')} - {val(row,'Cidade')}/{val(row,'Estado')}")],[cell('CEP',val(row,'CEP')),cell('Telefones',' | '.join(filter(None,[val(row,'Telefone 1'),val(row,'Telefone 2'),val(row,'Telefone 3')])))]],[CW*.54,CW*.46])
        sched=[('','')]*5+[(val(row,'Hora Inicio Sabado'),val(row,'Hora Termino Sabado')),('',''),('','')]
        top=hours(c,top,sched)
        problem=val(row,'Observações','Observações Atendimento')
        top=table_section(c,top,'3. SERVIÇO E DIAGNÓSTICO',[[cell('Serviço',val(row,'Serviço')),cell('Grupo',val(row,'Grupo Serviço'))],[cell('Problema relatado / Dados da OS',problem),'']],[CW*.55,CW*.45],spans=[((0,2),(1,2))],row_heights=[None,42])
        serial_cliente=val(row,'Serial Retirado') or val(row,'Serial Instalado')
        top=table_section(c,top,'3. EQUIPAMENTO E EXECUÇÃO',[[cell('Fabricante / modelo',val(row,'Modelo Retirado','Modelo Instalado')),cell('Número de série',val(row,'Serial Retirado','Serial Instalado'))],[cell('SERIAL DO CLIENTE / MÁQUINA ATUAL',serial_cliente),no_line_cell('Número do chip / operadora',val(row,'Operadora'))],[cell('Kit / Bobina',val(row,'Modelo Insumo','Modelo Instalado')),cell('Quantidade de kits',val(row,'Qtd. KIT'))],[no_line_cell('Solução aplicada',''),cell('Status final','□ Concluído  □ Insucesso  □ Reagendado')]], [CW*.56,CW*.44],row_heights=[None,38,34,40])
    else:
        code=val(row,'Código'); ext=val(row,'Id. Ext. do Contratante'); cont=val(row,'Contratante'); op=f'OGEA / {cont}'
        top=header(c,op,'OS',code,'INCIDENTE / TICKET',ext)
        top=strip(c,top,[('Status',val(row,'Status')),('Prazo',val(row,'Data Limite')),('Técnico',tech or val(row,'Técnico')),('Prestador',val(row,'Prestador de Serviço'))])
        top=table_section(c,top,'1. CLIENTE E LOCAL DE ATENDIMENTO',[[cell('Cliente',val(row,'Cliente')),cell('CNPJ / CPF',docfmt(val(row,'Documento')))],[cell('Endereço',f"{val(row,'Endereço')}, {val(row,'Número')}"),cell('Bairro / Cidade',f"{val(row,'Distrito')} - {val(row,'Cidade')}/{val(row,'Estado')}")],[cell('Contato',f"{val(row,'Contato')} - {val(row,'Número de Telefone')}"),cell('Código do estabelecimento',val(row,'Código do Cliente'))]],[CW*.56,CW*.44])
        top=hours(c,top,schedule_from_text(val(row,'Observação')))
        problem=val(row,'Defeito informado')
        obs=val(row,'Observação')
        if obs: problem=(problem+' | '+obs) if problem else obs
        top=table_section(c,top,'3. SERVIÇO E DIAGNÓSTICO',[[cell('Serviço',val(row,'Nome Alternativo do Serviço','Serviço')),cell('Grupo',val(row,'Grupo de Serviço'))],[cell('Problema relatado',problem),'']],[CW*.55,CW*.45],spans=[((0,2),(1,2))],row_heights=[None,44])
        serial=val(row,'Número de Série do Equipamento')
        top=table_section(c,top,'3. EQUIPAMENTO E EXECUÇÃO',[[cell('Fabricante / modelo',val(row,'Modelo do Equipamento')),cell('Número de série',serial)],[cell('Número lógico',val(row,'Número Lógico')),cell('Código de rede',val(row,'Código de Rede'))],[cell('SERIAL DO CLIENTE / MÁQUINA ATUAL',serial),no_line_cell('Número do chip / operadora',val(row,'ticketDetailsMobileOperator','Operadora'))],[no_line_cell('Serial retirado',''),no_line_cell('Serial instalado',val(row,'Número de Série do Novo Equipamento'))],[no_line_cell('Solução aplicada',''),cell('Status final','□ Concluído  □ Insucesso  □ Reagendado')]], [CW*.56,CW*.44],row_heights=[None,None,38,38,40])
    top=term(c,top); top=closing(c,top); footer(c); c.save()
def merge(paths,out):
    w=PdfWriter()
    for pth in paths:
        for pg in PdfReader(str(pth)).pages: w.add_page(pg)
    with open(out,'wb') as f: w.write(f)
def open_path(path):
    if os.name=='nt': os.startfile(str(path))
    elif sys.platform=='darwin': subprocess.Popen(['open',str(path)])
    else: subprocess.Popen(['xdg-open',str(path)])
def main():
    if not ROTEIRIZACAO.exists(): raise SystemExit('Gere a roteirização antes de executar o piloto.')
    mobf,ogeaf=find_reports()
    if not mobf and not ogeaf: raise SystemExit('Nenhum relatório atual foi encontrado em downloads/roteirizacao.')
    print('MODO PILOTO: o fluxo oficial não será alterado.')
    print('Mobyan:',mobf or 'não encontrado'); print('OGEA:',ogeaf or 'não encontrado')
    shutil.rmtree(OUT,ignore_errors=True); IND.mkdir(parents=True); POR_TEC.mkdir(parents=True)
    mob=read_report(mobf,'MOBYAN') if mobf else pd.DataFrame(); ogea=read_report(ogeaf,'OGEA') if ogeaf else pd.DataFrame()
    rot=pd.read_excel(ROTEIRIZACAO,sheet_name='Roteiro Geral',dtype=str)
    erros=[]; bytech={}; total=0
    for _,r in rot.iterrows():
        if norm(r.get('Resultado'))!='ROTEIRIZADA': continue
        kind=norm(r.get('Origem')); osnum=digits(r.get('OS')); tech=str(r.get('Técnico Roteirizado') or r.get('Técnico Atual') or 'Sem Técnico').strip()
        try:
            if kind=='MOBYAN':
                cand=mob[mob['Chamado'].astype(str).map(digits)==osnum] if not mob.empty else pd.DataFrame()
            elif kind=='OGEA':
                cand=ogea[ogea['Código'].astype(str).map(digits)==osnum] if not ogea.empty else pd.DataFrame()
            else: continue
            if cand.empty: raise ValueError('OS não encontrada no relatório bruto')
            out=IND/f'{kind} - {osnum}.pdf'; write_os(out,kind,cand.iloc[0],tech); bytech.setdefault(tech,[]).append(out); total+=1; print(f'OK {kind} {osnum} - {tech}')
        except Exception as e:
            erros.append({'Origem':kind,'OS':osnum,'Técnico':tech,'Erro':str(e)}); print(f'ERRO {kind} {osnum}: {e}')
    for tech,files in bytech.items():
        safe=re.sub(r'[\\/:*?"<>|]','-',tech); merge(files,POR_TEC/f'OS Visconde TESTE - {safe}.pdf')
    pd.DataFrame(erros,columns=['Origem','OS','Técnico','Erro']).to_excel(OUT/'relatorio_de_erros.xlsx',index=False)
    print('='*70); print('OSs Visconde geradas:',total); print('Pasta:',OUT); print('Erros:',len(erros)); print('O fluxo oficial permaneceu intacto.')
    open_path(OUT)
    return 0
if __name__=='__main__': raise SystemExit(main())
