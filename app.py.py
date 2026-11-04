"""GDS LOGÍSTICA | OPERAÇÃO ZARA - FREELANCERS v09
Executar: python -m streamlit run gds_extras_v09.py
Dependências: python -m pip install streamlit pandas openpyxl reportlab
Dados locais: gds_extras.db (mesmo banco das versões anteriores).
"""
import io
from decimal import Decimal, ROUND_HALF_UP
import re
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer

st.set_page_config(page_title="GDS | OPERAÇÃO ZARA - FREELANCERS", page_icon="🟠", layout="wide")
DB = Path(__file__).with_name('gds_extras.db')
ADMIN_PASSWORD = '230525'  # SOMENTE TESTE LOCAL; use autenticação segura antes de publicar.
VALOR_DIARIA = Decimal('170.00')
VALOR_HORA = Decimal('21.25')
INTERVALO_MINUTOS = 60
TURNOS = {'TURNO 1 — 10H ÀS 18H': ('10:00','18:00'), 'TURNO 2 — 15H ÀS 23H': ('15:00','23:00'), 'TURNO 3 — 19H ÀS 03H': ('19:00','03:00')}
STATUS = ['PENDENTE','CUMPRIU','HORÁRIO DIFERENTE','FALTA']

def upper(x): return ' '.join(str(x or '').upper().split())
def cpfnum(x): return re.sub(r'\D','',str(x or ''))
def valid_cpf(s):
    s=cpfnum(s)
    if len(s)!=11 or len(set(s))==1:return False
    for n in (9,10):
        d=(sum(int(s[i])*(n+1-i) for i in range(n))*10)%11
        if (0 if d==10 else d)!=int(s[n]):return False
    return True

def hora(s):
    t=str(s or '').strip().lower().replace('h',':').replace('.',':')
    if re.fullmatch(r'\d{1,2}',t):h,m=int(t),0
    elif re.fullmatch(r'\d{3,4}',t):t=t.zfill(4);h,m=int(t[:-2]),int(t[-2:])
    elif re.fullmatch(r'\d{1,2}:\d{1,2}',t):h,m=map(int,t.split(':'))
    else:raise ValueError('HORÁRIO INVÁLIDO. USE 10, 1030 OU 10:30.')
    if h>23 or m>59:raise ValueError('HORÁRIO DEVE ESTAR ENTRE 00:00 E 23:59.')
    return f'{h:02d}:{m:02d}'

def connect():
    c=sqlite3.connect(DB,timeout=20)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA busy_timeout=20000')
    c.execute('PRAGMA journal_mode=WAL')
    return c

def init():
    with connect() as c:
        c.execute('''CREATE TABLE IF NOT EXISTS prestadores(cpf TEXT PRIMARY KEY,nome TEXT NOT NULL,telefone TEXT NOT NULL DEFAULT '',tipo_pix TEXT NOT NULL DEFAULT '',pix TEXT NOT NULL DEFAULT '',diaria REAL NOT NULL DEFAULT 170,hora_extra REAL NOT NULL DEFAULT 21.25,ativo INTEGER NOT NULL DEFAULT 1)''')
        c.execute('''CREATE TABLE IF NOT EXISTS escalas(id INTEGER PRIMARY KEY AUTOINCREMENT,data TEXT NOT NULL,cpf TEXT NOT NULL,turno TEXT NOT NULL,status TEXT NOT NULL DEFAULT 'PENDENTE',entrada TEXT NOT NULL DEFAULT '',saida TEXT NOT NULL DEFAULT '',observacao TEXT NOT NULL DEFAULT '',extra INTEGER NOT NULL DEFAULT 0,atualizado TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,UNIQUE(data,cpf))''')
        # Mantém tabela de lançamentos das versões anteriores intacta.
        c.execute('CREATE INDEX IF NOT EXISTS ix_escalas_data ON escalas(data)')
init()

# Cadastro inicial: inserção idempotente, sem sobrescrever alterações posteriores.
# CPF ausente: identificador provisório interno, a completar pelo administrador.
CADASTRO_INICIAL = [
    ('PENDENTE:FABIANA', 'FABIANA CAROLINA DE BRITO VALENCIO', 'CELULAR', '19992939923'),
    ('74377493353', 'ANA MARY BEZERRA TABOSA', 'CPF', '74377493353'),
    ('29363599884', 'TEREZA NOBRE DA SILVA', 'CPF', '29363599884'),
    ('PENDENTE:DAIANE', 'DAIANE ROBERTA DE SOUZA VICENTE', 'CELULAR', '19991482771'),
    ('47798457852', 'INGLYD DA SILVA ADÃO', 'CPF', '47798457852'),
    ('25960641895', 'JESUS APARECIDO ADÃO', 'CELULAR', '19995117244'),
]
with connect() as c:
    for cpf_seed, nome_seed, tipo_seed, pix_seed in CADASTRO_INICIAL:
        c.execute('INSERT OR IGNORE INTO prestadores(cpf,nome,tipo_pix,pix) VALUES(?,?,?,?)',
                  (cpf_seed,nome_seed,tipo_seed,pix_seed))

def query(sql,args=()):
    with connect() as c:return pd.read_sql_query(sql,c,params=args)

def list_people():return query("SELECT cpf,nome FROM prestadores WHERE ativo=1 AND cpf NOT LIKE 'LEGADO:%' ORDER BY nome")

def save_scale(d,cpf,turno,extra=0):
    with connect() as c:
        c.execute('''INSERT INTO escalas(data,cpf,turno,extra) VALUES(?,?,?,?) ON CONFLICT(data,cpf) DO UPDATE SET turno=excluded.turno,atualizado=CURRENT_TIMESTAMP''',(d.isoformat(),cpf,turno,extra))

def scale_rows(d):
    return query('''SELECT e.id,e.data,e.cpf,p.nome,e.turno,e.status,e.entrada,e.saida,e.observacao,e.extra FROM escalas e JOIN prestadores p ON p.cpf=e.cpf WHERE e.data=? ORDER BY p.nome''',(d.isoformat(),))

def pdf_grade(d,df, titulo=None):
    buff=io.BytesIO()
    doc=SimpleDocTemplate(buff,pagesize=landscape(A4),leftMargin=26,rightMargin=26,topMargin=28,bottomMargin=25)
    styles=getSampleStyleSheet()
    content=[Paragraph('OPERAÇÃO ZARA - FREELANCERS',styles['Title']),Paragraph(titulo or ('DATA: '+d.strftime('%d/%m/%Y')),styles['Heading3']),Spacer(1,12)]
    head=['COLABORADOR','TURNO','ENTRADA','SAÍDA','STATUS','OBSERVAÇÃO']
    body=[head]
    for r in df.itertuples(index=False):
        e,s=TURNOS.get(r.turno,('',''))
        body.append([upper(r.nome),upper(r.turno),r.entrada if r.status in ('CUMPRIU','HORÁRIO DIFERENTE') else ('-' if r.status=='FALTA' else ''),r.saida if r.status in ('CUMPRIU','HORÁRIO DIFERENTE') else ('-' if r.status=='FALTA' else ''),upper(r.status),upper(r.observacao)])
    if len(body)==1:body.append(['']*len(head))
    widths=[175,150,66,66,105,185]
    table=Table([[Paragraph(escape(str(v)),styles['BodyText']) for v in row] for row in body],colWidths=widths,repeatRows=1,hAlign='LEFT')
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#F47721')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#B9B9B9')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F6F6F6')])]))
    # Paragraph headers need their own white text style.
    for cell in body[0]: pass
    from reportlab.lib.styles import ParagraphStyle
    white=ParagraphStyle('whiteheader',parent=styles['BodyText'],textColor=colors.white,fontName='Helvetica-Bold',fontSize=9)
    table._cellvalues[0]=[Paragraph(escape(v),white) for v in head]
    content.append(table)
    doc.build(content)
    return buff.getvalue()

def minutos_trabalhados(entrada, saida):
    a = datetime.strptime(entrada, '%H:%M')
    b = datetime.strptime(saida, '%H:%M')
    minutos = int((b-a).total_seconds() // 60)
    if minutos <= 0:
        minutos += 24 * 60
    return max(0, minutos - INTERVALO_MINUTOS)

def calcular_pagamento(status, entrada, saida):
    if status == 'FALTA':
        return (0, Decimal('0.00'), Decimal('0.00'))
    if status not in ('CUMPRIU', 'HORÁRIO DIFERENTE') or not entrada or not saida:
        return (None, None, None)
    try:
        minutos = minutos_trabalhados(hora(entrada), hora(saida))
    except (ValueError, TypeError):
        return (None, None, None)
    diferenca = (Decimal(minutos - 420) * VALOR_HORA / Decimal(60)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    total = max(Decimal('0.00'), VALOR_DIARIA + diferenca)
    return (minutos, diferenca, total)

def brl(valor):
    if valor is None: return 'PENDENTE'
    return 'R$ ' + f'{valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

def preparar_financeiro(df):
    if df.empty: return df.copy()
    resultado = df.copy()
    calculos = [calcular_pagamento(r.status, r.entrada, r.saida) for r in resultado.itertuples(index=False)]
    resultado['MINUTOS EFETIVOS'] = [v[0] for v in calculos]
    resultado['HORAS EFETIVAS'] = [f'{v[0]//60:02d}:{v[0]%60:02d}' if v[0] is not None else '' for v in calculos]
    resultado['AJUSTE (R$)'] = [float(v[1]) if v[1] is not None else None for v in calculos]
    resultado['TOTAL (R$)'] = [float(v[2]) if v[2] is not None else None for v in calculos]
    return resultado

def pdf_financeiro(df, titulo, consolidado=False):
    buff=io.BytesIO()
    doc=SimpleDocTemplate(buff,pagesize=landscape(A4),leftMargin=18,rightMargin=18,topMargin=26,bottomMargin=20)
    styles=getSampleStyleSheet()
    from reportlab.lib.styles import ParagraphStyle
    headerstyle=ParagraphStyle('finheader',parent=styles['BodyText'],textColor=colors.white,fontName='Helvetica-Bold',fontSize=7,leading=9)
    cellstyle=ParagraphStyle('fincell',parent=styles['BodyText'],fontSize=7,leading=9)
    if consolidado:
        cols=['COLABORADOR','CPF','PIX','DIÁRIAS','DATAS TRABALHADAS','HORA EXTRA','DIAS COM EXTRA','DESCONTOS','TOTAL A PAGAR']
        widths=[135,72,108,42,104,65,104,66,80]
    else:
        cols=['DATA','COLABORADOR','CPF','PIX','TURNO','ENTRADA','SAÍDA','HORAS','HORA EXTRA','DESCONTO','TOTAL']
        widths=[57,120,70,95,95,48,48,47,64,64,64]
    rows=[[Paragraph(escape(c),headerstyle) for c in cols]]
    for _,r in df.iterrows():
        rows.append([Paragraph(escape(str(r.get(k,'') if pd.notna(r.get(k,'')) else '')),cellstyle) for k in cols])
    t=Table(rows,colWidths=widths,repeatRows=1,hAlign='LEFT')
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#F47721')),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#BBBBBB')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F6F6F6')])]))
    doc.build([Paragraph('GDS LOGÍSTICA | RELATÓRIO ADMINISTRATIVO',styles['Title']),Paragraph(escape(titulo),styles['Heading3']),Spacer(1,12),t])
    return buff.getvalue()

def formatar_minutos(minutos):
    minutos=int(minutos)
    return f'{minutos//60}H{minutos%60:02d}' if minutos else '0H'

def relatorios_pagamento(financeiro):
    # Apenas presenças efetivamente confirmadas geram diárias.
    confirmados=financeiro[financeiro.status.isin(['CUMPRIU','HORÁRIO DIFERENTE']) & financeiro['TOTAL (R$)'].notna()].copy()
    confirmados['EXTRA MIN']=confirmados['MINUTOS EFETIVOS'].apply(lambda v:max(0,int(v)-420))
    confirmados['DESCONTO MIN']=confirmados['MINUTOS EFETIVOS'].apply(lambda v:max(0,420-int(v)))
    confirmados['DATA BR']=pd.to_datetime(confirmados.data).dt.strftime('%d/%m/%Y')
    detalhe=[]
    for r in confirmados.itertuples(index=False):
        extra=int(r[confirmados.columns.get_loc('EXTRA MIN')])
        desconto=int(r[confirmados.columns.get_loc('DESCONTO MIN')])
        detalhe.append({'DATA':r[confirmados.columns.get_loc('DATA BR')], 'COLABORADOR':r.nome,
            'CPF':r.cpf,'PIX':r.pix,'TURNO':r.turno,'ENTRADA':r.entrada,'SAÍDA':r.saida,
            'HORAS':formatar_minutos(r[confirmados.columns.get_loc('MINUTOS EFETIVOS')]),
            'HORA EXTRA':formatar_minutos(extra) if extra else 'NÃO',
            'DESCONTO':formatar_minutos(desconto) if desconto else 'NÃO',
            'TOTAL':brl(Decimal(str(r[confirmados.columns.get_loc('TOTAL (R$)')])) )})
    detalhado=pd.DataFrame(detalhe,columns=['DATA','COLABORADOR','CPF','PIX','TURNO','ENTRADA','SAÍDA','HORAS','HORA EXTRA','DESCONTO','TOTAL'])
    resumo=[]
    for (cpf,nome,pix),g in confirmados.groupby(['cpf','nome','pix'],dropna=False,sort=True):
        datas=sorted(g['DATA BR'].tolist(),key=lambda x:datetime.strptime(x,'%d/%m/%Y'))
        extra=g[g['EXTRA MIN']>0]
        extra_datas=sorted(extra['DATA BR'].tolist(),key=lambda x:datetime.strptime(x,'%d/%m/%Y'))
        total=sum((Decimal(str(v)) for v in g['TOTAL (R$)']),Decimal('0.00'))
        resumo.append({'COLABORADOR':nome,'CPF':cpf,'PIX':pix,'DIÁRIAS':len(g),
            'DATAS TRABALHADAS':', '.join(datas),'HORA EXTRA':formatar_minutos(extra['EXTRA MIN'].sum()) if len(extra) else 'NÃO',
            'DIAS COM EXTRA':', '.join(extra_datas) if extra_datas else '—',
            'DESCONTOS':formatar_minutos(g['DESCONTO MIN'].sum()) if g['DESCONTO MIN'].sum() else 'NÃO',
            'TOTAL A PAGAR':brl(total)})
    consolidado=pd.DataFrame(resumo,columns=['COLABORADOR','CPF','PIX','DIÁRIAS','DATAS TRABALHADAS','HORA EXTRA','DIAS COM EXTRA','DESCONTOS','TOTAL A PAGAR'])
    return consolidado,detalhado

def excel(df):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine='openpyxl') as w:
        df.to_excel(w,index=False,sheet_name='ESCALAS')
        ws=w.sheets['ESCALAS'];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width=min(45,max(15,max(len(str(c.value or '')) for c in col)+2))
    return b.getvalue()

st.markdown('''<style>.stApp{background:white}.block-container{padding-top:1rem;max-width:1250px}.stButton>button[kind="primary"]{background:#F47721;border-color:#F47721;color:white}h1,h2,h3{color:#161616}</style>''',unsafe_allow_html=True)
st.title('🟠 GDS LOGÍSTICA | OPERAÇÃO ZARA - FREELANCERS')
st.caption('ESCALA PLANEJADA → APONTAMENTO DA EXPEDIÇÃO → FOLHA EM PDF → RELATÓRIO POR PERÍODO')
if 'admin' not in st.session_state:st.session_state.admin=False
with st.sidebar:
    st.subheader('ACESSO ADMINISTRATIVO')
    if st.session_state.admin:
        st.success('ADMINISTRADOR ATIVO')
        if st.button('SAIR DO ADMINISTRADOR'):st.session_state.admin=False;st.rerun()
    else:
        senha=st.text_input('SENHA ADMIN',type='password')
        if st.button('ENTRAR COMO ADMIN'):
            if senha==ADMIN_PASSWORD:st.session_state.admin=True;st.rerun()
            else:st.error('SENHA INCORRETA')
    st.caption('VERSÃO PARA TESTE LOCAL. NÃO PUBLICAR NA INTERNET SEM AUTENTICAÇÃO ADEQUADA.')

aba_dia,aba_escala,aba_cadastro,aba_rel=st.tabs(['✅ OPERAÇÃO DO DIA','📅 CADASTRO DE ESCALA','👥 CADASTRO','📊 RELATÓRIOS'])

with aba_cadastro:
    st.subheader('CADASTRO DE COLABORADORES')
    st.caption('CADASTRE UMA VEZ. O CPF É ÚNICO. PIX É EXIBIDO APENAS AO ADMINISTRADOR.')
    with st.form('form_pessoa',clear_on_submit=True):
        a,b=st.columns(2)
        cpf=a.text_input('CPF');nome=b.text_input('NOME COMPLETO')
        if st.session_state.admin:
            tipo=st.selectbox('TIPO DE PIX',['CPF','CELULAR','E-MAIL','ALEATÓRIA'])
            pix=st.text_input('CHAVE PIX')
        else:tipo,pix=None,None
        submitted=st.form_submit_button('💾 SALVAR CADASTRO',type='primary')
    if submitted:
        cd=cpfnum(cpf);nm=upper(nome)
        if not valid_cpf(cd):st.error('CPF INVÁLIDO.')
        elif not nm:st.error('INFORME O NOME COMPLETO.')
        else:
            with connect() as c:
                if st.session_state.admin:
                    c.execute('''INSERT INTO prestadores(cpf,nome,tipo_pix,pix) VALUES(?,?,?,?) ON CONFLICT(cpf) DO UPDATE SET nome=excluded.nome,tipo_pix=excluded.tipo_pix,pix=excluded.pix''',(cd,nm,tipo,pix.strip()))
                else:
                    c.execute('''INSERT INTO prestadores(cpf,nome) VALUES(?,?) ON CONFLICT(cpf) DO UPDATE SET nome=excluded.nome''',(cd,nm))
            st.success('CADASTRO SALVO.');st.rerun()
    if st.session_state.admin:
        cad=query("SELECT nome,cpf,tipo_pix,pix FROM prestadores WHERE cpf NOT LIKE 'LEGADO:%' ORDER BY nome")
        st.dataframe(cad.rename(columns={'nome':'NOME','cpf':'CPF','tipo_pix':'TIPO PIX','pix':'CHAVE PIX'}),hide_index=True,use_container_width=True)
        st.caption('PARA ALTERAR PIX, CADASTRE NOVAMENTE O MESMO CPF COM A CHAVE ATUALIZADA.')
        pendentes=query("SELECT cpf,nome FROM prestadores WHERE cpf LIKE 'PENDENTE:%' ORDER BY nome")
        if not pendentes.empty:
            st.warning('EXISTEM COLABORADORES SEM CPF. ELES JÁ PODEM SER ESCALADOS; COMPLETE O CPF QUANDO DISPONÍVEL.')
            with st.form('regularizar_cpf'):
                escolha=st.selectbox('COLABORADOR COM CPF PENDENTE',pendentes['nome'].tolist())
                novo_cpf=st.text_input('CPF PARA COMPLETAR O CADASTRO')
                regularizar=st.form_submit_button('SALVAR CPF',type='primary')
            if regularizar:
                novo=cpfnum(novo_cpf)
                if not valid_cpf(novo):st.error('CPF INVÁLIDO.')
                else:
                    antigo=pendentes.loc[pendentes['nome']==escolha,'cpf'].iloc[0]
                    with connect() as c:
                        existe=c.execute('SELECT 1 FROM prestadores WHERE cpf=?',(novo,)).fetchone()
                        if existe:st.error('ESTE CPF JÁ ESTÁ CADASTRADO. VERIFIQUE ANTES DE UNIFICAR.')
                        else:
                            c.execute('UPDATE prestadores SET cpf=? WHERE cpf=?',(novo,antigo))
                            c.execute('UPDATE escalas SET cpf=? WHERE cpf=?',(novo,antigo))
                            st.success('CPF ATUALIZADO E ESCALAS PRESERVADAS.');st.rerun()
    else:
        cad=query("SELECT nome FROM prestadores WHERE cpf NOT LIKE 'LEGADO:%' ORDER BY nome")
        st.dataframe(cad.rename(columns={'nome':'COLABORADOR'}),hide_index=True,use_container_width=True)

with aba_escala:
    st.subheader('MONTAR ESCALA — NOME + TURNO')
    data_esc=st.date_input('DATA DA ESCALA',date.today(),format='DD/MM/YYYY',key='data_esc')
    pessoas=list_people()
    if pessoas.empty:st.warning('CADASTRE PRIMEIRO OS COLABORADORES NA ABA CADASTRO.')
    else:
        opcoes={f'{r.nome} — CPF PENDENTE' if r.cpf.startswith('PENDENTE:') else f'{r.nome} — FINAL CPF {r.cpf[-4:]}':r.cpf for r in pessoas.itertuples(index=False)}
        st.caption('CADASTRE UMA ÚNICA DATA OU UM PERÍODO COMPLETO. A ESCALA É PLANEJADA; NÃO GERA PAGAMENTO ATÉ O APONTAMENTO.')
        with st.form('form_escala',clear_on_submit=False):
            nome_sel=st.selectbox('SELECIONE O COLABORADOR',list(opcoes),index=None,placeholder='CLIQUE E DIGITE O NOME...')
            turno_sel=st.radio('TURNO',list(TURNOS),horizontal=True)
            modo=st.radio('TIPO DE ESCALA',['SOMENTE UMA DATA','PERÍODO DE DATAS'],horizontal=True)
            if modo=='PERÍODO DE DATAS':
                ca,cb=st.columns(2)
                inicio=ca.date_input('DATA INICIAL',date.today(),format='DD/MM/YYYY')
                fim=cb.date_input('DATA FINAL',date.today()+timedelta(days=30),format='DD/MM/YYYY')
                dias_escolhidos=st.multiselect('DIAS DA SEMANA', ['SEGUNDA','TERÇA','QUARTA','QUINTA','SEXTA','SÁBADO','DOMINGO'],default=['SEGUNDA','TERÇA','QUARTA','QUINTA','SEXTA','SÁBADO','DOMINGO'])
            else:
                inicio=fim=data_esc
                dias_escolhidos=['SEGUNDA','TERÇA','QUARTA','QUINTA','SEXTA','SÁBADO','DOMINGO']
            substituir=st.checkbox('ATUALIZAR TURNO DE DATAS JÁ ESCALADAS (SOMENTE SE AINDA NÃO HOUVER APONTAMENTO)',value=False)
            inserir=st.form_submit_button('💾 SALVAR ESCALA',type='primary')
        if inserir:
            if not nome_sel:st.error('SELECIONE O COLABORADOR.')
            elif fim<inicio:st.error('A DATA FINAL NÃO PODE SER ANTERIOR À INICIAL.')
            elif (fim-inicio).days>366:st.error('SELECIONE UM PERÍODO DE ATÉ 367 DIAS.')
            elif not dias_escolhidos:st.error('ESCOLHA PELO MENOS UM DIA DA SEMANA.')
            else:
                cpf_escolhido=opcoes[nome_sel]
                cont_novos=cont_atualizados=cont_ignorados=0
                with connect() as c:
                    for i in range((fim-inicio).days+1):
                        d=inicio+timedelta(days=i)
                        if ['SEGUNDA','TERÇA','QUARTA','QUINTA','SEXTA','SÁBADO','DOMINGO'][d.weekday()] not in dias_escolhidos:continue
                        existente=c.execute('SELECT id,status FROM escalas WHERE data=? AND cpf=?',(d.isoformat(),cpf_escolhido)).fetchone()
                        if existente is None:
                            c.execute('INSERT INTO escalas(data,cpf,turno) VALUES(?,?,?)',(d.isoformat(),cpf_escolhido,turno_sel))
                            cont_novos+=1
                        elif substituir and existente['status']=='PENDENTE':
                            c.execute('UPDATE escalas SET turno=?,atualizado=CURRENT_TIMESTAMP WHERE id=?',(turno_sel,existente['id']))
                            cont_atualizados+=1
                        else:cont_ignorados+=1
                st.success(f'ESCALA SALVA: {cont_novos} NOVA(S), {cont_atualizados} ATUALIZADA(S), {cont_ignorados} MANTIDA(S) SEM ALTERAÇÃO.')
                st.info('PARA CONFERIR A ESCALA, SELECIONE A DATA NA ÁREA ABAIXO.')
    es=scale_rows(data_esc)
    st.markdown(f'**ESCALA DE {data_esc.strftime("%d/%m/%Y")} — {len(es)} PESSOA(S)**')
    if es.empty:st.info('NENHUMA PESSOA ESCALADA NESTA DATA.')
    for r in es.itertuples(index=False):
        c1,c2,c3=st.columns([4,4,1])
        c1.write('**'+r.nome+'**');c2.write(r.turno)
        if c3.button('REMOVER',key=f'rem_{r.id}'):
            st.session_state['confirm_remove']=r.id
        if st.session_state.get('confirm_remove')==r.id:
            st.warning('REMOVER ESTA PESSOA DA ESCALA E SEUS APONTAMENTOS DESTE DIA?')
            a,b=st.columns(2)
            if a.button('CONFIRMAR REMOÇÃO',key=f'remy_{r.id}'):
                with connect() as c:c.execute('DELETE FROM escalas WHERE id=?',(r.id,))
                st.session_state.pop('confirm_remove',None);st.rerun()
            if b.button('CANCELAR',key=f'remn_{r.id}'):
                st.session_state.pop('confirm_remove',None);st.rerun()

with aba_dia:
    st.subheader('APONTAMENTO DA EXPEDIÇÃO')
    data_op=st.date_input('DATA DA OPERAÇÃO',date.today(),format='DD/MM/YYYY',key='data_op')
    op=scale_rows(data_op)
    if op.empty:st.info('NENHUM COLABORADOR ESCALADO PARA ESTA DATA. MONTE A ESCALA OU ADICIONE ALGUÉM ABAIXO.')
    else:
        st.caption('CLIQUE EM CUMPRIU, HORÁRIO DIFERENTE OU FALTA. CADA CONFIRMAÇÃO É SALVA IMEDIATAMENTE.')
        a,b,c,d=st.columns(4)
        a.metric('ESCALADOS',len(op));b.metric('CUMPRIU',int((op.status=='CUMPRIU').sum()));c.metric('HORÁRIO DIFERENTE',int((op.status=='HORÁRIO DIFERENTE').sum()));d.metric('FALTAS',int((op.status=='FALTA').sum()))
        for r in op.itertuples(index=False):
            prev_e,prev_s=TURNOS.get(r.turno,('',''))
            with st.container(border=True):
                c1,c2=st.columns([3,2])
                c1.markdown(f'**{r.nome}**'+('  ·  **EXTRA ADICIONADO NO DIA**' if r.extra else ''))
                c2.caption(f'{r.turno}  |  {prev_e} ÀS {prev_s}')
                a,b,c=st.columns(3)
                if a.button('✅ CUMPRIU',key=f'ok_{r.id}',use_container_width=True,type='primary' if r.status=='CUMPRIU' else 'secondary'):
                    with connect() as db:db.execute("UPDATE escalas SET status='CUMPRIU',entrada=?,saida=?,observacao='',atualizado=CURRENT_TIMESTAMP WHERE id=?",(prev_e,prev_s,r.id))
                    st.rerun()
                if b.button('🕒 HORÁRIO DIFERENTE',key=f'diff_{r.id}',use_container_width=True,type='primary' if r.status=='HORÁRIO DIFERENTE' else 'secondary'):
                    st.session_state[f'open_{r.id}']=True
                if c.button('❌ FALTA',key=f'falt_{r.id}',use_container_width=True,type='primary' if r.status=='FALTA' else 'secondary'):
                    with connect() as db:db.execute("UPDATE escalas SET status='FALTA',entrada='',saida='',observacao='',atualizado=CURRENT_TIMESTAMP WHERE id=?",(r.id,))
                    st.session_state.pop(f'open_{r.id}',None);st.rerun()
                if st.session_state.get(f'open_{r.id}'):
                    with st.form(f'diff_form_{r.id}'):
                        x,y=st.columns(2)
                        ent=x.text_input('ENTRADA REAL',value=r.entrada or prev_e)
                        sai=y.text_input('SAÍDA REAL',value=r.saida or prev_s)
                        obs=st.text_input('OBSERVAÇÃO (ATRASO, SAÍDA ANTECIPADA ETC.)',value=r.observacao)
                        save=st.form_submit_button('💾 SALVAR HORÁRIO DIFERENTE',type='primary')
                    if save:
                        try:
                            en,sa=hora(ent),hora(sai)
                            with connect() as db:db.execute("UPDATE escalas SET status='HORÁRIO DIFERENTE',entrada=?,saida=?,observacao=?,atualizado=CURRENT_TIMESTAMP WHERE id=?",(en,sa,upper(obs),r.id))
                            st.session_state.pop(f'open_{r.id}',None);st.rerun()
                        except ValueError as e:st.error(str(e))
                if r.status!='PENDENTE':
                    st.caption(f'SALVO: {r.status}'+(f' | {r.entrada} ÀS {r.saida}' if r.entrada else '')+(f' | {r.observacao}' if r.observacao else ''))
                else:st.caption('AGUARDANDO CONFIRMAÇÃO')
    with st.expander('➕ ADICIONAR OUTRA PESSOA À OPERAÇÃO DE HOJE'):
        pessoas=list_people()
        existentes=set(op.cpf) if not op.empty else set()
        livres={f'{r.nome} — CPF PENDENTE' if r.cpf.startswith('PENDENTE:') else f'{r.nome} — FINAL CPF {r.cpf[-4:]}':r.cpf for r in pessoas.itertuples(index=False) if r.cpf not in existentes}
        if not livres:st.info('TODOS OS CADASTRADOS JÁ ESTÃO NA ESCALA. CADASTRE OUTRA PESSOA NA ABA CADASTRO.')
        else:
            with st.form('extra_dia',clear_on_submit=True):
                nome_extra=st.selectbox('COLABORADOR',list(livres),index=None,placeholder='DIGITE PARA PESQUISAR...')
                turno_extra=st.radio('TURNO DO EXTRA',list(TURNOS),horizontal=True)
                adicionar=st.form_submit_button('ADICIONAR À OPERAÇÃO',type='primary')
            if adicionar:
                if not nome_extra:st.error('SELECIONE UMA PESSOA.')
                else:save_scale(data_op,livres[nome_extra],turno_extra,extra=1);st.rerun()
    st.divider()
    atual=scale_rows(data_op)
    st.subheader('GRADE DO DIA')
    if not atual.empty:
        grade=atual[['nome','turno','entrada','saida','status','observacao']].copy()
        grade.columns=['COLABORADOR','TURNO','ENTRADA','SAÍDA','STATUS','OBSERVAÇÃO']
        st.dataframe(grade,hide_index=True,use_container_width=True)
        st.download_button('🖨️ BAIXAR FOLHA DO DIA EM PDF',data=pdf_grade(data_op,atual),file_name=f'ZARA_FREELANCERS_{data_op.isoformat()}.pdf',mime='application/pdf',type='primary')
    else:st.info('NÃO HÁ GRADE PARA IMPRIMIR.')

with aba_rel:
    st.subheader('EXTRAIR RELATÓRIO POR PERÍODO')
    periodo=st.radio('SELECIONE O PERÍODO',['DIA','SEMANA','MÊS','PERSONALIZADO'],horizontal=True)
    ref=st.date_input('DATA DE REFERÊNCIA',date.today(),format='DD/MM/YYYY',key='data_ref')
    if periodo=='DIA':ini=fim=ref
    elif periodo=='SEMANA':ini=ref-timedelta(days=ref.weekday());fim=ini+timedelta(days=6)
    elif periodo=='MÊS':
        ini=ref.replace(day=1)
        proximo=ini.replace(year=ini.year+1,month=1) if ini.month==12 else ini.replace(month=ini.month+1)
        fim=proximo-timedelta(days=1)
    else:
        x,y=st.columns(2)
        ini=x.date_input('DATA INICIAL',ref.replace(day=1),format='DD/MM/YYYY')
        fim=y.date_input('DATA FINAL',ref,format='DD/MM/YYYY')
    if ini>fim:st.error('A DATA INICIAL NÃO PODE SER POSTERIOR À FINAL.')
    else:
        df=query('''SELECT e.data,p.nome,e.turno,e.status,e.entrada,e.saida,e.observacao,e.extra,e.cpf FROM escalas e JOIN prestadores p ON p.cpf=e.cpf WHERE e.data BETWEEN ? AND ? ORDER BY e.data,p.nome''',(ini.isoformat(),fim.isoformat()))
        if df.empty:
            st.info('SEM REGISTROS NESTE PERÍODO.')
        elif st.session_state.admin:
            st.info('RELATÓRIO ADMINISTRATIVO: CONTÉM CPF, PIX E VALORES. NÃO COMPARTILHE COM A EQUIPE OPERACIONAL.')
            financeiro=query('''SELECT e.data,p.nome,e.cpf,p.pix,e.turno,e.status,e.entrada,e.saida,e.observacao,e.extra
                FROM escalas e JOIN prestadores p ON p.cpf=e.cpf
                WHERE e.data BETWEEN ? AND ? ORDER BY e.data,p.nome''',(ini.isoformat(),fim.isoformat()))
            financeiro=preparar_financeiro(financeiro)
            consolidado,detalhe=relatorios_pagamento(financeiro)
            confirmados=financeiro[financeiro.status.isin(['CUMPRIU','HORÁRIO DIFERENTE']) & financeiro['TOTAL (R$)'].notna()]
            total=sum((Decimal(str(v)) for v in confirmados['TOTAL (R$)']),Decimal('0.00'))
            c1,c2,c3=st.columns(3)
            c1.metric('DIÁRIAS REALIZADAS',len(confirmados))
            c2.metric('TOTAL A PAGAR',brl(total))
            c3.metric('COLABORADORES',len(consolidado))
            st.caption('DIÁRIA COMPLETA: R$ 170,00 | 7H EFETIVAS | 1H DE INTERVALO | HORA ADICIONAL/DESCONTO: R$ 21,25 PROPORCIONAL.')
            pendentes=int((financeiro.status=='PENDENTE').sum())
            if pendentes:
                st.warning(f'ATENÇÃO: {pendentes} APONTAMENTO(S) AINDA NÃO CONFIRMADO(S). NÃO ENTRAM NO TOTAL A PAGAR.')
            st.subheader('RESUMO DE PAGAMENTOS — PIX, DIÁRIAS E HORAS EXTRAS')
            st.dataframe(consolidado,hide_index=True,use_container_width=True)
            st.download_button('📥 BAIXAR RESUMO EM EXCEL',excel(consolidado),file_name=f'GDS_RESUMO_PAGAMENTOS_{ini}_{fim}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.download_button('📄 BAIXAR RESUMO EM PDF',pdf_financeiro(consolidado,f'PAGAMENTOS: {ini.strftime("%d/%m/%Y")} A {fim.strftime("%d/%m/%Y")}',consolidado=True),file_name=f'GDS_RESUMO_PAGAMENTOS_{ini}_{fim}.pdf',mime='application/pdf')
            st.subheader('DETALHAMENTO — DATAS E HORAS DE CADA DIÁRIA')
            st.dataframe(detalhe,hide_index=True,use_container_width=True)
            st.download_button('📥 BAIXAR DETALHADO EM EXCEL',excel(detalhe),file_name=f'GDS_DETALHADO_{ini}_{fim}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.download_button('📄 BAIXAR DETALHADO EM PDF',pdf_financeiro(detalhe,f'PERÍODO: {ini.strftime("%d/%m/%Y")} A {fim.strftime("%d/%m/%Y")}'),file_name=f'GDS_DETALHADO_{ini}_{fim}.pdf',mime='application/pdf')
        else:
            st.metric('REGISTROS',len(df))
            st.caption(f'PRESENÇAS CONFIRMADAS: {int(df.status.isin(["CUMPRIU","HORÁRIO DIFERENTE"]).sum())} | FALTAS: {int((df.status=="FALTA").sum())} | PENDENTES: {int((df.status=="PENDENTE").sum())}')
            df=df.drop(columns=['cpf'])
            df['data']=pd.to_datetime(df['data']).dt.strftime('%d/%m/%Y')
            df['extra']=df['extra'].map({1:'SIM',0:'NÃO'})
            df.columns=['DATA','COLABORADOR','TURNO','STATUS','ENTRADA','SAÍDA','OBSERVAÇÃO','ADICIONADO NO DIA']
            st.dataframe(df,hide_index=True,use_container_width=True)
            st.download_button('📥 BAIXAR RELATÓRIO EM EXCEL',excel(df),file_name=f'RELATORIO_ZARA_{ini}_{fim}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
            st.download_button('📄 BAIXAR RELATÓRIO EM PDF',pdf_grade(ini,query('''SELECT p.nome,e.turno,e.status,e.entrada,e.saida,e.observacao FROM escalas e JOIN prestadores p ON p.cpf=e.cpf WHERE e.data BETWEEN ? AND ? ORDER BY e.data,p.nome''',(ini.isoformat(),fim.isoformat())),titulo=f'PERÍODO: {ini.strftime("%d/%m/%Y")} A {fim.strftime("%d/%m/%Y")}'),file_name=f'RELATORIO_ZARA_{ini}_{fim}.pdf',mime='application/pdf')
