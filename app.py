import streamlit as st
import pdfplumber
import pandas as pd
import datetime
import io
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Gerador de Escala Unificada",
    page_icon="🏥",
    layout="centered"
)

# --- CORES E CONSTANTES ---
COR_AZUL_CLARO = "CFE2F3"
COR_ROSA_CLARO = "F4CCCC"
COR_ROSA_ESCURO = "EA9999"

# --- FUNÇÕES AUXILIARES (LÓGICA DO SCRIPT) ---

def definir_cor_fundo(celula, cor_hex):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), cor_hex))
    celula._tc.get_or_add_tcPr().append(shading_elm)

def formatar_texto(run, tamanho=10, negrito=False):
    font = run.font
    font.size = Pt(tamanho)
    font.bold = negrito
    font.name = 'Arial'

def formatar_nome(nome_completo):
    if not isinstance(nome_completo, str): return ""
    partes = nome_completo.split()
    ignorar = ["ENF", "ENFERMEIRO", "CONTRATO", "EFETIVO", "TEC", "TECNICO", "MÉDIO", "MEDIO", "VÍNCULO", "FUNÇÃO", "COREN"]
    partes = [p for p in partes if p.upper() not in ignorar and len(p) > 2]
    if len(partes) > 1: return f"{partes[0]} {partes[-1]}".title()
    elif len(partes) == 1: return partes[0].title()
    return ""

def limpar_valor(val):
    return str(val).strip() if val is not None else ""

def detectar_metadados(pdf, nome_arquivo):
    # 1. Tenta definir TIPO pelo NOME DO ARQUIVO
    nome_upper = nome_arquivo.upper()
    tipo = "ENFERMEIROS" 
    
    if "TEC" in nome_upper or "TÉC" in nome_upper:
        tipo = "TÉCNICOS"
    elif "ENF" in nome_upper:
        tipo = "ENFERMEIROS"
    else:
        try:
            p0 = pdf.pages[0]
            texto_header = p0.crop((0, 0, p0.width, p0.height * 0.3)).extract_text() or ""
            if "TECNICO" in texto_header.upper() or "TÉCNICO" in texto_header.upper():
                tipo = "TÉCNICOS"
        except: pass

    # 2. Tenta definir DATA
    texto_completo = ""
    try:
        texto_completo = pdf.pages[0].extract_text().upper()
    except: pass
    
    meses = {r'\bJANEIRO\b':1, r'\bFEVEREIRO\b':2, r'\bMARÇO\b':3, r'\bMARCO\b':3, r'\bABRIL\b':4, r'\bMAIO\b':5, r'\bJUNHO\b':6, r'\bJULHO\b':7, r'\bAGOSTO\b':8, r'\bSETEMBRO\b':9, r'\bOUTUBRO\b':10, r'\bNOVEMBRO\b':11, r'\bDEZEMBRO\b':12}
    mes_detectado = 1
    for r_mes, n_mes in meses.items():
        if re.search(r_mes, texto_completo):
            mes_detectado = n_mes
            break
            
    ano_detectado = 2026
    match_ano = re.search(r'\b(202[3-9]|2030)\b', texto_completo)
    if match_ano: ano_detectado = int(match_ano.group(0))

    return tipo, ano_detectado, mes_detectado

def processar_pdf(file_obj, nome_arquivo):
    dados = []
    # Streamlit file_obj funciona como um arquivo aberto, pdfplumber aceita nativamente
    with pdfplumber.open(file_obj) as pdf:
        tipo, ano, mes = detectar_metadados(pdf, nome_arquivo)
        
        for page in pdf.pages:
            tabelas = page.extract_tables()
            for tabela in tabelas:
                df = pd.DataFrame(tabela)
                
                idx_cabecalho = -1
                mapa_dias = {}
                for idx, row in df.iterrows():
                    numeros = sum(1 for x in row if limpar_valor(x).isdigit() and 1 <= int(limpar_valor(x)) <= 31)
                    if numeros >= 5:
                        idx_cabecalho = idx
                        for c, v in enumerate(row):
                            vl = limpar_valor(v)
                            if vl.isdigit() and 1 <= int(vl) <= 31: mapa_dias[c] = int(vl)
                        break
                
                if idx_cabecalho == -1: continue
                
                df_dados = df.iloc[idx_cabecalho+1:].copy()
                col_nome = 1
                for i, col in enumerate(df.columns):
                    if "NOME" in str(col).upper(): col_nome = i; break
                
                for _, row in df_dados.iterrows():
                    if col_nome >= len(row): continue
                    nome = formatar_nome(limpar_valor(row.iloc[col_nome]))
                    if len(nome) < 3 or "TURNO" in nome.upper(): continue
                    
                    for c_idx, dia in mapa_dias.items():
                        if c_idx < len(row):
                            turno = limpar_valor(row.iloc[c_idx]).upper()
                            validos = ["M", "T", "N", "N1", "N2", "D", "SD", "SN", "LP", "MT", "TM", "MD"]
                            eh_valido = False
                            if len(turno) < 6:
                                for v in validos:
                                    if v in turno: eh_valido = True; break
                            
                            if eh_valido:
                                dados.append({"DIA": dia, "TURNO": turno, "NOME": nome})
                                
    if not dados:
        return pd.DataFrame(columns=['DIA', 'TURNO', 'NOME']), tipo, ano, mes
    return pd.DataFrame(dados), tipo, ano, mes

def adicionar_bloco_turno(table, df_filtrado, nome_turno, cor_lateral, start_row_idx):
    qtd = max(1, len(df_filtrado))
    for _ in range(qtd): table.add_row()
        
    c1 = table.rows[start_row_idx].cells[0]
    c2 = table.rows[start_row_idx + qtd - 1].cells[0]
    merged = c1.merge(c2)
    merged.text = nome_turno
    merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    merged.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    formatar_texto(merged.paragraphs[0].runs[0], negrito=True)
    definir_cor_fundo(merged, cor_lateral)
    
    if df_filtrado.empty:
        table.rows[start_row_idx].cells[1].text = "-"
        table.rows[start_row_idx].cells[2].text = "-"
    else:
        for i, (_, row) in enumerate(df_filtrado.iterrows()):
            r = table.rows[start_row_idx + i]
            r.cells[1].text = row['TURNO']
            r.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            formatar_texto(r.cells[1].paragraphs[0].runs[0], negrito=True)
            r.cells[2].text = row['NOME']
    return start_row_idx + qtd

def gerar_docx_completo(df_enf, df_tec, ano, mes):
    doc = Document()
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)

    df_enf['DIA'] = pd.to_numeric(df_enf['DIA'], errors='coerce')
    df_tec['DIA'] = pd.to_numeric(df_tec['DIA'], errors='coerce')
    
    dias_enf = set(df_enf['DIA'].dropna().unique())
    dias_tec = set(df_tec['DIA'].dropna().unique())
    dias_totais = sorted(list(dias_enf | dias_tec))
    
    dias_semana = {0: 'SEGUNDA-FEIRA', 1: 'TERÇA-FEIRA', 2: 'QUARTA-FEIRA', 3: 'QUINTA-FEIRA', 4: 'SEXTA-FEIRA', 5: 'SÁBADO', 6: 'DOMINGO'}

    for dia in dias_totais:
        try:
            dt = datetime.date(ano, mes, int(dia))
            txt_data = f"{dt.strftime('%d/%m/%Y')} {dias_semana[dt.weekday()]}"
        except: continue
        
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        
        # Data
        r = table.rows[0]
        c = r.cells[0].merge(r.cells[2])
        c.text = txt_data
        definir_cor_fundo(c, COR_AZUL_CLARO)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        formatar_texto(c.paragraphs[0].runs[0], tamanho=12, negrito=True)
        
        # --- ENFERMEIROS ---
        r = table.add_row()
        c = r.cells[0].merge(r.cells[2])
        c.text = "ENFERMEIROS"
        definir_cor_fundo(c, COR_ROSA_CLARO)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        formatar_texto(c.paragraphs[0].runs[0], tamanho=11, negrito=True)
        
        sub_enf = df_enf[df_enf['DIA'] == dia].sort_values('TURNO')
        not_enf = sub_enf[sub_enf['TURNO'].str.contains('N')]
        diu_enf = sub_enf[~sub_enf.index.isin(not_enf.index)]
        
        idx = 2
        idx = adicionar_bloco_turno(table, diu_enf, "DIURNO", COR_AZUL_CLARO, idx)
        idx = adicionar_bloco_turno(table, not_enf, "NOTURNO", COR_ROSA_ESCURO, idx)

        # --- TÉCNICOS ---
        r = table.add_row()
        c = r.cells[0].merge(r.cells[2])
        c.text = "TÉCNICOS"
        definir_cor_fundo(c, COR_ROSA_CLARO)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        formatar_texto(c.paragraphs[0].runs[0], tamanho=11, negrito=True)
        idx += 1
        
        sub_tec = df_tec[df_tec['DIA'] == dia].sort_values('TURNO')
        not_tec = sub_tec[sub_tec['TURNO'].str.contains('N')]
        diu_tec = sub_tec[~sub_tec.index.isin(not_tec.index)]
        
        idx = adicionar_bloco_turno(table, diu_tec, "DIURNO", COR_AZUL_CLARO, idx)
        idx = adicionar_bloco_turno(table, not_tec, "NOTURNO", COR_ROSA_ESCURO, idx)
        
        doc.add_paragraph("")

    return doc

# --- INTERFACE DO STREAMLIT ---

st.title("🏥 Gerador de Escala Unificada")
st.markdown("""
Este sistema converte as escalas de **Enfermeiros** e **Técnicos** (PDF) em um único documento Word formatado.
""")

with st.expander("ℹ️ Instruções (Clique para ler)"):
    st.write("""
    1. Arraste os arquivos PDF da escala abaixo (Enfermeiros e Técnicos).
    2. O sistema identificará automaticamente qual é qual.
    3. Clique no botão **Baixar Documento Word** quando aparecer.
    """)

# Upload de arquivos (permite múltiplos)
uploaded_files = st.file_uploader(
    "Arraste os arquivos PDF aqui (Enfermeiro e Técnico)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    # Botão para processar
    if st.button("🚀 Processar Escalas"):
        
        with st.spinner('Lendo arquivos e processando dados...'):
            dfs = {
                'ENFERMEIROS': pd.DataFrame(columns=['DIA', 'TURNO', 'NOME']), 
                'TÉCNICOS': pd.DataFrame(columns=['DIA', 'TURNO', 'NOME'])
            }
            meta_ano, meta_mes = 2026, 1
            
            sucesso = False
            
            for uploaded_file in uploaded_files:
                # Processa cada arquivo
                df, tipo, ano, mes = processar_pdf(uploaded_file, uploaded_file.name)
                
                if not df.empty:
                    dfs[tipo] = df
                    meta_ano, meta_mes = ano, mes
                    st.success(f"✅ Arquivo identificado: **{uploaded_file.name}** como _{tipo}_ ({mes}/{ano})")
                    sucesso = True
                else:
                    st.warning(f"⚠️ Não foi possível ler dados de: {uploaded_file.name}")

        if sucesso:
            with st.spinner('Gerando documento Word...'):
                doc_final = gerar_docx_completo(dfs['ENFERMEIROS'], dfs['TÉCNICOS'], meta_ano, meta_mes)
                
                # Salva em memória (buffer) para download
                buffer = io.BytesIO()
                doc_final.save(buffer)
                buffer.seek(0)
                
                st.markdown("---")
                st.write("### 🎉 Tudo pronto!")
                
                st.download_button(
                    label="📥 Baixar Escala Formatada (.docx)",
                    data=buffer,
                    file_name=f"Escala_Unificada_{meta_mes}_{meta_ano}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
