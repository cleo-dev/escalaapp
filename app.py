import streamlit as st
import pdfplumber
import pandas as pd
import datetime
import io
import re
import random
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Escala (V12 - Print Ready)", page_icon="🖨️", layout="centered")

# --- CORES ---
COR_AZUL_CLARO = "CFE2F3"
COR_ROSA_CLARO = "F4CCCC"
COR_ROSA_ESCURO = "EA9999"

# --- FUNÇÕES DE EXTRAÇÃO ---

def formatar_nome(nome_completo):
    if not isinstance(nome_completo, str): return ""
    partes = nome_completo.split()
    ignorar = ["ENF", "ENFERMEIRO", "CONTRATO", "EFETIVO", "TEC", "TECNICO", "MÉDIO", "MEDIO", "VÍNCULO", "FUNÇÃO", "COREN", "COREN-AP"]
    partes = [p for p in partes if p.upper() not in ignorar and len(p) > 2]
    
    # Retorna em MAIÚSCULO conforme pedido
    if len(partes) > 1: return f"{partes[0]} {partes[-1]}".upper()
    elif len(partes) == 1: return partes[0].upper()
    return ""

def limpar_valor(val):
    return str(val).strip() if val is not None else ""

def detectar_metadados(pdf, nome_arquivo):
    nome_upper = nome_arquivo.upper()
    tipo = "ENFERMEIROS"
    if "TEC" in nome_upper or "TÉC" in nome_upper: tipo = "TÉCNICOS"
    elif "ENF" in nome_upper: tipo = "ENFERMEIROS"
    
    texto_completo = ""
    try: texto_completo = pdf.pages[0].extract_text().upper()
    except: pass
    
    meses = {r'JANEIRO':1, r'FEVEREIRO':2, r'MARÇO':3, r'MARCO':3, r'ABRIL':4, r'MAIO':5, r'JUNHO':6, r'JULHO':7, r'AGOSTO':8, r'SETEMBRO':9, r'OUTUBRO':10, r'NOVEMBRO':11, r'DEZEMBRO':12}
    mes_detectado = 1
    for r_mes, n_mes in meses.items():
        if r_mes in texto_completo:
            mes_detectado = n_mes
            break
            
    ano_detectado = 2026
    match_ano = re.search(r'202[4-9]', texto_completo)
    if match_ano: ano_detectado = int(match_ano.group(0))

    return tipo, ano_detectado, mes_detectado

def processar_pdf(file_obj, nome_arquivo):
    dados = []
    with pdfplumber.open(file_obj) as pdf:
        tipo, ano, mes = detectar_metadados(pdf, nome_arquivo)
        
        for page in pdf.pages:
            tabelas = page.extract_tables()
            for tabela in tabelas:
                df = pd.DataFrame(tabela)
                idx_cabecalho = -1
                mapa_dias = {}
                
                for idx, row in df.iterrows():
                    numeros_validos = [int(limpar_valor(val)) for val in row if limpar_valor(val).isdigit() and 1 <= int(limpar_valor(val)) <= 31]
                    if len(numeros_validos) >= 5:
                        idx_cabecalho = idx
                        ultimo_dia_visto = 0
                        for c, v in enumerate(row):
                            vl = limpar_valor(v)
                            if vl.isdigit():
                                dia_num = int(vl)
                                if 1 <= dia_num <= 31:
                                    if dia_num < ultimo_dia_visto and ultimo_dia_visto > 20: continue
                                    mapa_dias[c] = dia_num
                                    ultimo_dia_visto = dia_num
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
                            if len(turno) < 7:
                                for v in validos:
                                    if v in turno: eh_valido = True; break
                            
                            if eh_valido:
                                dados.append({"DIA": dia, "TURNO": turno, "NOME": nome})
                                
    if not dados: return pd.DataFrame(columns=['DIA', 'TURNO', 'NOME']), tipo, ano, mes
    return pd.DataFrame(dados), tipo, ano, mes

# --- FUNÇÕES WORD ---

def definir_cor_fundo(celula, cor_hex):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), cor_hex))
    celula._tc.get_or_add_tcPr().append(shading_elm)

def formatar_texto(run, tamanho=10, negrito=False):
    font = run.font
    font.size = Pt(tamanho)
    font.bold = negrito
    font.name = 'Arial'

def set_col_widths_enf(table):
    # LARGURAS EXATAS PEDIDAS
    # Col 0: Label (Lateral) - Ajustamos para caber no A4 (~2.2cm)
    # Col 1: Turno - 1.8cm
    # Col 2: Nome - 4.0cm
    # Col 3: Função - 1.5cm
    # Col 4: Trocas - Restante (~6.5cm num A4 padrão)
    widths = [Cm(2.2), Cm(1.8), Cm(4.0), Cm(1.5), Cm(6.5)]
    for row in table.rows:
        for idx, width in enumerate(widths):
            if idx < len(row.cells): row.cells[idx].width = width

def definir_funcoes_aleatorias(df_turno):
    if df_turno.empty: return []
    total = len(df_turno)
    # Usa abreviação pedida
    funcoes = ["Clas."] * total
    if total > 0: funcoes[random.randint(0, total - 1)] = "Vol."
    return funcoes

def adicionar_bloco_turno_enf(table, df_filtrado, nome_turno, cor_lateral):
    start_row_idx = len(table.rows)
    qtd = max(1, len(df_filtrado))
    funcoes = definir_funcoes_aleatorias(df_filtrado)
    
    for _ in range(qtd): table.add_row()
        
    c1 = table.rows[start_row_idx].cells[0]
    c2 = table.rows[start_row_idx + qtd - 1].cells[0]
    merged = c1.merge(c2)
    merged.text = nome_turno
    merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    merged.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    formatar_texto(merged.paragraphs[0].runs[0], negrito=True, tamanho=9)
    definir_cor_fundo(merged, cor_lateral)
    
    if df_filtrado.empty:
        for i in range(1, 5): 
            if i < len(table.rows[start_row_idx].cells):
                table.rows[start_row_idx].cells[i].text = "-"
    else:
        for i, (_, row) in enumerate(df_filtrado.iterrows()):
            r = table.rows[start_row_idx + i]
            
            # Turno
            r.cells[1].text = row['TURNO']
            r.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            formatar_texto(r.cells[1].paragraphs[0].runs[0], negrito=True, tamanho=9)
            
            # Nome
            r.cells[2].text = row['NOME'] # Já vem em MAIUSCULO do formatar_nome
            formatar_texto(r.cells[2].paragraphs[0].runs[0], tamanho=8) # Fonte um pouco menor p/ caber
            
            # Função (Abreviada)
            r.cells[3].text = funcoes[i]
            r.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            formatar_texto(r.cells[3].paragraphs[0].runs[0], tamanho=8)

            # Trocas (Vazio)
            r.cells[4].text = ""

def adicionar_bloco_turno_tec(table, df_filtrado, nome_turno, cor_lateral):
    start_row_idx = len(table.rows)
    qtd = max(1, len(df_filtrado))
    
    for _ in range(qtd): table.add_row()
        
    c1 = table.rows[start_row_idx].cells[0]
    c2 = table.rows[start_row_idx + qtd - 1].cells[0]
    merged = c1.merge(c2)
    merged.text = nome_turno
    merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    merged.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    formatar_texto(merged.paragraphs[0].runs[0], negrito=True, tamanho=9)
    definir_cor_fundo(merged, cor_lateral)
    
    if df_filtrado.empty:
        if len(table.rows[start_row_idx].cells) > 2:
            table.rows[start_row_idx].cells[1].text = "-"
            table.rows[start_row_idx].cells[2].text = "-"
    else:
        for i, (_, row) in enumerate(df_filtrado.iterrows()):
            r = table.rows[start_row_idx + i]
            if len(r.cells) < 3: continue 
            
            # Turno
            r.cells[1].text = row['TURNO']
            r.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            formatar_texto(r.cells[1].paragraphs[0].runs[0], negrito=True, tamanho=9)
            
            # Nome
            r.cells[2].text = row['NOME']
            formatar_texto(r.cells[2].paragraphs[0].runs[0], tamanho=8)
            
            # Trocas (Mescla Col 3 e 4 se disponivel)
            if len(r.cells) >= 5:
                c_troca = r.cells[3].merge(r.cells[4])
                c_troca.text = ""
            elif len(r.cells) == 4:
                r.cells[3].text = ""

def gerar_docx_completo(df_enf, df_tec, ano, mes):
    doc = Document()
    doc.styles['Normal'].font.name = 'Arial'
    
    # Margens estreitas para caber tudo
    for section in doc.sections:
        section.top_margin = Cm(1.0); section.bottom_margin = Cm(1.0)
        section.left_margin = Cm(1.2); section.right_margin = Cm(1.2)

    df_enf['DIA'] = pd.to_numeric(df_enf['DIA'], errors='coerce')
    df_tec['DIA'] = pd.to_numeric(df_tec['DIA'], errors='coerce')
    dias_totais = sorted(list(set(df_enf['DIA'].dropna().unique()) | set(df_tec['DIA'].dropna().unique())))
    dias_semana = {0: 'SEG', 1: 'TER', 2: 'QUA', 3: 'QUI', 4: 'SEX', 5: 'SÁB', 6: 'DOM'}

    # CONTADOR DE DIAS PARA QUEBRA DE PÁGINA
    contador_dias = 0

    for dia in dias_totais:
        # Lógica de Quebra: A cada 2 dias (exceto no primeiro), cria nova página
        if contador_dias > 0 and contador_dias % 2 == 0:
            doc.add_page_break()
        
        contador_dias += 1

        try:
            dt = datetime.date(ano, mes, int(dia))
            txt_data = f"{dt.strftime('%d/%m/%Y')} {dias_semana[dt.weekday()]}"
        except: continue
        
        # Tabela Enfermeiros (5 colunas)
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        table.autofit = False 
        set_col_widths_enf(table)
        
        # Data Header
        r = table.rows[0]
        c = r.cells[0].merge(r.cells[4])
        c.text = txt_data
        definir_cor_fundo(c, COR_AZUL_CLARO)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        formatar_texto(c.paragraphs[0].runs[0], tamanho=11, negrito=True)
        
        # === ENFERMEIROS ===
        r = table.add_row()
        c = r.cells[0].merge(r.cells[4])
        c.text = "ENFERMEIROS"
        definir_cor_fundo(c, COR_ROSA_CLARO)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        formatar_texto(c.paragraphs[0].runs[0], tamanho=10, negrito=True)
        
        # Cabeçalho Colunas
        r_head = table.add_row()
        # Ajuste nomes para caber
        col_names = ["", "Turno", "Nome", "Func.", "Trocas"]
        for idx, nome in enumerate(col_names):
            r_head.cells[idx].text = nome
            r_head.cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if idx > 0: formatar_texto(r_head.cells[idx].paragraphs[0].runs[0], tamanho=8, negrito=True)
        
        sub_enf = df_enf[df_enf['DIA'] == dia].sort_values('TURNO')
        not_enf = sub_enf[sub_enf['TURNO'].str.contains('N')]
        diu_enf = sub_enf[~sub_enf.index.isin(not_enf.index)]
        
        adicionar_bloco_turno_enf(table, diu_enf, "DIURNO", COR_AZUL_CLARO)
        adicionar_bloco_turno_enf(table, not_enf, "NOTURNO", COR_ROSA_ESCURO)

        # === TÉCNICOS ===
        # Mesma tabela, adiciona linha de titulo
        r = table.add_row()
        c = r.cells[0].merge(r.cells[4])
        c.text = "TÉCNICOS"
        definir_cor_fundo(c, COR_ROSA_CLARO)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        formatar_texto(c.paragraphs[0].runs[0], tamanho=10, negrito=True)
        
        # Cabeçalho Técnicos (Mescla coluna 3 e 4 para trocas)
        r_head = table.add_row()
        col_names_tec = ["", "Turno", "Nome", "Trocas", ""]
        for idx, nome in enumerate(col_names_tec):
            if idx == 4: continue
            cell = r_head.cells[idx]
            if idx == 3: cell = cell.merge(r_head.cells[4])
            cell.text = nome
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if idx > 0: formatar_texto(cell.paragraphs[0].runs[0], tamanho=8, negrito=True)
            
        sub_tec = df_tec[df_tec['DIA'] == dia].sort_values('TURNO')
        not_tec = sub_tec[sub_tec['TURNO'].str.contains('N')]
        diu_tec = sub_tec[~sub_tec.index.isin(not_tec.index)]
        
        adicionar_bloco_turno_tec(table, diu_tec, "DIURNO", COR_AZUL_CLARO)
        adicionar_bloco_turno_tec(table, not_tec, "NOTURNO", COR_ROSA_ESCURO)
        
        # Espaço visual entre dias (paragrafo pequeno)
        doc.add_paragraph("")

    return doc

# --- INTERFACE ---
st.title("Gerador de Escala HEOC ")
st.markdown("""
- **Layout:** 2 dias por página.
- **Nomes:** MAIÚSCULOS.
- **Larguras:** Turno (1.8cm), Nome (4cm), Func (1.5cm).
- **Funções:** Clas. / Vol.
""")

uploaded_files = st.file_uploader("Arraste os PDFs aqui", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Processar"):
        dfs = {'ENFERMEIROS': pd.DataFrame(), 'TÉCNICOS': pd.DataFrame()}
        meta_ano, meta_mes = 2026, 1
        
        for f in uploaded_files:
            df, tipo, ano, mes = processar_pdf(f, f.name)
            if not df.empty:
                dfs[tipo] = df
                meta_ano, meta_mes = ano, mes
                st.success(f"Lido: {tipo} ({mes}/{ano})")

        if not dfs['ENFERMEIROS'].empty or not dfs['TÉCNICOS'].empty:
            doc = gerar_docx_completo(dfs['ENFERMEIROS'], dfs['TÉCNICOS'], meta_ano, meta_mes)
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button("📥 Baixar DOCX Formatado", data=bio.getvalue(), file_name=f"Escala_Oficial_{meta_mes}_{meta_ano}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.error("Nenhum dado encontrado nos arquivos.")
