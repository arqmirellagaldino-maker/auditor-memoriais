import io
import re
import os
import pandas as pd
import streamlit as st
import pypdf
from pypdf.generic import DictionaryObject, NameObject, TextStringObject, ArrayObject, FloatObject, BooleanObject
import docx
from docx.enum.text import WD_COLOR_INDEX

# ==============================================================================
# BASE DE DADOS TÉCNICA DA ECON CONSTRUTORA (FONTES: R96, BOOK R07 e BOOK R09)
# ==============================================================================

REGRAS_ECON = {
    "SUPER_ECONOMICO_E_ECONOMICO": {
        "cozinha_area_servico": {
            "revestimento_parede": {
                "regra": "Cerâmica sobre bancada e tanque até a altura mínima de 1,50m do piso. Não colocar cerâmica abaixo da bancada.",
                "palavras_chave": ["cerâmica", "1,50", "1.50", "bancada", "azulejo", "parede"],
                "nao_conformes_detectar": ["piso ao teto", "piso e teto", "todas as paredes", "gesso liso sem cerâmica"],
                "fonte": "R96 (Item 95) e Book R07"
            },
            "bancada_cozinha": {
                "regra": "Bancada em mármore sintético (bancada e cuba integrais). Não utilizar bancada em aço inox para novos empreendimentos.",
                "palavras_chave": ["mármore sintético", "marmore sintetico"],
                "nao_conformes_detectar": ["aço inox", "aco inox", "granito"],
                "fonte": "R96 / Book R07 (Pág. 33-34)"
            },
            "tanque": {
                "regra": "Tanque em mármore sintético (20 litros).",
                "palavras_chave": ["mármore sintético", "marmore sintetico", "tanque"],
                "nao_conformes_detectar": ["louça com coluna", "louca com coluna", "aço inox"],
                "fonte": "R96 / Book R07"
            },
            "metais_cozinha": {
                "regra": "Torneira de mesa bica móvel com saída/ponto para filtro (acionamento até 10N).",
                "palavras_chave": ["mesa", "filtro", "bica móvel", "bica movel"],
                "nao_conformes_detectar": ["parede", "misturador monocomando"],
                "fonte": "Book R07 / R96"
            }
        },
        "banheiros": {
            "louca_bacia": {
                "regra": "Bacia com caixa acoplada, acionamento duplo / Ecoflush.",
                "palavras_chave": ["caixa acoplada", "ecoflush", "duplo acionamento"],
                "nao_conformes_detectar": ["válvula de descarga", "valvula de descarga"],
                "fonte": "Book R07"
            },
            "louca_lavatorio": {
                "regra": "Lavatório com coluna ou coluna suspensa (unidades tipo). Para PCD, lavatório PNE sem coluna.",
                "palavras_chave": ["coluna", "coluna suspensa", "suspenso"],
                "nao_conformes_detectar": ["cuba de embutir", "bancada de granito"],
                "fonte": "Book R07"
            },
            "metais_banho": {
                "regra": "Torneira de lavatório de mesa. Para PCD, torneira com alavanca tipo Matic Clinic.",
                "palavras_chave": ["mesa", "matic", "alavanca"],
                "nao_conformes_detectar": ["parede", "misturador"],
                "fonte": "Book R07"
            }
        },
        "pisos_geral": {
            "banhos_e_servico": {
                "regra": "Cerâmica de piso 43x43cm branco acetinado (ex: Ceral ARQ White) com rodapé cortado in loco (h=7cm).",
                "palavras_chave": ["cerâmica", "ceramica", "43x43"],
                "nao_conformes_detectar": ["porcelanato", "cimentado"],
                "fonte": "Book R07 / Book R09"
            },
            "terraco_varanda": {
                "regra": "Cerâmica de piso 43x43cm cinza acetinado (ex: Ceral ARQ Cement). Pingadeira em ardósia.",
                "palavras_chave": ["cerâmica", "ceramica", "43x43", "ardósia", "ardosia", "pingadeira"],
                "nao_conformes_detectar": ["porcelanato 60x60"],
                "fonte": "Book R07 / R96 (Item 90)"
            }
        },
        "portas_e_eletrica": {
            "fechaduras": {
                "regra": "Fechadura com roseta e parafuso aparente em aço inox (ex: Arouca linha Abitare/Victória).",
                "palavras_chave": ["roseta", "inox", "arouca"],
                "nao_conformes_detectar": ["espelho cego", "fechadura digital"],
                "fonte": "Book R07"
            },
            "tomadas_interruptores": {
                "regra": "Interruptores e tomadas modulares na cor branca (ex: Alumbra Inova Pró / Gracia Maxx).",
                "palavras_chave": ["modular", "branca", "alumbra"],
                "nao_conformes_detectar": ["sobrepor", "linha antiga"],
                "fonte": "Book R07"
            }
        }
    },
    "MEDIO": {
        "cozinha_area_servico": {
            "bancada_cozinha": {
                "regra": "Bancada em granito Cinza Andorinha com cuba de embutir em aço inox.",
                "palavras_chave": ["granito", "cinza andorinha", "aço inox", "embutir"],
                "nao_conformes_detectar": ["mármore sintético"],
                "fonte": "R96 / Book R07"
            }
        },
        "banheiros": {
            "bancada_banho": {
                "regra": "Bancada em granito com cuba de embutir em aço inox/louça.",
                "palavras_chave": ["granito", "bancada"],
                "nao_conformes_detectar": ["lavatório com coluna simples"],
                "fonte": "R96 (Item 90)"
            }
        }
    }
}

ESTRUTURA_OBRIGATORIA_MEMORIAL = [
    "1. OBJETO E CARACTERÍSTICAS GERAIS",
    "2. REVESTIMENTOS DE PISO",
    "3. REVESTIMENTOS DE PAREDE E TETO",
    "4. BANCADAS, LOUÇAS E TANQUES",
    "5. METAIS E ACESSÓRIOS",
    "6. ESQUADRIAS, PORTAS E FERRAGENS",
    "7. INSTALAÇÕES ELÉTRICAS E HIDRÁULICAS"
]

# ==============================================================================
# FUNÇÕES DE EXTRAÇÃO DE TEXTO E AUDITORIA
# ==============================================================================

def extrair_texto_pdf(file_bytes):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    paginas_texto = []
    for i, page in enumerate(reader.pages):
        paginas_texto.append({
            "pagina": i + 1,
            "texto": page.extract_text() or ""
        })
    return paginas_texto

def extrair_texto_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    paragrafos = []
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            paragrafos.append({
                "indice": i,
                "texto": p.text
            })
    return paragrafos

def auditar_memorial(texto_completo, padrao, escopo, tipo_doc):
    apontamentos = []
    
    # Determinar dicionário de regras
    chave_padrao = "MEDIO" if padrao == "Médio" else "SUPER_ECONOMICO_E_ECONOMICO"
    regras = REGRAS_ECON.get(chave_padrao, REGRAS_ECON["SUPER_ECONOMICO_E_ECONOMICO"])
    
    texto_lower = texto_completo.lower()
    
    # 1. Checagem de Estrutura Padrão
    if tipo_doc == "Memorial CEF / Financiador":
        for secao in ESTRUTURA_OBRIGATORIA_MEMORIAL:
            secao_limpa = secao.split(".")[1].strip().lower()
            if secao_limpa not in texto_lower:
                apontamentos.append({
                    "Ambiente": "Geral / Estrutura",
                    "Item": f"Seção Obrigatória: {secao}",
                    "Descrito": "Seção não localizada explicitamente no documento.",
                    "Padrão Econ Exigido": f"Conter tópico '{secao}' conforme modelo oficial.",
                    "Status": "🟡 Atenção",
                    "Ação Recomendada": "Verificar se a seção técnica da Caixa/banco foi omitida.",
                    "Fonte": "Modelo Padronizado CEF"
                })

    # 2. Checagem Técnica de Acabamentos
    for ambiente, itens in regras.items():
        nom_ambiente = ambiente.replace("_", " ").title()
        for subitem, spec in itens.items():
            nom_item = subitem.replace("_", " ").title()
            
            # Verificar termos não conformes explícitos
            encontrou_nao_conforme = False
            for termo_proibido in spec.get("nao_conformes_detectar", []):
                if termo_proibido in texto_lower:
                    apontamentos.append({
                        "Ambiente": nom_ambiente,
                        "Item": nom_item,
                        "Descrito": f"Consta termo '{termo_proibido}' no texto enviado.",
                        "Padrão Econ Exigido": spec["regra"],
                        "Status": "🔴 Não Conforme",
                        "Ação Recomendada": f"Ajustar descrição para alinhar ao padrão Econ ({spec['fonte']}).",
                        "Fonte": spec["fonte"]
                    })
                    encontrou_nao_conforme = True
                    break
            
            # Se não detectou erro direto, checar se a palavra-chave esperada está presente
            if not encontrou_nao_conforme:
                kw_presente = any(kw in texto_lower for kw in spec["palavras_chave"])
                if kw_presente:
                    apontamentos.append({
                        "Ambiente": nom_ambiente,
                        "Item": nom_item,
                        "Descrito": "Especificação alinhada com as diretrizes da Econ.",
                        "Padrão Econ Exigido": spec["regra"],
                        "Status": "🟢 Conforme",
                        "Ação Recomendada": "Manter especificação.",
                        "Fonte": spec["fonte"]
                    })
                else:
                    apontamentos.append({
                        "Ambiente": nom_ambiente,
                        "Item": nom_item,
                        "Descrito": "Termo técnico padrão não identificado claramente no texto.",
                        "Padrão Econ Exigido": spec["regra"],
                        "Status": "🟡 Atenção",
                        "Ação Recomendada": "Confirmar se o detalhamento técnico atende aos requisitos funcionais exigidos.",
                        "Fonte": spec["fonte"]
                    })

    return pd.DataFrame(apontamentos)

# ==============================================================================
# FUNÇÕES DE ANOTAÇÃO E MARCAÇÃO NO DOCUMENTO
# ==============================================================================

def gerar_pdf_anotado(file_bytes, df_apontamentos):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    writer = pypdf.PdfWriter()
    
    for page in reader.pages:
        writer.add_page(page)
        
    # Adicionar anotações na primeira página com o resumo das não-conformidades
    nao_conformes = df_apontamentos[df_apontamentos["Status"] == "🔴 Não Conforme"]
    
    if not nao_conformes.empty:
        resumo_texto = "AUDITORIA ECON - NÃO CONFORMIDADES:\n\n"
        for _, row in nao_conformes.iterrows():
            resumo_texto += f"• [{row['Ambiente']} - {row['Item']}]: {row['Ação Recomendada']}\n"
            
        annotation = DictionaryObject({
            NameObject('/Type'): NameObject('/Annot'),
            NameObject('/Subtype'): NameObject('/Text'),
            NameObject('/Rect'): ArrayObject([FloatObject(50), FloatObject(700), FloatObject(80), FloatObject(730)]),
            NameObject('/Contents'): TextStringObject(resumo_texto),
            NameObject('/Open'): BooleanObject(False),
            NameObject('/Name'): NameObject('/Comment')
        })
        writer.add_annotation(page_number=0, annotation=annotation)
        
    output_pdf = io.BytesIO()
    writer.write(output_pdf)
    output_pdf.seek(0)
    return output_pdf

def gerar_docx_anotado(file_bytes, df_apontamentos):
    doc = docx.Document(io.BytesIO(file_bytes))
    
    nao_conformes = df_apontamentos[df_apontamentos["Status"] == "🔴 Não Conforme"]
    
    if not nao_conformes.empty:
        # Adicionar cabeçalho de aviso
        p_aviso = doc.add_paragraph()
        run_aviso = p_aviso.add_run("--- RELATÓRIO DE AUDITORIA AUTOMÁTICA DE COMPATIBILIZAÇÃO ECON ---")
        run_aviso.font.bold = True
        run_aviso.font.highlight_color = WD_COLOR_INDEX.YELLOW
        
        for _, row in nao_conformes.iterrows():
            p = doc.add_paragraph()
            r = p.add_run(f"🔴 DIVERGÊNCIA [{row['Ambiente']} - {row['Item']}]: {row['Descrito']} | EXIGIDO: {row['Padrão Econ Exigido']}")
            r.font.highlight_color = WD_COLOR_INDEX.RED
            
    output_docx = io.BytesIO()
    doc.save(output_docx)
    output_docx.seek(0)
    return output_docx

# ==============================================================================
# INTERFACE GRÁFICA STREAMLIT
# ==============================================================================

def main():
    st.set_page_config(
        page_title="Verificador de Memoriais Econ",
        page_icon="🏗️",
        layout="wide"
    )
    
    st.title("🏗️ Verificador Automático de Memoriais Descritivos")
    st.caption("Ferramenta de Auditoria de Compatibilização Técnica - Construtora Econ")
    st.markdown("---")
    
    # Barra Lateral
    st.sidebar.header("⚙️ Configurações da Análise")
    
    padrao = st.sidebar.selectbox(
        "Padrão do Empreendimento",
        ["Super Econômico", "Econômico", "Médio"],
        index=1
    )
    
    escopo = st.sidebar.selectbox(
        "Escopo da Verificação",
        ["Área Privativa", "Área Comum"],
        index=0
    )
    
    tipo_doc = st.sidebar.selectbox(
        "Tipo de Memorial",
        ["Memorial do Cliente (Comercial / Vendas)", "Memorial CEF / Financiador"],
        index=0
    )
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📄 Arquivo de Entrada")
    uploaded_file = st.sidebar.file_uploader("Selecione o Memorial Descritivo (.pdf ou .docx)", type=["pdf", "docx"])
    
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        file_ext = uploaded_file.name.split(".")[-1].lower()
        
        st.info(f"📁 **Arquivo carregado:** `{uploaded_file.name}` ({len(file_bytes) / 1024:.1f} KB)")
        
        if st.button("🚀 Executar Auditoria de Conformidade", type="primary"):
            with st.spinner("Analisando texto e comparando com as diretrizes técnicas..."):
                texto_completo = ""
                
                if file_ext == "pdf":
                    paginas = extrair_texto_pdf(file_bytes)
                    texto_completo = " ".join([p["texto"] for p in paginas])
                elif file_ext == "docx":
                    paragrafos = extrair_texto_docx(file_bytes)
                    texto_completo = " ".join([p["texto"] for p in paragrafos])
                
                if not texto_completo.strip():
                    st.error("⚠️ Não foi possível extrair texto do arquivo. Verifique se o documento é um PDF pesquisável ou arquivo Word válido.")
                    return
                
                # Executar auditoria
                df_resultados = auditar_memorial(texto_completo, padrao, escopo, tipo_doc)
                
                # Métricas
                total = len(df_resultados)
                n_nao_conf = len(df_resultados[df_resultados["Status"] == "🔴 Não Conforme"])
                n_atencao = len(df_resultados[df_resultados["Status"] == "🟡 Atenção"])
                n_conf = len(df_resultados[df_resultados["Status"] == "🟢 Conforme"])
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total de Itens Analisados", total)
                col2.metric("🔴 Não Conformidades", n_nao_conf)
                col3.metric("🟡 Alertas / Atenção", n_atencao)
                col4.metric("🟢 Conforme", n_conf)
                
                st.markdown("### 📋 Tabela Detalhada de Apontamentos")
                
                # Exibir tabela formatada
                st.dataframe(
                    df_resultados,
                    use_container_width=True,
                    hide_index=True
                )
                
                st.markdown("---")
                st.subheader("📥 Exportação de Resultados")
                
                col_exp1, col_exp2 = st.columns(2)
                
                # Botão Exportação Excel
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
                    df_resultados.to_excel(writer, index=False, sheet_name="Auditoria_Memoriais")
                buffer_excel.seek(0)
                
                col_exp1.download_button(
                    label="📊 Baixar Planilha de Divergências (.xlsx)",
                    data=buffer_excel,
                    file_name=f"Relatorio_Auditoria_{uploaded_file.name.split('.')[0]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Botão Exportação Documento Anotado
                if file_ext == "pdf":
                    pdf_anotado = gerar_pdf_anotado(file_bytes, df_resultados)
                    col_exp2.download_button(
                        label="📄 Baixar PDF com Anotações Técnicas",
                        data=pdf_anotado,
                        file_name=f"Anotado_{uploaded_file.name}",
                        mime="application/pdf"
                    )
                elif file_ext == "docx":
                    docx_anotado = gerar_docx_anotado(file_bytes, df_resultados)
                    col_exp2.download_button(
                        label="📝 Baixar Word com Comentários de Revisão",
                        data=docx_anotado,
                        file_name=f"Anotado_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
    else:
        st.warning("👈 Por favor, faça o upload do memorial descritivo no painel lateral para iniciar a verificação.")
        
        # Guia Informativo
        st.markdown("""
        ### 📖 Como funciona a verificação?
        
        1. **Seleção dos Parâmetros:** Escolha o padrão (*Super Econômico*, *Econômico* ou *Médio*), o escopo (*Área Privativa* ou *Comum*) e o tipo de memorial.
        2. **Leitura e Extração:** O sistema lê as cláusulas do seu arquivo PDF ou Word sem alterar a estrutura do documento original.
        3. **Cruzamento de Dados:** Cada item é comparado com os cadernos técnicos da Econ:
           - **Padrão de Acabamentos R96**
           - **Book de Acabamentos Área Privativa R07**
           - **Book de Acabamentos Áreas Comuns R09**
        4. **Relatório e Anotações:** Receba o diagnóstico diretamente na tela, exporte em Excel ou baixe o arquivo com marcadores visuais das divergências.
        """)

if __name__ == "__main__":
    main()
