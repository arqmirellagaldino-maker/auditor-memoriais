import io
import os
import re
import unicodedata
from pathlib import Path
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
import pypdf
import docx
from pypdf.generic import DictionaryObject, NameObject, TextStringObject, ArrayObject, FloatObject, BooleanObject
from docx.enum.text import WD_COLOR_INDEX

# ==============================================================================
# QUALITY HUB | MEMORIAIS — V3
# ------------------------------------------------------------------------------
# PRINCÍPIOS
# 1) O PADRÃO DE ACABAMENTOS R96 é a fonte técnica principal.
# 2) O mesmo motor técnico atende Memorial do Cliente e Memorial do Financiador.
# 3) O Memorial do Financiador recebe uma camada adicional de protocolo da
#    Coordenação (itens de atenção / confirmação em projeto / fora de escopo).
# 4) "Misto" não é um quarto padrão técnico: é uma configuração que roteia
#    trechos/grupos para Super Econômico, Econômico ou Médio.
# 5) Quando a base técnica consegue responder, o app apresenta a especificação
#    esperada. Não devolve apenas "verificar conforme planilhão".
# ==============================================================================

APP_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
NOME_BASE_PADRAO = "PADRÃO DE ACABAMENTOS-R96.xlsx"
CAMINHOS_BASE = [
    APP_DIR / NOME_BASE_PADRAO,
    Path.cwd() / NOME_BASE_PADRAO,
]

PADROES_TECNICOS = ["Super Econômico", "Econômico", "Médio"]
STATUS_OK = "🟢 Conforme"
STATUS_ERRO = "🔴 Divergência"
STATUS_ATENCAO = "🟡 Atenção do Coordenador"
STATUS_INFO = "⚪ Sem conferência"

# Termos usados apenas para evitar falsos positivos em comparação textual.
# A especificação oficial continua sendo lida diretamente do Excel R96.
MATERIAIS_RELEVANTES = [
    "ceramica", "porcelanato", "granito", "marmore sintetico", "marmore",
    "aco inox", "inox", "louca", "concreto desempenado", "cimentado",
    "vinilico", "laminado", "ardosia", "gesso", "textura acrilica",
    "monocapa", "aluminio", "ferro", "vidro", "pvc", "madeira",
    "intertravado", "pedra natural", "latex pva", "latex acrilica",
]

# Equivalências de nomes para localizar ambientes no memorial.
# Não mudam a regra da planilha; servem apenas para encontrar o contexto.
EQUIVALENCIAS_AMBIENTES = {
    "APA OU STUDIO": ["apa", "studio", "estudio"],
    "ÁREA DE SERVIÇO": ["area de servico", "a.s", "as", "lavanderia da unidade"],
    "BANHOS": ["banho", "banheiro", "banheiros", "wc"],
    "DORMITÓRIOS": ["dormitorio", "dormitorios", "quarto", "quartos"],
    "SALA": ["sala", "estar", "jantar"],
    "VARANDA COM A.S": ["varanda com a.s", "varanda com as", "varanda com area de servico"],
    "VARANDA SEM A.S": ["varanda sem a.s", "varanda sem as", "varanda"],
    "CIRCULAÇÃO": ["circulacao", "hall interno", "corredor"],
    "COZINHA": ["cozinha"],
}

# ==============================================================================
# PROTOCOLO ADICIONAL — MEMORIAL DO FINANCIADOR / CEF
# Fonte operacional: Orientações Equipe Projetos para Análise Memorial CEF.
# Estes itens NÃO substituem a base R96. Eles acrescentam alertas e limites
# de escopo quando a confirmação depende de projeto/coordenador.
# ==============================================================================

PROTOCOLO_CEF = [
    {
        "id": "identificacao",
        "titulo": "Dados de identificação do empreendimento",
        "gatilhos": ["identificacao do empreendimento", "objeto e caracteristicas gerais", "empreendimento"],
        "acao": "sem_conferencia",
        "orientacao": "Dados de identificação do empreendimento não são conferidos pela equipe de Projetos.",
        "fonte": "Orientações Coordenação CEF — pág. 2",
    },
    {
        "id": "estrutura",
        "titulo": "Estrutura das torres / garagem",
        "gatilhos": ["supraestrutura", "estrutura", "garagem", "pre moldada", "pré-moldada"],
        "acao": "atencao",
        "orientacao": "Confirmar em projeto o tipo de estrutura das torres e, se houver garagem, confirmar se a solução é pré-moldada ou não.",
        "fonte": "Orientações Coordenação CEF — pág. 2",
    },
    {
        "id": "ambientes_finais",
        "titulo": "Ambientes e finais de unidades",
        "gatilhos": ["revestimentos", "ambientes", "final", "finais", "unidade"],
        "acao": "atencao",
        "orientacao": "Confirmar no projeto se os ambientes citados estão corretos. Quando houver indicação de finais/unidades, conferir também a aplicação aos finais informados.",
        "fonte": "Orientações Coordenação CEF — págs. 4 a 6",
    },
    {
        "id": "portas_janelas_dimensao",
        "titulo": "Portas e janelas — dimensões",
        "gatilhos": ["portas", "janelas", "esquadrias"],
        "acao": "atencao",
        "orientacao": "Para dimensões, prever/solicitar a redação 'conforme projeto de arquitetura'. Para janelas, seguir a mesma orientação aplicável às portas.",
        "fonte": "Orientações Coordenação CEF — pág. 7",
    },
    {
        "id": "fachada_estrutura",
        "titulo": "Acabamento conforme sistema estrutural",
        "gatilhos": ["fachada", "textura", "monocapa", "alvenaria estrutural", "estrutura convencional"],
        "acao": "atencao",
        "orientacao": "Confirmar o sistema estrutural: em alvenaria estrutural, a orientação é monocapa; em estrutura convencional, textura acrílica, conforme aplicabilidade do projeto.",
        "fonte": "Orientações Coordenação CEF — pág. 7",
    },
    {
        "id": "marcas",
        "titulo": "Marcas / modelos",
        "gatilhos": ["marca", "modelo", "fabricante", "fechadura", "batente"],
        "acao": "sem_conferencia",
        "orientacao": "Marcas, modelos, fechaduras e batentes não são objeto de conferência técnica da equipe de Projetos neste protocolo.",
        "fonte": "Orientações Coordenação CEF — págs. 5, 7, 8, 13",
    },
    {
        "id": "gas",
        "titulo": "Rede e prumadas de gás",
        "gatilhos": ["gas", "gás", "prumada de gas", "rede enterrada"],
        "acao": "atencao",
        "orientacao": "Conferir o item. Para rede enterrada de gás, prever PEAD. Quando houver prumadas de gás, indicar que serão conforme projeto de hidráulica. Não conferir marcas.",
        "fonte": "Orientações Coordenação CEF — pág. 12",
    },
    {
        "id": "pressurizacao",
        "titulo": "Pressurização das escadas",
        "gatilhos": ["pressurizacao", "pressurização", "escada pressurizada", "sala de pressurizacao"],
        "acao": "atencao",
        "orientacao": "Confirmar no projeto se há pressurização das escadas. Quando não houver, informar a condição correspondente no memorial.",
        "fonte": "Orientações Coordenação CEF — pág. 12",
    },
    {
        "id": "aquecimento_solar",
        "titulo": "Aquecimento solar",
        "gatilhos": ["aquecimento solar", "solar"],
        "acao": "atencao",
        "orientacao": "Confirmar em projeto se o empreendimento possui ou não previsão de aquecimento solar.",
        "fonte": "Orientações Coordenação CEF — pág. 13",
    },
]

# Seções do protocolo original explicitamente fora do escopo da equipe de Projetos.
FORA_ESCOPO_CEF = [
    "quantidade de apartamentos",
    "estacionamento",
    "padrão do empreendimento",
    "padrao do empreendimento",
    "memorial de infraestrutura",
]

# ==============================================================================
# UTILITÁRIOS DE TEXTO
# ==============================================================================

def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).replace("\n", " ").strip().lower()
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def tokens_significativos(texto):
    stop = {
        "de", "da", "do", "das", "dos", "e", "em", "com", "ou", "para", "por",
        "quando", "onde", "no", "na", "nos", "nas", "um", "uma", "ao", "aos",
        "ser", "sera", "pode", "podera", "conforme", "aplicavel", "nao", "sim",
        "ponto", "pontos", "prever", "caso", "houver", "tipo", "sobre",
    }
    palavras = re.findall(r"[a-z0-9]+(?:,[0-9]+)?", normalizar(texto))
    return {p for p in palavras if len(p) >= 3 and p not in stop}


def similaridade_textual(a, b):
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return 0.0
    ta, tb = tokens_significativos(na), tokens_significativos(nb)
    jaccard = len(ta & tb) / max(1, len(ta | tb))
    seq = SequenceMatcher(None, na[:1200], nb[:1200]).ratio()
    return 0.65 * jaccard + 0.35 * seq


def materiais_presentes(texto):
    n = normalizar(texto)
    return {m for m in MATERIAIS_RELEVANTES if m in n}


def aliases_ambiente(nome):
    base = normalizar(nome)
    aliases = [base]
    for chave, equivalentes in EQUIVALENCIAS_AMBIENTES.items():
        if normalizar(chave) == base:
            aliases += [normalizar(x) for x in equivalentes]
    # termos do próprio nome composto ajudam em áreas comuns
    if len(base) > 8:
        aliases += [x.strip() for x in re.split(r"/| e/ou | ou |,|\(|\)", base) if len(x.strip()) > 4]
    return list(dict.fromkeys([a for a in aliases if a]))


def localizar_contexto(texto, ambiente, janela=1300):
    n = normalizar(texto)
    candidatos = []
    for alias in aliases_ambiente(ambiente):
        pos = n.find(alias)
        if pos >= 0:
            candidatos.append((pos, alias))
    if not candidatos:
        return "", False
    pos, alias = min(candidatos, key=lambda x: x[0])
    ini = max(0, pos - 200)
    fim = min(len(n), pos + max(janela, len(alias) + 400))
    return n[ini:fim], True


def localizar_melhor_trecho(contexto, item, especificacao):
    if not contexto:
        return ""
    partes = [p.strip() for p in re.split(r"(?<=[\.;:])\s+|\n+", contexto) if len(p.strip()) >= 8]
    if not partes:
        return contexto[:700]
    alvo = f"{item} {especificacao}"
    melhor = max(partes, key=lambda p: similaridade_textual(p, alvo))
    return melhor[:900]

# ==============================================================================
# LEITURA DA BASE R96
# ==============================================================================

def _achar_linha_ambientes(df):
    for i in range(min(len(df), 20)):
        if normalizar(df.iloc[i, 0]) == "ambientes":
            return i
    raise ValueError("Não foi encontrada a linha 'AMBIENTES' na planilha.")


def _eh_cabecalho_secao(valor, linha):
    v = normalizar(valor)
    if not v:
        return False
    if all((x is None or str(x).strip() == "" or pd.isna(x)) for x in linha[1:]):
        return v in {"instalacoes eletricas", "instalacoes hidraulicas", "acabamentos", "areas comuns", "area privativa"}
    return False


def parsear_matriz_planilha(df, padrao, escopo, nome_aba):
    regras = []
    idx_amb = _achar_linha_ambientes(df)
    ambientes = []
    for col in range(1, df.shape[1]):
        valor = df.iloc[idx_amb, col]
        ambientes.append(str(valor).strip() if valor is not None and not pd.isna(valor) else "")

    secao = "ACABAMENTOS"
    for r in range(idx_amb + 1, len(df)):
        item_raw = df.iloc[r, 0] if df.shape[1] else None
        item = "" if item_raw is None or pd.isna(item_raw) else str(item_raw).strip()
        if not item:
            continue
        if _eh_cabecalho_secao(item_raw, list(df.iloc[r, :])):
            secao = item.strip().upper()
            continue
        # descarta linhas administrativas / datas
        if normalizar(item) in {"atualizado em", "r", "r:"}:
            continue
        for col, ambiente in enumerate(ambientes, start=1):
            if not ambiente:
                continue
            valor = df.iloc[r, col] if col < df.shape[1] else None
            if valor is None or pd.isna(valor) or not str(valor).strip():
                continue
            regras.append({
                "padrao": padrao,
                "escopo": escopo,
                "ambiente": ambiente.strip(),
                "secao": secao,
                "item": item.strip(),
                "especificacao": str(valor).strip(),
                "fonte": f"Padrão de Acabamentos R96 — {nome_aba}",
            })
    return regras


@st.cache_data(show_spinner=False)
def carregar_base_r96(file_bytes=None):
    """Lê diretamente o XLSX R96. A planilha permanece a fonte de verdade."""
    if file_bytes is not None:
        origem = io.BytesIO(file_bytes)
    else:
        caminho = next((p for p in CAMINHOS_BASE if p.exists()), None)
        if caminho is None:
            raise FileNotFoundError(
                f"Base '{NOME_BASE_PADRAO}' não encontrada ao lado do app. Faça o upload da base na barra lateral."
            )
        origem = caminho

    # O app usa pandas para leitura dinâmica da fonte técnica. Nenhuma regra
    # de acabamento é hardcoded no Python.
    xls = pd.ExcelFile(origem)
    regras = []

    mapa_abas = [
        ("SUPER ECO-ECONOMICO_PRIVATIVA", "Super Econômico", "Área Privativa"),
        ("SUPER ECO-ECONOMICO_PRIVATIVA", "Econômico", "Área Privativa"),
        ("MÉDIO_PRIVATIVA", "Médio", "Área Privativa"),
        ("ÁREA COMUM pagina 1-2", "Todos", "Área Comum"),
        ("ÁREA COMUM pagina 2-2", "Todos", "Área Comum"),
    ]

    for aba, padrao, escopo in mapa_abas:
        if aba not in xls.sheet_names:
            continue
        df = pd.read_excel(xls, sheet_name=aba, header=None, dtype=object)
        regras.extend(parsear_matriz_planilha(df, padrao, escopo, aba))

    return pd.DataFrame(regras)

# ==============================================================================
# CONFIGURAÇÃO DE EMPREENDIMENTO MISTO
# ==============================================================================

def padrao_do_grupo(trecho, grupos_mistos, padrao_fallback=None):
    """Tenta rotear um trecho do memorial para o padrão cadastrado no início."""
    nt = normalizar(trecho)
    melhor = None
    melhor_score = 0
    for g in grupos_mistos or []:
        torre = normalizar(g.get("torre", ""))
        unidades = normalizar(g.get("unidades", ""))
        padrao = g.get("padrao", "")
        termos = [t for t in [torre, unidades] if t and t not in {"demais", "todas", "todos"}]
        score = sum(1 for t in termos if t in nt)
        # finais digitados como 01, 02, 05, 06
        nums = re.findall(r"\b\d{1,3}\b", unidades)
        score += sum(0.35 for num in nums if re.search(rf"\b0*{int(num)}\b", nt)) if nums else 0
        if score > melhor_score:
            melhor_score = score
            melhor = padrao
    return melhor or padrao_fallback


def descricao_config_mista(grupos):
    if not grupos:
        return "Nenhum grupo configurado."
    return " | ".join(f"{g['torre']} — {g['unidades']}: {g['padrao']}" for g in grupos)

# ==============================================================================
# EXTRAÇÃO DOS MEMORIAIS
# ==============================================================================

def extrair_texto_pdf(file_bytes):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    return [{"pagina": i + 1, "texto": page.extract_text() or ""} for i, page in enumerate(reader.pages)]


def extrair_texto_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    saida = []
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            saida.append({"indice": i, "texto": p.text})
    # inclui tabelas, muito comuns em memoriais CEF
    for ti, tabela in enumerate(doc.tables):
        for ri, row in enumerate(tabela.rows):
            linha = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if linha:
                saida.append({"indice": f"T{ti}-R{ri}", "texto": linha})
    return saida

# ==============================================================================
# MOTOR DE CONFERÊNCIA R96
# ==============================================================================

def avaliar_regra(trecho, regra):
    esperado = regra["especificacao"]
    ne = normalizar(esperado)
    nt = normalizar(trecho)

    if normalizar(esperado) in {"nao aplicavel", "não aplicável"}:
        # Sem contexto suficiente, não marcamos vermelho por ausência.
        if not trecho:
            return STATUS_OK, "Item previsto como não aplicável na base.", 1.0
        # se o trecho contém um material concreto, pede validação humana
        mats = materiais_presentes(trecho)
        if mats:
            return STATUS_ATENCAO, "A base indica 'Não aplicável', mas foi encontrada uma descrição no contexto. Validar aplicabilidade.", 0.4
        return STATUS_OK, "Não foram encontrados indícios de especificação conflitante.", 0.85

    if not trecho:
        return STATUS_ATENCAO, "Ambiente/descrição não localizado com segurança no memorial.", 0.0

    sim = similaridade_textual(trecho, esperado)
    mats_esp = materiais_presentes(esperado)
    mats_txt = materiais_presentes(trecho)

    # divergência de material é evidência mais forte que simples baixa similaridade
    if mats_esp and mats_txt and mats_esp.isdisjoint(mats_txt):
        return STATUS_ERRO, f"Material/solução encontrada parece divergir da base R96. Esperado: {', '.join(sorted(mats_esp))}.", sim

    # presença de materiais/termos esperados + similaridade razoável
    cobertura = 0.0
    tok_esp = tokens_significativos(esperado)
    tok_txt = tokens_significativos(trecho)
    if tok_esp:
        cobertura = len(tok_esp & tok_txt) / len(tok_esp)

    if sim >= 0.38 or cobertura >= 0.48 or (mats_esp and mats_esp.issubset(mats_txt)):
        return STATUS_OK, "Descrição compatível com os principais termos técnicos da base R96.", max(sim, cobertura)

    return STATUS_ATENCAO, "Não foi possível confirmar a aderência apenas pela leitura textual. Conferir o trecho indicado contra a especificação R96.", max(sim, cobertura)


def filtrar_regras(base, padrao, escopo):
    if escopo == "Área Comum":
        return base[base["escopo"] == "Área Comum"].copy()
    return base[(base["escopo"] == "Área Privativa") & (base["padrao"] == padrao)].copy()


def auditar_base_r96(texto, base, padrao, escopo, grupos_mistos=None):
    resultados = []

    if padrao != "Misto":
        conjuntos = [(padrao, texto, "Empreendimento")]
    else:
        conjuntos = []
        for g in grupos_mistos or []:
            # Busca um trecho mais amplo a partir da torre/unidades; quando não encontra,
            # mantém o texto completo, mas o resultado fica com identificação do grupo.
            termos = [g.get("torre", ""), g.get("unidades", "")]
            ntexto = normalizar(texto)
            posicoes = [ntexto.find(normalizar(t)) for t in termos if normalizar(t) and ntexto.find(normalizar(t)) >= 0]
            if posicoes:
                p = min(posicoes)
                trecho_grupo = ntexto[max(0, p - 500):min(len(ntexto), p + 9000)]
            else:
                trecho_grupo = texto
            conjuntos.append((g["padrao"], trecho_grupo, f"{g['torre']} — {g['unidades']}"))

    for padrao_tecnico, texto_grupo, grupo_nome in conjuntos:
        regras = filtrar_regras(base, padrao_tecnico, escopo)
        for _, regra in regras.iterrows():
            contexto, achou_ambiente = localizar_contexto(texto_grupo, regra["ambiente"])
            trecho = localizar_melhor_trecho(contexto, regra["item"], regra["especificacao"]) if achou_ambiente else ""
            status, obs, confianca = avaliar_regra(trecho, regra)

            resultados.append({
                "Grupo / Aplicação": grupo_nome,
                "Padrão aplicado": padrao_tecnico if escopo == "Área Privativa" else "Áreas Comuns R96",
                "Área": escopo,
                "Ambiente": regra["ambiente"],
                "Seção": regra["secao"],
                "Item": regra["item"],
                "Texto encontrado": trecho if trecho else "Não localizado com segurança",
                "Especificação prevista": regra["especificacao"],
                "Status": status,
                "Orientação / resposta prevista": regra["especificacao"],
                "Observação": obs,
                "Confiança": round(float(confianca), 2),
                "Fonte": regra["fonte"],
            })

    return resultados

# ==============================================================================
# CAMADAS COMPLEMENTARES DOS PROTOCOLOS
# ==============================================================================

def auditoria_especificacoes_gerais_cliente(texto, padrao, grupos_mistos):
    """Checklist humano já definido para Especificações Gerais do Memorial Cliente."""
    resultados = []
    regras = [
        ("Estrutura e Vedações", "Confirmar no projeto estrutural o sistema construtivo adotado."),
        ("Antena coletiva / TV por assinatura", "Confirmar em projeto as quantidades de pontos de TV a serem previstas."),
        ("Sistema de Telefonia", "Confirmar em projeto as quantidades de pontos de telefonia a serem previstas."),
        ("Elevadores", "Confirmar em projeto as quantidades de elevadores e de 'transfer' previstas."),
        ("Pressurização", "Confirmar se o projeto possui escada pressurizada ou ventilada."),
        ("Instalações Hidráulicas", "Confirmar diferenças entre grupos de unidades quanto a água quente, aquecimento a gás e/ou chuveiro elétrico."),
    ]
    for item, orient in regras:
        resultados.append({
            "Grupo / Aplicação": "Geral",
            "Padrão aplicado": padrao,
            "Área": "Especificações Gerais",
            "Ambiente": "Geral",
            "Seção": "ESPECIFICAÇÕES GERAIS",
            "Item": item,
            "Texto encontrado": "Verificação pontual / confirmação em projeto",
            "Especificação prevista": orient,
            "Status": STATUS_ATENCAO,
            "Orientação / resposta prevista": orient,
            "Observação": "Item dependente de confirmação do coordenador/projeto.",
            "Confiança": 1.0,
            "Fonte": "Protocolo Memorial do Cliente",
        })

    # Regras condicionais para Médio ou Misto contendo Médio
    tem_medio = padrao == "Médio" or (padrao == "Misto" and any(g.get("padrao") == "Médio" for g in grupos_mistos or []))
    if tem_medio:
        aplicacao = "Somente grupos/unidades de padrão Médio" if padrao == "Misto" else "Unidades padrão Médio"
        for item in ["Máquina de lavar louças", "Previsão de ar-condicionado"]:
            resultados.append({
                "Grupo / Aplicação": aplicacao,
                "Padrão aplicado": "Médio",
                "Área": "Especificações Gerais",
                "Ambiente": "Unidades",
                "Seção": "OBSERVAÇÕES GERAIS",
                "Item": item,
                "Texto encontrado": "Conferir atribuição no memorial",
                "Especificação prevista": f"O item deve ser atribuído apenas às unidades/grupos Médio ({aplicacao}).",
                "Status": STATUS_ATENCAO,
                "Orientação / resposta prevista": f"Confirmar que {item.lower()} aparece somente para as unidades de padrão Médio.",
                "Observação": "Em empreendimento misto, não generalizar este diferencial para unidades Econômico/Super Econômico.",
                "Confiança": 1.0,
                "Fonte": "Protocolo Memorial do Cliente",
            })
    return resultados


def auditoria_protocolo_cef(texto):
    nt = normalizar(texto)
    resultados = []
    for regra in PROTOCOLO_CEF:
        if not any(normalizar(g) in nt for g in regra["gatilhos"]):
            continue
        status = STATUS_ATENCAO if regra["acao"] == "atencao" else STATUS_INFO
        resultados.append({
            "Grupo / Aplicação": "Geral / CEF",
            "Padrão aplicado": "Conforme configuração do empreendimento",
            "Área": "Protocolo Financiador",
            "Ambiente": "Geral",
            "Seção": "ORIENTAÇÕES DA COORDENAÇÃO",
            "Item": regra["titulo"],
            "Texto encontrado": "Item/gatilho localizado no Memorial CEF",
            "Especificação prevista": regra["orientacao"],
            "Status": status,
            "Orientação / resposta prevista": regra["orientacao"],
            "Observação": "Camada adicional do protocolo CEF; não substitui a conferência técnica pela base R96.",
            "Confiança": 1.0,
            "Fonte": regra["fonte"],
        })
    return resultados


def auditar_memorial(texto, base, padrao, escopo, tipo_doc, grupos_mistos=None, incluir_gerais_cliente=True):
    resultados = auditar_base_r96(texto, base, padrao, escopo, grupos_mistos)

    if tipo_doc == "Memorial do Cliente (Comercial / Vendas)" and incluir_gerais_cliente:
        resultados.extend(auditoria_especificacoes_gerais_cliente(texto, padrao, grupos_mistos))

    if tipo_doc == "Memorial CEF / Financiador":
        resultados.extend(auditoria_protocolo_cef(texto))

    if not resultados:
        return pd.DataFrame(columns=[
            "Grupo / Aplicação", "Padrão aplicado", "Área", "Ambiente", "Seção", "Item",
            "Texto encontrado", "Especificação prevista", "Status", "Orientação / resposta prevista",
            "Observação", "Confiança", "Fonte"
        ])
    return pd.DataFrame(resultados)

# ==============================================================================
# DOCUMENTOS ANOTADOS
# ==============================================================================

def gerar_pdf_anotado(file_bytes, df):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    writer = pypdf.PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    pendencias = df[df["Status"].isin([STATUS_ERRO, STATUS_ATENCAO])]
    if not pendencias.empty:
        resumo = "QUALITY HUB - RESUMO DE AUDITORIA\n\n"
        for _, row in pendencias.head(60).iterrows():
            resumo += (
                f"[{row['Status']}] {row['Ambiente']} - {row['Item']}\n"
                f"Previsto: {row['Orientação / resposta prevista']}\n\n"
            )
        annotation = DictionaryObject({
            NameObject('/Type'): NameObject('/Annot'),
            NameObject('/Subtype'): NameObject('/Text'),
            NameObject('/Rect'): ArrayObject([FloatObject(50), FloatObject(700), FloatObject(80), FloatObject(730)]),
            NameObject('/Contents'): TextStringObject(resumo[:30000]),
            NameObject('/Open'): BooleanObject(False),
            NameObject('/Name'): NameObject('/Comment')
        })
        writer.add_annotation(page_number=0, annotation=annotation)

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output


def gerar_docx_anotado(file_bytes, df):
    doc = docx.Document(io.BytesIO(file_bytes))
    p = doc.add_paragraph()
    r = p.add_run("--- QUALITY HUB | RELATÓRIO DE AUDITORIA ---")
    r.font.bold = True
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW

    pendencias = df[df["Status"].isin([STATUS_ERRO, STATUS_ATENCAO])]
    for _, row in pendencias.head(100).iterrows():
        p = doc.add_paragraph()
        r = p.add_run(
            f"{row['Status']} [{row['Ambiente']} - {row['Item']}]\n"
            f"Encontrado: {row['Texto encontrado']}\n"
            f"Previsto: {row['Orientação / resposta prevista']}\n"
        )
        if row["Status"] == STATUS_ERRO:
            r.font.highlight_color = WD_COLOR_INDEX.RED
        else:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ==============================================================================
# INTERFACE
# ==============================================================================

def configurar_misto_sidebar():
    st.sidebar.markdown("### Distribuição do empreendimento misto")
    st.sidebar.caption("Cadastre os grupos. Ex.: Torre C | finais 01, 02, 05 e 06 | Médio")
    qtd = st.sidebar.number_input("Quantidade de grupos", min_value=2, max_value=20, value=2, step=1)
    grupos = []
    for i in range(int(qtd)):
        with st.sidebar.expander(f"Grupo {i+1}", expanded=(i < 2)):
            torre = st.text_input("Torre / bloco / grupo", value="Torre C" if i < 2 else "", key=f"torre_{i}")
            unidades = st.text_input(
                "Finais / tipologias / unidades",
                value="01, 02, 05 e 06" if i == 0 else ("demais finais" if i == 1 else ""),
                key=f"unidades_{i}",
            )
            padrao = st.selectbox("Padrão técnico", PADROES_TECNICOS, index=2 if i == 0 else 1, key=f"padrao_grupo_{i}")
            grupos.append({"torre": torre.strip(), "unidades": unidades.strip(), "padrao": padrao})
    return [g for g in grupos if g["torre"] and g["unidades"]]


def main():
    st.set_page_config(page_title="QUALITY HUB | Memoriais", page_icon="🏗️", layout="wide")

    st.title("QUALITY HUB")
    st.caption("Plataforma de Qualidade e Coordenação de Projetos | Módulo Memoriais")

    with st.sidebar:
        st.header("Configuração da análise")
        tipo_doc = st.selectbox(
            "Tipo de memorial",
            ["Memorial do Cliente (Comercial / Vendas)", "Memorial CEF / Financiador"],
        )
        padrao = st.selectbox(
            "Padrão / configuração do empreendimento",
            ["Super Econômico", "Econômico", "Médio", "Misto"],
            index=1,
        )

    grupos_mistos = configurar_misto_sidebar() if padrao == "Misto" else []

    with st.sidebar:
        escopo = st.selectbox("Escopo da conferência técnica", ["Área Privativa", "Área Comum"], index=0)
        if tipo_doc.startswith("Memorial do Cliente"):
            incluir_gerais_cliente = st.checkbox("Incluir checklist de Especificações Gerais", value=True)
        else:
            incluir_gerais_cliente = False

        st.markdown("---")
        st.subheader("Base técnica R96")
        base_upload = st.file_uploader(
            "Base de acabamentos (.xlsx) — opcional se o arquivo estiver junto do app",
            type=["xlsx"],
            key="base_r96",
        )
        st.caption(f"Arquivo esperado: {NOME_BASE_PADRAO}")

        st.subheader("Memorial para análise")
        memorial = st.file_uploader("PDF ou DOCX", type=["pdf", "docx"], key="memorial")

    # resumo da configuração
    c1, c2, c3 = st.columns(3)
    c1.metric("Tipo", "Cliente" if tipo_doc.startswith("Memorial do Cliente") else "Financiador / CEF")
    c2.metric("Configuração", padrao)
    c3.metric("Escopo", escopo)

    if padrao == "Misto":
        st.info("**Distribuição cadastrada:** " + descricao_config_mista(grupos_mistos))
        if len(grupos_mistos) < 2:
            st.warning("Cadastre pelo menos dois grupos para a configuração Misto.")

    if memorial is None:
        st.markdown("### Fluxo desta versão")
        st.markdown(
            """
            **R96 → ambiente → item → especificação prevista → trecho do memorial → resultado.**  
            Para empreendimentos **Mistos**, o grupo cadastrado roteia a conferência para o padrão técnico correspondente.  
            No **Memorial CEF**, as orientações da Coordenação entram como uma camada adicional de alertas de projeto/escopo.
            """
        )
        return

    if st.button("Executar conferência", type="primary", use_container_width=True):
        if padrao == "Misto" and len(grupos_mistos) < 2:
            st.error("Configure os grupos do empreendimento misto antes de executar.")
            return

        with st.spinner("Lendo a base R96 e analisando o memorial..."):
            try:
                base_bytes = base_upload.getvalue() if base_upload else None
                base = carregar_base_r96(base_bytes)
            except Exception as e:
                st.error(f"Não foi possível carregar a base R96: {e}")
                return

            file_bytes = memorial.getvalue()
            ext = memorial.name.rsplit(".", 1)[-1].lower()
            try:
                if ext == "pdf":
                    partes = extrair_texto_pdf(file_bytes)
                else:
                    partes = extrair_texto_docx(file_bytes)
                texto = "\n".join(p["texto"] for p in partes)
            except Exception as e:
                st.error(f"Erro ao ler o memorial: {e}")
                return

            if not texto.strip():
                st.error("O memorial não contém texto pesquisável suficiente para a conferência automática.")
                return

            df = auditar_memorial(
                texto=texto,
                base=base,
                padrao=padrao,
                escopo=escopo,
                tipo_doc=tipo_doc,
                grupos_mistos=grupos_mistos,
                incluir_gerais_cliente=incluir_gerais_cliente,
            )
            st.session_state["resultado_memorial"] = df
            st.session_state["memorial_bytes"] = file_bytes
            st.session_state["memorial_ext"] = ext
            st.session_state["memorial_nome"] = memorial.name

    df = st.session_state.get("resultado_memorial")
    if df is None:
        return

    st.markdown("---")
    st.subheader("Resultado da conferência")
    total = len(df)
    n_erro = int((df["Status"] == STATUS_ERRO).sum())
    n_at = int((df["Status"] == STATUS_ATENCAO).sum())
    n_ok = int((df["Status"] == STATUS_OK).sum())
    n_info = int((df["Status"] == STATUS_INFO).sum())

    a, b, c, d, e = st.columns(5)
    a.metric("Itens", total)
    b.metric("Divergências", n_erro)
    c.metric("Atenção", n_at)
    d.metric("Conforme", n_ok)
    e.metric("Sem conferência", n_info)

    filtro_status = st.multiselect(
        "Filtrar status",
        [STATUS_ERRO, STATUS_ATENCAO, STATUS_OK, STATUS_INFO],
        default=[STATUS_ERRO, STATUS_ATENCAO, STATUS_OK],
    )
    vis = df[df["Status"].isin(filtro_status)] if filtro_status else df

    st.dataframe(
        vis,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Confiança": st.column_config.ProgressColumn("Confiança", min_value=0, max_value=1, format="%.2f"),
            "Especificação prevista": st.column_config.TextColumn("Especificação prevista", width="large"),
            "Texto encontrado": st.column_config.TextColumn("Texto encontrado", width="large"),
            "Orientação / resposta prevista": st.column_config.TextColumn("Orientação / resposta prevista", width="large"),
        },
    )

    st.caption(
        "Importante: vermelho é reservado para divergência com evidência textual forte. "
        "Quando o contexto é insuficiente ou a decisão depende de projeto, o sistema mantém o item em amarelo."
    )

    st.markdown("### Exportações")
    col1, col2 = st.columns(2)
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Auditoria")
        if padrao == "Misto":
            pd.DataFrame(grupos_mistos).to_excel(writer, index=False, sheet_name="Config_Misto")
    buffer_excel.seek(0)
    col1.download_button(
        "Baixar relatório XLSX",
        data=buffer_excel,
        file_name=f"QUALITY_HUB_Auditoria_{Path(st.session_state['memorial_nome']).stem}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    mbytes = st.session_state["memorial_bytes"]
    ext = st.session_state["memorial_ext"]
    if ext == "pdf":
        anotado = gerar_pdf_anotado(mbytes, df)
        col2.download_button(
            "Baixar PDF anotado",
            data=anotado,
            file_name=f"Anotado_{st.session_state['memorial_nome']}",
            mime="application/pdf",
            use_container_width=True,
        )
    else:
        anotado = gerar_docx_anotado(mbytes, df)
        col2.download_button(
            "Baixar DOCX anotado",
            data=anotado,
            file_name=f"Anotado_{st.session_state['memorial_nome']}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
