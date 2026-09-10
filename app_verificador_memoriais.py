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
import fitz  # PyMuPDF: recortes visuais e PDF revisado sem sobreposição
from pypdf.generic import DictionaryObject, NameObject, TextStringObject, ArrayObject, FloatObject, BooleanObject
from docx.enum.text import WD_COLOR_INDEX

# ==============================================================================
# MIA | MEMORIAIS — V5
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
STATUS_INFO = "⚪ Não verificado"

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

# ------------------------------------------------------------------------------
# LOCALIZAÇÃO VISUAL NO PDF
# ------------------------------------------------------------------------------

def localizar_no_pdf(file_bytes, ambiente, item, trecho=""):
    """Retorna página (1-based) e retângulo aproximado do item no PDF.
    A busca é deliberadamente conservadora: primeiro ambiente, depois rótulo do item.
    """
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        aliases_amb = aliases_ambiente(ambiente)
        aliases_it = _aliases_item_rotulo(item) if '_aliases_item_rotulo' in globals() else [normalizar(item)]
        for pno, page in enumerate(doc):
            txt = normalizar(page.get_text("text"))
            if not any(a in txt for a in aliases_amb):
                continue
            amb_rects = []
            for a in aliases_amb:
                amb_rects += page.search_for(a, quads=False)
            for it in aliases_it:
                rects = page.search_for(it, quads=False)
                if rects:
                    # prefere ocorrência abaixo do cabeçalho do ambiente
                    if amb_rects:
                        ay = min(r.y0 for r in amb_rects)
                        abaixo = [r for r in rects if r.y0 >= ay - 4]
                        if abaixo:
                            r = min(abaixo, key=lambda x: x.y0)
                            return pno + 1, (r.x0, r.y0, r.x1, r.y1)
                    r = rects[0]
                    return pno + 1, (r.x0, r.y0, r.x1, r.y1)
        return None, None
    except Exception:
        return None, None


def recorte_ocorrencia_pdf(file_bytes, pagina, bbox=None, zoom=1.7):
    if not pagina:
        return None
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        page = doc[pagina - 1]
        if bbox:
            r = fitz.Rect(*bbox)
            clip = fitz.Rect(max(0, r.x0 - 55), max(0, r.y0 - 90), min(page.rect.width, r.x1 + 360), min(page.rect.height, r.y1 + 135))
        else:
            clip = page.rect
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
        return pix.tobytes("png")
    except Exception:
        return None

# ==============================================================================
# MOTOR DE CONFERÊNCIA R96
# ==============================================================================

# ==============================================================================
# MOTOR DE PAREAMENTO ESTRUTURAL — V5
# ==============================================================================

ITEM_EQUIVALENCIAS = {
    "piso": ["piso", "pisos"],
    "parede": ["parede", "paredes", "revestimento de parede"],
    "teto": ["teto", "forro", "revestimento de teto"],
    "rodape": ["rodape", "rodapé"],
    "bancada": ["bancada", "bancadas"],
    "louca": ["louca", "louças", "louca sanitaria", "louças sanitárias"],
    "metais": ["metal", "metais", "torneira", "misturador", "registro"],
    "porta": ["porta", "portas"],
    "janela": ["janela", "janelas"],
    "peitoril": ["peitoril", "peitoris"],
    "soleira": ["soleira", "soleiras"],
    "guarda corpo": ["guarda corpo", "guarda-corpo"],
    "portao": ["portao", "portão", "portoes", "portões"],
    "ponto de luz": ["ponto de luz", "pontos de luz", "iluminacao", "iluminação"],
    "interruptor": ["interruptor", "interruptores"],
    "tomada": ["tomada", "tomadas", "ponto de forca", "ponto de força"],
    "agua fria": ["agua fria", "água fria"],
    "agua quente": ["agua quente", "água quente"],
    "esgoto": ["esgoto", "ponto de esgoto"],
    "gas": ["gas", "gás", "ponto de gas", "ponto de gás"],
}


def aliases_item(item):
    n = normalizar(item)
    aliases = [n]
    for chave, vals in ITEM_EQUIVALENCIAS.items():
        if chave in n or n in chave:
            aliases.extend(normalizar(v) for v in vals)
    # O primeiro núcleo nominal costuma ser suficiente para títulos de linhas da matriz.
    for chave, vals in ITEM_EQUIVALENCIAS.items():
        if chave in n:
            aliases.extend(normalizar(v) for v in vals)
    return list(dict.fromkeys(a for a in aliases if len(a) >= 3))


def _linhas_texto(texto):
    linhas = []
    for i, raw in enumerate(str(texto).splitlines()):
        limpo = re.sub(r"\s+", " ", raw).strip()
        if limpo:
            linhas.append({"i": i, "raw": limpo, "norm": normalizar(limpo)})
    return linhas


def _linha_eh_ambiente(linha_norm, ambiente):
    # Ambiente precisa aparecer como cabeçalho/linha curta, não perdido dentro de uma descrição.
    # Remove marcadores típicos dos memoriais comerciais (•, ▪, hífen).
    ln = re.sub(r"^[^a-z0-9]+", "", linha_norm).strip()
    for alias in aliases_ambiente(ambiente):
        if ln == alias:
            return True
        if len(ln) <= 120 and (ln.startswith(alias + " ") or ln.startswith(alias + " -") or ln.startswith(alias + " –")):
            return True
    return False


def _linha_parece_novo_ambiente(raw):
    txt = raw.strip()
    if txt.startswith(("•", "▪", "- ")) and len(txt) <= 140:
        return True
    letras = [c for c in txt if c.isalpha()]
    if len(letras) >= 4 and len(txt) <= 95:
        prop = sum(c.isupper() for c in letras) / len(letras)
        if prop >= 0.82 and not re.match(r"^(PISO|PAREDE|PAREDES|TETO|FORRO|RODAP[EÉ]|BANCADA|LOU[CÇ]A|METAIS?|PORTA|JANELA|PEITORIL|SOLEIRA|EQUIPAMENTOS?)\b", txt, re.I):
            return True
    return False


def _aliases_item_rotulo(item):
    # Núcleos de item aceitos somente quando usados como RÓTULO.
    n = normalizar(item)
    nucleos = []
    for chave, vals in ITEM_EQUIVALENCIAS.items():
        if chave in n or n in chave:
            nucleos.extend([chave] + vals)
    # Casos compostos da planilha.
    if "bancad" in n or "louca" in n or "tanque" in n:
        nucleos += ["bancada", "bancadas", "louca", "loucas", "tanque", "tanques"]
    if "parede" in n or "sanca" in n:
        nucleos += ["parede", "paredes", "sanca", "sancas", "revestimento"]
    if "piso" in n:
        nucleos += ["piso", "pisos"]
    if "rodape" in n:
        nucleos += ["rodape", "rodapes"]
    if "soleira" in n:
        nucleos += ["soleira", "soleiras"]
    if not nucleos:
        nucleos = [n]
    return list(dict.fromkeys(normalizar(x) for x in nucleos if len(normalizar(x)) >= 3))


def _linha_tem_rotulo_item(linha_norm, item):
    for alias in _aliases_item_rotulo(item):
        # Fundamental: o nome do item deve estar no começo da linha/campo e seguido por ':' ou '-'.
        if re.match(r"^" + re.escape(alias) + r"\s*[:\-–—]", linha_norm):
            return True
        # Em algumas extrações o rótulo vem sozinho na linha.
        if linha_norm == alias:
            return True
    return False


def extrair_contexto_hierarquico(texto, ambiente, item, janela_ambiente=0):
    """Pareamento estrutural V5.

    1) encontra o AMBIENTE como cabeçalho;
    2) limita o bloco até o PRÓXIMO cabeçalho de ambiente;
    3) encontra o ITEM somente como rótulo explícito dentro desse bloco;
    4) captura o valor até o próximo rótulo de item.

    Assim, a palavra 'bancada' dentro da descrição de PAREDE nunca vira o item Bancada.
    """
    linhas = _linhas_texto(texto)
    if not linhas:
        return "", False, False

    pos_ambientes = [k for k, ln in enumerate(linhas) if _linha_eh_ambiente(ln["norm"], ambiente)]
    if not pos_ambientes:
        return "", False, False

    for pos in pos_ambientes:
        fim = len(linhas)
        for j in range(pos + 1, len(linhas)):
            if _linha_parece_novo_ambiente(linhas[j]["raw"]):
                fim = j
                break
        bloco = linhas[pos + 1:fim]
        if not bloco:
            continue

        idx_item = next((j for j, ln in enumerate(bloco) if _linha_tem_rotulo_item(ln["norm"], item)), None)
        if idx_item is None:
            continue

        capt = [bloco[idx_item]["raw"]]
        for j in range(idx_item + 1, len(bloco)):
            ln = bloco[j]
            # Próximo rótulo técnico encerra o campo atual.
            if any(_linha_tem_rotulo_item(ln["norm"], chave) for chave in ITEM_EQUIVALENCIAS):
                break
            if re.match(r"^(Piso|Paredes?|Teto|Forro|Rodap[eé]|Bancadas?|Lou[cç]as?|Tanque|Metais?|Portas?|Janelas?|Peitoris?|Soleiras?|Equipamentos?)\s*[:\-–—]", ln["raw"], re.I):
                break
            capt.append(ln["raw"])
            if len(" ".join(capt)) > 950:
                break
        trecho = " ".join(capt).strip()
        return trecho, True, True

    return "", True, False


def avaliar_regra_v4(trecho, regra):
    esp = str(regra["especificacao"]).strip()
    nt, ne = normalizar(trecho), normalizar(esp)
    if not trecho:
        return STATUS_INFO, "Item não localizado com vínculo seguro entre ambiente e item.", 0.0

    # Não aplicável só é conforme quando o próprio memorial também o declara.
    if "nao aplicavel" in ne:
        if "nao aplicavel" in nt:
            return STATUS_OK, "Memorial e R96 indicam item não aplicável.", 1.0
        return STATUS_INFO, "A base indica condição não aplicável, mas o trecho não permite confirmar a mesma condição.", 0.3

    sim = similaridade_textual(nt, ne)
    te, tt = tokens_significativos(ne), tokens_significativos(nt)
    cobertura = len(te & tt) / max(1, len(te))

    # Correspondência literal ou técnica forte => conforme.
    if ne in nt or nt in ne or cobertura >= 0.62 or sim >= 0.56:
        return STATUS_OK, "Descrição tecnicamente compatível com a base R96.", max(sim, cobertura)

    mats_e = materiais_presentes(ne)
    mats_t = materiais_presentes(nt)
    # Se o memorial escolhe uma das alternativas materiais explicitamente previstas, considera conforme.
    if mats_e and mats_t and (mats_t <= mats_e or any(m in nt and m in ne for m in sorted(mats_e, key=len, reverse=True))):
        extras = {m for m in mats_t if m not in mats_e}
        if not extras:
            return STATUS_OK, "Material/solução encontrada está entre as alternativas previstas no R96.", max(sim, cobertura, 0.75)
    # Vermelho somente com conflito material explícito, no ambiente + item já ancorados.
    if mats_e and mats_t and mats_e.isdisjoint(mats_t):
        return STATUS_ERRO, f"Divergência material objetiva: memorial indica {', '.join(sorted(mats_t))}; R96 prevê {', '.join(sorted(mats_e))}.", max(sim, cobertura)

    # Ausência de prova de conformidade NÃO vira atenção do coordenador.
    return STATUS_INFO, "Não foi possível concluir a comparação com segurança; não classificado como divergência.", max(sim, cobertura)


def filtrar_regras_v4(base, padrao, escopo):
    if escopo == "Área Comum":
        return base[base["escopo"] == "Área Comum"].copy()
    return base[(base["escopo"] == "Área Privativa") & (base["padrao"] == padrao)].copy()



def recortar_texto_por_escopo(texto, escopo):
    """Separa Área Comum e Área Privativa quando o memorial possui um divisor claro.
    Evita, por exemplo, que SALA privativa seja pareada com SALÃO/SALA de área comum.
    Em memoriais CEF, que repetem tabelas de área privativa/comum em várias seções,
    mantém o texto integral para não perder blocos técnicos.
    """
    n = normalizar(texto)
    marcadores_priv = ["unidades autonomas residenciais", "unidades autonomas", "apartamentos - area privativa"]
    pos = -1
    for m in marcadores_priv:
        p = n.find(m)
        if p >= 0:
            pos = p if pos < 0 else min(pos, p)
    if pos < 0:
        return texto
    # converte posição normalizada em aproximação na string original via busca sem acentos simples
    # Para o memorial comercial, o marcador aparece literalmente em linha própria.
    linhas = str(texto).splitlines()
    idx = next((i for i,l in enumerate(linhas) if any(m in normalizar(l) for m in marcadores_priv)), None)
    if idx is None:
        return texto
    if escopo == "Área Privativa":
        fim = next((j for j in range(idx+1, len(linhas)) if "especificacoes gerais" in normalizar(linhas[j])), len(linhas))
        return "\n".join(linhas[idx:fim])
    return "\n".join(linhas[:idx])

def auditar_conjunto_r96(texto, base, padrao_tecnico, escopo, grupo_nome):
    resultados = []
    regras = filtrar_regras_v4(base, padrao_tecnico, escopo)
    for _, regra in regras.iterrows():
        trecho, achou_amb, achou_item = extrair_contexto_hierarquico(texto, regra["ambiente"], regra["item"])
        if not achou_amb or not achou_item:
            # Não polui a lista com toda a matriz R96. O não-verificado fica disponível
            # apenas quando o ambiente existe mas o item não pôde ser vinculado.
            if achou_amb:
                resultados.append({
                    "Grupo / Aplicação": grupo_nome, "Padrão aplicado": padrao_tecnico,
                    "Área": escopo, "Ambiente": regra["ambiente"], "Seção": regra["secao"], "Item": regra["item"],
                    "Texto encontrado": "Não localizado com vínculo seguro", "Especificação prevista": regra["especificacao"],
                    "Status": STATUS_INFO, "Orientação / resposta prevista": "Sem ação automática - revisar somente se necessário.",
                    "Observação": "Ambiente localizado, mas o item não foi identificado com segurança.", "Confiança": 0.0,
                    "Fonte": regra["fonte"],
                })
            continue
        status, obs, conf = avaliar_regra_v4(trecho, regra)
        orient = "Nenhuma ação necessária." if status == STATUS_OK else (regra["especificacao"] if status == STATUS_ERRO else "Sem ação automática - vínculo insuficiente.")
        resultados.append({
            "Grupo / Aplicação": grupo_nome, "Padrão aplicado": padrao_tecnico,
            "Área": escopo, "Ambiente": regra["ambiente"], "Seção": regra["secao"], "Item": regra["item"],
            "Texto encontrado": trecho, "Especificação prevista": regra["especificacao"], "Status": status,
            "Orientação / resposta prevista": orient, "Observação": obs, "Confiança": round(float(conf), 2), "Fonte": regra["fonte"],
        })
    return resultados


def auditar_base_r96_v4(texto, base, padrao_predominante, excecoes=None):
    resultados = []
    # Sempre verifica área privativa + área comum. Escopo não é mais escolha do usuário.
    for escopo in ["Área Privativa", "Área Comum"]:
        texto_escopo = recortar_texto_por_escopo(texto, escopo)
        resultados.extend(auditar_conjunto_r96(texto_escopo, base, padrao_predominante, escopo, "Regra geral"))

    # Exceções de padrão: avalia somente um recorte identificado pela descrição fornecida.
    nt = normalizar(texto)
    for ex in excecoes or []:
        aplic = ex.get("aplicacao", "").strip()
        if not aplic:
            continue
        termos = [t.strip() for t in re.split(r"[-–—,:]", aplic) if len(t.strip()) >= 3]
        pos = next((nt.find(normalizar(t)) for t in termos if normalizar(t) and nt.find(normalizar(t)) >= 0), -1)
        if pos < 0:
            continue
        recorte = nt[max(0, pos-400):min(len(nt), pos+9000)]
        resultados.extend(auditar_conjunto_r96(recorte, base, ex["padrao"], "Área Privativa", aplic))
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


def auditar_memorial(texto, base, padrao, tipo_doc, excecoes=None, incluir_gerais_cliente=True):
    resultados = auditar_base_r96_v4(texto, base, padrao, excecoes)

    if tipo_doc == "Memorial do Cliente (Comercial / Vendas)" and incluir_gerais_cliente:
        resultados.extend(auditoria_especificacoes_gerais_cliente(texto, padrao, []))
    if tipo_doc == "Memorial CEF / Financiador":
        resultados.extend(auditoria_protocolo_cef(texto))

    cols = ["Grupo / Aplicação", "Padrão aplicado", "Área", "Ambiente", "Seção", "Item",
            "Texto encontrado", "Especificação prevista", "Status", "Orientação / resposta prevista",
            "Observação", "Confiança", "Fonte"]
    return pd.DataFrame(resultados, columns=cols) if resultados else pd.DataFrame(columns=cols)

# ==============================================================================
# DOCUMENTOS ANOTADOS
# ==============================================================================

def _texto_caixa(row):
    if row["Status"] == STATUS_ERRO:
        return (f"DIVERGÊNCIA — {row['Ambiente']} / {row['Item']}\n"
                f"Encontrado: {str(row['Texto encontrado'])[:180]}\n"
                f"Corrigir para: {str(row['Especificação prevista'])[:220]}")
    return (f"ATENÇÃO — {row['Item']}\n"
            f"Confirmar: {str(row['Orientação / resposta prevista'])[:320]}")


def gerar_pdf_anotado(file_bytes, df):
    """Cria PDF revisado com o memorial INTACTO à esquerda e faixa MIA à direita.
    Nenhuma caixa cobre o documento original.
    """
    src = fitz.open(stream=file_bytes, filetype="pdf")
    out = fitz.open()
    painel = 250

    # Agrupa ocorrências por página previamente calculada; sem página vai para a primeira.
    por_pg = {}
    for _, row in df[df["Status"].isin([STATUS_ERRO, STATUS_ATENCAO])].iterrows():
        pg = row.get("Página", None)
        try:
            pg = int(pg) if pg and not pd.isna(pg) else 1
        except Exception:
            pg = 1
        por_pg.setdefault(max(1, min(pg, len(src))), []).append(row)

    for pno, sp in enumerate(src, start=1):
        op = out.new_page(width=sp.rect.width + painel, height=sp.rect.height)
        op.show_pdf_page(fitz.Rect(0, 0, sp.rect.width, sp.rect.height), src, pno - 1)
        op.draw_line((sp.rect.width, 0), (sp.rect.width, sp.rect.height), color=(0.82,0.82,0.82), width=0.7)
        op.insert_text((sp.rect.width + 16, 24), "MIA · REVISÃO", fontsize=9, fontname="helv", color=(0.25,0.25,0.25))
        y = 42
        for row in por_pg.get(pno, []):
            texto = _texto_caixa(row)
            iserr = row["Status"] == STATUS_ERRO
            fill = (1.0,0.94,0.94) if iserr else (1.0,0.97,0.83)
            stroke = (0.80,0.12,0.12) if iserr else (0.62,0.43,0.0)
            h = 108 if iserr else 88
            if y + h > sp.rect.height - 18:
                # Não sobrepõe: interrompe e registra continuação compacta no rodapé.
                op.insert_text((sp.rect.width + 16, sp.rect.height - 16), "Demais apontamentos: consultar a tela/relatório MIA.", fontsize=6.5, fontname="helv", color=(0.35,0.35,0.35))
                break
            box = fitz.Rect(sp.rect.width + 12, y, sp.rect.width + painel - 12, y + h)
            op.draw_rect(box, color=stroke, fill=fill, width=0.8)
            op.insert_textbox(box + (8,8,-8,-8), texto, fontsize=7.1, fontname="helv", color=(0.10,0.10,0.10), lineheight=1.12)
            # linha de referência quando houver bbox
            bb = row.get("BBox", None)
            if isinstance(bb, (tuple, list)) and len(bb) == 4:
                try:
                    rr = fitz.Rect(*bb)
                    yy = min(max(rr.y0 + rr.height/2, 10), sp.rect.height-10)
                    op.draw_line((rr.x1 + 4, yy), (sp.rect.width + 12, y + 15), color=stroke, width=0.6)
                except Exception:
                    pass
            y += h + 10

    return out.tobytes(garbage=3, deflate=True)


def gerar_docx_anotado(file_bytes, df):
    doc = docx.Document(io.BytesIO(file_bytes))
    p = doc.add_paragraph(); r = p.add_run("--- MIA | AUDITORIA ---"); r.font.bold = True
    pendencias = df[df["Status"].isin([STATUS_ERRO, STATUS_ATENCAO])]
    for _, row in pendencias.head(100).iterrows():
        p = doc.add_paragraph(); r = p.add_run(_texto_caixa(row))
        r.font.highlight_color = WD_COLOR_INDEX.RED if row["Status"] == STATUS_ERRO else WD_COLOR_INDEX.YELLOW
    output = io.BytesIO(); doc.save(output); output.seek(0); return output

# ==============================================================================
# INTERFACE
# ==============================================================================

def configurar_excecoes_sidebar():
    st.sidebar.markdown("### Padrões diferentes (opcional)")
    possui = st.sidebar.checkbox("O empreendimento possui unidades com padrão diferente?", value=False)
    if not possui:
        return []
    qtd = st.sidebar.number_input("Quantidade de exceções", min_value=1, max_value=12, value=1, step=1)
    excecoes = []
    for i in range(int(qtd)):
        with st.sidebar.expander(f"Exceção {i+1}", expanded=True):
            aplicacao = st.text_input("Aplicação", placeholder="Ex.: Torre C - finais 01, 02, 05 e 06", key=f"aplic_{i}")
            pad = st.selectbox("Padrão técnico", PADROES_TECNICOS, index=2, key=f"pad_ex_{i}")
            if aplicacao.strip(): excecoes.append({"aplicacao": aplicacao.strip(), "padrao": pad})
    return excecoes


def main():
    st.set_page_config(page_title="MIA | Memoriais", page_icon="M", layout="wide")
    st.title("MIA")
    st.caption("Coordenação e Qualidade de Projetos | Memoriais")

    with st.sidebar:
        st.header("Configuração da análise")
        tipo_doc = st.selectbox("Tipo de memorial", ["Memorial do Cliente (Comercial / Vendas)", "Memorial CEF / Financiador"])
        padrao = st.selectbox("Padrão predominante do empreendimento", PADROES_TECNICOS, index=1)
    excecoes = configurar_excecoes_sidebar()
    with st.sidebar:
        if tipo_doc.startswith("Memorial do Cliente"):
            incluir_gerais_cliente = st.checkbox("Incluir checklist de Especificações Gerais", value=True)
        else:
            incluir_gerais_cliente = False
        st.markdown("---")
        st.caption("Base técnica R96 carregada automaticamente pelo sistema.")
        memorial = st.file_uploader("Memorial para análise (PDF ou DOCX)", type=["pdf", "docx"], key="memorial")

    c1,c2,c3 = st.columns(3)
    c1.metric("Tipo", "Cliente" if tipo_doc.startswith("Memorial do Cliente") else "Financiador / CEF")
    c2.metric("Padrão predominante", padrao)
    c3.metric("Exceções", len(excecoes))

    if memorial is None:
        st.info("Envie um memorial. A MIA verificará automaticamente Área Privativa e Área Comum contra a base R96.")
        return

    if st.button("Executar conferência", type="primary", use_container_width=True):
        with st.spinner("Analisando memorial e cruzando com a base R96..."):
            try:
                base = carregar_base_r96(None)
            except Exception as e:
                st.error(f"Não foi possível carregar a base R96 do repositório: {e}"); return
            file_bytes = memorial.getvalue(); ext = memorial.name.rsplit(".",1)[-1].lower()
            try:
                partes = extrair_texto_pdf(file_bytes) if ext == "pdf" else extrair_texto_docx(file_bytes)
                texto = "\n".join(p["texto"] for p in partes)
            except Exception as e:
                st.error(f"Erro ao ler o memorial: {e}"); return
            if not texto.strip(): st.error("O memorial não contém texto pesquisável suficiente."); return
            df = auditar_memorial(texto, base, padrao, tipo_doc, excecoes, incluir_gerais_cliente)
            if ext == "pdf" and not df.empty:
                paginas_loc, bboxes = [], []
                for _, rr in df.iterrows():
                    if rr["Status"] in [STATUS_ERRO, STATUS_ATENCAO] and rr["Área"] not in ["Especificações Gerais", "Protocolo Financiador"]:
                        pg, bb = localizar_no_pdf(file_bytes, str(rr["Ambiente"]), str(rr["Item"]), str(rr["Texto encontrado"]))
                    else:
                        pg, bb = (None, None)
                    paginas_loc.append(pg); bboxes.append(bb)
                df["Página"] = paginas_loc; df["BBox"] = bboxes
            st.session_state.update(resultado_memorial=df, memorial_bytes=file_bytes, memorial_ext=ext, memorial_nome=memorial.name)

    df = st.session_state.get("resultado_memorial")
    if df is None: return
    st.markdown("---"); st.subheader("Resultado da conferência")
    nerr=int((df["Status"]==STATUS_ERRO).sum()); natt=int((df["Status"]==STATUS_ATENCAO).sum()); ninfo=int((df["Status"]==STATUS_INFO).sum()); nok=int((df["Status"]==STATUS_OK).sum())
    a,b,c,d=st.columns(4); a.metric("Divergências",nerr); b.metric("Atenções",natt); c.metric("Não verificados",ninfo); d.metric("Conformes",nok)

    modo = st.radio("Exibir", ["Itens que exigem ação", "Não verificados", "Conformes", "Todos"], horizontal=True)
    if modo=="Itens que exigem ação": vis=df[df["Status"].isin([STATUS_ERRO,STATUS_ATENCAO])]
    elif modo=="Não verificados": vis=df[df["Status"]==STATUS_INFO]
    elif modo=="Conformes": vis=df[df["Status"]==STATUS_OK]
    else: vis=df

    if vis.empty: st.success("Nenhum item nesta categoria.")
    else:
        for idx,row in vis.head(150).iterrows():
            titulo=f"{row['Status']}  {row['Ambiente']} - {row['Item']}"
            with st.expander(titulo, expanded=row["Status"] in [STATUS_ERRO,STATUS_ATENCAO]):
                st.markdown(f"**Área:** {row['Área']}  |  **Padrão:** {row['Padrão aplicado']}")
                if row["Status"]==STATUS_ERRO:
                    if st.session_state.get("memorial_ext") == "pdf" and row.get("Página", None):
                        img = recorte_ocorrencia_pdf(st.session_state["memorial_bytes"], int(row["Página"]), row.get("BBox", None))
                        if img:
                            st.caption(f"Trecho do memorial · página {int(row['Página'])}")
                            st.image(img, use_container_width=False, width=760)
                    st.markdown(f"**Encontrado:** {row['Texto encontrado']}")
                    st.markdown(f"**Previsto:** {row['Especificação prevista']}")
                    st.markdown(f"**O que corrigir:** {row['Orientação / resposta prevista']}")
                elif row["Status"]==STATUS_ATENCAO:
                    if st.session_state.get("memorial_ext") == "pdf" and row.get("Página", None):
                        img = recorte_ocorrencia_pdf(st.session_state["memorial_bytes"], int(row["Página"]), row.get("BBox", None))
                        if img:
                            st.caption(f"Trecho do memorial · página {int(row['Página'])}")
                            st.image(img, use_container_width=False, width=760)
                    st.markdown(f"**O que confirmar:** {row['Orientação / resposta prevista']}")
                else:
                    st.markdown(f"**Encontrado:** {row['Texto encontrado']}")
                    st.markdown(f"**Previsto:** {row['Especificação prevista']}")
                st.caption(row["Fonte"])

    st.markdown("### Exportações")
    col1,col2=st.columns(2)
    xbio=io.BytesIO()
    with pd.ExcelWriter(xbio, engine="openpyxl") as wr: df.to_excel(wr,index=False,sheet_name="Auditoria")
    col1.download_button("Baixar relatório Excel", xbio.getvalue(), file_name=f"MIA_Auditoria_{Path(st.session_state['memorial_nome']).stem}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    if st.session_state["memorial_ext"]=="pdf":
        anot=gerar_pdf_anotado(st.session_state["memorial_bytes"],df)
        col2.download_button("Baixar PDF revisado com caixas de texto", anot, file_name=f"MIA_Revisado_{st.session_state['memorial_nome']}", mime="application/pdf", use_container_width=True)
    else:
        anot=gerar_docx_anotado(st.session_state["memorial_bytes"],df)
        col2.download_button("Baixar DOCX revisado", anot, file_name=f"MIA_Revisado_{st.session_state['memorial_nome']}", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)


if __name__ == "__main__":
    main()
