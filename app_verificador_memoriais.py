import io
import re
import unicodedata
from pathlib import Path
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
import pypdf
import docx
import fitz  # PyMuPDF
from docx.enum.text import WD_COLOR_INDEX

# ==============================================================================
# MIA | MEMORIAIS — V9
# Foco: associação técnica geral Área > Ambiente > Item > Subitem > R96,
# comparação por requisitos técnicos e preservação de evidência/página.
# ==============================================================================

APP_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
NOME_BASE_PADRAO = "PADRÃO DE ACABAMENTOS-R96.xlsx"
CAMINHOS_BASE = [APP_DIR / NOME_BASE_PADRAO, Path.cwd() / NOME_BASE_PADRAO]

PADROES_TECNICOS = ["Super Econômico", "Econômico", "Médio"]
STATUS_OK = "🟢 Conforme"
STATUS_ERRO = "🔴 Divergência"
STATUS_ATENCAO = "🟡 Atenção do Coordenador"
STATUS_INFO = "⚪ Não verificado"

# ------------------------------------------------------------------------------
# Protocolo adicional do financiador (não substitui o R96)
# ------------------------------------------------------------------------------
PROTOCOLO_CEF = [
    ("Estrutura das torres / garagem", ["estrutura", "supraestrutura", "garagem"],
     "Confirmar em projeto o tipo de estrutura das torres e, quando houver garagem, a solução adotada.",
     "Orientações Coordenação CEF — pág. 2"),
    ("Ambientes e finais de unidades", ["finais", "final", "ambientes"],
     "Confirmar em projeto os ambientes citados e, quando houver finais/unidades, a aplicação indicada.",
     "Orientações Coordenação CEF — págs. 4 a 6"),
    ("Portas e janelas — dimensões", ["portas", "janelas", "esquadrias"],
     "Confirmar no projeto de arquitetura as dimensões e a aplicabilidade das esquadrias.",
     "Orientações Coordenação CEF — pág. 7"),
    ("Acabamento conforme sistema estrutural", ["monocapa", "textura acrilica", "fachada"],
     "Confirmar o sistema estrutural para validar a solução de acabamento aplicável.",
     "Orientações Coordenação CEF — pág. 7"),
    ("Rede e prumadas de gás", ["rede de gas", "prumada", "gas"],
     "Confirmar em projeto hidráulico a rede/prumadas de gás e a solução aplicável.",
     "Orientações Coordenação CEF — pág. 12"),
    ("Pressurização das escadas", ["pressurizacao", "escada pressurizada", "sala de pressurizacao"],
     "Confirmar no projeto se a escada é pressurizada ou ventilada e ajustar a redação quando necessário.",
     "Orientações Coordenação CEF — pág. 12"),
    ("Aquecimento solar", ["aquecimento solar"],
     "Confirmar em projeto se existe previsão de aquecimento solar.",
     "Orientações Coordenação CEF — pág. 13"),
]

# ------------------------------------------------------------------------------
# Equivalências controladas de ambientes
# ------------------------------------------------------------------------------
# A associação automática só é feita quando existe evidência lexical real.
# Estes aliases resolvem nomes usuais dos memoriais que não são idênticos ao R96.
ALIASES_AMBIENTES_PRIV = {
    "COZINHA": ["cozinha"],
    "APA OU STUDIO": ["apa", "studio", "estudio"],
    "ÁREA DE SERVIÇO": ["area de servico", "lavanderia da unidade"],
    "LAVATÓRIO EXTERNO": ["lavatorio externo"],
    "BANHOS": ["banheiro", "banheiros", "banho", "banhos", "wc"],
    "DORMITÓRIOS": ["dormitorio", "dormitorios", "quarto", "quartos"],
    "SALA": ["sala", "sala de jantar/estar", "sala de jantar/estar e circulacao", "sala de estar", "sala de jantar"],
    "VARANDA COM A.S": ["varanda com a.s", "varanda com as", "varanda com area de servico", "terraco com area de servico", "terraço com área de serviço", "terraco com a.s", "terraço com a.s"],
    "VARANDA SEM A.S": ["varanda sem a.s", "varanda sem as", "varanda", "terraco sem area de servico", "terraço sem área de serviço", "terraco", "terraço"],
    "CIRCULAÇÃO": ["circulacao"],
}

ALIASES_AMBIENTES_COMUM = {
    "PORTARIA/ADM": ["portaria", "administracao", "administração", "adm"],
    "BANHEIROS DE ÁREAS COMUNS E PORTARIA (torres e churrasqueiras)": [
        "sanitario da portaria", "sanitários da portaria", "sanitarios da portaria",
        "banheiro da portaria", "banheiros das areas comuns", "sanitarios das areas comuns",
        "sanitários e sanitários pcd das áreas comuns", "sanitarios e sanitarios pcd das areas comuns",
    ],
    "ACESSO DE PEDESTRES E DE VEÍCULOS": ["acesso de pedestres e acesso de veiculos", "acesso de pedestres e de veiculos"],
    "CALÇADA EXTERNA AO EMPREENDIMENTO E ESTACIONAMENTO DE VISITANTES": ["passeio externo de pedestres", "calcada", "calçada externa ao empreendimento"],
    "ÁREA EXTERNA DA TORRE, HALLS EXTERNOS DE ACESSO ÀS TORRES E ÁREAS COBERTAS SEM FECHAMENTO DAS TORRES E MUROS DE DIVISA. (voltadas para a área externa)": [
        "circulacao externa das torres", "área externa da torre", "area externa da torre"
    ],
    "ÁREA DA PISCINA": ["piscina adulto", "piscina infantil", "solario das piscinas", "solário das piscinas", "area da piscina"],
    "CHURRASQUEIRA E/OU ESPAÇO GOURMET EXTERNO": ["churrasqueira", "espaco gourmet externo", "espaço gourmet externo"],
    "COPA ou APA DO SALÃO DE FESTAS E/OU DO ESPAÇO GOURMET": ["copa do salao de festas", "apa do salao de festas", "copa ou apa do salao de festas"],
    "COPA ou APA DE FUNCIONÁRIOS": ["apa de funcionario", "apa de funcionários", "apa de funcionarios", "copa de funcionarios"],
    "COWORKING": ["coworking"],
    "BEAUTY CARE": ["beauty care"],
    "BICICLETÁRIO": ["bicicletario", "bicicletários", "bicicletarios"],
    "DEPÓSITO DE MATERIAL DE LIMPEZA": ["deposito de material de limpeza", "dml"],
    "DEPÓSITO DE LIXO": ["deposito de lixo"],
    "DEPÓSITO PRIVATIVO": ["deposito privativo"],
    "PET CARE": ["pet care"],
    "ESPAÇO DELIVERY": ["delivery", "espaco delivery"],
    "HALL DOS ANDARES E DO EDIFÍCIO GARAGEM": ["hall social e circulacao social dos andares", "hall dos andares"],
    "HALLS SOCIAIS E CIRCULAÇÃO DOS TÉRREOS, SALÕES DE JOGOS, SALA DE POKER E FESTAS, ESPAÇO GOURMET, OFFICE, SPORTS BAR E SALAS DAS ÁREAS COMUNS": [
        "salao de festas", "salão de festas", "salao de jogos", "salão de jogos", "espaco gourmet", "espaço gourmet", "office", "sports bar"
    ],
    "LAVANDERIA COLETIVA": ["lavanderia coletiva"],
    "MINI MARKET": ["mini market", "mini mercado", "minimercado"],
    "OFICINA BIKE": ["oficina de bike", "oficina bike"],
    "SALA DE GINÁSTICA,  ESPAÇO PILATES E DEMAIS ÁREAS DE PRÁTICA DE ESPORTES E BRINQUEDOTECA": [
        "fitness", "espaco pilates", "espaço pilates", "brinquedoteca", "sala de ginastica", "sala de ginástica"
    ],
    "SALA DE PRESSURIZAÇÃO DAS ESCADAS": ["sala de pressurizacao", "sala de pressurização"],
    "SAUNA E SALA DE DESCANSO": ["sauna", "sala de descanso"],
    "VESTIÁRIOS E BANHEIROS DAS ÁREAS TÉCNICAS": ["vestiario de funcionario", "vestiário de funcionário", "vestiarios", "vestiários"],
    "CENTRO DE MEDIÇÃO, DG, OUTRAS ÁREAS TÉCNICAS E DEPÓSITOS e ETE": ["centro de medicao", "centro de medição", "dg", "ete", "areas tecnicas", "áreas técnicas"],
    "BARRILETE (RESERVATÓRIOS SUPERIORES), POÇO DE ELEVADOR E CASA DE BOMBAS DA PISCINA E ÁREAS SEM USO.": [
        "barrilete", "reservatorio superior", "reservatórios superiores", "poco de elevador", "poço de elevador", "casa de bombas"
    ],
    "RESERVATÓRIOS INFERIORES": ["reservatorio inferior", "reservatórios inferiores"],
    "ESCADA COBERTA (TORRE )": ["escada e circulacao tecnica da torre", "escada coberta torre"],
    "ESCADAS E RAMPAS DE PEDESTRES DESCOBERTAS": ["escadas e rampas de pedestres descobertas"],
}

# ------------------------------------------------------------------------------
# Utilitários
# ------------------------------------------------------------------------------
def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).replace("\n", " ").strip().lower()
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def tokens(texto):
    stop = {"de","da","do","das","dos","e","em","com","ou","para","por","quando","onde","no","na","nos","nas","um","uma","ao","aos","conforme","aplicavel","ser","sera","caso","houver","sobre"}
    return {x for x in re.findall(r"[a-z0-9]+", normalizar(texto)) if len(x) >= 3 and x not in stop}


def _limpar_cabecalho(txt):
    txt = re.sub(r"^[•▪●\-–—\s]+", "", str(txt)).strip()
    txt = re.sub(r"\s+[–—-]\s+TORRE\s+[A-Z0-9]+.*$", "", txt, flags=re.I)
    # qualificadores de aplicação não fazem parte do nome técnico do ambiente
    txt = re.sub(r"\s*\((?:UNIDADES?|FINAIS?|TORRES?).*?\)\s*$", "", txt, flags=re.I)
    return txt.strip(" .:-–—")


def _bbox_union(rects):
    if not rects:
        return None
    r = fitz.Rect(rects[0])
    for rr in rects[1:]:
        r.include_rect(fitz.Rect(rr))
    return (r.x0, r.y0, r.x1, r.y1)


def pagina_int(valor):
    try:
        if valor is None or pd.isna(valor):
            return None
        v = int(float(valor))
        return v if v > 0 else None
    except Exception:
        return None

# ==============================================================================
# Base R96
# ==============================================================================
def _achar_linha_ambientes(df):
    for i in range(min(len(df), 20)):
        if normalizar(df.iloc[i, 0]) == "ambientes":
            return i
    raise ValueError("Não foi encontrada a linha AMBIENTES na planilha R96.")


def _eh_cabecalho_secao(valor, linha):
    v = normalizar(valor)
    if not v:
        return False
    vazios = all((x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == "") for x in linha[1:])
    if not vazios:
        return False
    return any(k in v for k in ["instalacoes eletricas", "instalacoes hidraulicas", "acabamentos"])


def parsear_matriz_planilha(df, padrao, escopo, nome_aba):
    regras = []
    idx = _achar_linha_ambientes(df)
    ambientes = []
    for c in range(1, df.shape[1]):
        v = df.iloc[idx, c]
        ambientes.append("" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip())
    secao = "ACABAMENTOS"
    for r in range(idx + 1, len(df)):
        raw = df.iloc[r, 0] if df.shape[1] else None
        item = "" if raw is None or (isinstance(raw, float) and pd.isna(raw)) else str(raw).strip()
        if not item:
            continue
        if _eh_cabecalho_secao(raw, list(df.iloc[r, :])):
            secao = item.strip().upper()
            continue
        if normalizar(item) in {"atualizado em", "r", "r:"}:
            continue
        for c, ambiente in enumerate(ambientes, start=1):
            if not ambiente:
                continue
            v = df.iloc[r, c] if c < df.shape[1] else None
            if v is None or (isinstance(v, float) and pd.isna(v)) or not str(v).strip():
                continue
            regras.append({
                "padrao": padrao, "escopo": escopo, "ambiente": ambiente,
                "secao": secao, "item": item, "especificacao": str(v).strip(),
                "fonte": f"Padrão de Acabamentos R96 — {nome_aba}",
            })
    return regras


@st.cache_data(show_spinner=False)
def carregar_base_r96():
    caminho = next((p for p in CAMINHOS_BASE if p.exists()), None)
    if caminho is None:
        raise FileNotFoundError(f"Base {NOME_BASE_PADRAO} não encontrada no repositório.")
    xls = pd.ExcelFile(caminho)
    regras = []
    mapa = [
        ("SUPER ECO-ECONOMICO_PRIVATIVA", "Super Econômico", "Área Privativa"),
        ("SUPER ECO-ECONOMICO_PRIVATIVA", "Econômico", "Área Privativa"),
        ("MÉDIO_PRIVATIVA", "Médio", "Área Privativa"),
        ("ÁREA COMUM pagina 1-2", "Todos", "Área Comum"),
        ("ÁREA COMUM pagina 2-2", "Todos", "Área Comum"),
    ]
    for aba, pad, esc in mapa:
        if aba in xls.sheet_names:
            regras.extend(parsear_matriz_planilha(pd.read_excel(xls, aba, header=None, dtype=object), pad, esc, aba))
    return pd.DataFrame(regras)

# ==============================================================================
# Canonização e associação Ambiente / Item
# ==============================================================================
BUILD_VERSION = "V9.0"

def canon_item(nome):
    """Canonização V9: famílias técnicas distintas e combinações controladas."""
    n = normalizar(nome)
    # itens combinados da matriz: escolhe família dominante sem apagar subitens no fallback
    if "parede" in n and "sanca" in n: return "parede"
    if "tanque" in n and not "bancad" in n: return "tanque"
    if "bancad" in n: return "bancada"
    if "louca" in n or "bacia" in n or "lavatorio" in n: return "louca"
    if "guarda corpo" in n or "guarda-corpo" in n: return "guarda_corpo"
    if "portao" in n: return "portao"
    if "janela" in n: return "janela"
    if "porta" in n: return "porta"
    if "esquadr" in n and any(x in n for x in ["especial", "vp", "ventilacao", "veneziana", "caixilho"]): return "esquadria_especial"
    if "esquadr" in n: return "esquadria_geral"
    if "peitoril" in n: return "peitoril"
    if any(x in n for x in ["soleira", "baguete", "tento", "meia soleira", " bit "]): return "soleira_baguete"
    if "rodape" in n: return "rodape"
    if "parede" in n or "revestimento" in n: return "parede"
    if "sanca" in n: return "sanca"
    if "teto" in n or "forro" in n: return "teto"
    if "piso" in n: return "piso"
    if "metal" in n or "torneira" in n or "misturador" in n or "registro" in n: return "metais"
    if "agua fria" in n: return "agua_fria"
    if "agua quente" in n: return "agua_quente"
    if "esgoto" in n: return "esgoto"
    if "ponto" in n and "luz" in n: return "pontos_luz"
    if "interrupt" in n: return "interruptor"
    if any(x in n for x in ["tomada", "forca", "força"]): return "tomada_forca"
    if "telecom" in n: return "telecom"
    if "interfone" in n: return "interfone"
    if "cigarra" in n: return "cigarra"
    if "gas" in n: return "gas"
    return n[:100]

def canon_item_contexto(rotulo, valor):
    """Refina rótulos genéricos usando SOMENTE o conteúdo do próprio campo."""
    base = canon_item(rotulo)
    nv = normalizar(valor)
    if base == "esquadria_geral":
        if "guarda corpo" in nv or "guarda-corpo" in nv:
            return "guarda_corpo"
        if re.search(r"\bporta\b", nv):
            return "porta"
        if re.search(r"\bjanela\b", nv):
            return "janela"
        if any(x in nv for x in ["ventilacao", "veneziana", "caixilho", " vp ", "vp inferior", "vp superior"]):
            return "esquadria_especial"
    return base

def familias_item_regra(nome, especificacao=""):
    """Famílias possíveis de uma linha R96 sem fundir subitens distintos."""
    n = normalizar(nome); e = normalizar(especificacao)
    combinado_blt = ("bancad" in n and "louca" in n) or ("bancad" in n and "tanque" in n) or ("louca" in n and "tanque" in n)
    fam = set() if combinado_blt else {canon_item(nome)}
    if "parede" in n and "sanca" in n:
        fam.update({"parede", "sanca"})
    if any(x in n for x in ["bancad", "louca", "tanque"]):
        # Para linhas combinadas, a especificação decide o subitem.
        if "bancada" in e or (not combinado_blt and "bancad" in n):
            fam.add("bancada")
        if "tanque" in e or (not combinado_blt and "tanque" in n):
            fam.add("tanque")
        if any(x in e for x in ["bacia", "lavatorio", "louca sanitaria"]) or (not combinado_blt and "louca" in n):
            fam.add("louca")
        # Quando a descrição é inequivocamente de bancada mas omite a palavra bancada.
        if combinado_blt and not fam and any(x in e for x in ["granito com cuba", "marmore com cuba", "cuba de embutir"]):
            fam.add("bancada")
    if "esquadr" in n:
        if any(x in e for x in ["vp", "ventilacao", "veneziana", "caixilho"]): fam.add("esquadria_especial")
        if "janela" in e: fam.add("janela")
        if "porta" in e: fam.add("porta")
        if "guarda corpo" in e or "guarda-corpo" in e: fam.add("guarda_corpo")
    return fam


ROTULOS_CLIENTE = [
    "Piso Veículos", "Piso Pedestres", "Piso", "Paredes", "Parede", "Teto", "Forro", "Rodapé",
    "Bancada", "Bancadas", "Louça", "Louças", "Tanque", "Tanques", "Metais", "Peitoril", "Peitoris",
    "Soleira", "Soleiras", "Baguete", "Baguetes", "Tento", "Tentos", "Janela", "Janelas", "Porta", "Portas", "Esquadrias",
]
ROT_RE = re.compile(r"^(" + "|".join(re.escape(x) for x in sorted(ROTULOS_CLIENTE, key=len, reverse=True)) + r")\s*[:\-–—]\s*(.*)$", re.I)
STOP_RE = re.compile(r"^(Equipamentos?|Fechamento|Borda|Revestimento|Comunicação|Interfone|Ar[ -]?condicionado)\s*[:\-–—]", re.I)


def aliases_do_ambiente(base_nome, escopo):
    mapa = ALIASES_AMBIENTES_PRIV if escopo == "Área Privativa" else ALIASES_AMBIENTES_COMUM
    vals = [base_nome]
    vals += mapa.get(base_nome, [])
    return list(dict.fromkeys(normalizar(v) for v in vals if v))


def score_ambiente(doc_nome, base_nome, escopo):
    a, b = normalizar(_limpar_cabecalho(doc_nome)), normalizar(base_nome)
    if not a or not b:
        return 0.0
    aliases = aliases_do_ambiente(base_nome, escopo)
    if a in aliases:
        return 1.0
    # contenção só serve como apoio; quanto mais específico o alias, maior o peso
    cont_scores=[]
    for x in aliases:
        if len(x) >= 5 and (x in a or a in x):
            # qualificadores mudam o ambiente: "fitness externo" não é automaticamente
            # a sala de ginástica interna só porque contém a palavra fitness.
            if "extern" in a and "extern" not in x and a != x:
                continue
            coverage = min(len(x), len(a)) / max(len(x), len(a))
            cont_scores.append(0.78 + 0.18 * coverage)
    if cont_scores:
        return max(cont_scores)
    ta, tb = tokens(a), tokens(b)
    inter = ta & tb
    if not inter:
        return 0.0
    jac = len(inter) / max(1, len(ta | tb))
    cov = len(inter) / max(1, min(len(ta), len(tb)))
    seq = SequenceMatcher(None, a, b).ratio()
    return 0.45 * cov + 0.35 * jac + 0.20 * seq

def parear_ambiente(doc_nome, regras, escopo):
    nomes = list(dict.fromkeys(str(x) for x in regras["ambiente"].dropna()))
    if not nomes:
        return None, 0.0
    a = normalizar(_limpar_cabecalho(doc_nome))
    # 1) alias exato: decisão determinística
    exatos=[]
    for n in nomes:
        als=aliases_do_ambiente(n, escopo)
        if a in als:
            exatos.append((max(len(x) for x in als if x==a), n))
    if exatos:
        return sorted(exatos, reverse=True)[0][1], 1.0
    # 2) demais casos com margem de segurança
    scored = sorted(((score_ambiente(doc_nome, n, escopo), n) for n in nomes), reverse=True)
    s1, n1 = scored[0]
    s2 = scored[1][0] if len(scored) > 1 else 0.0
    if s1 < 0.80 or (s1 < 0.96 and s1 - s2 < 0.12):
        return None, s1
    return n1, s1

def parear_regra_item(regras_amb, canon, valor_encontrado=""):
    """Escolhe a regra do MESMO subitem; similaridade textual nunca troca a família técnica."""
    if regras_amb.empty:
        return None
    candidatos = []
    nf = normalizar(valor_encontrado)
    for idx, row in regras_amb.iterrows():
        fam = familias_item_regra(row.get("item", ""), row.get("especificacao", ""))
        if canon not in fam:
            continue
        score = 1.0 if canon_item(row.get("item", "")) == canon else 0.72
        ni = normalizar(row.get("item", "")); ne = normalizar(row.get("especificacao", ""))
        # reforços do subitem explícito; evitam Tanque <- Bancada e Esquadria especial <- Guarda-corpo
        pistas = {
            "tanque": ["tanque"], "bancada": ["bancad", "granito", "marmore", "inox"],
            "louca": ["louca", "bacia", "lavatorio"], "janela": ["janela"], "porta": ["porta"],
            "guarda_corpo": ["guarda corpo", "guarda-corpo"],
            "esquadria_especial": ["vp", "ventilacao", "veneziana", "caixilho"],
        }.get(canon, [])
        if pistas:
            score += 0.12 * sum(1 for x in pistas if x in ni or x in ne)
        # pequeno desempate por vocabulário do trecho, nunca suficiente para mudar a família
        if nf:
            score += min(0.12, 0.12 * similaridade(nf, ne))
        candidatos.append((score, idx))
    if not candidatos:
        return None
    candidatos.sort(reverse=True)
    return regras_amb.loc[candidatos[0][1]]


# ==============================================================================
# Extração PDF Cliente — com página + bbox + escopo real
# ==============================================================================
def linhas_pdf(file_bytes):
    """Reconstrói linhas lógicas do PDF.

    Em muitos memoriais o rótulo e o valor são dois blocos na MESMA altura
    (ex.: 'Piso:' em x=99 e 'Cerâmica' em x=191). A V6 tratava isso como
    duas linhas distintas e perdia quase todos os itens. Aqui agrupamos
    fragmentos com y semelhante e concatenamos da esquerda para a direita.
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    saida = []
    for pno, page in enumerate(doc, start=1):
        d = page.get_text("dict")
        frags = []
        for block in d.get("blocks", []):
            if "lines" not in block:
                continue
            for line in block["lines"]:
                spans = line.get("spans", [])
                txt = "".join(s.get("text", "") for s in spans).strip()
                if not txt:
                    continue
                bb = _bbox_union([sp.get("bbox") for sp in spans if sp.get("bbox")])
                if bb:
                    frags.append({"raw": re.sub(r"\s+", " ", txt).strip(), "bbox": bb})
        frags.sort(key=lambda x: (x["bbox"][1], x["bbox"][0]))
        grupos = []
        tol_y = 2.2
        for f in frags:
            yc = (f["bbox"][1] + f["bbox"][3]) / 2
            alvo = None
            for g in reversed(grupos[-5:]):
                if abs(yc - g["yc"]) <= tol_y:
                    alvo = g
                    break
            if alvo is None:
                grupos.append({"yc": yc, "frags": [f]})
            else:
                alvo["frags"].append(f)
                ys = [((z["bbox"][1]+z["bbox"][3])/2) for z in alvo["frags"]]
                alvo["yc"] = sum(ys)/len(ys)
        for g in grupos:
            fs = sorted(g["frags"], key=lambda x: x["bbox"][0])
            raw = " ".join(x["raw"] for x in fs).strip()
            bb = _bbox_union([x["bbox"] for x in fs])
            saida.append({"page": pno, "raw": raw, "bbox": bb})
    return saida


def _marcador_escopo(n):
    if any(x in n for x in ["unidades autonomas residenciais", "area privativa"]):
        return "Área Privativa"
    if any(x in n for x in ["areas comuns sociais", "area comum", "area uso comum"]):
        return "Área Comum"
    if "especificacoes gerais" in n:
        return "Geral"
    return None


def _parece_ambiente_cliente(raw):
    txt = _limpar_cabecalho(raw)
    if not txt or len(txt) > 130:
        return False
    n = normalizar(txt)
    if n in {"areas externas", "areas internas", "areas comuns sociais", "areas comuns", "unidades autonomas residenciais", "unidades autonomas", "especificacoes gerais"}:
        return False
    letras = [c for c in txt if c.isalpha()]
    caps = (sum(c.isupper() for c in letras) / max(1, len(letras))) if letras else 0
    return raw.strip().startswith(("•","▪","●")) or caps >= 0.78


def extrair_cliente_pdf(file_bytes, base, padrao):
    linhas = linhas_pdf(file_bytes)
    resultados = []
    escopo = "Área Comum"  # memorial comercial normalmente inicia nas áreas comuns
    ambiente_doc = None
    ambiente_base = None
    score_amb = 0.0
    atual = None

    def fechar():
        nonlocal atual
        if atual and atual["valor"].strip() and ambiente_base:
            atual["item_canon"] = canon_item_contexto(atual.get("rotulo", ""), atual.get("valor", ""))
            resultados.append(atual)
        atual = None

    for ln in linhas:
        raw, n = ln["raw"], normalizar(ln["raw"])
        mk = _marcador_escopo(n)
        if mk:
            fechar()
            escopo = mk
            ambiente_doc = ambiente_base = None
            score_amb = 0.0
            continue
        if escopo == "Geral":
            continue

        m = ROT_RE.match(_limpar_cabecalho(raw))
        if m and ambiente_base:
            fechar()
            atual = {
                "escopo": escopo, "ambiente_doc": ambiente_doc, "ambiente_base": ambiente_base,
                "score_ambiente": score_amb, "item_canon": canon_item_contexto(m.group(1), m.group(2).strip()),
                "rotulo": m.group(1), "valor": f"{m.group(1)}: {m.group(2).strip()}".strip(),
                "page": ln["page"], "bbox": ln["bbox"],
            }
            continue

        if _parece_ambiente_cliente(raw):
            regras_esc = base[base["escopo"] == escopo]
            if escopo == "Área Privativa":
                regras_esc = regras_esc[regras_esc["padrao"] == padrao]
            nome, sc = parear_ambiente(_limpar_cabecalho(raw), regras_esc, escopo)
            fechar()
            ambiente_doc = _limpar_cabecalho(raw)
            ambiente_base = nome
            score_amb = sc
            continue

        if atual:
            # não absorve outro campo (Equipamentos, Fechamento, Borda...) no valor atual.
            if ROT_RE.match(_limpar_cabecalho(raw)) or STOP_RE.match(_limpar_cabecalho(raw)) or _parece_ambiente_cliente(raw):
                fechar()
            else:
                atual["valor"] += " " + raw.strip()
                if atual["bbox"] and ln["page"] == atual["page"] and ln["bbox"]:
                    atual["bbox"] = _bbox_union([atual["bbox"], ln["bbox"]])
                if len(atual["valor"]) > 1300:
                    fechar()
    fechar()
    return resultados

# ==============================================================================
# Extração PDF CEF — usa as tabelas do próprio PDF
# ==============================================================================
def _split_piso_rodape_soleira(texto):
    """Separa a célula combinada do CEF em itens técnicos quando o texto permite."""
    t = re.sub(r"\s+", " ", str(texto or "")).strip()
    if not t:
        return []
    partes = re.split(r"(?<=[\.;])\s+", t)
    grupos = {"piso": [], "rodape": [], "soleira_baguete": []}
    for p in partes:
        np = normalizar(p)
        if any(x in np for x in ["rodape"]): grupos["rodape"].append(p)
        elif any(x in np for x in ["soleira", "baguete", "tento"]): grupos["soleira_baguete"].append(p)
        else: grupos["piso"].append(p)
    out = []
    for k, arr in grupos.items():
        if arr:
            out.append((k, " ".join(arr)))
    return out


def _header_canon(h):
    n = normalizar(h)
    if "piso" in n and ("rodape" in n or "soleira" in n): return "piso_combo"
    if n == "parede" or "parede" in n: return "parede"
    if n == "teto" or "teto" in n: return "teto"
    if "peitoril" in n: return "peitoril"
    if "agua fria" in n: return "agua_fria"
    if "agua quente" in n: return "agua_quente"
    if "esgoto" in n: return "esgoto"
    if "pontos de luz" in n or "ponto de luz" in n: return "pontos_luz"
    if "interrupt" in n: return "interruptor"
    if "forca" in n or "tomada" in n: return "tomada_forca"
    if "telecom" in n: return "telecom"
    if "interfone" in n: return "interfone"
    if "cigarra" in n: return "cigarra"
    return None


def extrair_cef_pdf(file_bytes, base, padrao):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    registros = []
    for pno, page in enumerate(doc, start=1):
        try:
            finder = page.find_tables()
            tabelas = finder.tables
        except Exception:
            tabelas = []
        for tab in tabelas:
            data = tab.extract()
            if not data or len(data) < 3:
                continue
            # localiza linha de cabeçalho AMBIENTE
            hi = next((i for i, r in enumerate(data[:5]) if any("ambiente" == normalizar(c) for c in r if c)), None)
            if hi is None:
                continue
            hdr = data[hi]
            if len(hdr) < 3:
                continue
            # normaliza cabeçalhos; a coluna 1 normalmente é o ambiente (col 0 carrega o escopo)
            item_cols = {}
            for ci, h in enumerate(hdr):
                if ci < 2 or not h:
                    continue
                can = _header_canon(h)
                if can:
                    item_cols[ci] = can
            if not item_cols:
                continue
            escopo = None
            for row in data[hi+1:]:
                if len(row) < 2:
                    continue
                esc_cell = normalizar(row[0]) if row[0] else ""
                if "area privativa" in esc_cell:
                    escopo = "Área Privativa"
                elif "area comum" in esc_cell or "area uso comum" in esc_cell:
                    escopo = "Área Comum"
                if escopo not in ["Área Privativa", "Área Comum"]:
                    continue
                amb_doc = re.sub(r"\s+", " ", str(row[1] or "")).strip()
                if not amb_doc:
                    continue
                regras_esc = base[base["escopo"] == escopo]
                if escopo == "Área Privativa":
                    regras_esc = regras_esc[regras_esc["padrao"] == padrao]
                amb_base, sc = parear_ambiente(amb_doc, regras_esc, escopo)
                if not amb_base:
                    continue
                # bbox de evidência: ambiente na página; suficiente para recorte visual do card
                rects = page.search_for(amb_doc.replace("\n", " "))
                bb = _bbox_union(rects[:1]) if rects else None
                for ci, can in item_cols.items():
                    if ci >= len(row) or not row[ci]:
                        continue
                    cell = re.sub(r"\s+", " ", str(row[ci])).strip()
                    if not cell:
                        continue
                    itens = _split_piso_rodape_soleira(cell) if can == "piso_combo" else [(can, cell)]
                    for item_can, valor in itens:
                        registros.append({
                            "escopo": escopo, "ambiente_doc": amb_doc, "ambiente_base": amb_base,
                            "score_ambiente": sc, "item_canon": item_can, "rotulo": item_can,
                            "valor": valor, "page": pno, "bbox": bb,
                        })
    return registros

# ==============================================================================
# Comparação técnica
# ==============================================================================
MATERIAIS = [
    "ceramica", "porcelanato", "granito", "marmore sintetico", "marmore", "aco inox", "louca",
    "concreto desempenado", "cimentado", "vinilico", "laminado", "ardosia", "gesso", "textura acrilica",
    "monocapa", "aluminio", "ferro", "vidro", "madeira", "intertravado", "pedra natural", "caiação", "caiacao",
]


def materiais(texto):
    n = normalizar(texto)
    # inox é sinônimo de aço inox para comparação
    if "inox" in n and "aco inox" not in n:
        n += " aco inox"
    return {m for m in MATERIAIS if m in n}


def caracteristicas_revestimento(texto):
    n = normalizar(texto)
    feats = set()
    if re.search(r"piso\s+ao\s+teto|do\s+piso\s+ao\s+teto", n): feats.add("piso_teto")
    if re.search(r"1[,\.]?50\s*m|1\s*,\s*50", n): feats.add("altura_150")
    if "acima da bancada" in n: feats.add("acima_bancada")
    if re.search(r"2\s*(?:ou|a)\s*3\s*fiadas|duas\s*(?:ou|a)\s*tres\s*fiadas", n): feats.add("fiadas_2_3")
    if "todas as paredes" in n: feats.add("todas_paredes")
    if "parede hidraulica" in n: feats.add("parede_hidraulica")
    if "box" in n and "parede" in n: feats.add("zona_box")
    if "lateral do shaft" in n or "parede do lavatorio" in n or "parede da bacia" in n: feats.add("zonas_especificas")
    return feats


def conflito_geometria_revestimento(found, expected):
    f, e = caracteristicas_revestimento(found), caracteristicas_revestimento(expected)
    if not f or not e:
        return None
    # Piso-teto é diferente de meia altura / fiadas.
    meia = {"altura_150", "acima_bancada", "fiadas_2_3"}
    if ("piso_teto" in e and f & meia) or ("piso_teto" in f and e & meia):
        return "Extensão do revestimento divergente (piso ao teto x revestimento parcial/acima da bancada)."
    # Todas as paredes é mais abrangente que parede hidráulica / zonas específicas.
    restr = {"parede_hidraulica", "zona_box", "zonas_especificas"}
    if ("todas_paredes" in e and f & restr) or ("todas_paredes" in f and e & restr):
        return "Área de aplicação do revestimento divergente (todas as paredes x paredes/zonas específicas)."
    return None


def similaridade(a, b):
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return 0.0
    ta, tb = tokens(na), tokens(nb)
    jac = len(ta & tb) / max(1, len(ta | tb))
    cov = len(ta & tb) / max(1, min(len(ta), len(tb)))
    seq = SequenceMatcher(None, na[:1400], nb[:1400]).ratio()
    return 0.40 * jac + 0.35 * cov + 0.25 * seq


def atributos_tecnicos(texto, item_canon=""):
    """Extrai requisitos objetivos; complementos não conflitantes não viram erro."""
    n = normalizar(texto)
    a = {
        "materiais": materiais(n),
        "revestimento": caracteristicas_revestimento(n),
        "quantidades": set(re.findall(r"\b\d+\b", n)),
        "condicional": any(x in n for x in ["quando houver", "quando aplicavel", "conforme projeto", "caso haja"]),
    }
    if "cuba" in n: a["cuba"] = True
    if "coluna" in n: a["coluna"] = True
    if "embutir" in n: a["embutir"] = True
    if "ventilacao" in n or re.search(r"\bvp\b", n): a["ventilacao"] = True
    return a


def materiais_nucleo(item_canon, mats):
    """Remove materiais acessórios quando o item principal já está identificado."""
    mats=set(mats)
    if item_canon == "bancada":
        # louça/vidro/inox podem descrever cuba/acessório; o material da bancada é prioritário.
        principais={m for m in mats if m in {"granito","marmore sintetico","marmore","aco inox","concreto desempenado","cimentado"}}
        return principais or mats
    if item_canon == "tanque":
        principais={m for m in mats if m in {"marmore sintetico","marmore","louca","aco inox"}}
        return principais or mats
    return mats


def _sistemas_superficie(texto):
    n=normalizar(texto); out=set()
    if "gesso liso" in n: out.add("gesso_liso")
    if "bloco aparente" in n or "blocos aparentes" in n: out.add("bloco_aparente")
    if "caiacao" in n: out.add("caiacao")
    if "textura acrilica" in n: out.add("textura_acrilica")
    if "monocapa" in n: out.add("monocapa")
    if "ceramica" in n: out.add("ceramica")
    if "pintura" in n: out.add("pintura")
    if "concreto aparente" in n: out.add("concreto_aparente")
    return out


def _segmento_principal_regra(expected, found=""):
    """Usa o requisito padrão antes das exceções marcadas com *; só abre exceção se o trecho a invocar."""
    txt=str(expected or "")
    nf=normalizar(found)
    # Se o próprio trecho identifica PCD/PNE/condição especial, mantém a regra completa.
    if any(x in nf for x in ["pcd", "pne", "32m", "32 m", "quando aplicavel", "quando houver"]):
        return txt
    # notas condicionais no R96 normalmente começam por asterisco; não devem tornar
    # automaticamente compatível uma solução que contradiz o requisito principal.
    principal=re.split(r"\n?\s*\*+", txt, maxsplit=1)[0].strip()
    return principal or txt

def requisitos_compativeis(found, expected, item_canon):
    expected_cmp = _segmento_principal_regra(expected, found)
    af, ae = atributos_tecnicos(found, item_canon), atributos_tecnicos(expected_cmp, item_canon)
    mf = materiais_nucleo(item_canon, af["materiais"]); me = materiais_nucleo(item_canon, ae["materiais"])
    if item_canon == "parede":
        cg = conflito_geometria_revestimento(found, expected_cmp)
        if cg: return False, cg
        sf,se=_sistemas_superficie(found),_sistemas_superficie(expected_cmp)
        fortes={"gesso_liso","bloco_aparente","caiacao","textura_acrilica","monocapa","ceramica","concreto_aparente"}
        ff,ee=sf&fortes,se&fortes
        if ff and ee and ff.isdisjoint(ee):
            return False, f"Sistema/acabamento de parede divergente: memorial indica {', '.join(sorted(ff))}; R96 prevê {', '.join(sorted(ee))}."
        if ff and ee and ff & ee:
            return True, "Sistema principal de parede compatível; diferenças complementares de redação não alteram a conformidade."
    if mf and me and mf.isdisjoint(me):
        return False, f"Material/solução divergente: memorial indica {', '.join(sorted(mf))}; R96 prevê {', '.join(sorted(me))}."
    if mf and me and (mf & me):
        return True, "Requisito técnico essencial compatível; diferenças de redação/complementos não alteram a conformidade."
    # Condicionais: se o trecho encontrado satisfaz uma alternativa explícita do esperado, considera compatível
    if any(x in normalizar(expected_cmp) for x in ["quando aplicavel", "quando houver", "se for coberto", "ou,"]):
        sf,se=_sistemas_superficie(found),_sistemas_superficie(expected_cmp)
        if sf and se and sf & se:
            return True, "O memorial atende a uma das condições/alternativas previstas no R96."
    return None, ""

def avaliar(found, expected, item_canon):
    nf, ne = normalizar(found), normalizar(expected)
    if not nf:
        return STATUS_INFO, "Trecho não localizado com segurança.", 0.0

    mats_ne = materiais(ne)
    esperado_na_puro = ("nao aplicavel" in ne and not mats_ne and len(ne) <= 180)
    if esperado_na_puro:
        equivalentes_na = ["nao aplicavel", "sem rodape", "sem soleira", "sem baguete", "sem tento", "sem peitoril", "sem janela", "sem janelas"]
        if any(x in nf for x in equivalentes_na) or nf.strip() in {"0", "0.", "zero"}:
            return STATUS_OK, "Memorial e R96 indicam ausência/não aplicabilidade equivalente.", 1.0
        return STATUS_INFO, "R96 indica não aplicável/condicional; o trecho requer confirmação de aplicabilidade.", 0.35

    # Primeiro compara requisitos técnicos; não exige redação idêntica.
    comp, motivo = requisitos_compativeis(found, expected, item_canon)
    if comp is False:
        return STATUS_ERRO, motivo, 0.96
    if comp is True:
        return STATUS_OK, motivo, 0.92

    sim = similaridade(nf, ne)
    if ne in nf or nf in ne:
        return STATUS_OK, "Descrição compatível com o R96.", 1.0

    # Núcleo textual é fallback apenas dentro do mesmo Ambiente+Item já validado.
    ff = re.split(r"[.;]", nf)[0].strip(); ee = re.split(r"[.;]", ne)[0].strip()
    tf, te = tokens(ff), tokens(ee)
    nucleo = len(tf & te) / max(1, min(len(tf), len(te))) if tf and te else 0.0
    if nucleo >= 0.68:
        return STATUS_OK, "Núcleo técnico da descrição compatível com o R96.", max(sim, nucleo)
    if sim >= 0.52:
        return STATUS_OK, "Descrição tecnicamente compatível com o R96.", sim
    return STATUS_INFO, "Item identificado, mas a equivalência técnica não pôde ser concluída automaticamente.", sim


def regras_para(base, padrao, escopo):
    if escopo == "Área Comum":
        return base[base["escopo"] == "Área Comum"].copy()
    return base[(base["escopo"] == "Área Privativa") & (base["padrao"] == padrao)].copy()



def padrao_para_registro(registro, padrao_predominante, excecoes):
    """Roteia apenas itens privativos quando o próprio trecho identifica uma exceção.
    Se não houver evidência do grupo, mantém o padrão predominante.
    """
    if registro.get("escopo") != "Área Privativa" or not excecoes:
        return padrao_predominante
    txt = normalizar(f"{registro.get('ambiente_doc','')} {registro.get('valor','')}")
    melhor = (0.0, padrao_predominante)
    for ex in excecoes:
        aplic = normalizar(ex.get("aplicacao", ""))
        if not aplic:
            continue
        score = 0.0
        # Torres explicitadas
        for torre in re.findall(r"torre\s+([a-z0-9]+)", aplic):
            if re.search(rf"torre\s+{re.escape(torre)}\b", txt):
                score += 2.0
        # Finais / unidades explicitados
        nums = re.findall(r"\b\d{1,3}\b", aplic)
        if nums:
            acertos = 0
            for num in nums:
                try:
                    if re.search(rf"\b0*{int(num)}\b", txt):
                        acertos += 1
                except Exception:
                    pass
            score += 0.45 * acertos
        # Tipologias / termos textuais relevantes
        for termo in ["pcd", "pne", "studio", "apa"]:
            if termo in aplic and termo in txt:
                score += 1.0
        if score > melhor[0]:
            melhor = (score, ex.get("padrao", padrao_predominante))
    return melhor[1] if melhor[0] >= 1.0 else padrao_predominante

def nome_item_exibicao(registro, regra=None):
    can = registro.get("item_canon", "")
    mapa = {
        "guarda_corpo": "GUARDA-CORPO", "porta": "PORTAS", "janela": "JANELAS",
        "esquadria_especial": "ESQUADRIAS ESPECIAIS", "tanque": "TANQUE",
        "bancada": "BANCADA", "louca": "LOUÇAS", "parede": "PAREDE/ SANCAS",
        "teto": "TETO", "piso": "PISO", "metais": "METAIS"
    }
    return mapa.get(can, str(regra.get("item")) if regra is not None else registro.get("rotulo", can))

def auditar_registros(registros, base, padrao, grupo="Regra geral", excecoes=None):
    out = []
    vistos = set()
    for r in registros:
        esc = r["escopo"]
        padrao_item = padrao_para_registro(r, padrao, excecoes)
        regras = regras_para(base, padrao_item, esc)
        amb = r["ambiente_base"]
        regra = parear_regra_item(regras[regras["ambiente"] == amb], r["item_canon"], r.get("valor", ""))
        chave = (esc, amb, r["item_canon"], r["page"], normalizar(r["valor"])[:160])
        if chave in vistos:
            continue
        vistos.add(chave)
        if regra is None:
            out.append({
                "Grupo / Aplicação": grupo, "Padrão aplicado": padrao_item, "Área": esc,
                "Ambiente": amb, "Seção": "MAPEAMENTO", "Item": r["rotulo"],
                "Texto encontrado": r["valor"], "Especificação prevista": "Sem item equivalente no R96 para este ambiente",
                "Status": STATUS_INFO, "Orientação / resposta prevista": "Sem ação automática.",
                "Observação": "O item foi localizado no memorial, mas não existe correspondência segura no R96.",
                "Confiança": round(r["score_ambiente"], 2), "Fonte": "Padrão de Acabamentos R96",
                "Página": r["page"], "BBox": r["bbox"],
            })
            continue
        stt, obs, conf = avaliar(r["valor"], regra["especificacao"], r["item_canon"])
        orient = "Nenhuma ação necessária." if stt == STATUS_OK else (regra["especificacao"] if stt == STATUS_ERRO else "Comparação inconclusiva; revisar apenas se necessário.")
        out.append({
            "Grupo / Aplicação": grupo, "Padrão aplicado": padrao_item, "Área": esc,
            "Ambiente": regra["ambiente"], "Seção": regra["secao"], "Item": nome_item_exibicao(r, regra),
            "Texto encontrado": r["valor"], "Especificação prevista": regra["especificacao"],
            "Status": stt, "Orientação / resposta prevista": orient, "Observação": obs,
            "Confiança": round(min(r["score_ambiente"], max(conf, 0.01)), 2), "Fonte": regra["fonte"],
            "Página": r["page"], "BBox": r["bbox"],
        })
    return out

# ==============================================================================
# Protocolo / Especificações Gerais
# ==============================================================================
def localizar_gatilho_pdf(file_bytes, gatilhos):
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for pno, page in enumerate(doc, start=1):
            txt = normalizar(page.get_text("text"))
            for g in gatilhos:
                ng = normalizar(g)
                if ng and ng in txt:
                    rects = page.search_for(g)
                    return pno, (_bbox_union(rects[:1]) if rects else None), g
    except Exception:
        pass
    return None, None, ""


def auditoria_gerais_cliente(file_bytes, padrao):
    regras = [
        ("Estrutura e Vedações", ["estrutura e vedacoes", "estrutura e vedações"], "Confirmar no projeto estrutural o sistema construtivo adotado."),
        ("Antena coletiva / TV por assinatura", ["antena coletiva", "tv por assinatura"], "Confirmar em projeto as quantidades de pontos de TV a serem previstas."),
        ("Sistema de Telefonia", ["sistema de telefonia", "telefonia"], "Confirmar em projeto as quantidades de pontos de telefonia a serem previstas."),
        ("Elevadores", ["elevadores"], "Confirmar em projeto as quantidades de elevadores e de 'transfer' previstas."),
        ("Pressurização", ["pressurizacao", "pressurização"], "Confirmar se o projeto possui escada pressurizada ou ventilada."),
        ("Instalações Hidráulicas", ["instalacoes hidraulicas", "instalações hidráulicas"], "Confirmar diferenças entre grupos de unidades quanto a água quente, aquecimento a gás e/ou chuveiro elétrico."),
    ]
    out = []
    for item, gat, orient in regras:
        pg, bb, achado = localizar_gatilho_pdf(file_bytes, gat)
        if not pg:
            continue
        out.append({
            "Grupo / Aplicação":"Geral", "Padrão aplicado":padrao, "Área":"Especificações Gerais", "Ambiente":"Geral",
            "Seção":"ESPECIFICAÇÕES GERAIS", "Item":item,
            "Texto encontrado": f"Trecho localizado: {achado}", "Especificação prevista":orient,
            "Status":STATUS_ATENCAO, "Orientação / resposta prevista":orient,
            "Observação":"Item dependente de confirmação do coordenador/projeto.", "Confiança":1.0,
            "Fonte":"Protocolo Memorial do Cliente", "Página":pg, "BBox":bb,
        })
    return out


def auditoria_cef_protocolo(file_bytes, padrao):
    out = []
    for item, gat, orient, fonte in PROTOCOLO_CEF:
        pg, bb, achado = localizar_gatilho_pdf(file_bytes, gat)
        if not pg:
            continue
        out.append({
            "Grupo / Aplicação":"Geral / CEF", "Padrão aplicado":padrao, "Área":"Protocolo Financiador", "Ambiente":"Geral",
            "Seção":"ORIENTAÇÕES DA COORDENAÇÃO", "Item":item,
            "Texto encontrado":f"Trecho localizado: {achado}", "Especificação prevista":orient,
            "Status":STATUS_ATENCAO, "Orientação / resposta prevista":orient,
            "Observação":"Alerta adicional do protocolo CEF; não substitui a conferência do R96.", "Confiança":1.0,
            "Fonte":fonte, "Página":pg, "BBox":bb,
        })
    return out

# ==============================================================================
# Recorte / PDF revisado
# ==============================================================================
def recorte_ocorrencia_pdf(file_bytes, pagina, bbox=None, zoom=1.7):
    pg = pagina_int(pagina)
    if not pg:
        return None
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        if pg > len(doc): return None
        page = doc[pg-1]
        if bbox and isinstance(bbox, (list, tuple)) and len(bbox) == 4:
            r = fitz.Rect(*bbox)
            clip = fitz.Rect(max(0, r.x0-80), max(0, r.y0-95), min(page.rect.width, r.x1+420), min(page.rect.height, r.y1+160))
        else:
            clip = page.rect
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
        return pix.tobytes("png")
    except Exception:
        return None


def _texto_caixa(row):
    trecho = str(row.get("Texto encontrado", "")).replace("\n", " ")[:145]
    if row["Status"] == STATUS_ERRO:
        prev = str(row.get("Especificação prevista", "")).replace("\n", " ")[:175]
        return f"DIVERGÊNCIA · {row['Ambiente']} / {row['Item']}\nTrecho: {trecho}\nCorrigir: {prev}"
    return f"ATENÇÃO · {row['Item']}\nTrecho: {trecho}\nConfirmar: {str(row.get('Orientação / resposta prevista',''))[:185]}"


def gerar_pdf_anotado(file_bytes, df):
    src = fitz.open(stream=file_bytes, filetype="pdf")
    out = fitz.open()
    painel = 250
    por_pg = {}
    for _, row in df[df["Status"].isin([STATUS_ERRO, STATUS_ATENCAO])].iterrows():
        pg = pagina_int(row.get("Página"))
        if pg:
            por_pg.setdefault(pg, []).append(row)
    for pno, sp in enumerate(src, start=1):
        op = out.new_page(width=sp.rect.width + painel, height=sp.rect.height)
        op.show_pdf_page(fitz.Rect(0,0,sp.rect.width,sp.rect.height), src, pno-1)
        op.draw_line((sp.rect.width,0),(sp.rect.width,sp.rect.height), color=(0.84,0.84,0.84), width=0.6)
        if pno not in por_pg:
            continue
        op.insert_text((sp.rect.width+12,18), "MIA · REVISÃO", fontsize=7.5, fontname="helv", color=(0.30,0.30,0.30))
        y = 28
        for row in por_pg[pno]:
            texto = _texto_caixa(row)
            iserr = row["Status"] == STATUS_ERRO
            fill = (1.0,0.94,0.94) if iserr else (1.0,0.98,0.86)
            stroke = (0.80,0.12,0.12) if iserr else (0.62,0.43,0.0)
            # altura dinâmica e compacta
            chars = max(1, len(texto))
            h = min(118, max(54, 42 + 9 * (chars // 70)))
            if y + h > sp.rect.height - 10:
                break
            box = fitz.Rect(sp.rect.width+10, y, sp.rect.width+painel-10, y+h)
            op.draw_rect(box, color=stroke, fill=fill, width=0.7)
            op.insert_textbox(box + (7,6,-7,-6), texto, fontsize=6.6, fontname="helv", color=(0.08,0.08,0.08), lineheight=1.08)
            bb = row.get("BBox")
            if isinstance(bb,(tuple,list)) and len(bb)==4:
                try:
                    rr = fitz.Rect(*bb)
                    yy = min(max(rr.y0 + rr.height/2, 8), sp.rect.height-8)
                    op.draw_line((rr.x1+3,yy),(sp.rect.width+10,y+12), color=stroke, width=0.5)
                except Exception:
                    pass
            y += h + 6
    return out.tobytes(garbage=3, deflate=True)


def gerar_docx_anotado(file_bytes, df):
    doc = docx.Document(io.BytesIO(file_bytes))
    p = doc.add_paragraph(); p.add_run("--- MIA | AUDITORIA ---").bold = True
    for _, row in df[df["Status"].isin([STATUS_ERRO, STATUS_ATENCAO])].head(100).iterrows():
        p = doc.add_paragraph(); r = p.add_run(_texto_caixa(row))
        r.font.highlight_color = WD_COLOR_INDEX.RED if row["Status"] == STATUS_ERRO else WD_COLOR_INDEX.YELLOW
    out = io.BytesIO(); doc.save(out); return out.getvalue()

# ==============================================================================
# Empreendimento com exceções de padrão
# ==============================================================================
def configurar_excecoes_sidebar():
    st.sidebar.markdown("### Padrões diferentes (opcional)")
    possui = st.sidebar.checkbox("O empreendimento possui unidades com padrão diferente?", value=False)
    if not possui:
        return []
    qtd = st.sidebar.number_input("Quantidade de exceções", min_value=1, max_value=12, value=1, step=1)
    out = []
    for i in range(int(qtd)):
        with st.sidebar.expander(f"Exceção {i+1}", expanded=True):
            aplic = st.text_input("Aplicação", placeholder="Ex.: Torre C - finais 01, 02, 05 e 06", key=f"aplic_{i}")
            pad = st.selectbox("Padrão técnico", PADROES_TECNICOS, index=2, key=f"pad_ex_{i}")
            if aplic.strip(): out.append({"aplicacao":aplic.strip(), "padrao":pad})
    return out

# ==============================================================================
# Main
# ==============================================================================
def main():
    st.set_page_config(page_title="MIA | Memoriais", page_icon="M", layout="wide")
    st.title("MIA")
    st.caption(f"Coordenação e Qualidade de Projetos | Memoriais · Motor {BUILD_VERSION}")

    with st.sidebar:
        st.header("Configuração da análise")
        tipo_doc = st.selectbox("Tipo de memorial", ["Memorial do Cliente (Comercial / Vendas)", "Memorial CEF / Financiador"])
        padrao = st.selectbox("Padrão predominante do empreendimento", PADROES_TECNICOS, index=1)
    excecoes = configurar_excecoes_sidebar()
    with st.sidebar:
        incluir_gerais = st.checkbox("Incluir checklist de Especificações Gerais", value=True) if tipo_doc.startswith("Memorial do Cliente") else False
        st.markdown("---")
        st.caption("Base técnica R96 carregada automaticamente.")
        memorial = st.file_uploader("Memorial para análise (PDF ou DOCX)", type=["pdf","docx"])

    a,b,c = st.columns(3)
    a.metric("Tipo", "Cliente" if tipo_doc.startswith("Memorial do Cliente") else "Financiador / CEF")
    b.metric("Padrão predominante", padrao)
    c.metric("Exceções", len(excecoes))

    if memorial is None:
        st.info("Envie um memorial para iniciar a conferência.")
        return

    if st.button("Executar conferência", type="primary", use_container_width=True):
        with st.spinner("Lendo estrutura do memorial e cruzando Área > Ambiente > Item com o R96..."):
            try:
                base = carregar_base_r96()
                file_bytes = memorial.getvalue()
                ext = memorial.name.rsplit(".",1)[-1].lower()
                if ext == "pdf":
                    if tipo_doc.startswith("Memorial do Cliente"):
                        regs = extrair_cliente_pdf(file_bytes, base, padrao)
                    else:
                        regs = extrair_cef_pdf(file_bytes, base, padrao)
                    resultados = auditar_registros(regs, base, padrao, excecoes=excecoes)
                    if tipo_doc.startswith("Memorial do Cliente") and incluir_gerais:
                        resultados += auditoria_gerais_cliente(file_bytes, padrao)
                    if tipo_doc == "Memorial CEF / Financiador":
                        resultados += auditoria_cef_protocolo(file_bytes, padrao)
                else:
                    # DOCX: mantém suporte básico; PDF é o fluxo com evidência visual precisa.
                    doc = docx.Document(io.BytesIO(file_bytes))
                    texto = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
                    st.warning("Para o pareamento visual completo e recortes, prefira PDF. DOCX está em modo de compatibilidade básica.")
                    resultados = []
                cols = ["Grupo / Aplicação","Padrão aplicado","Área","Ambiente","Seção","Item","Texto encontrado","Especificação prevista","Status","Orientação / resposta prevista","Observação","Confiança","Fonte","Página","BBox"]
                df = pd.DataFrame(resultados, columns=cols) if resultados else pd.DataFrame(columns=cols)
                st.session_state.update(resultado_memorial=df, memorial_bytes=file_bytes, memorial_ext=ext, memorial_nome=memorial.name)
            except Exception as e:
                st.exception(e)
                return

    df = st.session_state.get("resultado_memorial")
    if df is None:
        return

    st.markdown("---")
    st.subheader("Resultado da conferência")
    nerr=int((df["Status"]==STATUS_ERRO).sum()); natt=int((df["Status"]==STATUS_ATENCAO).sum()); ninfo=int((df["Status"]==STATUS_INFO).sum()); nok=int((df["Status"]==STATUS_OK).sum())
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Divergências",nerr); c2.metric("Atenções",natt); c3.metric("Não verificados",ninfo); c4.metric("Conformes",nok)
    st.caption(f"{BUILD_VERSION}: {len(df)} itens localizados no próprio memorial. A análise preserva Área > Ambiente > Item > Subitem e compara requisitos técnicos, não redações idênticas.")

    modo=st.radio("Exibir",["Itens que exigem ação","Não verificados","Conformes","Todos"],horizontal=True)
    if modo=="Itens que exigem ação": vis=df[df["Status"].isin([STATUS_ERRO,STATUS_ATENCAO])]
    elif modo=="Não verificados": vis=df[df["Status"]==STATUS_INFO]
    elif modo=="Conformes": vis=df[df["Status"]==STATUS_OK]
    else: vis=df

    if vis.empty:
        st.success("Nenhum item nesta categoria.")
    else:
        for _, row in vis.head(180).iterrows():
            titulo=f"{row['Status']}  {row['Ambiente']} - {row['Item']}"
            with st.expander(titulo, expanded=row["Status"] in [STATUS_ERRO,STATUS_ATENCAO]):
                st.markdown(f"**Área:** {row['Área']}  |  **Padrão:** {row['Padrão aplicado']}")
                pg = pagina_int(row.get("Página"))
                if st.session_state.get("memorial_ext") == "pdf" and pg:
                    img = recorte_ocorrencia_pdf(st.session_state["memorial_bytes"], pg, row.get("BBox"))
                    if img:
                        st.caption(f"Trecho do memorial · página {pg}")
                        st.image(img, width=760)
                st.markdown(f"**Encontrado:** {row['Texto encontrado']}")
                st.markdown(f"**Previsto:** {row['Especificação prevista']}")
                if row["Status"] == STATUS_ERRO:
                    st.markdown(f"**O que corrigir:** {row['Orientação / resposta prevista']}")
                elif row["Status"] == STATUS_ATENCAO:
                    st.markdown(f"**O que confirmar:** {row['Orientação / resposta prevista']}")
                elif row["Status"] == STATUS_OK:
                    st.markdown("**Resultado:** descrição compatível com a base.")
                st.caption(row["Fonte"])

    st.markdown("### Exportações")
    xbio=io.BytesIO()
    with pd.ExcelWriter(xbio, engine="openpyxl") as wr:
        df.to_excel(wr,index=False,sheet_name="Auditoria")
    e1,e2=st.columns(2)
    e1.download_button("Baixar relatório Excel",xbio.getvalue(),file_name=f"MIA_Auditoria_{Path(st.session_state['memorial_nome']).stem}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    if st.session_state.get("memorial_ext") == "pdf":
        pdf=gerar_pdf_anotado(st.session_state["memorial_bytes"],df)
        e2.download_button("Baixar PDF revisado com caixas de texto",pdf,file_name=f"MIA_Revisado_{st.session_state['memorial_nome']}",mime="application/pdf",use_container_width=True)
    else:
        docx_out=gerar_docx_anotado(st.session_state["memorial_bytes"],df)
        e2.download_button("Baixar DOCX revisado",docx_out,file_name=f"MIA_Revisado_{st.session_state['memorial_nome']}",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)


if __name__ == "__main__":
    main()
