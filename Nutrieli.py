import os
import io
import re
import sys
import json
import hmac
import math
import inspect
import secrets
import hashlib
from datetime import datetime, date, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Nutrieli", page_icon="🌿", layout="wide")

try:
    from zoneinfo import ZoneInfo
    _FUSO = ZoneInfo("America/Sao_Paulo")
except Exception:
    _FUSO = None

FMT_DH = "%d/%m/%Y %H:%M:%S"
FMT_D = "%d/%m/%Y"


def agora() -> datetime:
    """Data/hora de Brasília (o servidor do Streamlit Cloud roda em UTC)."""
    if _FUSO is not None:
        return datetime.now(_FUSO).replace(tzinfo=None)
    return datetime.now()


# Compatibilidade entre versões do Streamlit (use_container_width x width="stretch")
_ST_NOVO = "width" in inspect.signature(st.button).parameters


def larg() -> dict:
    return {"width": "stretch"} if _ST_NOVO else {"use_container_width": True}


def tabela(df: pd.DataFrame):
    try:
        st.dataframe(df, width="stretch", hide_index=True)
    except Exception:
        st.dataframe(df, use_container_width=True, hide_index=True)


# =====================================================================
# ARQUIVOS
# =====================================================================
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARQ_PACIENTES = os.path.join(BASE_DIR, "pacientes_nutrieli.xlsx")
ARQ_AVALIACOES = os.path.join(BASE_DIR, "avaliacoes_nutrieli.xlsx")
ARQ_USUARIOS = os.path.join(BASE_DIR, "usuarios_nutrieli.json")
ARQ_CATALOGO = os.path.join(BASE_DIR, "catalogo_enterais_nutrieli.json")

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_JSON = "application/json"

# =====================================================================
# ESTILO
# =====================================================================
CSS_GLOBAL = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1300px; }
h1, h2, h3 { color: #0F766E; letter-spacing: -0.01em; }
h2 { font-size: 1.6rem !important; }
footer { visibility: hidden; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #F0FDFA 0%, #E6F4FA 100%);
    border-right: 1px solid #CCFBF1;
}
[data-testid="stSidebar"] [role="radiogroup"] label {
    padding: 6px 10px; border-radius: 8px; width: 100%;
}
[data-testid="stSidebar"] [role="radiogroup"] label:hover { background: rgba(13,148,136,.10); }

[data-testid="stMetric"] {
    background: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #0D9488;
    border-radius: 10px; padding: 12px 16px;
}
[data-testid="stMetricLabel"] p { font-size: .78rem; color: #475569; font-weight: 600; }
[data-testid="stMetricValue"] { font-weight: 700; color: #0F172A; }

.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
    border-radius: 10px; font-weight: 600;
}
button[kind="primary"], button[kind="primaryFormSubmit"] {
    background: linear-gradient(135deg, #0D9488 0%, #0369A1 100%); border: none; color: #fff;
}
.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] { padding: 8px 16px; font-weight: 600; border-radius: 8px 8px 0 0; }

.nt-topo { display:flex; align-items:center; justify-content:space-between; gap:14px;
           padding: 4px 0 14px 0; border-bottom: 1px solid #E2E8F0; margin-bottom: 18px; }
.nt-topo h1 { margin:0; padding:0; font-size:1.9rem; line-height:1.1; }
.nt-topo p { margin:2px 0 0 0; font-size:.88rem; color:#475569; }
.nt-chip { display:inline-block; background:#CCFBF1; color:#0F766E; font-size:.78rem; font-weight:600;
           padding:4px 12px; border-radius:999px; }
.nt-ficha { background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:10px 14px;
            font-size:.88rem; color:#334155; margin: 6px 0 4px 0; }
.nt-ficha b { color:#0F172A; }
.login-header-card { background: linear-gradient(135deg, #0F766E 0%, #0369A1 100%);
    padding: 26px 24px 22px 24px; border-radius: 16px; text-align: center; color: white;
    margin-bottom: 18px; box-shadow: 0 8px 24px rgba(13,148,136,.18); }
.login-title { font-size: 2rem; font-weight: 700; margin: 8px 0 2px 0; color: #FFFFFF; }
.login-subtitle { font-size: .92rem; color: #CCFBF1; margin: 0; font-weight: 500; }
.login-footer { text-align:center; color:#64748B; font-size:.8rem; margin-top:14px; }
</style>
"""
st.markdown(CSS_GLOBAL, unsafe_allow_html=True)


def logo_svg(size: int = 42, uid: str = "a") -> str:
    return f"""
<svg width="{size}" height="{size}" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="gc{uid}" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0D9488"/><stop offset="100%" stop-color="#0369A1"/>
    </linearGradient>
    <linearGradient id="gf{uid}" x1="0%" y1="100%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#10B981"/><stop offset="100%" stop-color="#34D399"/>
    </linearGradient>
  </defs>
  <circle cx="50" cy="50" r="46" fill="url(#gc{uid})" stroke="#CCFBF1" stroke-width="3"/>
  <rect x="43" y="20" width="14" height="60" rx="4" fill="#FFFFFF" fill-opacity="0.18"/>
  <rect x="20" y="43" width="60" height="14" rx="4" fill="#FFFFFF" fill-opacity="0.18"/>
  <path d="M50 74C50 74 28 62 28 41C28 28 40 23 50 32C60 23 72 28 72 41C72 62 50 74 50 74Z" fill="url(#gf{uid})"/>
  <path d="M24 52H37L42 41L49 61L55 46L59 52H76" stroke="#FFFFFF" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
</svg>"""


# =====================================================================
# USUÁRIOS E AUTENTICAÇÃO
# =====================================================================
PERFIS = {
    "master": "Master (administrador)",
    "submaster": "Submaster (nutricionista responsável)",
    "normal": "Normal (lançamento de dados)",
}


def gerar_hash(senha: str, salt: str = None) -> str:
    salt = salt or secrets.token_hex(8)
    h = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), salt.encode("utf-8"), 120_000).hex()
    return f"{salt}${h}"


def conferir_senha(senha: str, armazenado: str) -> bool:
    try:
        salt, h = armazenado.split("$", 1)
    except ValueError:
        return False
    return hmac.compare_digest(gerar_hash(senha, salt).split("$", 1)[1], h)


def salvar_usuarios(usuarios: dict):
    with open(ARQ_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=2)


def carregar_usuarios() -> dict:
    usuarios = {}
    existia = os.path.exists(ARQ_USUARIOS)
    if existia:
        try:
            with open(ARQ_USUARIOS, "r", encoding="utf-8") as f:
                usuarios = json.load(f)
        except Exception:
            usuarios = {}
    alterado = False
    if not existia:
        usuarios["nutrieli"] = {
            "nome": "Nutrição Clínica", "perfil": "submaster", "ativo": True,
            "senha": gerar_hash("nutri@2026"), "criado_em": agora().strftime(FMT_DH),
        }
        alterado = True
    if "master" not in usuarios:
        usuarios["master"] = {
            "nome": "Administrador", "perfil": "master", "ativo": True,
            "senha": gerar_hash("1234"), "criado_em": agora().strftime(FMT_DH),
        }
        alterado = True
    if alterado:
        salvar_usuarios(usuarios)
    return usuarios


if "usuario" not in st.session_state:
    st.session_state.usuario = None


def tela_login():
    st.markdown("<br>", unsafe_allow_html=True)
    _, centro, _ = st.columns([1, 1.35, 1])
    with centro:
        st.markdown(
            f"""
            <div class="login-header-card">
                <div>{logo_svg(76, "login")}</div>
                <div class="login-title">Nutrieli</div>
                <p class="login-subtitle">Nutrição Clínica Hospitalar, Pediátrica e Lactário</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("form_login"):
            st.markdown("##### Identificação profissional")
            usuario = st.text_input("Usuário", placeholder="Digite seu usuário")
            senha = st.text_input("Senha", type="password", placeholder="••••••••")
            entrar = st.form_submit_button("Entrar", type="primary", **larg())
            if entrar:
                u = carregar_usuarios().get(usuario.strip().lower())
                if u and u.get("ativo", True) and conferir_senha(senha, u["senha"]):
                    st.session_state.usuario = usuario.strip().lower()
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos. Confira os dados e tente novamente.")
        st.markdown(
            '<div class="login-footer">Protocolos baseados em OMS • DRI • ESPGHAN • ASPEN • SBC</div>',
            unsafe_allow_html=True,
        )


USUARIOS = carregar_usuarios()
if not st.session_state.usuario:
    tela_login()
    st.stop()

LOGIN = st.session_state.usuario
USER = USUARIOS.get(LOGIN)
if not USER or not USER.get("ativo", True):
    st.session_state.usuario = None
    st.rerun()

PERFIL = USER["perfil"]
NOME_USUARIO = USER["nome"]
E_MASTER = PERFIL == "master"
E_SUB = PERFIL == "submaster"

# =====================================================================
# ROTULOS, LISTAS E UTILITÁRIOS
# =====================================================================
ROTULOS = {
    "ID_Paciente": "ID", "Data_Cadastro": "Cadastrado em", "Tipo_Paciente": "Tipo",
    "Nome_Paciente": "Paciente", "Data_Nascimento": "Data de nascimento", "Nome_Mae": "Nome da mãe",
    "Prontuario": "Prontuário", "Clinica": "Clínica", "Enfermaria": "Enfermaria", "Leito": "Leito",
    "Situacao": "Situação", "Diagnostico_Clinico_Inicial": "Diagnóstico clínico inicial",
    "Comorbidades": "Comorbidades", "Patologia": "Patologia",
    "Diagnostico_Nutricional": "Diagnóstico nutricional",
    "Via_Terapia_Nutricional": "Via de terapia nutricional", "Cadastrado_Por": "Cadastrado por",
    "Ultima_Edicao_Por": "Última edição por", "Ultima_Edicao_Em": "Última edição em",
    "ID_Avaliacao": "ID da avaliação", "Data_Hora": "Data e hora", "Registrado_Por": "Registrado por",
    "Modulo_Calculo": "Protocolo", "Sexo": "Sexo", "Faixa_Etaria_Condicao": "Faixa etária / condição",
    "Peso_kg": "Peso (kg)", "Estatura_cm": "Estatura (cm)", "VET_kcal_dia": "VET (kcal/dia)",
    "Proteina_g_dia": "Proteína (g/dia)", "Carboidrato_g_dia": "Carboidrato (g/dia)",
    "Lipidio_g_dia": "Lipídio (g/dia)", "Nec_Hidrica_ml_dia": "Necessidade hídrica (ml/dia)",
    "Vol_Recebido_ml_dia": "Volume recebido (ml/dia)", "Taxa_Hidrica_pct": "Adequação hídrica (%)",
    "Fibras_g_dia": "Fibras (g/dia)", "Acucar_Max_g_dia": "Açúcar máx. (g/dia)",
    "TMB_Schofield_kcal": "TMB Schofield (kcal)", "TMB_OMS_kcal": "TMB OMS (kcal)",
    "Formula_Ou_Item": "Fórmula / item", "Qtd_Po_g_Ou_Vol_ml": "Quantidade",
    "Detalhes_Protocolo": "Detalhes do protocolo", "Observacoes_Clinicas": "Observações da avaliação",
}

SITUACOES = ["Internado", "Alta", "Transferência", "Óbito"]
DIAG_NUTRICIONAL = [
    "A definir", "Eutrofia", "Risco nutricional", "Desnutrição leve", "Desnutrição moderada",
    "Desnutrição grave", "Sobrepeso", "Obesidade", "Obesidade grave", "Outro (descrever nas observações)",
]
VIAS_TN = [
    "A definir", "Via oral", "Sonda nasogástrica (SNG)", "Sonda nasoenteral (SNE)",
    "Sonda orogástrica (SOG)", "Gastrostomia (GTT)", "Jejunostomia", "Nutrição parenteral",
    "Via oral + sonda", "Jejum",
]


def vazio(x) -> bool:
    return x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() in ("", "nan", "None")


def val(x, padrao="-"):
    return padrao if vazio(x) else x


def exibir(df: pd.DataFrame, colunas=None) -> pd.DataFrame:
    d = df.copy()
    if colunas:
        d = d[colunas]
    return d.fillna("").astype(str).rename(columns=ROTULOS)


def xlsx_bytes(df: pd.DataFrame, aba="Dados") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=aba)
    return buf.getvalue()


def parse_data(txt):
    try:
        return datetime.strptime(str(txt).strip()[:10], FMT_D).date()
    except Exception:
        return None


def idade_em_meses(dn: date) -> int:
    ref = agora().date()
    m = (ref.year - dn.year) * 12 + (ref.month - dn.month)
    if ref.day < dn.day:
        m -= 1
    return max(m, 0)


def texto_idade(dn) -> str:
    if not dn:
        return "-"
    m = idade_em_meses(dn)
    a, r = divmod(m, 12)
    if a >= 2:
        return f"{a} anos" + (f" e {r} m" if r else "")
    return f"{m} meses"


def localizacao(r) -> str:
    partes = []
    if not vazio(r.get("Clinica")):
        partes.append(f"Clínica {r['Clinica']}")
    if not vazio(r.get("Enfermaria")):
        partes.append(f"Enfermaria {r['Enfermaria']}")
    if not vazio(r.get("Leito")):
        partes.append(f"Leito {r['Leito']}")
    return " • ".join(partes) if partes else "-"


def rotulo_paciente(r) -> str:
    partes = [str(r["Nome_Paciente"]), f"DN {r['Data_Nascimento']}"]
    if not vazio(r.get("Nome_Mae")):
        partes.append(f"Mãe: {r['Nome_Mae']}")
    if not vazio(r.get("Prontuario")):
        partes.append(f"Pront. {r['Prontuario']}")
    return " | ".join(partes)


def idx(lista, valor, padrao=0):
    return lista.index(valor) if valor in lista else padrao


# =====================================================================
# BANCO DE DADOS EM PLANILHA
# =====================================================================
COLUNAS_PACIENTES = [
    "ID_Paciente", "Data_Cadastro", "Tipo_Paciente", "Nome_Paciente", "Data_Nascimento", "Nome_Mae",
    "Prontuario", "Clinica", "Enfermaria", "Leito", "Situacao", "Diagnostico_Clinico_Inicial",
    "Comorbidades", "Patologia", "Diagnostico_Nutricional", "Via_Terapia_Nutricional",
    "Cadastrado_Por", "Ultima_Edicao_Por", "Ultima_Edicao_Em",
]

COLUNAS_AVALIACOES = [
    "ID_Avaliacao", "Data_Hora", "ID_Paciente", "Registrado_Por", "Modulo_Calculo", "Sexo",
    "Faixa_Etaria_Condicao", "Peso_kg", "Estatura_cm", "VET_kcal_dia", "Proteina_g_dia",
    "Carboidrato_g_dia", "Lipidio_g_dia", "Nec_Hidrica_ml_dia", "Vol_Recebido_ml_dia",
    "Taxa_Hidrica_pct", "Fibras_g_dia", "Acucar_Max_g_dia", "TMB_Schofield_kcal", "TMB_OMS_kcal",
    "Formula_Ou_Item", "Qtd_Po_g_Ou_Vol_ml", "Detalhes_Protocolo", "Observacoes_Clinicas",
]


def salvar_xlsx(df: pd.DataFrame, caminho: str):
    df.to_excel(caminho, index=False)


def carregar_pacientes() -> pd.DataFrame:
    if os.path.exists(ARQ_PACIENTES):
        df = pd.read_excel(ARQ_PACIENTES, dtype=str)
        return df.reindex(columns=COLUNAS_PACIENTES).fillna("")
    return pd.DataFrame(columns=COLUNAS_PACIENTES)


def carregar_avaliacoes() -> pd.DataFrame:
    if os.path.exists(ARQ_AVALIACOES):
        df = pd.read_excel(ARQ_AVALIACOES, dtype={"ID_Avaliacao": str, "ID_Paciente": str})
        df = df.reindex(columns=COLUNAS_AVALIACOES)
        df["Data_Hora"] = df["Data_Hora"].astype(str)
        return df
    return pd.DataFrame(columns=COLUNAS_AVALIACOES)


def carregar_historico_completo() -> pd.DataFrame:
    av = carregar_avaliacoes()
    pac = carregar_pacientes()
    m = av.merge(pac, on="ID_Paciente", how="left")
    for c in pac.columns:
        if c != "ID_Paciente":
            m[c] = m[c].fillna("")
    return m


def proximo_id(serie: pd.Series, prefixo: str, largura: int) -> str:
    nums = pd.to_numeric(serie.astype(str).str.replace(prefixo, "", regex=False), errors="coerce").dropna()
    prox = int(nums.max()) + 1 if len(nums) else 1
    return f"{prefixo}{prox:0{largura}d}"


def salvar_avaliacao(dados: dict, id_paciente: str):
    df = carregar_avaliacoes()
    linha = {c: dados.get(c, "") for c in COLUNAS_AVALIACOES}
    linha.update({
        "ID_Avaliacao": proximo_id(df["ID_Avaliacao"], "A", 6),
        "ID_Paciente": id_paciente,
        "Data_Hora": agora().strftime(FMT_DH),
        "Registrado_Por": LOGIN,
    })
    salvar_xlsx(pd.concat([df, pd.DataFrame([linha])], ignore_index=True), ARQ_AVALIACOES)


def excluir_avaliacoes(ids: list):
    df = carregar_avaliacoes()
    salvar_xlsx(df[~df["ID_Avaliacao"].isin(ids)].reset_index(drop=True), ARQ_AVALIACOES)


def excluir_paciente(id_paciente: str):
    p = carregar_pacientes()
    a = carregar_avaliacoes()
    salvar_xlsx(p[p["ID_Paciente"] != id_paciente].reset_index(drop=True), ARQ_PACIENTES)
    salvar_xlsx(a[a["ID_Paciente"] != id_paciente].reset_index(drop=True), ARQ_AVALIACOES)


# =====================================================================
# CATÁLOGO DE DIETAS ENTERAIS (ADULTO)
# =====================================================================
CATALOGO_PADRAO = [
    {"nome": "Padrão normocalórica 1,0 kcal/ml (genérica)", "kcal": 100.0, "ptn": 4.0, "cho": 12.8, "lip": 3.4, "agua_pct": 85.0},
    {"nome": "Padrão hipercalórica 1,5 kcal/ml (genérica)", "kcal": 150.0, "ptn": 6.0, "cho": 18.8, "lip": 5.8, "agua_pct": 76.0},
    {"nome": "Hiperproteica normocalórica 1,0 kcal/ml (genérica)", "kcal": 100.0, "ptn": 6.3, "cho": 11.5, "lip": 3.0, "agua_pct": 84.0},
]


def carregar_catalogo() -> list:
    if os.path.exists(ARQ_CATALOGO):
        try:
            with open(ARQ_CATALOGO, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return list(CATALOGO_PADRAO)


def salvar_catalogo(cat: list):
    with open(ARQ_CATALOGO, "w", encoding="utf-8") as f:
        json.dump(cat, f, ensure_ascii=False, indent=2)


# =====================================================================
# TABELAS E FÓRMULAS DE REFERÊNCIA
# =====================================================================
def holliday_segar(peso: float) -> float:
    if peso <= 10:
        return 100 * peso
    elif peso <= 20:
        return 1000 + 50 * (peso - 10)
    return 1500 + 20 * (peso - 20)


def calcular_tmb(faixa: str, sexo: str, peso: float, altura_cm: float):
    """Equações de Schofield (altura em cm) e OMS."""
    if faixa == "0 a 3 anos":
        if sexo == "M":
            schofield = 0.167 * peso + 15.174 * altura_cm - 617.6
            oms = 60.9 * peso - 54
        else:
            schofield = 16.252 * peso + 10.232 * altura_cm - 413.5
            oms = 61.0 * peso - 51
    elif faixa == "3 a 10 anos":
        if sexo == "M":
            schofield = 19.59 * peso + 1.303 * altura_cm + 414.9
            oms = 22.7 * peso + 495
        else:
            schofield = 16.969 * peso + 1.618 * altura_cm + 371.2
            oms = 22.5 * peso + 499
    else:
        if sexo == "M":
            schofield = 16.25 * peso + 1.372 * altura_cm + 515.5
            oms = 17.5 * peso + 651
        else:
            schofield = 8.365 * peso + 4.65 * altura_cm + 200
            oms = 12.2 * peso + 746
    return schofield, oms


TABELA_ENERGIA_OMS = {
    "0-1 meses": (113, 107), "1-2 meses": (104, 101), "2-3 meses": (95, 94),
    "3-4 meses": (82, 84), "4-5 meses": (81, 83), "5-6 meses": (81, 82),
    "6-7 meses": (79, 78), "7-8 meses": (79, 78), "8-9 meses": (80, 78),
    "9-10 meses": (80, 79), "10-11 meses": (80, 79), "11-12 meses": (81, 79),
    "0-2 anos": (82, 80), "2-3 anos": (84, 81), "3-4 anos": (80, 77),
    "4-5 anos": (77, 74), "5-6 anos": (74, 72), "6-7 anos": (73, 69),
    "7-8 anos": (71, 67), "8-9 anos": (69, 64), "9-10 anos": (67, 61),
    "10-11 anos": (65, 58), "11-12 anos": (62, 55), "12-13 anos": (60, 52),
    "13-14 anos": (58, 49), "14-15 anos": (56, 47), "15-16 anos": (53, 45),
    "16-17 anos": (52, 44), "17-18 anos": (50, 44),
}

TABELA_PROTEINA_DRI = {
    "0-6 meses": 1.52, "7-12 meses": 1.50, "1-3 anos": 1.10,
    "4-8 anos": 0.95, "9-13 anos": 0.96, "14-18 anos": 0.85, "19-30 anos": 0.80,
}


def classificar_z(z: float, cortes: list, ultimo: str) -> str:
    """cortes: lista de (limite_superior_inclusivo_ou_exclusivo, rotulo) já na ordem."""
    for limite, rotulo in cortes:
        if z < limite:
            return rotulo
    return ultimo


def class_imc_i(z: float, idade_meses: int) -> str:
    if idade_meses < 60:
        if z < -3:
            return "Magreza acentuada"
        if z < -2:
            return "Magreza"
        if z <= 1:
            return "Eutrofia"
        if z <= 2:
            return "Risco de sobrepeso"
        if z <= 3:
            return "Sobrepeso"
        return "Obesidade"
    if z < -3:
        return "Magreza acentuada"
    if z < -2:
        return "Magreza"
    if z <= 1:
        return "Eutrofia"
    if z <= 2:
        return "Sobrepeso"
    if z <= 3:
        return "Obesidade"
    return "Obesidade grave"


def class_peso_i(z: float) -> str:
    if z < -3:
        return "Muito baixo peso para a idade"
    if z < -2:
        return "Baixo peso para a idade"
    if z <= 2:
        return "Peso adequado para a idade"
    return "Peso elevado para a idade"


def class_estatura_i(z: float) -> str:
    if z < -3:
        return "Muito baixa estatura para a idade"
    if z < -2:
        return "Baixa estatura para a idade"
    return "Estatura adequada para a idade"


def class_pc_i(z: float) -> str:
    if z < -2:
        return "Perímetro cefálico reduzido (investigar)"
    if z > 2:
        return "Perímetro cefálico aumentado (investigar)"
    return "Perímetro cefálico adequado"


def class_imc_adulto(imc: float, idade: int) -> str:
    if idade >= 60:  # Lipschitz
        if imc < 22:
            return "Baixo peso"
        if imc <= 27:
            return "Eutrofia"
        return "Excesso de peso"
    if imc < 18.5:
        return "Baixo peso"
    if imc < 25:
        return "Eutrofia"
    if imc < 30:
        return "Sobrepeso"
    if imc < 35:
        return "Obesidade grau I"
    if imc < 40:
        return "Obesidade grau II"
    return "Obesidade grau III"


def class_adequacao(pct: float) -> str:  # Blackburn & Thornton
    if pct > 90:
        return "Eutrofia"
    if pct >= 80:
        return "Desnutrição leve"
    if pct >= 70:
        return "Desnutrição moderada"
    return "Desnutrição grave"


def altura_chumlea(sexo: str, idade: int, aj: float) -> float:
    if idade >= 60:
        return (64.19 - 0.04 * idade + 2.02 * aj) if sexo == "M" else (84.88 - 0.24 * idade + 1.83 * aj)
    return (71.85 + 1.88 * aj) if sexo == "M" else (70.25 + 1.87 * aj - 0.06 * idade)


def peso_chumlea(sexo: str, aj: float, cb: float) -> float:
    if sexo == "M":
        return aj * 1.10 + cb * 3.07 - 75.81
    return aj * 1.09 + cb * 2.68 - 65.51


# =====================================================================
# MENUS POR PERFIL
# =====================================================================
MODULOS_CALCULO = [
    "Pediatria Saudável e TMB",
    "Neonatologia",
    "Neuropediatria",
    "Oncologia e Cardiologia",
    "Antropometria Pediátrica",
    "Antropometria do Adulto Crítico",
    "Estatura Pediátrica e Desnutrição Grave",
    "Dietas Enterais para Adultos",
    "Diluição de Fórmulas Infantis",
    "Composição Nutricional do Lactário",
]
MENUS = {
    "master": ["Gerenciar Usuários"],
    "submaster": ["Painel de Gestão", "Cadastro e Prontuário"] + MODULOS_CALCULO + ["Histórico e Relatórios"],
    "normal": ["Cadastro e Prontuário"] + MODULOS_CALCULO + ["Histórico e Relatórios"],
}

# =====================================================================
# BARRA LATERAL
# =====================================================================
st.sidebar.markdown(
    f"""
    <div style="display:flex; align-items:center; gap:10px; padding:6px 4px 10px 4px;">
        <div>{logo_svg(42, "side")}</div>
        <div>
            <div style="font-weight:700; font-size:1.15rem; color:#0F766E; line-height:1.1;">Nutrieli</div>
            <div style="font-size:.78rem; color:#475569;">{NOME_USUARIO}</div>
            <div style="font-size:.72rem; color:#0F766E; font-weight:600;">{PERFIS[PERFIL].split(" (")[0]}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
if st.sidebar.button("Sair", **larg()):
    st.session_state.usuario = None
    st.rerun()

with st.sidebar.expander("🔑 Alterar minha senha"):
    with st.form("form_troca_senha", clear_on_submit=True):
        s_atual = st.text_input("Senha atual", type="password")
        s_nova = st.text_input("Nova senha (mínimo 4 caracteres)", type="password")
        s_conf = st.text_input("Confirmar nova senha", type="password")
        if st.form_submit_button("Salvar nova senha", **larg()):
            if not conferir_senha(s_atual, USER["senha"]):
                st.error("A senha atual está incorreta.")
            elif len(s_nova) < 4:
                st.error("A nova senha precisa ter pelo menos 4 caracteres.")
            elif s_nova != s_conf:
                st.error("A confirmação não confere com a nova senha.")
            else:
                USUARIOS[LOGIN]["senha"] = gerar_hash(s_nova)
                salvar_usuarios(USUARIOS)
                st.success("Senha alterada com sucesso.")


def bloco_backup(itens):
    with st.sidebar.expander("☁️ Backup e restauração"):
        st.caption("No Streamlit Cloud os arquivos podem ser apagados em reinicializações. Baixe cópias com frequência.")
        for rotulo, caminho, mime, ext in itens:
            if os.path.exists(caminho):
                with open(caminho, "rb") as f:
                    st.download_button(f"⬇️ Baixar {rotulo}", f.read(), file_name=os.path.basename(caminho),
                                       mime=mime, key=f"dl_{rotulo}", **larg())
            up = st.file_uploader(f"Restaurar {rotulo}", type=[ext], key=f"up_{rotulo}")
            if up is not None and st.button(f"⬆️ Restaurar {rotulo}", key=f"rs_{rotulo}", **larg()):
                try:
                    dados = up.getvalue()
                    if ext == "json":
                        json.loads(dados.decode("utf-8"))
                    else:
                        pd.read_excel(io.BytesIO(dados))
                    with open(caminho, "wb") as f:
                        f.write(dados)
                    st.success(f"{rotulo} restaurado.")
                    st.rerun()
                except Exception:
                    st.error("Arquivo inválido. Envie uma cópia gerada pelo próprio sistema.")


if E_MASTER:
    bloco_backup([("usuários", ARQ_USUARIOS, MIME_JSON, "json")])
elif E_SUB:
    bloco_backup([
        ("pacientes", ARQ_PACIENTES, MIME_XLSX, "xlsx"),
        ("avaliações", ARQ_AVALIACOES, MIME_XLSX, "xlsx"),
        ("catálogo de dietas", ARQ_CATALOGO, MIME_JSON, "json"),
    ])

st.sidebar.divider()
modulo = st.sidebar.radio("Menu", MENUS[PERFIL], label_visibility="collapsed")

st.markdown(
    f"""
    <div class="nt-topo">
        <div style="display:flex; align-items:center; gap:14px;">
            <div>{logo_svg(42, "top")}</div>
            <div>
                <h1>Nutrieli</h1>
                <p>Sistema integrado de nutrição clínica — protocolos hospitalares, pediatria e lactário</p>
            </div>
        </div>
        <span class="nt-chip">{NOME_USUARIO} · {PERFIS[PERFIL].split(" (")[0]}</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# =====================================================================
# SELEÇÃO DO PACIENTE NOS MÓDULOS DE CÁLCULO
# =====================================================================
PAC_ID, PAC_ROW, OBS_AVAL = None, None, ""


def bloco_paciente():
    df_p = carregar_pacientes()
    with st.container(border=True):
        st.markdown("**Paciente desta avaliação**")
        if df_p.empty:
            st.warning("Nenhum paciente cadastrado ainda. Cadastre o paciente em **Cadastro e Prontuário** para salvar avaliações.")
            return None, None, ""
        registros = df_p.sort_values("Nome_Paciente").to_dict("records")
        ids = [r["ID_Paciente"] for r in registros]
        rot = {r["ID_Paciente"]: rotulo_paciente(r) for r in registros}
        atual = st.session_state.get("_pac_atual")
        sel = st.selectbox(
            "Paciente", ids, index=ids.index(atual) if atual in ids else None,
            format_func=lambda i: rot[i], placeholder="Selecione o paciente", label_visibility="collapsed",
        )
        st.session_state["_pac_atual"] = sel
        obs = st.text_input("Observações da avaliação", placeholder="Ex.: jejum para exame, restrição hídrica, intolerância...")
        row = next((r for r in registros if r["ID_Paciente"] == sel), None) if sel else None
        if row:
            dn = parse_data(row["Data_Nascimento"])
            st.markdown(
                f"""<div class="nt-ficha"><b>{row['Nome_Paciente']}</b> · {texto_idade(dn)} · {row['Tipo_Paciente'] or '-'}
                · Pront. {val(row['Prontuario'])} · {localizacao(row)}<br>
                Dx nutricional: <b>{val(row['Diagnostico_Nutricional'])}</b> · Via: <b>{val(row['Via_Terapia_Nutricional'])}</b>
                · Patologia: {val(row['Patologia'])}</div>""",
                unsafe_allow_html=True,
            )
        return sel, row, obs


if modulo in MODULOS_CALCULO:
    PAC_ID, PAC_ROW, OBS_AVAL = bloco_paciente()


def botao_salvar_registro(registro: dict):
    st.markdown("---")
    if st.button("💾 Salvar avaliação no prontuário", type="primary", key="btn_salvar_avaliacao"):
        if not PAC_ID:
            st.warning("Selecione o paciente no topo da tela antes de salvar.")
        else:
            registro["Observacoes_Clinicas"] = OBS_AVAL.strip()
            salvar_avaliacao(registro, PAC_ID)
            st.success(f"Avaliação de **{PAC_ROW['Nome_Paciente']}** salva com sucesso.")


# =====================================================================
# CADASTRO E PRONTUÁRIO
# =====================================================================
def campos_prontuario(prefixo: str, v: dict, pediatrico: bool) -> dict:
    """Desenha os campos da ficha dentro de um st.form e devolve os valores."""
    d = {}
    c1, c2 = st.columns([2, 1])
    d["Nome_Paciente"] = c1.text_input("Nome do paciente *", value=v.get("Nome_Paciente", ""), key=f"{prefixo}_nome")
    d["_dn"] = c2.date_input(
        "Data de nascimento *", value=parse_data(v.get("Data_Nascimento", "")),
        min_value=date(1900, 1, 1), max_value=agora().date(), format="DD/MM/YYYY", key=f"{prefixo}_dn",
    )
    if pediatrico:
        d["Nome_Mae"] = st.text_input("Nome da mãe *", value=v.get("Nome_Mae", ""), key=f"{prefixo}_mae")
    else:
        d["Nome_Mae"] = ""

    l1, l2, l3, l4 = st.columns([1, 1.3, 1.3, 1])
    d["Prontuario"] = l1.text_input("Prontuário * (somente números)", value=v.get("Prontuario", ""), key=f"{prefixo}_pront")
    d["Clinica"] = l2.text_input("Clínica", value=v.get("Clinica", ""), key=f"{prefixo}_clin")
    d["Enfermaria"] = l3.text_input("Enfermaria", value=v.get("Enfermaria", ""), key=f"{prefixo}_enf")
    d["Leito"] = l4.text_input("Leito", value=v.get("Leito", ""), key=f"{prefixo}_leito")

    s1, s2, s3 = st.columns(3)
    d["Situacao"] = s1.selectbox("Situação", SITUACOES, index=idx(SITUACOES, v.get("Situacao")), key=f"{prefixo}_sit")
    d["Diagnostico_Nutricional"] = s2.selectbox(
        "Diagnóstico nutricional", DIAG_NUTRICIONAL, index=idx(DIAG_NUTRICIONAL, v.get("Diagnostico_Nutricional")), key=f"{prefixo}_dxn")
    d["Via_Terapia_Nutricional"] = s3.selectbox(
        "Via de terapia nutricional", VIAS_TN, index=idx(VIAS_TN, v.get("Via_Terapia_Nutricional")), key=f"{prefixo}_via")

    t1, t2 = st.columns(2)
    d["Diagnostico_Clinico_Inicial"] = t1.text_area(
        "Diagnóstico clínico inicial", value=v.get("Diagnostico_Clinico_Inicial", ""), height=90, key=f"{prefixo}_dxc")
    d["Comorbidades"] = t2.text_area("Comorbidades", value=v.get("Comorbidades", ""), height=90, key=f"{prefixo}_com")
    d["Patologia"] = st.text_input("Patologia", value=v.get("Patologia", ""), key=f"{prefixo}_pat")
    return d


def validar_prontuario(d: dict, pediatrico: bool, df_p: pd.DataFrame, id_atual: str = None) -> list:
    erros = []
    nome = d["Nome_Paciente"].strip()
    pront = d["Prontuario"].strip()
    if not nome:
        erros.append("Informe o nome do paciente.")
    if d["_dn"] is None:
        erros.append("Informe a data de nascimento.")
    if pediatrico and not d["Nome_Mae"].strip():
        erros.append("Informe o nome da mãe (paciente pediátrico).")
    ok_pront = bool(re.fullmatch(r"[0-9]+", pront))
    if not ok_pront:
        erros.append("O prontuário deve conter somente números, sem letras, espaços ou símbolos.")
    outros = df_p[df_p["ID_Paciente"] != id_atual] if id_atual else df_p
    if ok_pront and (outros["Prontuario"] == pront).any():
        erros.append(f"Já existe um paciente cadastrado com o prontuário {pront}.")
    if nome and d["_dn"] is not None:
        dn = d["_dn"].strftime(FMT_D)
        if ((outros["Nome_Paciente"].str.lower() == nome.lower()) & (outros["Data_Nascimento"] == dn)).any():
            erros.append("Já existe um paciente com este nome e esta data de nascimento.")
    return erros


def dados_para_linha(d: dict) -> dict:
    campos = ["Nome_Paciente", "Nome_Mae", "Prontuario", "Clinica", "Enfermaria", "Leito", "Situacao",
              "Diagnostico_Clinico_Inicial", "Comorbidades", "Patologia", "Diagnostico_Nutricional",
              "Via_Terapia_Nutricional"]
    linha = {c: d[c].strip() if isinstance(d[c], str) else d[c] for c in campos}
    linha["Data_Nascimento"] = d["_dn"].strftime(FMT_D)
    return linha


def ficha_markdown(r: dict):
    dn = parse_data(r["Data_Nascimento"])
    st.markdown(f"### {r['Nome_Paciente']}")
    a, b, c, e = st.columns(4)
    a.metric("Prontuário", val(r["Prontuario"]))
    b.metric("Data de nascimento", f"{r['Data_Nascimento']}", texto_idade(dn))
    c.metric("Situação", val(r["Situacao"]))
    e.metric("Tipo", val(r["Tipo_Paciente"]))
    st.markdown(
        f"**Nome da mãe:** {val(r['Nome_Mae'])}  \n"
        f"**Localização:** {localizacao(r)}  \n"
        f"**Diagnóstico clínico inicial:** {val(r['Diagnostico_Clinico_Inicial'])}  \n"
        f"**Comorbidades:** {val(r['Comorbidades'])}  \n"
        f"**Patologia:** {val(r['Patologia'])}  \n"
        f"**Diagnóstico nutricional:** {val(r['Diagnostico_Nutricional'])}  \n"
        f"**Via de terapia nutricional:** {val(r['Via_Terapia_Nutricional'])}  \n"
        f"**Cadastro:** {val(r['Data_Cadastro'])} por {val(r['Cadastrado_Por'])}"
        + (f"  \n**Última edição:** {r['Ultima_Edicao_Em']} por {r['Ultima_Edicao_Por']}" if not vazio(r["Ultima_Edicao_Em"]) else "")
    )


if modulo == "Cadastro e Prontuário":
    st.header("Cadastro e Prontuário de Pacientes")
    abas = ["Novo paciente", "Consultar prontuários"] + (["Editar prontuário"] if E_SUB else [])
    tabs = st.tabs(abas)

    with tabs[0]:
        if "_ok_cadastro" in st.session_state:
            st.success(st.session_state.pop("_ok_cadastro"))
        st.session_state.setdefault("n_form_novo", 0)
        tipo = st.radio("Tipo de paciente", ["Pediátrico", "Adulto"], horizontal=True, key="novo_tipo")
        ped = tipo == "Pediátrico"
        with st.form("form_novo_paciente"):
            dados = campos_prontuario(f"novo{st.session_state.n_form_novo}", {}, ped)
            enviar = st.form_submit_button("Cadastrar paciente", type="primary", **larg())
        if enviar:
            df_p = carregar_pacientes()
            erros = validar_prontuario(dados, ped, df_p)
            if erros:
                for e in erros:
                    st.error(e)
            else:
                linha = {c: "" for c in COLUNAS_PACIENTES}
                linha.update(dados_para_linha(dados))
                linha.update({
                    "ID_Paciente": proximo_id(df_p["ID_Paciente"], "P", 5),
                    "Data_Cadastro": agora().strftime(FMT_DH),
                    "Tipo_Paciente": tipo,
                    "Cadastrado_Por": LOGIN,
                })
                salvar_xlsx(pd.concat([df_p, pd.DataFrame([linha])], ignore_index=True), ARQ_PACIENTES)
                st.session_state.n_form_novo += 1
                st.session_state["_ok_cadastro"] = f"Paciente **{linha['Nome_Paciente']}** cadastrado com sucesso."
                st.rerun()

    with tabs[1]:
        df_p = carregar_pacientes()
        if df_p.empty:
            st.info("Nenhum paciente cadastrado ainda. Use a aba **Novo paciente** para começar.")
        else:
            b1, b2, b3 = st.columns([2, 1.5, 1])
            busca = b1.text_input("Buscar por nome, mãe ou prontuário")
            f_sit = b2.multiselect("Situação", SITUACOES, default=["Internado"])
            f_tipo = b3.selectbox("Tipo", ["Todos", "Pediátrico", "Adulto"])
            df_f = df_p.copy()
            if busca.strip():
                alvo = df_f["Nome_Paciente"] + " " + df_f["Nome_Mae"] + " " + df_f["Prontuario"]
                df_f = df_f[alvo.str.contains(busca.strip(), case=False, na=False, regex=False)]
            if f_sit:
                df_f = df_f[df_f["Situacao"].isin(f_sit)]
            if f_tipo != "Todos":
                df_f = df_f[df_f["Tipo_Paciente"] == f_tipo]
            tabela(exibir(df_f, ["Prontuario", "Nome_Paciente", "Data_Nascimento", "Nome_Mae", "Clinica", "Enfermaria",
                                 "Leito", "Situacao", "Diagnostico_Nutricional", "Via_Terapia_Nutricional"]))
            st.caption(f"{len(df_f)} paciente(s) encontrado(s).")
            regs = df_f.to_dict("records")
            if regs:
                rot = {r["ID_Paciente"]: rotulo_paciente(r) for r in regs}
                pid = st.selectbox("Abrir ficha completa", list(rot.keys()), index=None,
                                   format_func=lambda i: rot[i], placeholder="Selecione o paciente")
                if pid:
                    with st.container(border=True):
                        ficha_markdown(next(r for r in regs if r["ID_Paciente"] == pid))
            if not E_SUB:
                st.caption("A edição de prontuários é restrita à nutricionista responsável.")

    if E_SUB:
        with tabs[2]:
            df_p = carregar_pacientes()
            if "_ok_edicao" in st.session_state:
                st.success(st.session_state.pop("_ok_edicao"))
            if df_p.empty:
                st.info("Nenhum paciente cadastrado ainda.")
            else:
                rot = {r["ID_Paciente"]: rotulo_paciente(r) for r in df_p.sort_values("Nome_Paciente").to_dict("records")}
                pid = st.selectbox("Paciente a editar", list(rot.keys()), index=None,
                                   format_func=lambda i: rot[i], placeholder="Selecione o paciente", key="edit_sel")
                if pid:
                    reg = df_p[df_p["ID_Paciente"] == pid].iloc[0].to_dict()
                    tipo_e = st.radio("Tipo de paciente", ["Pediátrico", "Adulto"], horizontal=True,
                                      index=1 if reg["Tipo_Paciente"] == "Adulto" else 0, key=f"edit_tipo_{pid}")
                    ped_e = tipo_e == "Pediátrico"
                    with st.form(f"form_edit_{pid}"):
                        dados = campos_prontuario(f"edit_{pid}", reg, ped_e)
                        salvar = st.form_submit_button("Salvar alterações", type="primary", **larg())
                    if salvar:
                        erros = validar_prontuario(dados, ped_e, df_p, id_atual=pid)
                        if erros:
                            for e in erros:
                                st.error(e)
                        else:
                            mask = df_p["ID_Paciente"] == pid
                            novos = dados_para_linha(dados)
                            novos.update({"Tipo_Paciente": tipo_e, "Ultima_Edicao_Por": LOGIN,
                                          "Ultima_Edicao_Em": agora().strftime(FMT_DH)})
                            for k, v in novos.items():
                                df_p.loc[mask, k] = v
                            salvar_xlsx(df_p, ARQ_PACIENTES)
                            st.session_state["_ok_edicao"] = "Prontuário atualizado com sucesso."
                            st.rerun()
                    with st.expander("Excluir paciente"):
                        n_av = int((carregar_avaliacoes()["ID_Paciente"] == pid).sum())
                        st.warning(f"Esta ação remove o paciente e suas {n_av} avaliação(ões) de forma permanente.")
                        conf = st.checkbox("Confirmo a exclusão definitiva deste paciente", key=f"conf_del_{pid}")
                        if st.button("Excluir paciente", disabled=not conf, key=f"btn_del_{pid}"):
                            excluir_paciente(pid)
                            st.session_state["_ok_edicao"] = "Paciente excluído."
                            st.session_state.pop("edit_sel", None)
                            st.rerun()

# =====================================================================
# PAINEL DE GESTÃO (SUBMASTER)
# =====================================================================
elif modulo == "Painel de Gestão":
    st.header("Painel de Gestão")
    df_p = carregar_pacientes()
    df_a = carregar_avaliacoes()
    if df_p.empty:
        st.info("O painel será preenchido assim que houver pacientes cadastrados e avaliações salvas.")
    else:
        agora_ = agora()
        with st.container(border=True):
            f1, f2, f3, f4 = st.columns(4)
            periodo = f1.selectbox("Período de análise", ["Últimos 7 dias", "Últimos 30 dias", "Últimos 90 dias", "Todo o período"], index=1)
            prazo = f2.number_input("Prazo para reavaliação (dias)", min_value=1, max_value=60, value=7, step=1)
            tipo_f = f3.selectbox("Tipo de paciente", ["Todos", "Pediátrico", "Adulto"])
            clinicas = sorted([c for c in df_p["Clinica"].unique() if str(c).strip()])
            clin_f = f4.selectbox("Clínica", ["Todas"] + clinicas)

        dias_per = {"Últimos 7 dias": 7, "Últimos 30 dias": 30, "Últimos 90 dias": 90, "Todo o período": None}[periodo]
        inicio = agora_ - timedelta(days=dias_per) if dias_per else None

        pac = df_p.copy()
        if tipo_f != "Todos":
            pac = pac[pac["Tipo_Paciente"] == tipo_f]
        if clin_f != "Todas":
            pac = pac[pac["Clinica"] == clin_f]
        pac["dt_cad"] = pd.to_datetime(pac["Data_Cadastro"], format=FMT_DH, errors="coerce")

        av = df_a[df_a["ID_Paciente"].isin(pac["ID_Paciente"])].copy()
        av["dt"] = pd.to_datetime(av["Data_Hora"], format=FMT_DH, errors="coerce")
        av_per = av[av["dt"] >= inicio] if inicio is not None else av

        ultima = av.groupby("ID_Paciente")["dt"].max()
        primeira = av.groupby("ID_Paciente")["dt"].min()

        internados = pac[pac["Situacao"] == "Internado"].copy()
        internados["ultima"] = pd.to_datetime(internados["ID_Paciente"].map(ultima), errors="coerce")
        referencia = internados["ultima"].fillna(internados["dt_cad"])
        internados["dias_sem"] = (agora_ - referencia).dt.total_seconds() / 86400
        pendentes = internados[internados["dias_sem"] > prazo]

        n_int = len(internados)
        cobertura = 100 * (1 - len(pendentes) / n_int) if n_int else None
        avaliados_per = av_per["ID_Paciente"].nunique()

        base_cad = pac.set_index("ID_Paciente")["dt_cad"]
        horas = ((primeira - base_cad.reindex(primeira.index)).dt.total_seconds() / 3600).dropna()
        horas = horas[horas >= 0]
        t_primeira = horas.mean() if len(horas) else None

        av_ord = av.dropna(subset=["dt"]).copy()
        av_ord["dia"] = av_ord["dt"].dt.normalize()
        dias_av = av_ord.drop_duplicates(["ID_Paciente", "dia"]).sort_values(["ID_Paciente", "dia"])
        inter = dias_av.groupby("ID_Paciente")["dia"].diff().dt.total_seconds() / 86400
        intervalo = inter.dropna().mean() if inter.notna().any() else None

        dx_ok = (~internados["Diagnostico_Nutricional"].isin(["", "A definir"])).mean() * 100 if n_int else None

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Pacientes internados", f"{n_int}")
        k2.metric("Pacientes avaliados no período", f"{avaliados_per}")
        k3.metric("Reavaliação em dia", f"{cobertura:.0f}%" if cobertura is not None else "-", f"prazo de {prazo} dia(s)", delta_color="off")
        k4.metric("Reavaliações pendentes", f"{len(pendentes)}")

        k5, k6, k7, k8 = st.columns(4)
        k5.metric("Avaliações no período", f"{len(av_per)}")
        k6.metric("Média de avaliações por paciente", f"{len(av_per) / avaliados_per:.1f}" if avaliados_per else "-")
        k7.metric("Tempo médio até a 1ª avaliação", f"{t_primeira:.1f} h" if t_primeira is not None else "-")
        k8.metric("Intervalo médio entre reavaliações", f"{intervalo:.1f} dias" if intervalo is not None else "-")

        st.caption(
            f"Diagnóstico nutricional definido em {dx_ok:.0f}% dos internados." if dx_ok is not None else ""
        )
        st.divider()

        def contagem(serie: pd.Series) -> pd.Series:
            return serie.replace("", "Não informado").value_counts()

        g1, g2 = st.columns(2)
        with g1:
            st.subheader("Diagnóstico nutricional (internados)")
            if n_int:
                st.bar_chart(contagem(internados["Diagnostico_Nutricional"]))
        with g2:
            st.subheader("Via de terapia nutricional (internados)")
            if n_int:
                st.bar_chart(contagem(internados["Via_Terapia_Nutricional"]))

        g3, g4 = st.columns(2)
        with g3:
            st.subheader("Avaliações por dia")
            if len(av_per.dropna(subset=["dt"])):
                st.line_chart(av_per.dropna(subset=["dt"]).groupby(av_per["dt"].dt.normalize()).size())
            else:
                st.caption("Sem avaliações no período.")
        with g4:
            st.subheader("Avaliações por protocolo")
            if len(av_per):
                st.bar_chart(av_per["Modulo_Calculo"].value_counts())
            else:
                st.caption("Sem avaliações no período.")

        g5, g6 = st.columns(2)
        with g5:
            st.subheader("Avaliações por profissional")
            if len(av_per):
                nomes = {k: v["nome"] for k, v in USUARIOS.items()}
                st.bar_chart(av_per["Registrado_Por"].map(lambda x: nomes.get(x, x)).value_counts())
            else:
                st.caption("Sem avaliações no período.")
        with g6:
            st.subheader("Internados por clínica")
            if n_int:
                st.bar_chart(contagem(internados["Clinica"]))

        st.divider()
        st.subheader(f"Pacientes com reavaliação pendente (mais de {prazo} dia(s))")
        if pendentes.empty:
            st.success("Nenhuma reavaliação pendente com os parâmetros atuais.")
        else:
            t = pendentes.sort_values("dias_sem", ascending=False)
            tab_pend = pd.DataFrame({
                "Paciente": t["Nome_Paciente"].values,
                "Prontuário": t["Prontuario"].values,
                "Localização": [localizacao(r) for r in t.to_dict("records")],
                "Via de terapia": t["Via_Terapia_Nutricional"].values,
                "Diagnóstico nutricional": t["Diagnostico_Nutricional"].values,
                "Última avaliação": t["ultima"].dt.strftime(FMT_DH).fillna("Sem avaliação").values,
                "Dias sem reavaliação": t["dias_sem"].round(1).values,
            })
            tabela(tab_pend.astype(str))
            st.download_button("Baixar lista de pendências (.xlsx)", xlsx_bytes(tab_pend, "Pendencias"),
                               file_name="reavaliacoes_pendentes.xlsx", mime=MIME_XLSX)

# =====================================================================
# PEDIATRIA SAUDÁVEL E TMB
# =====================================================================
elif modulo == "Pediatria Saudável e TMB":
    st.header("Pediatria Saudável e TMB")
    col1, col2, col3 = st.columns(3)
    with col1:
        sexo = st.selectbox("Sexo", ["M", "F"])
        idade_energia = st.selectbox("Faixa etária (energia OMS)", list(TABELA_ENERGIA_OMS.keys()))
        idade_ptn = st.selectbox("Faixa etária (proteína DRI 2001)", list(TABELA_PROTEINA_DRI.keys()))
    with col2:
        peso = st.number_input("Peso (kg)", min_value=0.0, value=12.0, step=0.1)
        altura_cm = st.number_input("Estatura (cm)", min_value=0.0, value=85.0, step=0.5)
        vol_recebido = st.number_input("Volume hídrico recebido (ml/dia)", min_value=0.0, value=1000.0, step=10.0)
    with col3:
        idade_anos = st.number_input("Idade em anos completos (fibras e açúcar)", min_value=0, max_value=18, value=2)
        faixa_tmb = st.selectbox("Faixa etária (equações de TMB)", ["0 a 3 anos", "3 a 10 anos", "10 a 18 anos"])
        fator_injuria = st.number_input("Fator de injúria / atividade", min_value=1.0, value=1.0, step=0.05)

    kcal_kg = TABELA_ENERGIA_OMS[idade_energia][0 if sexo == "M" else 1]
    vet_oms = peso * kcal_kg * fator_injuria
    ptn_dia = peso * TABELA_PROTEINA_DRI[idade_ptn]
    nec_hidrica = holliday_segar(peso)
    taxa_hidrica_pct = (vol_recebido * 100 / nec_hidrica) if nec_hidrica > 0 else 0
    taxa_hidrica_ml_kg = (vol_recebido / peso) if peso > 0 else 0
    fibras = idade_anos + 5
    acucar_max_g = (0.25 * vet_oms) / 4
    tmb_schofield, tmb_oms = calcular_tmb(faixa_tmb, sexo, peso, altura_cm)

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("VET recomendado (OMS)", f"{vet_oms:.1f} kcal/dia", f"{kcal_kg} kcal/kg")
    r2.metric("Necessidade proteica (DRI)", f"{ptn_dia:.1f} g/dia", f"{TABELA_PROTEINA_DRI[idade_ptn]} g/kg")
    r3.metric("Necessidade hídrica", f"{nec_hidrica:.0f} ml/dia", f"Recebido: {taxa_hidrica_pct:.1f}% ({taxa_hidrica_ml_kg:.1f} ml/kg)")
    r4.metric("Fibras / limite de açúcar", f"{fibras} g fibras/dia", f"Máx. açúcar (2-6a): {acucar_max_g:.1f} g")
    st.info(f"**TMB estimada:** Schofield = **{tmb_schofield:.1f} kcal/dia** | OMS = **{tmb_oms:.1f} kcal/dia**")

    botao_salvar_registro({
        "Modulo_Calculo": "Pediatria Saudável e TMB", "Sexo": sexo,
        "Faixa_Etaria_Condicao": f"{idade_energia} | PTN: {idade_ptn}",
        "Peso_kg": round(peso, 2), "Estatura_cm": round(altura_cm, 1),
        "VET_kcal_dia": round(vet_oms, 1), "Proteina_g_dia": round(ptn_dia, 1),
        "Nec_Hidrica_ml_dia": round(nec_hidrica, 1), "Vol_Recebido_ml_dia": round(vol_recebido, 1),
        "Taxa_Hidrica_pct": round(taxa_hidrica_pct, 1), "Fibras_g_dia": fibras,
        "Acucar_Max_g_dia": round(acucar_max_g, 1), "TMB_Schofield_kcal": round(tmb_schofield, 1),
        "TMB_OMS_kcal": round(tmb_oms, 1),
        "Detalhes_Protocolo": f"Fator injúria: {fator_injuria} | {taxa_hidrica_ml_kg:.1f} ml/kg/dia",
    })

# =====================================================================
# NEONATOLOGIA
# =====================================================================
elif modulo == "Neonatologia":
    st.header("Necessidades Nutricionais e Hídricas em Neonatologia")
    col1, col2 = st.columns(2)
    with col1:
        tipo_rn = st.radio("Classificação", ["Recém-nascido a termo (RNT)", "Recém-nascido pré-termo (RNPT / MBP)"])
        peso_neo = st.number_input("Peso (kg)", min_value=0.0, value=2.5, step=0.05)
        vol_neo = st.number_input("Volume recebido (ml/dia)", min_value=0.0, value=300.0, step=10.0)

    ptn_str, cho_str, lip_str = "", "", ""
    with col2:
        if tipo_rn.startswith("Recém-nascido a termo"):
            taxa_kcal = st.selectbox("Taxa calórica RNT (kcal/kg/dia)", [100, 110, 120])
            taxa_agua = 120
            ptn_str = f"{1.5 * peso_neo:.1f} a {2.5 * peso_neo:.1f}"
            st.write(f"**Proteína para lactentes (1,5–2,5 g/kg):** {ptn_str} g/dia")
        else:
            taxa_kcal = st.selectbox("Taxa calórica RNPT (kcal/kg/dia)", [110, 120, 130])
            taxa_agua = st.selectbox("Taxa hídrica RNPT (ml/kg/dia)", [135, 200])
            ptn_str = f"{3.5 * peso_neo:.1f} a {4.5 * peso_neo:.1f}"
            cho_str = f"{11.6 * peso_neo:.1f} a {13.2 * peso_neo:.1f}"
            lip_str = f"{4.8 * peso_neo:.1f} a {6.6 * peso_neo:.1f}"
            st.write("**Macronutrientes RNPT / MBP (Koletzko / ESPGHAN):**")
            st.write(f"- Proteína (3,5–4,5 g/kg): **{ptn_str} g/dia**")
            st.write(f"- Carboidrato (11,6–13,2 g/kg): **{cho_str} g/dia**")
            st.write(f"- Lipídio (4,8–6,6 g/kg): **{lip_str} g/dia**")

    vet_neo = peso_neo * taxa_kcal
    rec_hidrica_neo = peso_neo * taxa_agua
    pct_hidrica = (vol_neo * 100 / rec_hidrica_neo) if rec_hidrica_neo > 0 else 0

    c1, c2 = st.columns(2)
    c1.metric("VET calculado", f"{vet_neo:.1f} kcal/dia", f"{taxa_kcal} kcal/kg")
    c2.metric("Recomendação hídrica", f"{rec_hidrica_neo:.0f} ml/dia", f"Adequação: {pct_hidrica:.1f}%")

    botao_salvar_registro({
        "Modulo_Calculo": "Neonatologia (RNT/RNPT)", "Faixa_Etaria_Condicao": tipo_rn,
        "Peso_kg": round(peso_neo, 3), "VET_kcal_dia": round(vet_neo, 1), "Proteina_g_dia": ptn_str,
        "Carboidrato_g_dia": cho_str, "Lipidio_g_dia": lip_str, "Nec_Hidrica_ml_dia": round(rec_hidrica_neo, 1),
        "Vol_Recebido_ml_dia": round(vol_neo, 1), "Taxa_Hidrica_pct": round(pct_hidrica, 1),
        "Detalhes_Protocolo": f"{taxa_kcal} kcal/kg | {taxa_agua} ml/kg",
    })

# =====================================================================
# NEUROPEDIATRIA
# =====================================================================
elif modulo == "Neuropediatria":
    st.header("Estimativa Energética e Hídrica para Pacientes Neuropatas")
    col1, col2 = st.columns(2)
    with col1:
        peso_n = st.number_input("Peso (kg)", min_value=0.0, value=15.0, step=0.1)
        altura_n = st.number_input("Altura (cm)", min_value=0.0, value=100.0, step=0.5)
        condicao_motora = st.selectbox(
            "Condição motora (estimativa pela altura)",
            ["Sem disfunção motora (15 kcal/cm)", "Com disfunção, mas deambula (14 kcal/cm)", "Não deambula (11 kcal/cm)"],
        )
    with col2:
        taxa_idade_n = st.number_input("Taxa calórica para a idade (kcal/kg)", min_value=0.0, value=80.0, step=1.0)
        tmb_n = st.number_input("TMB / GEB calculada (kcal)", min_value=0.0, value=600.0, step=10.0)
        vol_n = st.number_input("Volume hídrico recebido (ml/dia)", min_value=0.0, value=1000.0, step=10.0)

    fator_alt = 15 if "15" in condicao_motora else (14 if "14" in condicao_motora else 11)
    vet_altura = altura_n * fator_alt
    vet_70 = peso_n * taxa_idade_n * 0.7
    vet_geb = tmb_n * 1.1
    hs_base = holliday_segar(peso_n)
    nec_hidrica_neuro = hs_base * 0.9
    pct_hid_n = (vol_n * 100 / nec_hidrica_neuro) if nec_hidrica_neuro > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("NE pela altura", f"{vet_altura:.1f} kcal/dia")
    c2.metric("70% da taxa calórica da idade", f"{vet_70:.1f} kcal/dia")
    c3.metric("GET (GEB × 1,1)", f"{vet_geb:.1f} kcal/dia")
    st.info(
        f"**Necessidade hídrica ajustada (90% H&S):** {nec_hidrica_neuro:.0f} ml/dia | "
        f"**BPI (+20%):** {hs_base * 1.2:.0f} ml | **MBPI (+50%):** {hs_base * 1.5:.0f} ml"
    )

    botao_salvar_registro({
        "Modulo_Calculo": "Neuropediatria", "Faixa_Etaria_Condicao": condicao_motora,
        "Peso_kg": round(peso_n, 2), "Estatura_cm": round(altura_n, 1), "VET_kcal_dia": round(vet_altura, 1),
        "Nec_Hidrica_ml_dia": round(nec_hidrica_neuro, 1), "Vol_Recebido_ml_dia": round(vol_n, 1),
        "Taxa_Hidrica_pct": round(pct_hid_n, 1),
        "Detalhes_Protocolo": f"VET altura: {vet_altura:.0f} | 70% idade: {vet_70:.0f} | GET: {vet_geb:.0f} | BPI: {hs_base * 1.2:.0f} ml | MBPI: {hs_base * 1.5:.0f} ml",
    })

# =====================================================================
# ONCOLOGIA E CARDIOLOGIA
# =====================================================================
elif modulo == "Oncologia e Cardiologia":
    st.header("Cálculo Calórico e Proteico em Oncologia e Cardiologia")
    especialidade = st.radio("Especialidade clínica", ["Oncologia (ASPEN, 2002)", "Cardiologia (SBC, 2014 / AMB)"], horizontal=True)

    if especialidade.startswith("Oncologia"):
        peso_o = st.number_input("Peso do paciente oncológico (kg)", min_value=0.0, value=20.0, step=0.1)
        faixas_onco = {
            "0 a 1 ano (90–120 kcal/kg)": (90, 120), "1 a 7 anos (75–90 kcal/kg)": (75, 90),
            "7 a 12 anos (60–75 kcal/kg)": (60, 75), "12 a 18 anos (30–60 kcal/kg)": (30, 60),
            "18 a 25 anos (25–30 kcal/kg)": (25, 30),
        }
        faixa_o = st.selectbox("Faixa etária (ASPEN)", list(faixas_onco.keys()))
        acrescimo_desnut = st.slider("Acréscimo por perda de peso / desnutrição (%)", 0, 50, 0)
        k_min, k_max = faixas_onco[faixa_o]
        fator_d = 1 + (acrescimo_desnut / 100)
        vet_str = f"{peso_o * k_min * fator_d:.1f} a {peso_o * k_max * fator_d:.1f}"
        st.success(f"**VET estimado:** {vet_str} kcal/dia")

        botao_salvar_registro({
            "Modulo_Calculo": "Oncologia (ASPEN)", "Faixa_Etaria_Condicao": faixa_o,
            "Peso_kg": round(peso_o, 2), "VET_kcal_dia": vet_str,
            "Detalhes_Protocolo": f"Acréscimo por desnutrição: +{acrescimo_desnut}%",
        })
    else:
        peso_c = st.number_input("Peso do paciente cardiopata (kg)", min_value=0.0, value=8.0, step=0.1)
        gravidade = st.selectbox("Classificação (0 a 2 anos — AMB/CFM)",
                                 ["Leve a moderado (120–150 kcal/kg)", "Moderado a grave (170–180 kcal/kg)"])
        c_min, c_max = (120, 150) if gravidade.startswith("Leve") else (170, 180)
        vet_c_str = f"{peso_c * c_min:.1f} a {peso_c * c_max:.1f}"
        ptn_c_str = f"{peso_c * 3:.1f} a {peso_c * 4:.1f}"
        st.success(f"**VET do cardiopata:** {vet_c_str} kcal/dia | **Proteína (3,0–4,0 g/kg):** {ptn_c_str} g/dia")

        botao_salvar_registro({
            "Modulo_Calculo": "Cardiologia (SBC/AMB)", "Faixa_Etaria_Condicao": gravidade,
            "Peso_kg": round(peso_c, 2), "VET_kcal_dia": vet_c_str, "Proteina_g_dia": ptn_c_str,
            "Detalhes_Protocolo": "Protocolo de cardiopatia de 0 a 2 anos",
        })

# =====================================================================
# ANTROPOMETRIA PEDIÁTRICA (NOVO)
# =====================================================================
elif modulo == "Antropometria Pediátrica":
    st.header("Antropometria Pediátrica")
    st.caption("Indicadores de crescimento classificados por escore-z (Ministério da Saúde / OMS).")
    dn_pac = parse_data(PAC_ROW["Data_Nascimento"]) if PAC_ROW else None
    meses_pac = min(idade_em_meses(dn_pac), 228) if dn_pac else 24

    c1, c2, c3 = st.columns(3)
    with c1:
        sexo_a = st.selectbox("Sexo", ["M", "F"])
        idade_m = st.number_input("Idade (meses completos)", min_value=0, max_value=228, value=meses_pac, step=1)
    with c2:
        peso_a = st.number_input("Peso (kg)", min_value=0.0, max_value=150.0, value=12.0, step=0.05)
        est_a = st.number_input("Comprimento / estatura (cm)", min_value=0.0, max_value=220.0, value=85.0, step=0.1)
    with c3:
        pc_a = st.number_input("Perímetro cefálico (cm) — opcional", min_value=0.0, max_value=70.0, value=0.0, step=0.1)
        cb_a = st.number_input("Circunferência do braço (cm) — opcional", min_value=0.0, max_value=60.0, value=0.0, step=0.1)
        dct_a = st.number_input("Dobra cutânea tricipital (mm) — opcional", min_value=0.0, max_value=60.0, value=0.0, step=0.5)

    imc_a = peso_a / ((est_a / 100) ** 2) if peso_a > 0 and est_a > 0 else None
    m1, m2 = st.columns(2)
    m1.metric("IMC", f"{imc_a:.2f} kg/m²" if imc_a else "-")
    m2.metric("Idade", f"{idade_m // 12} anos e {idade_m % 12} meses")

    st.markdown("##### Escores-z dos indicadores")
    st.caption("Obtenha o escore-z em tabelas da OMS ou no WHO Anthro e informe abaixo. Deixe em branco o que não for avaliado.")
    z1, z2, z3, z4, z5 = st.columns(5)
    z_pi = z1.number_input("Peso/Idade", value=None, step=0.1, format="%.2f")
    z_ei = z2.number_input("Estatura/Idade", value=None, step=0.1, format="%.2f")
    z_imc = z3.number_input("IMC/Idade", value=None, step=0.1, format="%.2f")
    z_pe = z4.number_input("Peso/Estatura (até 5 anos)", value=None, step=0.1, format="%.2f")
    z_pc = z5.number_input("Perímetro cefálico/Idade", value=None, step=0.1, format="%.2f")

    linhas = []
    if z_pi is not None:
        linhas.append(("Peso/Idade", z_pi, class_peso_i(z_pi) + (" (referência válida até 10 anos)" if idade_m > 120 else "")))
    if z_ei is not None:
        linhas.append(("Estatura/Idade", z_ei, class_estatura_i(z_ei)))
    if z_imc is not None:
        linhas.append(("IMC/Idade", z_imc, class_imc_i(z_imc, idade_m)))
    if z_pe is not None:
        linhas.append(("Peso/Estatura", z_pe, class_imc_i(z_pe, 0)))
    if z_pc is not None:
        linhas.append(("Perímetro cefálico/Idade", z_pc, class_pc_i(z_pc)))
    if linhas:
        tabela(pd.DataFrame([(a, f"{b:+.2f}", c) for a, b, c in linhas], columns=["Indicador", "Escore-z", "Classificação"]))
    else:
        st.info("Informe ao menos um escore-z para ver a classificação.")

    with st.expander("Calcular escore-z a partir de L, M e S (tabelas da OMS)"):
        l1, l2, l3, l4 = st.columns(4)
        x_med = l1.number_input("Medida do paciente (X)", value=0.0, format="%.2f")
        l_lms = l2.number_input("L", value=0.0, format="%.4f")
        m_lms = l3.number_input("M", value=0.0, format="%.4f")
        s_lms = l4.number_input("S", value=0.0, format="%.5f")
        if x_med > 0 and m_lms > 0 and s_lms > 0:
            z_calc = ((x_med / m_lms) ** l_lms - 1) / (l_lms * s_lms) if abs(l_lms) > 1e-9 else math.log(x_med / m_lms) / s_lms
            l4.metric("Escore-z", f"{z_calc:+.2f}")
            st.caption("Para |z| acima de 3 a OMS aplica ajuste específico; confirme no WHO Anthro.")

    reserva = ""
    if cb_a > 0:
        cmb = cb_a - 0.314 * dct_a
        amb = (cmb ** 2) / (4 * math.pi)
        agb = (cb_a * (dct_a / 10)) / 2 - (math.pi * (dct_a / 10) ** 2) / 4
        r1, r2, r3 = st.columns(3)
        r1.metric("Circunferência muscular do braço", f"{cmb:.1f} cm")
        r2.metric("Área muscular do braço", f"{amb:.1f} cm²")
        r3.metric("Área de gordura do braço", f"{agb:.1f} cm²" if dct_a > 0 else "-")
        reserva = f" | CB {cb_a:.1f} cm, DCT {dct_a:.1f} mm, CMB {cmb:.1f} cm, AMB {amb:.1f} cm²"

    detalhes = " | ".join(f"{a}: z {b:+.2f} ({c})" for a, b, c in linhas)
    if pc_a > 0:
        detalhes += f" | PC {pc_a:.1f} cm"
    botao_salvar_registro({
        "Modulo_Calculo": "Antropometria Pediátrica", "Sexo": sexo_a,
        "Faixa_Etaria_Condicao": f"{idade_m // 12} anos e {idade_m % 12} meses",
        "Peso_kg": round(peso_a, 2), "Estatura_cm": round(est_a, 1),
        "Detalhes_Protocolo": (f"IMC {imc_a:.2f} | " if imc_a else "") + detalhes + reserva,
    })

# =====================================================================
# ANTROPOMETRIA DO ADULTO CRÍTICO (NOVO)
# =====================================================================
elif modulo == "Antropometria do Adulto Crítico":
    st.header("Antropometria do Adulto Crítico")
    st.caption("Estimativas para pacientes acamados. Confirme as equações adotadas com o protocolo do seu serviço.")
    dn_pac = parse_data(PAC_ROW["Data_Nascimento"]) if PAC_ROW else None
    idade_pac = min(max(idade_em_meses(dn_pac) // 12, 18), 110) if dn_pac else 60

    c1, c2, c3 = st.columns(3)
    with c1:
        sexo_c = st.selectbox("Sexo", ["M", "F"])
        idade_c = st.number_input("Idade (anos)", min_value=18, max_value=110, value=idade_pac, step=1)
        peso_af = st.number_input("Peso aferido (kg) — 0 se não aferido", min_value=0.0, max_value=350.0, value=0.0, step=0.5)
        alt_af = st.number_input("Altura aferida (cm) — 0 se não aferida", min_value=0.0, max_value=230.0, value=0.0, step=0.5)
    with c2:
        aj = st.number_input("Altura do joelho (cm)", min_value=0.0, max_value=80.0, value=50.0, step=0.5)
        cb_c = st.number_input("Circunferência do braço (cm)", min_value=0.0, max_value=70.0, value=28.0, step=0.5)
        dct_c = st.number_input("Dobra cutânea tricipital (mm)", min_value=0.0, max_value=60.0, value=12.0, step=0.5)
        cp_c = st.number_input("Circunferência da panturrilha (cm)", min_value=0.0, max_value=70.0, value=33.0, step=0.5)
    with c3:
        peso_hab = st.number_input("Peso habitual (kg) — opcional", min_value=0.0, max_value=350.0, value=0.0, step=0.5)
        periodo_perda = st.selectbox("Período da perda de peso", ["1 semana", "1 mês", "3 meses", "6 meses"], index=1)
        st.caption("Percentil 50 de referência (opcional), para o cálculo de adequação:")
        p50_cb = st.number_input("CB — percentil 50 (cm)", min_value=0.0, value=0.0, step=0.1)
        p50_dct = st.number_input("DCT — percentil 50 (mm)", min_value=0.0, value=0.0, step=0.1)
        p50_cmb = st.number_input("CMB — percentil 50 (cm)", min_value=0.0, value=0.0, step=0.1)

    alt_est = altura_chumlea(sexo_c, int(idade_c), aj) if aj > 0 else None
    peso_est = peso_chumlea(sexo_c, aj, cb_c) if aj > 0 and cb_c > 0 else None
    peso_uso = peso_af if peso_af > 0 else peso_est
    alt_uso = alt_af if alt_af > 0 else alt_est

    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Altura estimada (Chumlea)", f"{alt_est:.1f} cm" if alt_est else "-")
    e2.metric("Peso estimado (Chumlea)", f"{peso_est:.1f} kg" if peso_est else "-")
    e3.metric("Peso utilizado", f"{peso_uso:.1f} kg" if peso_uso else "-", "aferido" if peso_af > 0 else "estimado", delta_color="off")
    e4.metric("Altura utilizada", f"{alt_uso:.1f} cm" if alt_uso else "-", "aferida" if alt_af > 0 else "estimada", delta_color="off")

    detalhes = []
    imc_c = None
    if peso_uso and alt_uso:
        imc_c = peso_uso / ((alt_uso / 100) ** 2)
        cl_imc = class_imc_adulto(imc_c, int(idade_c))
        st.info(f"**IMC:** {imc_c:.1f} kg/m² — **{cl_imc}** ({'Lipschitz, idosos' if idade_c >= 60 else 'OMS, adultos'})")
        detalhes.append(f"IMC {imc_c:.1f} ({cl_imc})")

    if cb_c > 0 and dct_c > 0:
        cmb_c = cb_c - 0.314 * dct_c
        amb_c = (cmb_c ** 2) / (4 * math.pi)
        amb_corr = amb_c - (10 if sexo_c == "M" else 6.5)
        agb_c = (cb_c * (dct_c / 10)) / 2 - (math.pi * (dct_c / 10) ** 2) / 4
        a1, a2, a3, a4 = st.columns(4)
        a1.metric("CMB", f"{cmb_c:.1f} cm")
        a2.metric("AMB", f"{amb_c:.1f} cm²")
        a3.metric("AMB corrigida", f"{amb_corr:.1f} cm²")
        a4.metric("AGB", f"{agb_c:.1f} cm²")
        detalhes.append(f"CB {cb_c:.1f} cm, DCT {dct_c:.1f} mm, CMB {cmb_c:.1f} cm, AMB corr. {amb_corr:.1f} cm², AGB {agb_c:.1f} cm²")

        adeq = []
        for nome_i, valor_i, p50 in [("CB", cb_c, p50_cb), ("DCT", dct_c, p50_dct), ("CMB", cmb_c, p50_cmb)]:
            if p50 > 0:
                pct = valor_i / p50 * 100
                adeq.append((nome_i, f"{pct:.1f}%", class_adequacao(pct)))
                detalhes.append(f"Adequação {nome_i} {pct:.0f}% ({class_adequacao(pct)})")
        if adeq:
            tabela(pd.DataFrame(adeq, columns=["Medida", "Adequação ao percentil 50", "Classificação (Blackburn & Thornton)"]))

    if cp_c > 0:
        if cp_c < 31:
            st.warning(f"**Circunferência da panturrilha:** {cp_c:.1f} cm — abaixo de 31 cm, sugere redução de massa muscular.")
            detalhes.append(f"CP {cp_c:.1f} cm (reduzida)")
        else:
            st.success(f"**Circunferência da panturrilha:** {cp_c:.1f} cm — adequada (31 cm ou mais).")
            detalhes.append(f"CP {cp_c:.1f} cm (adequada)")

    if peso_hab > 0 and peso_uso:
        perda = (peso_hab - peso_uso) / peso_hab * 100
        limites = {"1 semana": (1, 2), "1 mês": (5, 5), "3 meses": (7.5, 7.5), "6 meses": (10, 10)}
        lo, hi = limites[periodo_perda]
        if perda < lo:
            cl_perda = "perda não significativa" if perda > 0 else "sem perda"
        elif perda <= hi:
            cl_perda = "perda significativa"
        else:
            cl_perda = "perda grave"
        st.metric(f"Variação do peso em {periodo_perda}", f"{perda:.1f}%", cl_perda, delta_color="off")
        detalhes.append(f"Perda de peso {perda:.1f}% em {periodo_perda} ({cl_perda})")

    botao_salvar_registro({
        "Modulo_Calculo": "Antropometria do Adulto Crítico", "Sexo": sexo_c,
        "Faixa_Etaria_Condicao": f"{int(idade_c)} anos",
        "Peso_kg": round(peso_uso, 1) if peso_uso else "", "Estatura_cm": round(alt_uso, 1) if alt_uso else "",
        "Detalhes_Protocolo": " | ".join(detalhes),
    })

# =====================================================================
# ESTATURA PEDIÁTRICA E DESNUTRIÇÃO GRAVE
# =====================================================================
elif modulo == "Estatura Pediátrica e Desnutrição Grave":
    st.header("Estimativa de Estatura e Protocolo de Desnutrição Grave")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Estimativa de estatura (Stevenson)")
        medida_tipo = st.selectbox("Segmento corporal", [
            "Comprimento superior do braço (Stevenson)", "Comprimento tibial (Stevenson)",
            "Altura do joelho pediátrica (Stevenson)",
        ])
        medida_cm = st.number_input("Medida do segmento (cm)", min_value=0.0, value=30.0, step=0.1)
        if "braço" in medida_tipo:
            est = 4.35 * medida_cm + 21.8
        elif "tibial" in medida_tipo:
            est = 3.26 * medida_cm + 30.8
        else:
            est = 2.69 * medida_cm + 24.2
        st.metric("Estatura estimada", f"{est:.1f} cm")
        st.caption("Para adultos acamados, use o módulo Antropometria do Adulto Crítico.")

    with col2:
        st.subheader("Protocolo de desnutrição grave (OMS, 2005)")
        peso_desn = st.number_input("Peso da criança (kg)", min_value=0.0, value=6.0, step=0.1)
        fase_atual = st.selectbox("Fase atual do protocolo", ["Fase 1 (2 dias)", "Fase 2 (3 dias)", "Fase 3 (3 dias)", "Fase 4 (3 dias*)", "Fase 5 (3 dias*)"])
        fases_df = pd.DataFrame({
            "Fase": ["Fase 1", "Fase 2", "Fase 3", "Fase 4", "Fase 5"],
            "VET (kcal/dia)": [75 * peso_desn, 100 * peso_desn, 150 * peso_desn, 200 * peso_desn, 250 * peso_desn],
            "Proteína (g/dia)": [1.5 * peso_desn, 2.5 * peso_desn, 3.5 * peso_desn, 4.5 * peso_desn, 5.5 * peso_desn],
            "Evolução": ["2 dias", "3 dias", "3 dias", "3 dias*", "3 dias*"],
        })
        tabela(fases_df)

    idx_fase = int(fase_atual.split()[1]) - 1
    botao_salvar_registro({
        "Modulo_Calculo": "Estatura e Desnutrição Grave", "Faixa_Etaria_Condicao": f"{medida_tipo} | {fase_atual}",
        "Peso_kg": round(peso_desn, 2), "Estatura_cm": round(est, 1),
        "VET_kcal_dia": round(fases_df.loc[idx_fase, "VET (kcal/dia)"], 1),
        "Proteina_g_dia": round(fases_df.loc[idx_fase, "Proteína (g/dia)"], 1),
        "Detalhes_Protocolo": f"Segmento: {medida_cm} cm → estatura estimada: {est:.1f} cm",
    })

# =====================================================================
# DIETAS ENTERAIS PARA ADULTOS (NOVO)
# =====================================================================
elif modulo == "Dietas Enterais para Adultos":
    st.header("Dietas Enterais para Adultos")
    st.caption("Os valores das dietas genéricas são aproximados. Para produtos comerciais, cadastre a composição do rótulo no catálogo.")
    catalogo = carregar_catalogo()

    c1, c2, c3 = st.columns(3)
    peso_e = c1.number_input("Peso (kg)", min_value=20.0, max_value=250.0, value=70.0, step=0.5)
    kcal_kg_e = c2.number_input("Meta calórica (kcal/kg/dia)", min_value=10.0, max_value=50.0, value=25.0, step=1.0)
    ptn_kg_e = c3.number_input("Meta proteica (g/kg/dia)", min_value=0.5, max_value=3.0, value=1.5, step=0.1)

    opcoes = [d["nome"] for d in catalogo] + ["Personalizada (informar rótulo)"]
    dieta_sel = st.selectbox("Dieta enteral", opcoes)
    if dieta_sel.startswith("Personalizada"):
        p1, p2, p3, p4, p5 = st.columns(5)
        dieta = {
            "nome": "Personalizada",
            "kcal": p1.number_input("kcal / 100 ml", min_value=1.0, value=100.0, step=1.0),
            "ptn": p2.number_input("Proteína g / 100 ml", min_value=0.0, value=4.0, step=0.1),
            "cho": p3.number_input("Carboidrato g / 100 ml", min_value=0.0, value=12.8, step=0.1),
            "lip": p4.number_input("Lipídio g / 100 ml", min_value=0.0, value=3.4, step=0.1),
            "agua_pct": p5.number_input("Água na dieta (%)", min_value=0.0, max_value=100.0, value=85.0, step=1.0),
        }
    else:
        dieta = next(d for d in catalogo if d["nome"] == dieta_sel)

    meta_kcal = peso_e * kcal_kg_e
    meta_ptn = peso_e * ptn_kg_e
    d1, d2 = st.columns(2)
    modo = d1.radio("Definir o volume por", ["Atingir a meta calórica", "Informar volume prescrito (ml/dia)"], horizontal=True)
    if modo.startswith("Atingir"):
        vol_e = meta_kcal / (dieta["kcal"] / 100)
        d1.caption(f"Volume calculado: {vol_e:.0f} ml/dia")
    else:
        vol_e = d1.number_input("Volume prescrito (ml/dia)", min_value=0.0, max_value=5000.0, value=1500.0, step=50.0)
    nec_hid_kg = d2.number_input("Necessidade hídrica total (ml/kg/dia)", min_value=20.0, max_value=45.0, value=30.0, step=1.0)

    i1, i2 = st.columns(2)
    infusao = i1.radio("Forma de infusão", ["Contínua (bomba)", "Intermitente / fracionada"], horizontal=True)
    if infusao.startswith("Contínua"):
        horas = i2.number_input("Horas de infusão por dia", min_value=1.0, max_value=24.0, value=20.0, step=1.0)
        txt_inf = f"{vol_e / horas:.0f} ml/h por {horas:.0f} h"
    else:
        etapas = i2.number_input("Número de etapas por dia", min_value=1, max_value=12, value=6, step=1)
        txt_inf = f"{vol_e / etapas:.0f} ml por etapa, {etapas}×/dia"

    kcal_e = vol_e * dieta["kcal"] / 100
    ptn_e = vol_e * dieta["ptn"] / 100
    cho_e = vol_e * dieta["cho"] / 100
    lip_e = vol_e * dieta["lip"] / 100
    agua_dieta = vol_e * dieta["agua_pct"] / 100
    agua_livre = max(0.0, peso_e * nec_hid_kg - agua_dieta)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Volume da dieta", f"{vol_e:.0f} ml/dia", txt_inf, delta_color="off")
    m2.metric("Calorias", f"{kcal_e:.0f} kcal/dia", f"{kcal_e / meta_kcal * 100:.0f}% da meta ({meta_kcal:.0f})", delta_color="off")
    m3.metric("Proteína", f"{ptn_e:.1f} g/dia", f"{ptn_e / meta_ptn * 100:.0f}% da meta ({meta_ptn:.0f} g)", delta_color="off")
    m4.metric("Água livre a complementar", f"{agua_livre:.0f} ml/dia", f"Dieta fornece {agua_dieta:.0f} ml", delta_color="off")
    st.info(f"**CHO:** {cho_e:.1f} g/dia | **LIP:** {lip_e:.1f} g/dia | **{kcal_e / peso_e:.1f} kcal/kg** | **{ptn_e / peso_e:.2f} g de proteína/kg**")

    botao_salvar_registro({
        "Modulo_Calculo": "Dietas Enterais (Adulto)", "Peso_kg": round(peso_e, 1),
        "VET_kcal_dia": round(kcal_e, 0), "Proteina_g_dia": round(ptn_e, 1),
        "Carboidrato_g_dia": round(cho_e, 1), "Lipidio_g_dia": round(lip_e, 1),
        "Nec_Hidrica_ml_dia": round(peso_e * nec_hid_kg, 0), "Formula_Ou_Item": dieta["nome"],
        "Qtd_Po_g_Ou_Vol_ml": f"{vol_e:.0f} ml/dia",
        "Detalhes_Protocolo": f"Meta {kcal_kg_e:.0f} kcal/kg e {ptn_kg_e:.1f} g/kg | {txt_inf} | água livre {agua_livre:.0f} ml/dia",
    })

    if E_SUB:
        with st.expander("Gerenciar catálogo de dietas"):
            with st.form("form_catalogo", clear_on_submit=True):
                n1, n2 = st.columns([2, 1])
                nome_d = n1.text_input("Nome da dieta")
                kcal_d = n2.number_input("kcal / 100 ml", min_value=0.0, value=100.0, step=1.0)
                q1, q2, q3, q4 = st.columns(4)
                ptn_d = q1.number_input("Proteína g / 100 ml", min_value=0.0, value=4.0, step=0.1)
                cho_d = q2.number_input("Carboidrato g / 100 ml", min_value=0.0, value=12.8, step=0.1)
                lip_d = q3.number_input("Lipídio g / 100 ml", min_value=0.0, value=3.4, step=0.1)
                agua_d = q4.number_input("Água na dieta (%)", min_value=0.0, max_value=100.0, value=85.0, step=1.0)
                if st.form_submit_button("Adicionar ao catálogo", type="primary"):
                    if not nome_d.strip() or kcal_d <= 0:
                        st.error("Informe o nome da dieta e as calorias por 100 ml.")
                    elif any(d["nome"].lower() == nome_d.strip().lower() for d in catalogo):
                        st.error("Já existe uma dieta com este nome no catálogo.")
                    else:
                        catalogo.append({"nome": nome_d.strip(), "kcal": kcal_d, "ptn": ptn_d, "cho": cho_d, "lip": lip_d, "agua_pct": agua_d})
                        salvar_catalogo(catalogo)
                        st.rerun()
            remover = st.selectbox("Remover dieta do catálogo", [d["nome"] for d in catalogo], index=None, placeholder="Selecione")
            if remover and st.button("Remover dieta selecionada"):
                salvar_catalogo([d for d in catalogo if d["nome"] != remover])
                st.rerun()

# =====================================================================
# DILUIÇÃO DE FÓRMULAS INFANTIS
# =====================================================================
elif modulo == "Diluição de Fórmulas Infantis":
    st.header("Diluição e Pesagem de Fórmulas Infantis")
    FORMULAS_G_100ML = {
        "PRE NAN": 17.9, "APTAMIL PRÉ": 18.3, "ENFAMIL PRÉ PREMIUM": 15.0,
        "NAN SUPREME 1": 14.7, "MILUPA 1": 14.7, "NESTOGENO 1": 14.7,
        "ENFAMIL PREMIUM 1": 14.3, "NAN COMFOR 1": 14.7, "NAN PRO 1": 14.3,
        "APTAMIL PREMIUM 1": 15.3, "APTAMIL PROFUTURA 1": 15.3, "FÓRMULA 1 - 20%": 19.7,
        "NESTOGENO 2": 15.3, "NAN COMFOR 2": 15.3, "NAN PRO 2": 15.3,
        "NAN SUPREME 2": 15.0, "MILUPA 2": 15.3, "APTAMIL PROFUTURA 2": 16.7,
        "APTAMIL 3": 16.3, "NAN COMFOR 3": 15.3, "NESLAC": 15.7,
        "FORTINI": 20.3, "APTAMIL PREMIUM 2": 17.7, "MILNUTRI": 16.7,
        "PURAMINO": 15.0, "PURAMINO 20%": 20.0, "APTAMIL AR": 15.7,
        "APTAMIL H.A Proexpert": 15.3, "APTAMIL SEM LACTOSE": 14.3,
        "APTAMIL PEPTI": 15.0, "ALFAMINO": 15.3, "ALFAMINO 20%": 23.3,
        "ALFARÉ": 15.0, "ALTHÉRA": 14.5, "NEOCATE LCP": 16.7,
        "NEO ADVANCED": 25.0, "NAN AR": 14.3, "PREGOMIN": 14.3,
        "INFATRINI": 20.0, "INFATRINI 1.5": 30.0, "PEPTAMEN JUNIOR PÓ": 26.2,
    }
    formula_sel = st.selectbox("Fórmula", list(FORMULAS_G_100ML.keys()))
    vol_desejado = st.number_input("Volume prescrito (ml)", min_value=10, max_value=2000, value=150, step=10)
    gramas_po = (vol_desejado * FORMULAS_G_100ML[formula_sel]) / 100
    st.metric(f"Quantidade de pó ({formula_sel})", f"{gramas_po:.2f} g", f"Concentração padrão: {FORMULAS_G_100ML[formula_sel]} g / 100 ml")

    botao_salvar_registro({
        "Modulo_Calculo": "Diluição de Fórmulas Infantis", "Vol_Recebido_ml_dia": vol_desejado,
        "Formula_Ou_Item": formula_sel, "Qtd_Po_g_Ou_Vol_ml": f"{gramas_po:.2f} g de pó para {vol_desejado} ml",
        "Detalhes_Protocolo": f"Diluição padrão: {FORMULAS_G_100ML[formula_sel]}%",
    })

# =====================================================================
# COMPOSIÇÃO NUTRICIONAL DO LACTÁRIO
# =====================================================================
elif modulo == "Composição Nutricional do Lactário":
    st.header("Composição Nutricional do Lactário")
    ITENS_LACTARIO = {
        "Fórmula de partida / seguimento padrão (15%)": (70.0, 0.9, 7.3, 4.2),
        "Alfamino a 15%": (70.0, 1.9, 7.9, 3.4),
        "Mingau / Mucilon / papa / suco (base lactário)": (70.0, 0.9, 7.3, 4.2),
    }
    item = st.selectbox("Preparação / item", list(ITENS_LACTARIO.keys()))
    qtd = st.number_input("Quantidade (ml ou g)", min_value=0.0, value=200.0, step=10.0)
    kcal_100, ptn_100, cho_100, lip_100 = ITENS_LACTARIO[item]
    kcal_tot, ptn_tot = kcal_100 * qtd / 100, ptn_100 * qtd / 100
    cho_tot, lip_tot = cho_100 * qtd / 100, lip_100 * qtd / 100

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Calorias", f"{kcal_tot:.1f} kcal")
    c2.metric("Proteína", f"{ptn_tot:.2f} g")
    c3.metric("Carboidrato", f"{cho_tot:.2f} g")
    c4.metric("Lipídio", f"{lip_tot:.2f} g")

    botao_salvar_registro({
        "Modulo_Calculo": "Composição Nutricional do Lactário", "VET_kcal_dia": round(kcal_tot, 1),
        "Proteina_g_dia": round(ptn_tot, 2), "Carboidrato_g_dia": round(cho_tot, 2),
        "Lipidio_g_dia": round(lip_tot, 2), "Formula_Ou_Item": item, "Qtd_Po_g_Ou_Vol_ml": f"{qtd} ml/g",
    })

# =====================================================================
# HISTÓRICO E RELATÓRIOS
# =====================================================================
elif modulo == "Histórico e Relatórios":
    st.header("Histórico e Relatórios Clínicos")
    df_hist = carregar_historico_completo()
    df_pac_all = carregar_pacientes()

    if df_hist.empty:
        st.info("Nenhuma avaliação registrada ainda. Selecione um paciente em qualquer módulo de cálculo e use **Salvar avaliação no prontuário**.")
    else:
        tab_rel, tab_geral = st.tabs(["Relatório por paciente", "Planilha geral"])

        with tab_rel:
            com_aval = df_pac_all[df_pac_all["ID_Paciente"].isin(df_hist["ID_Paciente"])]
            regs = com_aval.sort_values("Nome_Paciente").to_dict("records")
            rot = {r["ID_Paciente"]: rotulo_paciente(r) for r in regs}
            s1, s2 = st.columns([2, 1])
            pid = s1.selectbox("Paciente", list(rot.keys()), format_func=lambda i: rot[i])
            df_pac = df_hist[df_hist["ID_Paciente"] == pid].copy()
            mods = ["Todos"] + sorted(df_pac["Modulo_Calculo"].dropna().unique().tolist())
            f_mod = s2.selectbox("Protocolo", mods)
            if f_mod != "Todos":
                df_pac = df_pac[df_pac["Modulo_Calculo"] == f_mod]
            df_pac = df_pac.iloc[::-1].reset_index(drop=True)
            reg = next(r for r in regs if r["ID_Paciente"] == pid)

            with st.container(border=True):
                ficha_markdown(reg)

            ult = df_pac.iloc[0]
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Avaliações", f"{len(df_pac)}")
            m2.metric("Último peso / estatura", f"{val(ult['Peso_kg'])} kg | {val(ult['Estatura_cm'])} cm")
            m3.metric("Último VET", f"{val(ult['VET_kcal_dia'])} kcal/d")
            m4.metric("Última proteína", f"{val(ult['Proteina_g_dia'])} g/d")
            st.divider()

            texto = [
                "====================================================================",
                "NUTRIELI - RELATÓRIO INDIVIDUAL DE AVALIAÇÃO E PRESCRIÇÃO NUTRICIONAL",
                f"Data de emissão: {agora().strftime('%d/%m/%Y %H:%M')} | Emitido por: {NOME_USUARIO}",
                f"Paciente: {reg['Nome_Paciente']} | Data de nascimento: {reg['Data_Nascimento']}",
                f"Nome da mãe: {val(reg['Nome_Mae'])} | Prontuário: {val(reg['Prontuario'])}",
                f"Localização: {localizacao(reg)}",
                f"Diagnóstico clínico inicial: {val(reg['Diagnostico_Clinico_Inicial'])}",
                f"Comorbidades: {val(reg['Comorbidades'])} | Patologia: {val(reg['Patologia'])}",
                f"Diagnóstico nutricional: {val(reg['Diagnostico_Nutricional'])} | Via de terapia nutricional: {val(reg['Via_Terapia_Nutricional'])}",
                f"Total de avaliações no relatório: {len(df_pac)}",
                "====================================================================\n",
            ]
            for _, row in df_pac.iterrows():
                with st.expander(f"{row['Data_Hora']} — {row['Modulo_Calculo']}", expanded=False):
                    a, b, c, d = st.columns(4)
                    a.write(f"**Peso:** {val(row['Peso_kg'])} kg")
                    b.write(f"**Estatura:** {val(row['Estatura_cm'])} cm")
                    c.write(f"**Condição:** {val(row['Faixa_Etaria_Condicao'])}")
                    d.write(f"**Registrado por:** {val(row['Registrado_Por'])}")
                    st.write(
                        f"**VET:** {val(row['VET_kcal_dia'])} kcal/dia | **Proteína:** {val(row['Proteina_g_dia'])} g/dia | "
                        f"**Nec. hídrica:** {val(row['Nec_Hidrica_ml_dia'])} ml/dia ({val(row['Taxa_Hidrica_pct'])}%)"
                    )
                    if not vazio(row["Carboidrato_g_dia"]):
                        st.write(f"**Carboidrato:** {row['Carboidrato_g_dia']} g/dia | **Lipídio:** {val(row['Lipidio_g_dia'])} g/dia")
                    if not vazio(row["TMB_Schofield_kcal"]):
                        st.write(f"**TMB Schofield:** {row['TMB_Schofield_kcal']} kcal/dia | **TMB OMS:** {val(row['TMB_OMS_kcal'])} kcal/dia")
                    if not vazio(row["Fibras_g_dia"]):
                        st.write(f"**Fibras:** {row['Fibras_g_dia']} g/dia | **Açúcar máx.:** {val(row['Acucar_Max_g_dia'])} g/dia")
                    if not vazio(row["Formula_Ou_Item"]):
                        st.write(f"**Fórmula / item:** {row['Formula_Ou_Item']} — {val(row['Qtd_Po_g_Ou_Vol_ml'])}")
                    if not vazio(row["Detalhes_Protocolo"]):
                        st.write(f"**Detalhes do protocolo:** {row['Detalhes_Protocolo']}")
                    if not vazio(row["Observacoes_Clinicas"]):
                        st.write(f"**Observações:** {row['Observacoes_Clinicas']}")
                texto.extend([
                    f"Data/hora: {row['Data_Hora']} | Protocolo: {row['Modulo_Calculo']} | Registrado por: {val(row['Registrado_Por'])}",
                    f"Peso: {val(row['Peso_kg'])} kg | Estatura: {val(row['Estatura_cm'])} cm | Condição: {val(row['Faixa_Etaria_Condicao'])}",
                    f"VET: {val(row['VET_kcal_dia'])} kcal/dia | Proteína: {val(row['Proteina_g_dia'])} g/dia | CHO: {val(row['Carboidrato_g_dia'])} g/dia | LIP: {val(row['Lipidio_g_dia'])} g/dia",
                    f"Nec. hídrica: {val(row['Nec_Hidrica_ml_dia'])} ml/dia | Vol. recebido: {val(row['Vol_Recebido_ml_dia'])} ml/dia ({val(row['Taxa_Hidrica_pct'])}%)",
                    f"TMB Schofield: {val(row['TMB_Schofield_kcal'])} kcal | TMB OMS: {val(row['TMB_OMS_kcal'])} kcal | Fibras: {val(row['Fibras_g_dia'])} g",
                    f"Fórmula/item: {val(row['Formula_Ou_Item'])} ({val(row['Qtd_Po_g_Ou_Vol_ml'])})",
                    f"Detalhes: {val(row['Detalhes_Protocolo'])}",
                    f"Observações: {val(row['Observacoes_Clinicas'])}",
                    "--------------------------------------------------------------------\n",
                ])

            nome_arq = re.sub(r"\W+", "_", reg["Nome_Paciente"])
            b1, b2 = st.columns(2)
            b1.download_button("🖨️ Baixar relatório (.txt para prontuário)", "\n".join(texto),
                               file_name=f"Relatorio_{nome_arq}.txt", mime="text/plain", **larg())
            b2.download_button("📊 Baixar planilha do paciente (.xlsx)", xlsx_bytes(exibir(df_pac), "Paciente"),
                               file_name=f"Historico_{nome_arq}.xlsx", mime=MIME_XLSX, **larg())

        with tab_geral:
            g1, g2 = st.columns(2)
            f_nome = g1.text_input("Filtrar por nome do paciente")
            f_mods = g2.multiselect("Filtrar por protocolo", options=sorted(df_hist["Modulo_Calculo"].dropna().unique()))
            df_f = df_hist.copy()
            if f_nome.strip():
                df_f = df_f[df_f["Nome_Paciente"].str.contains(f_nome.strip(), case=False, na=False, regex=False)]
            if f_mods:
                df_f = df_f[df_f["Modulo_Calculo"].isin(f_mods)]
            df_f = df_f.iloc[::-1]
            colunas_geral = ["Data_Hora", "Nome_Paciente", "Prontuario", "Clinica", "Enfermaria", "Leito"] + [
                c for c in COLUNAS_AVALIACOES if c not in ("Data_Hora", "ID_Paciente")]
            tabela(exibir(df_f, colunas_geral))
            st.download_button("📥 Baixar planilha geral completa (.xlsx)", xlsx_bytes(exibir(df_f, colunas_geral), "Avaliacoes"),
                               file_name="historico_geral_nutrieli.xlsx", mime=MIME_XLSX)

            if E_SUB:
                st.divider()
                st.subheader("Excluir avaliações")
                opcoes_av = {
                    r["ID_Avaliacao"]: f"{r['Data_Hora']} — {r['Nome_Paciente']} ({r['Modulo_Calculo']})"
                    for r in df_hist.iloc[::-1].to_dict("records")
                }
                sel_av = st.multiselect("Avaliações a excluir", list(opcoes_av.keys()), format_func=lambda i: opcoes_av[i])
                if sel_av:
                    conf = st.checkbox("Confirmo a exclusão permanente das avaliações selecionadas", key="chk_del_av")
                    if st.button("Excluir avaliações selecionadas", type="primary", disabled=not conf):
                        excluir_avaliacoes(sel_av)
                        st.success("Avaliações excluídas.")
                        st.rerun()

# =====================================================================
# GERENCIAR USUÁRIOS (MASTER)
# =====================================================================
elif modulo == "Gerenciar Usuários":
    st.header("Gerenciar Usuários")
    if "_ok_usuario" in st.session_state:
        st.success(st.session_state.pop("_ok_usuario"))

    st.subheader("Novo usuário")
    with st.form("form_novo_usuario", clear_on_submit=True):
        u1, u2 = st.columns(2)
        novo_login = u1.text_input("Usuário (login)", placeholder="Letras minúsculas, números, ponto ou hífen")
        novo_nome = u2.text_input("Nome de exibição")
        u3, u4, u5 = st.columns(3)
        novo_perfil = u3.selectbox("Perfil", ["normal", "submaster"], format_func=lambda x: PERFIS[x])
        nova_senha = u4.text_input("Senha (mínimo 6 caracteres)", type="password")
        conf_senha = u5.text_input("Confirmar senha", type="password")
        criar = st.form_submit_button("Criar usuário", type="primary", **larg())
    if criar:
        login_n = novo_login.strip().lower()
        if not re.fullmatch(r"[a-z0-9._-]{3,30}", login_n):
            st.error("O login deve ter de 3 a 30 caracteres, usando letras minúsculas, números, ponto, hífen ou sublinhado.")
        elif login_n in USUARIOS:
            st.error("Já existe um usuário com este login.")
        elif not novo_nome.strip():
            st.error("Informe o nome de exibição.")
        elif len(nova_senha) < 6:
            st.error("A senha precisa ter pelo menos 6 caracteres.")
        elif nova_senha != conf_senha:
            st.error("A confirmação não confere com a senha.")
        else:
            USUARIOS[login_n] = {"nome": novo_nome.strip(), "perfil": novo_perfil, "ativo": True,
                                 "senha": gerar_hash(nova_senha), "criado_em": agora().strftime(FMT_DH)}
            salvar_usuarios(USUARIOS)
            st.session_state["_ok_usuario"] = f"Usuário **{login_n}** criado com o perfil {PERFIS[novo_perfil]}."
            st.rerun()

    st.divider()
    st.subheader("Usuários cadastrados")
    tabela(pd.DataFrame([
        {"Usuário": k, "Nome": v["nome"], "Perfil": PERFIS[v["perfil"]],
         "Situação": "Ativo" if v.get("ativo", True) else "Desativado", "Criado em": v.get("criado_em", "")}
        for k, v in USUARIOS.items()
    ]))

    gerenciaveis = [k for k in USUARIOS if k != "master"]
    if gerenciaveis:
        st.subheader("Alterar usuário")
        alvo = st.selectbox("Usuário", gerenciaveis, index=None, placeholder="Selecione o usuário",
                            format_func=lambda k: f"{k} — {USUARIOS[k]['nome']}")
        if alvo:
            ativo = USUARIOS[alvo].get("ativo", True)
            a1, a2 = st.columns(2)
            with a1:
                with st.form(f"form_reset_{alvo}", clear_on_submit=True):
                    ns = st.text_input("Nova senha (mínimo 6 caracteres)", type="password")
                    if st.form_submit_button("Redefinir senha", **larg()):
                        if len(ns) < 6:
                            st.error("A senha precisa ter pelo menos 6 caracteres.")
                        else:
                            USUARIOS[alvo]["senha"] = gerar_hash(ns)
                            salvar_usuarios(USUARIOS)
                            st.session_state["_ok_usuario"] = f"Senha de **{alvo}** redefinida."
                            st.rerun()
            with a2:
                perfil_novo = st.selectbox("Perfil", ["normal", "submaster"], format_func=lambda x: PERFIS[x],
                                           index=0 if USUARIOS[alvo]["perfil"] == "normal" else 1, key=f"perfil_{alvo}")
                if st.button("Salvar perfil", key=f"sp_{alvo}", **larg()):
                    USUARIOS[alvo]["perfil"] = perfil_novo
                    salvar_usuarios(USUARIOS)
                    st.session_state["_ok_usuario"] = f"Perfil de **{alvo}** atualizado."
                    st.rerun()
                if st.button("Desativar usuário" if ativo else "Reativar usuário", key=f"at_{alvo}", **larg()):
                    USUARIOS[alvo]["ativo"] = not ativo
                    salvar_usuarios(USUARIOS)
                    st.session_state["_ok_usuario"] = f"Usuário **{alvo}** {'desativado' if ativo else 'reativado'}."
                    st.rerun()
            with st.expander("Excluir usuário"):
                conf = st.checkbox(f"Confirmo a exclusão definitiva de {alvo}", key=f"cd_{alvo}")
                if st.button("Excluir usuário", disabled=not conf, key=f"del_{alvo}"):
                    USUARIOS.pop(alvo)
                    salvar_usuarios(USUARIOS)
                    st.session_state["_ok_usuario"] = f"Usuário **{alvo}** excluído."
                    st.rerun()