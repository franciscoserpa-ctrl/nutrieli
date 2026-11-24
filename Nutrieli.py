import os
import hashlib
from datetime import datetime
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Nutrieli", layout="wide")

import sys
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_HISTORICO = os.path.join(BASE_DIR, "historico_pacientes_nutrieli.xlsx")

# =====================================================================
# CONFIGURAÇÃO DE LOGIN E AUTENTICAÇÃO (CREDENCIAIS PADRÃO)
# =====================================================================
def gerar_hash(senha: str) -> str:
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()

USUARIO_PADRAO = "nutrieli"
SENHA_PADRAO_HASH = gerar_hash("nutri@2026")
NOME_EXIBICAO = "Nutrição Clínica"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False


LOGO_CLINICO_SVG = """
<svg width="76" height="76" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="gradClinico" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0D9488"/>
      <stop offset="100%" stop-color="#0369A1"/>
    </linearGradient>
    <linearGradient id="gradFolha" x1="0%" y1="100%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#10B981"/>
      <stop offset="100%" stop-color="#34D399"/>
    </linearGradient>
  </defs>
  <!-- Escudo / Círculo Clínico -->
  <circle cx="50" cy="50" r="46" fill="url(#gradClinico)" stroke="#CCFBF1" stroke-width="3"/>
  <!-- Cruz Médica / Hospitalar sutil ao fundo -->
  <rect x="43" y="20" width="14" height="60" rx="4" fill="#FFFFFF" fill-opacity="0.18"/>
  <rect x="20" y="43" width="60" height="14" rx="4" fill="#FFFFFF" fill-opacity="0.18"/>
  <!-- Folha / Broto Nutricional -->
  <path d="M50 74C50 74 28 62 28 41C28 28 40 23 50 32C60 23 72 28 72 41C72 62 50 74 50 74Z" fill="url(#gradFolha)"/>
  <!-- Linha de Pulso Clínico / Vitalidade -->
  <path d="M24 52H37L42 41L49 61L55 46L59 52H76" stroke="#FFFFFF" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
</svg>
"""

LOGO_MINI_SVG = """
<svg width="42" height="42" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="gradClinicoMini" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0D9488"/>
      <stop offset="100%" stop-color="#0369A1"/>
    </linearGradient>
    <linearGradient id="gradFolhaMini" x1="0%" y1="100%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#10B981"/>
      <stop offset="100%" stop-color="#34D399"/>
    </linearGradient>
  </defs>
  <circle cx="50" cy="50" r="46" fill="url(#gradClinicoMini)" stroke="#CCFBF1" stroke-width="3"/>
  <rect x="43" y="20" width="14" height="60" rx="4" fill="#FFFFFF" fill-opacity="0.18"/>
  <rect x="20" y="43" width="60" height="14" rx="4" fill="#FFFFFF" fill-opacity="0.18"/>
  <path d="M50 74C50 74 28 62 28 41C28 28 40 23 50 32C60 23 72 28 72 41C72 62 50 74 50 74Z" fill="url(#gradFolhaMini)"/>
  <path d="M24 52H37L42 41L49 61L55 46L59 52H76" stroke="#FFFFFF" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
</svg>
"""

def tela_login():
    st.markdown(
        """
        <style>
        .login-header-card {
            background: linear-gradient(135deg, #0F766E 0%, #0369A1 100%);
            padding: 26px 24px 22px 24px;
            border-radius: 16px;
            text-align: center;
            color: white;
            margin-bottom: 18px;
            box-shadow: 0 8px 24px rgba(13, 148, 136, 0.18);
        }
        .login-title {
            font-size: 2rem;
            font-weight: 700;
            margin: 8px 0 2px 0;
            letter-spacing: 0.4px;
            color: #FFFFFF;
        }
        .login-subtitle {
            font-size: 0.92rem;
            color: #CCFBF1;
            margin: 0;
            font-weight: 500;
        }
        .login-badge {
            display: inline-block;
            background: rgba(255, 255, 255, 0.16);
            color: #F0FDFA;
            font-size: 0.75rem;
            padding: 4px 12px;
            border-radius: 999px;
            margin-top: 10px;
            letter-spacing: 0.3px;
        }
        .login-footer {
            text-align: center;
            color: #64748B;
            font-size: 0.8rem;
            margin-top: 14px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<br>", unsafe_allow_html=True)
    col_esq, col_centro, col_dir = st.columns([1, 1.35, 1])

    with col_centro:
        st.markdown(
            f"""
            <div class="login-header-card">
                <div>{LOGO_CLINICO_SVG}</div>
                <div class="login-title">Nutrieli</div>
                <p class="login-subtitle">Nutrição Clínica Hospitalar, Pediátrica e Lactário</p>
                <span class="login-badge">⚕️ Sistema de Avaliação e Prescrição Nutricional</span>
            </div>
            """,
            unsafe_allow_html=True
        )

        with st.form("form_login"):
            st.markdown("##### 🔐 Identificação Profissional")
            usuario = st.text_input("Usuário", placeholder="Digite seu usuário clínico")
            senha = st.text_input("Senha", type="password", placeholder="••••••••")
            entrar = st.form_submit_button("Acessar Prontuário e Protocolos", type="primary", use_container_width=True)

            if entrar:
                if (
                    usuario.strip().lower() == USUARIO_PADRAO.lower()
                    and gerar_hash(senha) == SENHA_PADRAO_HASH
                ):
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("Credenciais inválidas. Verifique usuário e senha.")

        st.markdown(
            '<div class="login-footer">Protocolos baseados em OMS • DRI • ESPGHAN • ASPEN • SBC</div>',
            unsafe_allow_html=True
        )

if not st.session_state.autenticado:
    tela_login()
    st.stop()

# =====================================================================
# FUNÇÕES DE BANCO DE DADOS EM PLANILHA (HISTÓRICO EM LINHAS)
# =====================================================================
COLUNAS_PLANILHA = [
    "Data_Hora",
    "Nome_Paciente",
    "Prontuario_Leito",
    "Modulo_Calculo",
    "Sexo",
    "Faixa_Etaria_Condicao",
    "Peso_kg",
    "Estatura_cm",
    "VET_kcal_dia",
    "Proteina_g_dia",
    "Carboidrato_g_dia",
    "Lipidio_g_dia",
    "Nec_Hidrica_ml_dia",
    "Vol_Recebido_ml_dia",
    "Taxa_Hidrica_pct",
    "Fibras_g_dia",
    "Acucar_Max_g_dia",
    "TMB_Schofield_kcal",
    "TMB_OMS_kcal",
    "Formula_Ou_Item",
    "Qtd_Po_g_Ou_Vol_ml",
    "Detalhes_Protocolo",
    "Observacoes_Clinicas"
]

def carregar_historico() -> pd.DataFrame:
    if os.path.exists(ARQUIVO_HISTORICO):
        return pd.read_excel(ARQUIVO_HISTORICO)
    return pd.DataFrame(columns=COLUNAS_PLANILHA)

def salvar_registro_paciente(dados: dict):
    df_atual = carregar_historico()
    nova_linha = {col: dados.get(col, "") for col in COLUNAS_PLANILHA}
    nova_linha["Data_Hora"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    df_novo = pd.concat([df_atual, pd.DataFrame([nova_linha])], ignore_index=True)
    df_novo.to_excel(ARQUIVO_HISTORICO, index=False)

def excluir_registro_por_indice(indices_remover: list):
    df_atual = carregar_historico()
    df_novo = df_atual.drop(index=indices_remover).reset_index(drop=True)
    df_novo.to_excel(ARQUIVO_HISTORICO, index=False)

def excluir_todos_registros_paciente(nome_pac: str):
    df_atual = carregar_historico()
    df_novo = df_atual[df_atual["Nome_Paciente"] != nome_pac].reset_index(drop=True)
    df_novo.to_excel(ARQUIVO_HISTORICO, index=False)

# =====================================================================
# FUNÇÕES AUXILIARES E TABELAS DE REFERÊNCIA
# =====================================================================
def holliday_segar(peso: float) -> float:
    if peso <= 10:
        return 100 * peso
    elif peso <= 20:
        return 1000 + 50 * (peso - 10)
    else:
        return 1500 + 20 * (peso - 20)

def calcular_tmb(faixa: str, sexo: str, peso: float, altura_m: float):
    if faixa == "0 a 3 anos":
        if sexo == "M":
            schofield = 0.167 * peso + 1517.4 * altura_m - 617.6
            oms = 60.9 * peso - 54
        else:
            schofield = 16.25 * peso + 1023.2 * altura_m - 413.5
            oms = 61.0 * peso - 51
    elif faixa == "3 a 10 anos":
        if sexo == "M":
            schofield = 19.59 * peso + 130.3 * altura_m + 414.9
            oms = 22.7 * peso + 495
        else:
            schofield = 16.969 * peso + 1.618 * altura_m + 371.2
            oms = 22.5 * peso + 499
    else:
        if sexo == "M":
            schofield = 16.25 * peso + 1.372 * altura_m + 515.5
            oms = 17.5 * peso + 651
        else:
            schofield = 8.365 * peso + 4.65 * altura_m + 200
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
    "16-17 anos": (52, 44), "17-18 anos": (50, 44)
}

TABELA_PROTEINA_DRI = {
    "0-6 meses": 1.52, "7-12 meses": 1.50, "1-3 anos": 1.10,
    "4-8 anos": 0.95, "9-13 anos": 0.96, "14-18 anos": 0.85, "19-30 anos": 0.80
}

# =====================================================================
# MENU LATERAL COM NOMES DESCRITIVOS DOS MÓDULOS
# =====================================================================
st.sidebar.markdown(
    f"""
    <div style="display:flex; align-items:center; gap:10px; padding:6px 4px 10px 4px;">
        <div>{LOGO_MINI_SVG}</div>
        <div>
            <div style="font-weight:700; font-size:1.15rem; color:#0F766E; line-height:1.1;">Nutrieli</div>
            <div style="font-size:0.78rem; color:#475569;">⚕️ {NOME_EXIBICAO}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)
if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state.autenticado = False
    st.rerun()

st.sidebar.divider()

with st.sidebar.expander("☁️ Backup / Restaurar Planilha", expanded=False):
    st.caption("Na versão web, baixe a planilha ao final do expediente ou restaure uma planilha anterior:")
    df_side = carregar_historico()
    if not df_side.empty and os.path.exists(ARQUIVO_HISTORICO):
        with open(ARQUIVO_HISTORICO, "rb") as f_bkp:
            st.download_button(
                label="⬇️ Baixar Backup (.xlsx)",
                data=f_bkp,
                file_name="historico_pacientes_nutrieli.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    arquivo_upload = st.file_uploader("Restaurar Planilha (.xlsx)", type=["xlsx"], key="upload_bkp_xlsx")
    if arquivo_upload is not None:
        if st.button("⬆️ Carregar Planilha Enviada", use_container_width=True):
            df_restaurado = pd.read_excel(arquivo_upload)
            df_restaurado.to_excel(ARQUIVO_HISTORICO, index=False)
            st.success("Planilha restaurada com sucesso!")
            st.rerun()

st.sidebar.divider()

modulo = st.sidebar.radio(
    "Selecione o Módulo:",
    [
        "1. Dashboard Geral dos Atendimentos e Preview Completo",
        "2. Cálculo Energético, Proteico, Hídrico e TMB (Pediatria Saudável)",
        "3. Cálculo de Necessidades Nutricionais e Hídricas em Neonatologia",
        "4. Estimativa Energética e Hídrica para Pacientes Neuropatas",
        "5. Cálculo Calórico e Proteico em Oncologia e Cardiologia",
        "6. Estimativa de Estatura Corporal e Protocolo de Desnutrição Grave",
        "7. Cálculo de Diluição e Pesagem de Fórmulas Infantis",
        "8. Cálculo de Composição Nutricional do Lactário",
        "9. Histórico de Pacientes e Relatórios Clínicos"
    ]
)

st.markdown(
    f"""
    <div style="display:flex; align-items:center; gap:14px; margin-bottom:6px;">
        <div>{LOGO_MINI_SVG}</div>
        <div>
            <h1 style="margin:0; padding:0; font-size:2rem; color:#0F766E; line-height:1.1;">Nutrieli</h1>
            <p style="margin:2px 0 0 0; font-size:0.9rem; color:#475569;">
                Sistema Integrado de Nutrição Clínica — Protocolos Hospitalares, Pediatria e Lactário
            </p>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# =====================================================================
# BLOCO FIXO DE IDENTIFICAÇÃO DO PACIENTE (NOS MÓDULOS DE CÁLCULO)
# =====================================================================
if not modulo.startswith("1.") and not modulo.startswith("9."):
    with st.container(border=True):
        st.markdown("**Identificação do Paciente para Registro na Planilha**")
        p1, p2, p3 = st.columns([2, 1, 2])
        nome_paciente = p1.text_input("Nome do Paciente *", placeholder="Digite o nome completo do paciente")
        prontuario = p2.text_input("Prontuário / Leito", placeholder="Ex: 12345 / Enf. B")
        obs_clinica = p3.text_input("Observações Clínicas", placeholder="Ex: Via oral, SNE, restrição hídrica...")

def botao_salvar_registro(registro_dict: dict):
    st.markdown("---")
    if st.button("💾 Salvar Avaliação na Planilha de Pacientes", type="primary"):
        if not nome_paciente.strip():
            st.warning("⚠️ Informe o **Nome do Paciente** no topo da tela antes de salvar.")
        else:
            registro_dict["Nome_Paciente"] = nome_paciente.strip()
            registro_dict["Prontuario_Leito"] = prontuario.strip()
            registro_dict["Observacoes_Clinicas"] = obs_clinica.strip()
            salvar_registro_paciente(registro_dict)
            st.success(f"✅ Avaliação de **{nome_paciente.strip()}** salva com sucesso em `{ARQUIVO_HISTORICO}`!")

# =====================================================================
# MÓDULO 1: DASHBOARD DOS ATENDIMENTOS E PREVIEW COMPLETO
# =====================================================================
if modulo.startswith("1."):
    st.header("📊 Dashboard Geral dos Atendimentos e Preview Completo")
    df_dash = carregar_historico()

    if df_dash.empty:
        st.info("Ainda não há atendimentos salvos na planilha. Registre avaliações nos módulos clínicos (2 a 8) para visualizar os indicadores e o preview completo aqui.")
        
        st.subheader("Pré-visualização da Estrutura da Planilha de Pacientes")
        st.caption("Abaixo estão as colunas que serão preenchidas em linha a cada novo atendimento salvo:")
        st.dataframe(pd.DataFrame(columns=COLUNAS_PLANILHA), use_container_width=True)
    else:
        # Conversões numéricas seguras para métricas e gráficos
        df_dash["Peso_num"] = pd.to_numeric(df_dash["Peso_kg"], errors="coerce")
        df_dash["Taxa_Hidrica_num"] = pd.to_numeric(df_dash["Taxa_Hidrica_pct"], errors="coerce")
        df_dash["Data_Apenas"] = df_dash["Data_Hora"].astype(str).str.slice(0, 10)

        # Filtros rápidos do Dashboard
        with st.expander("🔍 Filtros do Dashboard", expanded=False):
            f_col1, f_col2, f_col3 = st.columns(3)
            modulos_disp = ["Todos"] + sorted(df_dash["Modulo_Calculo"].dropna().unique().tolist())
            mod_filtro = f_col1.selectbox("Filtrar por Protocolo/Módulo:", modulos_disp)
            pac_filtro = f_col2.text_input("Buscar Paciente:")
            data_filtro = f_col3.selectbox("Filtrar por Data:", ["Todas"] + sorted(df_dash["Data_Apenas"].dropna().unique().tolist(), reverse=True))

            if mod_filtro != "Todos":
                df_dash = df_dash[df_dash["Modulo_Calculo"] == mod_filtro]
            if pac_filtro.strip():
                df_dash = df_dash[df_dash["Nome_Paciente"].str.contains(pac_filtro.strip(), case=False, na=False)]
            if data_filtro != "Todas":
                df_dash = df_dash[df_dash["Data_Apenas"] == data_filtro]

        # Linha 1: Indicadores Principais (KPIs)
        total_atend = len(df_dash)
        total_pacientes = df_dash["Nome_Paciente"].nunique()
        media_peso = df_dash["Peso_num"].mean()
        media_hidrica = df_dash["Taxa_Hidrica_num"].mean()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total de Avaliações", f"{total_atend}")
        k2.metric("Pacientes Únicos Atendidos", f"{total_pacientes}")
        k3.metric("Peso Médio Avaliado", f"{media_peso:.2f} kg" if pd.notna(media_peso) else "-")
        k4.metric("Adequação Hídrica Média", f"{media_hidrica:.1f}%" if pd.notna(media_hidrica) else "-")

        st.divider()

        # Linha 2: Gráficos de Distribuição
        g1, g2 = st.columns(2)
        with g1:
            st.subheader("Atendimentos por Módulo / Protocolo")
            contagem_mod = df_dash["Modulo_Calculo"].value_counts()
            st.bar_chart(contagem_mod)

        with g2:
            st.subheader("Evolução Diária de Atendimentos")
            contagem_data = df_dash.groupby("Data_Apenas").size()
            st.line_chart(contagem_data)

        st.divider()

        # Linha 3: Preview Completo dos Atendimentos
        st.subheader("👁️ Preview Completo dos Atendimentos (Visão em Linha e Ficha Detalhada)")
        
        tab_prev_ficha, tab_prev_tabela = st.tabs(["🗂️ Preview Completo por Ficha (Últimos Atendimentos)", "📑 Preview Completo da Planilha (Todas as Colunas)"])

        with tab_prev_ficha:
            df_recentes = df_dash.iloc[::-1].reset_index(drop=True)
            for idx, row in df_recentes.iterrows():
                with st.container(border=True):
                    c_top1, c_top2, c_top3 = st.columns([2, 1.5, 1])
                    c_top1.markdown(f"### 🧑‍⚕️ {row['Nome_Paciente']}")
                    c_top2.markdown(f"**Protocolo:** `{row['Modulo_Calculo']}`")
                    c_top3.markdown(f"🕒 **{row['Data_Hora']}**")

                    d1, d2, d3, d4, d5 = st.columns(5)
                    d1.metric("Prontuário / Leito", f"{row['Prontuario_Leito'] or '-'}")
                    d2.metric("Peso / Estatura", f"{row['Peso_kg'] or '-'} kg | {row['Estatura_cm'] or '-'} cm")
                    d3.metric("VET Calculado", f"{row['VET_kcal_dia'] or '-'} kcal/d")
                    d4.metric("Proteína", f"{row['Proteina_g_dia'] or '-'} g/d")
                    d5.metric("Taxa Hídrica", f"{row['Nec_Hidrica_ml_dia'] or '-'} ml ({row['Taxa_Hidrica_pct'] or '-'}%)")

                    extras = []
                    if pd.notna(row["Faixa_Etaria_Condicao"]) and str(row["Faixa_Etaria_Condicao"]).strip():
                        extras.append(f"**Faixa/Condição:** {row['Faixa_Etaria_Condicao']}")
                    if pd.notna(row["Carboidrato_g_dia"]) and str(row["Carboidrato_g_dia"]).strip():
                        extras.append(f"**CHO:** {row['Carboidrato_g_dia']} g/d | **LIP:** {row['Lipidio_g_dia']} g/d")
                    if pd.notna(row["TMB_Schofield_kcal"]) and str(row["TMB_Schofield_kcal"]).strip():
                        extras.append(f"**TMB Schofield:** {row['TMB_Schofield_kcal']} kcal | **TMB OMS:** {row['TMB_OMS_kcal']} kcal")
                    if pd.notna(row["Fibras_g_dia"]) and str(row["Fibras_g_dia"]).strip():
                        extras.append(f"**Fibras:** {row['Fibras_g_dia']} g/d | **Açúcar Máx:** {row['Acucar_Max_g_dia']} g/d")
                    if pd.notna(row["Formula_Ou_Item"]) and str(row["Formula_Ou_Item"]).strip():
                        extras.append(f"**Fórmula/Lactário:** {row['Formula_Ou_Item']} ({row['Qtd_Po_g_Ou_Vol_ml']})")
                    if pd.notna(row["Detalhes_Protocolo"]) and str(row["Detalhes_Protocolo"]).strip():
                        extras.append(f"**Parâmetros:** {row['Detalhes_Protocolo']}")
                    if pd.notna(row["Observacoes_Clinicas"]) and str(row["Observacoes_Clinicas"]).strip():
                        extras.append(f"**Obs. Clínicas:** {row['Observacoes_Clinicas']}")

                    if extras:
                        st.markdown(" • ".join(extras))

        with tab_prev_tabela:
            st.caption("Visualização integral de todas as 23 colunas registradas em uma única linha por atendimento:")
            df_exibir = df_dash.drop(columns=["Peso_num", "Taxa_Hidrica_num", "Data_Apenas"], errors="ignore").iloc[::-1]
            st.dataframe(df_exibir, use_container_width=True)

# =====================================================================
# MÓDULO 2: CHD SAUDÁVEL & TMB
# =====================================================================
elif modulo.startswith("2."):
    st.header("2. Cálculo Energético, Proteico, Hídrico e TMB (Pediatria Saudável)")
    col1, col2, col3 = st.columns(3)
    with col1:
        sexo = st.selectbox("Sexo", ["M", "F"])
        idade_energia = st.selectbox("Faixa Etária (Energia OMS)", list(TABELA_ENERGIA_OMS.keys()))
        idade_ptn = st.selectbox("Faixa Etária (Proteína DRI 2001)", list(TABELA_PROTEINA_DRI.keys()))
    with col2:
        peso = st.number_input("Peso (kg)", min_value=0.0, value=12.0, step=0.1)
        altura_cm = st.number_input("Estatura (cm)", min_value=0.0, value=85.0, step=0.5)
        vol_recebido = st.number_input("Volume Hídrico Recebido (ml/dia)", min_value=0.0, value=1000.0, step=10.0)
    with col3:
        idade_anos = st.number_input("Idade em anos completos (Fibras/Açúcar)", min_value=0, max_value=18, value=2)
        faixa_tmb = st.selectbox("Faixa Etária (Equações TMB)", ["0 a 3 anos", "3 a 10 anos", "10 a 18 anos"])
        fator_injuria = st.number_input("Fator Injúria / Atividade", min_value=1.0, value=1.0, step=0.05)

    kcal_kg = TABELA_ENERGIA_OMS[idade_energia][0 if sexo == "M" else 1]
    vet_oms = peso * kcal_kg * fator_injuria
    ptn_dia = peso * TABELA_PROTEINA_DRI[idade_ptn]
    nec_hidrica = holliday_segar(peso)
    taxa_hidrica_pct = (vol_recebido * 100 / nec_hidrica) if nec_hidrica > 0 else 0
    taxa_hidrica_ml_kg = (vol_recebido / peso) if peso > 0 else 0
    fibras = idade_anos + 5
    acucar_max_g = (0.25 * vet_oms) / 4
    tmb_schofield, tmb_oms = calcular_tmb(faixa_tmb, sexo, peso, altura_cm / 100)

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("VET Recomendado (OMS)", f"{vet_oms:.1f} kcal/dia", f"{kcal_kg} kcal/kg")
    r2.metric("Necessidade Proteica (DRI)", f"{ptn_dia:.1f} g/dia", f"{TABELA_PROTEINA_DRI[idade_ptn]} g/kg")
    r3.metric("Necessidade Hídrica", f"{nec_hidrica:.0f} ml/dia", f"Recebido: {taxa_hidrica_pct:.1f}% ({taxa_hidrica_ml_kg:.1f} ml/kg)")
    r4.metric("Fibras / Limite Açúcar", f"{fibras} g fibras/dia", f"Máx. açúcar (2-6a): {acucar_max_g:.1f} g")
    st.info(f"**TMB Estimada:** Schofield = **{tmb_schofield:.1f} kcal/dia** | OMS = **{tmb_oms:.1f} kcal/dia**")

    botao_salvar_registro({
        "Modulo_Calculo": "Pediatria Saudável & TMB",
        "Sexo": sexo,
        "Faixa_Etaria_Condicao": f"{idade_energia} | PTN: {idade_ptn}",
        "Peso_kg": round(peso, 2),
        "Estatura_cm": round(altura_cm, 1),
        "VET_kcal_dia": round(vet_oms, 1),
        "Proteina_g_dia": round(ptn_dia, 1),
        "Nec_Hidrica_ml_dia": round(nec_hidrica, 1),
        "Vol_Recebido_ml_dia": round(vol_recebido, 1),
        "Taxa_Hidrica_pct": round(taxa_hidrica_pct, 1),
        "Fibras_g_dia": fibras,
        "Acucar_Max_g_dia": round(acucar_max_g, 1),
        "TMB_Schofield_kcal": round(tmb_schofield, 1),
        "TMB_OMS_kcal": round(tmb_oms, 1),
        "Detalhes_Protocolo": f"Fator injúria: {fator_injuria} | {taxa_hidrica_ml_kg:.1f} ml/kg/dia"
    })

# =====================================================================
# MÓDULO 3: NEONATOLOGIA
# =====================================================================
elif modulo.startswith("3."):
    st.header("3. Cálculo de Necessidades Nutricionais e Hídricas em Neonatologia")
    col1, col2 = st.columns(2)
    with col1:
        tipo_rn = st.radio("Classificação", ["Recém-Nascido a Termo (RNT)", "Recém-Nascido Pré-Termo (RNPT / MBP)"])
        peso_neo = st.number_input("Peso (kg)", min_value=0.0, value=2.5, step=0.05)
        vol_neo = st.number_input("Volume Recebido (ml/dia)", min_value=0.0, value=300.0, step=10.0)

    ptn_str, cho_str, lip_str = "", "", ""
    with col2:
        if tipo_rn == "Recém-Nascido a Termo (RNT)":
            taxa_kcal = st.selectbox("Taxa Calórica RNT (kcal/kg/dia)", [100, 110, 120])
            taxa_agua = 120
            ptn_str = f"{1.5*peso_neo:.1f} a {2.5*peso_neo:.1f}"
            st.write(f"**Proteína Lactentes (1,5–2,5 g/kg):** {ptn_str} g/dia")
        else:
            taxa_kcal = st.selectbox("Taxa Calórica RNPT (kcal/kg/dia)", [110, 120, 130])
            taxa_agua = st.selectbox("Taxa Hídrica RNPT (ml/kg/dia)", [135, 200])
            ptn_str = f"{3.5*peso_neo:.1f} a {4.5*peso_neo:.1f}"
            cho_str = f"{11.6*peso_neo:.1f} a {13.2*peso_neo:.1f}"
            lip_str = f"{4.8*peso_neo:.1f} a {6.6*peso_neo:.1f}"
            st.write("**Macronutrientes RNPT / MBP (Koletzko/ESPGHAN):**")
            st.write(f"- Proteína (3,5–4,5 g/kg): **{ptn_str} g/dia**")
            st.write(f"- Carboidrato (11,6–13,2 g/kg): **{cho_str} g/dia**")
            st.write(f"- Lipídio (4,8–6,6 g/kg): **{lip_str} g/dia**")

    vet_neo = peso_neo * taxa_kcal
    rec_hidrica_neo = peso_neo * taxa_agua
    pct_hidrica = (vol_neo * 100 / rec_hidrica_neo) if rec_hidrica_neo > 0 else 0

    c1, c2 = st.columns(2)
    c1.metric("VET Calculado", f"{vet_neo:.1f} kcal/dia", f"{taxa_kcal} kcal/kg")
    c2.metric("Recomendação Hídrica", f"{rec_hidrica_neo:.0f} ml/dia", f"Adequação: {pct_hidrica:.1f}%")

    botao_salvar_registro({
        "Modulo_Calculo": "Neonatologia (RNT/RNPT)",
        "Faixa_Etaria_Condicao": tipo_rn,
        "Peso_kg": round(peso_neo, 3),
        "VET_kcal_dia": round(vet_neo, 1),
        "Proteina_g_dia": ptn_str,
        "Carboidrato_g_dia": cho_str,
        "Lipidio_g_dia": lip_str,
        "Nec_Hidrica_ml_dia": round(rec_hidrica_neo, 1),
        "Vol_Recebido_ml_dia": round(vol_neo, 1),
        "Taxa_Hidrica_pct": round(pct_hidrica, 1),
        "Detalhes_Protocolo": f"{taxa_kcal} kcal/kg | {taxa_agua} ml/kg"
    })

# =====================================================================
# MÓDULO 4: NEUROPEDIATRIA
# =====================================================================
elif modulo.startswith("4."):
    st.header("4. Estimativa Energética e Hídrica para Pacientes Neuropatas")
    col1, col2 = st.columns(2)
    with col1:
        peso_n = st.number_input("Peso (kg)", min_value=0.0, value=15.0, step=0.1)
        altura_n = st.number_input("Altura (cm)", min_value=0.0, value=100.0, step=0.5)
        condicao_motora = st.selectbox(
            "Condição Motora (Estimativa pela Altura)",
            ["Sem disfunção motora (15 kcal/cm)", "Com disfunção, mas deambula (14 kcal/cm)", "Não deambula (11 kcal/cm)"]
        )
    with col2:
        taxa_idade_n = st.number_input("Taxa Calórica para Idade (kcal/kg)", min_value=0.0, value=80.0, step=1.0)
        tmb_n = st.number_input("TMB / GEB Calculada (kcal)", min_value=0.0, value=600.0, step=10.0)
        vol_n = st.number_input("Volume Hídrico Recebido (ml/dia)", min_value=0.0, value=1000.0, step=10.0)

    fator_alt = 15 if "15" in condicao_motora else (14 if "14" in condicao_motora else 11)
    vet_altura = altura_n * fator_alt
    vet_70 = peso_n * taxa_idade_n * 0.7
    vet_geb = tmb_n * 1.1
    hs_base = holliday_segar(peso_n)
    nec_hidrica_neuro = hs_base * 0.9
    pct_hid_n = (vol_n * 100 / nec_hidrica_neuro) if nec_hidrica_neuro > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("NE pela Altura", f"{vet_altura:.1f} kcal/dia")
    c2.metric("70% Taxa Calórica Idade", f"{vet_70:.1f} kcal/dia")
    c3.metric("GET (GEB x 1,1)", f"{vet_geb:.1f} kcal/dia")
    st.info(f"**Necessidade Hídrica Ajustada (90% H&S):** {nec_hidrica_neuro:.0f} ml/dia | **BPI (+20%):** {hs_base*1.2:.0f} kcal | **MBPI (+50%):** {hs_base*1.5:.0f} kcal")

    botao_salvar_registro({
        "Modulo_Calculo": "Neuropediatria",
        "Faixa_Etaria_Condicao": condicao_motora,
        "Peso_kg": round(peso_n, 2),
        "Estatura_cm": round(altura_n, 1),
        "VET_kcal_dia": round(vet_altura, 1),
        "Nec_Hidrica_ml_dia": round(nec_hidrica_neuro, 1),
        "Vol_Recebido_ml_dia": round(vol_n, 1),
        "Taxa_Hidrica_pct": round(pct_hid_n, 1),
        "Detalhes_Protocolo": f"VET Altura: {vet_altura:.0f} | 70% Idade: {vet_70:.0f} | GET: {vet_geb:.0f} | BPI: {hs_base*1.2:.0f} | MBPI: {hs_base*1.5:.0f}"
    })

# =====================================================================
# MÓDULO 5: ONCOLOGIA E CARDIOLOGIA
# =====================================================================
elif modulo.startswith("5."):
    st.header("5. Cálculo Calórico e Proteico em Oncologia e Cardiologia")
    especialidade = st.radio("Especialidade Clínica", ["Oncologia (ASPEN, 2002)", "Cardiologia (SBC, 2014 / AMB)"])

    if "Oncologia" in especialidade:
        peso_o = st.number_input("Peso Paciente Oncológico (kg)", min_value=0.0, value=20.0, step=0.1)
        faixa_o = st.selectbox("Faixa Etária (ASPEN)", [
            "0 a 1 ano (90–120 kcal/kg)", "1 a 7 anos (75–90 kcal/kg)",
            "7 a 12 anos (60–75 kcal/kg)", "12 a 18 anos (30–60 kcal/kg)",
            "18 a 25 anos (25–30 kcal/kg)"
        ])
        acrescimo_desnut = st.slider("Acréscimo por Perda de Peso / Desnutrição (%)", 0, 50, 0)
        ranges = {"0 a 1": (90, 120), "1 a 7": (75, 90), "7 a 12": (60, 75), "12 a 18": (30, 60), "18 a 25": (25, 30)}
        k_min, k_max = next(v for k, v in ranges.items() if k in faixa_o)
        fator_d = 1 + (acrescimo_desnut / 100)
        vet_str = f"{peso_o * k_min * fator_d:.1f} a {peso_o * k_max * fator_d:.1f}"
        st.success(f"**VET Estimado:** {vet_str} kcal/dia")

        botao_salvar_registro({
            "Modulo_Calculo": "Oncologia (ASPEN)",
            "Faixa_Etaria_Condicao": faixa_o,
            "Peso_kg": round(peso_o, 2),
            "VET_kcal_dia": vet_str,
            "Detalhes_Protocolo": f"Acréscimo desnutrição: +{acrescimo_desnut}%"
        })
    else:
        peso_c = st.number_input("Peso Paciente Cardiopata (kg)", min_value=0.0, value=8.0, step=0.1)
        gravidade = st.selectbox("Classificação (0 a 2 anos - AMB/CFM)", [
            "Leve a Moderado (120–150 kcal/kg)", "Moderado a Grave (170–180 kcal/kg)"
        ])
        c_min, c_max = (120, 150) if "Leve" in gravidade else (170, 180)
        vet_c_str = f"{peso_c * c_min:.1f} a {peso_c * c_max:.1f}"
        ptn_c_str = f"{peso_c * 3:.1f} a {peso_c * 4:.1f}"
        st.success(f"**VET Cardiopata:** {vet_c_str} kcal/dia | **Proteína (3,0–4,0 g/kg):** {ptn_c_str} g/dia")

        botao_salvar_registro({
            "Modulo_Calculo": "Cardiologia (SBC/AMB)",
            "Faixa_Etaria_Condicao": gravidade,
            "Peso_kg": round(peso_c, 2),
            "VET_kcal_dia": vet_c_str,
            "Proteina_g_dia": ptn_c_str,
            "Detalhes_Protocolo": "Protocolo Cardiopatia 0 a 2 anos"
        })

# =====================================================================
# MÓDULO 6: ESTIMATIVA DE ESTATURA E DESNUTRIÇÃO GRAVE
# =====================================================================
elif modulo.startswith("6."):
    st.header("6. Estimativa de Estatura Corporal e Protocolo de Desnutrição Grave")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Estimativa de Estatura (Stevenson & Chumlea)")
        medida_tipo = st.selectbox("Segmento Corporal", [
            "Comprimento Superior do Braço (Stevenson)",
            "Comprimento Tibial (Stevenson)",
            "Altura do Joelho Pediátrica (Stevenson)",
            "Altura do Joelho Adulto/Idoso (Chumlea)"
        ])
        medida_cm = st.number_input("Medida do Segmento (cm)", min_value=0.0, value=30.0, step=0.1)
        if "Braço" in medida_tipo:
            est = 4.35 * medida_cm + 21.8
        elif "Tibial" in medida_tipo:
            est = 3.26 * medida_cm + 30.8
        elif "Pediátrica" in medida_tipo:
            est = 2.69 * medida_cm + 24.2
        else:
            sexo_ch = st.radio("Sexo (Chumlea)", ["M", "F"])
            idade_ch = st.number_input("Idade (anos)", min_value=18, max_value=110, value=65)
            est = (64.19 - 0.04 * idade_ch + 2.02 * medida_cm) if sexo_ch == "M" else (84.88 - 0.24 * idade_ch + 1.83 * medida_cm)
        st.metric("Estatura Estimada", f"{est:.1f} cm")

    with col2:
        st.subheader("Protocolo de Desnutrição Grave (OMS, 2005)")
        peso_desn = st.number_input("Peso da Criança (kg)", min_value=0.0, value=6.0, step=0.1)
        fase_atual = st.selectbox("Fase Atual do Protocolo", ["Fase 1 (2 dias)", "Fase 2 (3 dias)", "Fase 3 (3 dias)", "Fase 4 (3 dias*)", "Fase 5 (3 dias*)"])
        fases_df = pd.DataFrame({
            "Fase": ["Fase 1", "Fase 2", "Fase 3", "Fase 4", "Fase 5"],
            "VET (kcal/dia)": [75 * peso_desn, 100 * peso_desn, 150 * peso_desn, 200 * peso_desn, 250 * peso_desn],
            "Proteína (g/dia)": [1.5 * peso_desn, 2.5 * peso_desn, 3.5 * peso_desn, 4.5 * peso_desn, 5.5 * peso_desn],
            "Evolução": ["2 dias", "3 dias", "3 dias", "3 dias*", "3 dias*"]
        })
        st.dataframe(fases_df, use_container_width=True)

    idx_fase = int(fase_atual.split()[1]) - 1
    vet_fase = fases_df.loc[idx_fase, "VET (kcal/dia)"]
    ptn_fase = fases_df.loc[idx_fase, "Proteína (g/dia)"]

    botao_salvar_registro({
        "Modulo_Calculo": "Antropometria & Desnutrição Grave",
        "Faixa_Etaria_Condicao": f"{medida_tipo} | {fase_atual}",
        "Peso_kg": round(peso_desn, 2),
        "Estatura_cm": round(est, 1),
        "VET_kcal_dia": round(vet_fase, 1),
        "Proteina_g_dia": round(ptn_fase, 1),
        "Detalhes_Protocolo": f"Segmento: {medida_cm} cm -> Estatura est.: {est:.1f} cm"
    })

# =====================================================================
# MÓDULO 7: DILUIÇÃO PADRÃO DE FÓRMULAS INFANTIS
# =====================================================================
elif modulo.startswith("7."):
    st.header("7. Cálculo de Diluição e Pesagem de Fórmulas Infantis")
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
        "INFATRINI": 20.0, "INFATRINI 1.5": 30.0, "PEPTAMEN JUNIOR PÓ": 26.2
    }
    formula_sel = st.selectbox("Escolha a Fórmula", list(FORMULAS_G_100ML.keys()))
    vol_desejado = st.number_input("Volume Prescrito (ml)", min_value=10, max_value=2000, value=150, step=10)
    gramas_po = (vol_desejado * FORMULAS_G_100ML[formula_sel]) / 100
    st.metric(f"Quantidade de Pó ({formula_sel})", f"{gramas_po:.2f} g", f"Concentração padrão: {FORMULAS_G_100ML[formula_sel]} g / 100 ml")

    botao_salvar_registro({
        "Modulo_Calculo": "Diluição de Fórmulas Infantis",
        "Vol_Recebido_ml_dia": vol_desejado,
        "Formula_Ou_Item": formula_sel,
        "Qtd_Po_g_Ou_Vol_ml": f"{gramas_po:.2f} g de pó para {vol_desejado} ml",
        "Detalhes_Protocolo": f"Diluição padrão: {FORMULAS_G_100ML[formula_sel]}%"
    })

# =====================================================================
# MÓDULO 8: COMPOSIÇÃO NUTRICIONAL DO LACTÁRIO
# =====================================================================
elif modulo.startswith("8."):
    st.header("8. Cálculo de Composição Nutricional do Lactário")
    ITENS_LACTARIO = {
        "FL Partida / Seguimento Padrão (15%)": (70.0, 0.9, 7.3, 4.2),
        "Alfamino a 15%": (70.0, 1.9, 7.9, 3.4),
        "Mingau / Mucilon / Papa / Suco (Base Lactário)": (70.0, 0.9, 7.3, 4.2),
    }
    item = st.selectbox("Preparação / Item", list(ITENS_LACTARIO.keys()))
    qtd = st.number_input("Quantidade (ml ou g)", min_value=0.0, value=200.0, step=10.0)
    kcal_100, ptn_100, cho_100, lip_100 = ITENS_LACTARIO[item]

    kcal_tot = kcal_100 * qtd / 100
    ptn_tot = ptn_100 * qtd / 100
    cho_tot = cho_100 * qtd / 100
    lip_tot = lip_100 * qtd / 100

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Calorias", f"{kcal_tot:.1f} kcal")
    c2.metric("Proteína", f"{ptn_tot:.2f} g")
    c3.metric("Carboidrato", f"{cho_tot:.2f} g")
    c4.metric("Lipídio", f"{lip_tot:.2f} g")

    botao_salvar_registro({
        "Modulo_Calculo": "Composição Nutricional Lactário",
        "VET_kcal_dia": round(kcal_tot, 1),
        "Proteina_g_dia": round(ptn_tot, 2),
        "Carboidrato_g_dia": round(cho_tot, 2),
        "Lipidio_g_dia": round(lip_tot, 2),
        "Formula_Ou_Item": item,
        "Qtd_Po_g_Ou_Vol_ml": f"{qtd} ml/g"
    })

# =====================================================================
# MÓDULO 9: HISTÓRICO DE PACIENTES E RELATÓRIOS
# =====================================================================
elif modulo.startswith("9."):
    st.header("9. Histórico de Pacientes e Relatórios Clínicos")
    df_hist = carregar_historico()

    if df_hist.empty:
        st.info("Nenhum atendimento foi registrado ainda. Preencha o nome do paciente em qualquer módulo de cálculo e clique em **Salvar Avaliação na Planilha de Pacientes**.")
    else:
        tab_relatorio, tab_tabela = st.tabs([
            "📄 Relatório Completo por Paciente",
            "📋 Planilha Geral de Histórico (Todas as Linhas)"
        ])

        with tab_relatorio:
            pacientes_unicos = sorted(df_hist["Nome_Paciente"].dropna().unique().tolist())
            col_sel1, col_sel2 = st.columns([2, 1])
            paciente_sel = col_sel1.selectbox("Selecione o Paciente para Gerar o Relatório:", pacientes_unicos)

            df_pac = df_hist[df_hist["Nome_Paciente"] == paciente_sel].copy()
            modulos_pac = ["Todos"] + sorted(df_pac["Modulo_Calculo"].dropna().unique().tolist())
            filtro_mod_pac = col_sel2.selectbox("Filtrar Protocolo deste Paciente:", modulos_pac)

            if filtro_mod_pac != "Todos":
                df_pac = df_pac[df_pac["Modulo_Calculo"] == filtro_mod_pac]

            df_pac = df_pac.iloc[::-1].reset_index(drop=True)

            # Resumo Clínico Consolidado do Paciente
            st.subheader(f"📋 Relatório Individual — {paciente_sel}")
            ultimo_reg = df_pac.iloc[0]
            pront_atual = ultimo_reg["Prontuario_Leito"] if pd.notna(ultimo_reg["Prontuario_Leito"]) and str(ultimo_reg["Prontuario_Leito"]).strip() else "-"
            peso_atual = ultimo_reg["Peso_kg"] if pd.notna(ultimo_reg["Peso_kg"]) and str(ultimo_reg["Peso_kg"]).strip() else "-"
            est_atual = ultimo_reg["Estatura_cm"] if pd.notna(ultimo_reg["Estatura_cm"]) and str(ultimo_reg["Estatura_cm"]).strip() else "-"
            vet_atual = ultimo_reg["VET_kcal_dia"] if pd.notna(ultimo_reg["VET_kcal_dia"]) and str(ultimo_reg["VET_kcal_dia"]).strip() else "-"
            ptn_atual = ultimo_reg["Proteina_g_dia"] if pd.notna(ultimo_reg["Proteina_g_dia"]) and str(ultimo_reg["Proteina_g_dia"]).strip() else "-"

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total de Avaliações", f"{len(df_pac)}")
            m2.metric("Prontuário / Leito", f"{pront_atual}")
            m3.metric("Último Peso / Estatura", f"{peso_atual} kg | {est_atual} cm")
            m4.metric("Último VET", f"{vet_atual} kcal/d")
            m5.metric("Última Proteína", f"{ptn_atual} g/d")

            st.divider()

            # Tabela Resumo da Evolução do Paciente
            st.markdown("**Resumo Cronológico em Linha (Paciente Selecionado)**")
            st.dataframe(df_pac, use_container_width=True)

            # Detalhamento Ficha a Ficha + Montagem do Arquivo de Impressão
            st.markdown("**Histórico Clínico Detalhado**")
            texto_relatorio = [
                "====================================================================",
                "NUTRIELI - RELATÓRIO INDIVIDUAL DE AVALIAÇÃO E PRESCRIÇÃO NUTRICIONAL",
                f"Data de Emissão: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                f"Paciente: {paciente_sel}",
                f"Prontuário / Leito (Último registro): {pront_atual}",
                f"Total de Avaliações no Relatório: {len(df_pac)}",
                "====================================================================\n"
            ]

            for idx, row in df_pac.iterrows():
                with st.expander(f"📅 {row['Data_Hora']} — {row['Modulo_Calculo']}", expanded=True):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.write(f"**Prontuário/Leito:** {row['Prontuario_Leito'] or '-'}")
                    c2.write(f"**Peso:** {row['Peso_kg'] or '-'} kg")
                    c3.write(f"**Estatura:** {row['Estatura_cm'] or '-'} cm")
                    c4.write(f"**Condição:** {row['Faixa_Etaria_Condicao'] or '-'}")

                    st.write(
                        f"**VET:** {row['VET_kcal_dia'] or '-'} kcal/dia | "
                        f"**Proteína:** {row['Proteina_g_dia'] or '-'} g/dia | "
                        f"**Nec. Hídrica:** {row['Nec_Hidrica_ml_dia'] or '-'} ml/dia ({row['Taxa_Hidrica_pct'] or '-'}%)"
                    )
                    if pd.notna(row["Carboidrato_g_dia"]) and str(row["Carboidrato_g_dia"]).strip():
                        st.write(f"**Carboidrato:** {row['Carboidrato_g_dia']} g/dia | **Lipídio:** {row['Lipidio_g_dia']} g/dia")
                    if pd.notna(row["TMB_Schofield_kcal"]) and str(row["TMB_Schofield_kcal"]).strip():
                        st.write(f"**TMB Schofield:** {row['TMB_Schofield_kcal']} kcal/dia | **TMB OMS:** {row['TMB_OMS_kcal']} kcal/dia")
                    if pd.notna(row["Fibras_g_dia"]) and str(row["Fibras_g_dia"]).strip():
                        st.write(f"**Fibras:** {row['Fibras_g_dia']} g/dia | **Açúcar Máx.:** {row['Acucar_Max_g_dia']} g/dia")
                    if pd.notna(row["Formula_Ou_Item"]) and str(row["Formula_Ou_Item"]).strip():
                        st.write(f"**Fórmula / Lactário:** {row['Formula_Ou_Item']} — {row['Qtd_Po_g_Ou_Vol_ml']}")
                    if pd.notna(row["Detalhes_Protocolo"]) and str(row["Detalhes_Protocolo"]).strip():
                        st.write(f"**Detalhes do Protocolo:** {row['Detalhes_Protocolo']}")
                    if pd.notna(row["Observacoes_Clinicas"]) and str(row["Observacoes_Clinicas"]).strip():
                        st.write(f"**Observações Clínicas:** {row['Observacoes_Clinicas']}")

                texto_relatorio.extend([
                    f"Data/Hora: {row['Data_Hora']} | Módulo: {row['Modulo_Calculo']}",
                    f"Prontuário/Leito: {row['Prontuario_Leito']} | Peso: {row['Peso_kg']} kg | Estatura: {row['Estatura_cm']} cm",
                    f"Condição/Faixa: {row['Faixa_Etaria_Condicao']}",
                    f"VET: {row['VET_kcal_dia']} kcal/dia | Proteína: {row['Proteina_g_dia']} g/dia | CHO: {row['Carboidrato_g_dia']} g/d | LIP: {row['Lipidio_g_dia']} g/d",
                    f"Nec. Hídrica: {row['Nec_Hidrica_ml_dia']} ml/dia | Vol. Recebido: {row['Vol_Recebido_ml_dia']} ml/dia ({row['Taxa_Hidrica_pct']}%)",
                    f"TMB Schofield: {row['TMB_Schofield_kcal']} kcal | TMB OMS: {row['TMB_OMS_kcal']} kcal | Fibras: {row['Fibras_g_dia']} g",
                    f"Fórmula/Item: {row['Formula_Ou_Item']} ({row['Qtd_Po_g_Ou_Vol_ml']})",
                    f"Detalhes: {row['Detalhes_Protocolo']}",
                    f"Observações Clínicas: {row['Observacoes_Clinicas']}",
                    "--------------------------------------------------------------------\n"
                ])

            # Botões de Exportação Individual do Paciente (.TXT e .XLSX)
            import io
            buffer_pac_xlsx = io.BytesIO()
            with pd.ExcelWriter(buffer_pac_xlsx, engine="openpyxl") as writer:
                df_pac.to_excel(writer, index=False, sheet_name="Relatorio_Paciente")
            buffer_pac_xlsx.seek(0)

            b_col1, b_col2 = st.columns(2)
            b_col1.download_button(
                label="🖨️ Baixar Relatório do Paciente (.txt para Prontuário)",
                data="\n".join(texto_relatorio),
                file_name=f"Relatorio_{paciente_sel.replace(' ', '_')}.txt",
                mime="text/plain",
                use_container_width=True
            )
            b_col2.download_button(
                label="📊 Baixar Planilha Individual do Paciente (.xlsx)",
                data=buffer_pac_xlsx,
                file_name=f"Historico_{paciente_sel.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.markdown("---")
            with st.expander(f"🗑️ Excluir Paciente da Planilha ({paciente_sel})", expanded=False):
                st.warning(f"Atenção: Esta ação removerá permanentemente todas as avaliações de **{paciente_sel}** da planilha.")
                confirmar_del_pac = st.checkbox(
                    f"Confirmo que desejo apagar todos os registros de {paciente_sel}",
                    key="chk_del_pac_rel"
                )
                if st.button("🗑️ Apagar Todos os Registros deste Paciente", type="secondary", disabled=not confirmar_del_pac, key="btn_del_pac_rel"):
                    excluir_todos_registros_paciente(paciente_sel)
                    st.success(f"Paciente **{paciente_sel}** removido da planilha com sucesso!")
                    st.rerun()

        with tab_tabela:
            col_f1, col_f2 = st.columns(2)
            filtro_nome = col_f1.text_input("Filtrar por Nome do Paciente:")
            filtro_mod = col_f2.multiselect("Filtrar por Módulo:", options=df_hist["Modulo_Calculo"].dropna().unique())

            df_filtrado = df_hist.copy()
            if filtro_nome:
                df_filtrado = df_filtrado[df_filtrado["Nome_Paciente"].str.contains(filtro_nome, case=False, na=False)]
            if filtro_mod:
                df_filtrado = df_filtrado[df_filtrado["Modulo_Calculo"].isin(filtro_mod)]

            st.dataframe(df_filtrado, use_container_width=True)

            with open(ARQUIVO_HISTORICO, "rb") as f_excel:
                st.download_button(
                    label="📥 Baixar Planilha Geral Completa (.xlsx)",
                    data=f_excel,
                    file_name=ARQUIVO_HISTORICO,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.divider()
            st.subheader("🗑️ Gerenciar / Apagar Registros da Planilha")
            modo_exclusao = st.radio(
                "Escolha o tipo de exclusão:",
                ["Apagar atendimento(s) específico(s) (linha individual)", "Apagar paciente completo (todas as linhas do paciente)"],
                horizontal=True
            )

            if modo_exclusao == "Apagar atendimento(s) específico(s) (linha individual)":
                opcoes_linhas = {
                    idx: f"Linha {idx + 1} | {row['Data_Hora']} — {row['Nome_Paciente']} ({row['Modulo_Calculo']})"
                    for idx, row in df_hist.iterrows()
                }
                linhas_sel = st.multiselect(
                    "Selecione o(s) atendimento(s) que deseja apagar:",
                    options=list(opcoes_linhas.keys()),
                    format_func=lambda x: opcoes_linhas[x]
                )
                if linhas_sel:
                    confirmar_linhas = st.checkbox("Confirmo a exclusão permanente da(s) linha(s) selecionada(s)", key="chk_del_linhas")
                    if st.button("🗑️ Apagar Linha(s) Selecionada(s)", type="primary", disabled=not confirmar_linhas):
                        excluir_registro_por_indice(linhas_sel)
                        st.success("Atendimento(s) excluído(s) com sucesso da planilha!")
                        st.rerun()
            else:
                lista_pacs_del = sorted(df_hist["Nome_Paciente"].dropna().unique().tolist())
                pac_del_sel = st.selectbox("Selecione o paciente que deseja apagar completamente:", lista_pacs_del, key="sel_pac_del_tab")
                qtd_reg_pac = len(df_hist[df_hist["Nome_Paciente"] == pac_del_sel])
                st.caption(f"Este paciente possui **{qtd_reg_pac}** registro(s) salvo(s) na planilha.")
                confirmar_pac_tab = st.checkbox(f"Confirmo que desejo excluir definitivamente {pac_del_sel}", key="chk_del_pac_tab")
                if st.button("🗑️ Apagar Paciente da Planilha", type="primary", disabled=not confirmar_pac_tab, key="btn_del_pac_tab"):
                    excluir_todos_registros_paciente(pac_del_sel)
                    st.success(f"Todos os registros de **{pac_del_sel}** foram removidos com sucesso!")
                    st.rerun()
