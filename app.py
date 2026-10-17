import streamlit as st
import pandas as pd
import os
from datetime import datetime

EXCEL_FILE = "RelacaoVeiculos_APP.xlsx"
BACKUP_DIR = "backups"

st.set_page_config(page_title="Gestão de Vagas e Veículos", layout="wide")

def calcular_setor_vaga(comprimento, peso):
    """Aplica a regra exata da fórmula do Excel baseada em Comprimento e Peso"""
    try:
        comp = float(comprimento)
        p = float(peso) if peso != "" and peso is not None else 0.0
    except (ValueError, TypeError):
        return ""

    if comp <= 4.10 and p <= 1500:
        return "PALETE"
    elif comp <= 4.20:
        return "SETOR: G / H"
    elif comp <= 4.40:
        return "SETOR: I / J / K / L / M"
    elif comp <= 4.50:
        return "SETOR: A / C"
    else:
        return "SETOR: B / E / F"

@st.cache_data(ttl=1)
def load_data():
    if not os.path.exists(EXCEL_FILE):
        st.error(f"Arquivo {EXCEL_FILE} não encontrado na pasta!")
        return None, None
    xls = pd.ExcelFile(EXCEL_FILE)
    
    # Carrega Veículos
    df_veic = pd.read_excel(xls, sheet_name="CadastroVeiculos")
    df_veic = df_veic.dropna(how="all").reset_index(drop=True)
    
    # Carrega Vagas
    df_vagas = pd.read_excel(xls, sheet_name="CadastroVaga")
    
    if "Bloco" not in df_vagas.columns:
        for i in range(len(df_vagas)):
            if "Bloco" in df_vagas.iloc[i].values:
                df_vagas = pd.read_excel(xls, sheet_name="CadastroVaga", skiprows=i+1)
                break

    cols_vagas = [
        "Bloco", "Apto", "Proprietario_Morador", "CPF", "Contato", "Vinculo",
        "Modelo", "Marca", "Placa", "Cor", "Comprimento", "TAG", "Data",
        "Alugada", "De_Quem", "Locatario_CPF", "Locatario_Contato", "Observacao", "Local"
    ]
    
    # Adiciona colunas faltantes automaticamente se não existirem no Excel
    for col in cols_vagas:
        if col not in df_vagas.columns:
            df_vagas[col] = ""
    
    df_vagas = df_vagas[cols_vagas]
    df_vagas = df_vagas.dropna(how="all").reset_index(drop=True)
    df_vagas = df_vagas.astype(str).replace(["nan", "None", "NaN"], "")
    
    return df_veic, df_vagas

def save_data(df_veic, df_vagas):
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
        
    if os.path.exists(EXCEL_FILE):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(BACKUP_DIR, f"RelacaoVeiculos_backup_{timestamp}.xlsx")
        try:
            with open(EXCEL_FILE, "rb") as f_in, open(backup_file, "wb") as f_out:
                f_out.write(f_in.read())
        except Exception:
            pass

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_veic.to_excel(writer, sheet_name="CadastroVeiculos", index=False)
        df_vagas.to_excel(writer, sheet_name="CadastroVaga", index=False)
    st.cache_data.clear()

df_veic, df_vagas = load_data()

st.title("🚗 GESTÃO DE VAGAS E VEÍCULOS")

if df_veic is not None and df_vagas is not None:
    col_mod = [c for c in df_veic.columns if "Modelo" in c][0] if any("Modelo" in c for c in df_veic.columns) else df_veic.columns[0]
    col_mar = [c for c in df_veic.columns if "Marca" in c][0] if any("Marca" in c for c in df_veic.columns) else df_veic.columns[1]
    col_pes = [c for c in df_veic.columns if "Peso" in c][0] if any("Peso" in c for c in df_veic.columns) else df_veic.columns[2]
    col_lar = [c for c in df_veic.columns if "Largura" in c][0] if any("Largura" in c for c in df_veic.columns) else (df_veic.columns[3] if len(df_veic.columns) > 3 else "Largura")
    col_com = [c for c in df_veic.columns if "Comprimento" in c][0] if any("Comprimento" in c for c in df_veic.columns) else df_veic.columns[4]

    tab_vagas, tab_veic = st.tabs(["📋 GESTÃO DE VAGAS & MORADORES", "🚘 CADASTRO DE VEÍCULOS"])

    # =========================================================================
    # ---------------- ABA 1: VAGAS E MORADORES -------------------------------
    # =========================================================================
    with tab_vagas:
        sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs(["🔍 CONSULTAR", "➕ NOVO CADASTRO", "✏️ EDITAR", "🗑️ EXCLUIR"])
        
        # --- SUB-ABA 1: CONSULTAR ---
        with sub_tab1:
            st.subheader("🔍 CONSULTAR MORADOR / VAGA")
            
            if not df_vagas.empty:
                opcoes_dict = {
                    f"Bloco {row['Bloco']} - Apto {row['Apto']} | {row['Proprietario_Morador']} | Placa: {row['Placa']}": idx
                    for idx, row in df_vagas.iterrows()
                }
                
                # Inicializa chave de estado para controle do selectbox
                if "sb_consultar_key" not in st.session_state:
                    st.session_state["sb_consultar_key"] = "sb_consultar_0"

                c_sel, c_btn = st.columns([5, 1])
                
                with c_btn:
                    st.write("") # Espaçamento para alinhar com a caixa
                    st.write("")
                    if st.button("🗑️ Limpar Busca", key="btn_limpar_busca"):
                        # Altera a chave para resetar o selectbox
                        st.session_state["sb_consultar_key"] = f"sb_consultar_{datetime.now().timestamp()}"
                        st.rerun()

                with c_sel:
                    selecionado = st.selectbox(
                        "Pesquise ou selecione um morador:", 
                        [""] + list(opcoes_dict.keys()), 
                        key=st.session_state["sb_consultar_key"]
                    )
                
                if selecionado and selecionado in opcoes_dict:
                    idx_sel = opcoes_dict[selecionado]
                    m = df_vagas.loc[idx_sel].copy()
                    
                    mod_morador = str(m.get("Modelo", "")).strip().upper()
                    info_veic = df_veic[df_veic[col_mod].astype(str).str.strip().str.upper() == mod_morador]
                    
                    if not info_veic.empty:
                        v_row = info_veic.iloc[0]
                        marca_dinamica = str(v_row[col_mar])
                        comp_dinamico = str(v_row[col_com])
                        peso_dinamico = v_row[col_pes]
                        setor_dinamico = calcular_setor_vaga(comp_dinamico, peso_dinamico)
                    else:
                        marca_dinamica = m.get("Marca", "-")
                        comp_dinamico = m.get("Comprimento", "-")
                        setor_dinamico = m.get("Local", "-")

                    st.divider()
                    st.markdown(f"### 👤 Ficha do Morador: **{m.get('Proprietario_Morador', '')}**")
                    
                    c1, c2, c3 = st.columns(3)
                    
                    with c1:
                        st.info("**🏠 Unidade / Identificação**")
                        st.write(f"**Bloco:** {m.get('Bloco', '-')}")
                        st.write(f"**Apto:** {m.get('Apto', '-')}")
                        st.write(f"**Proprietário / Morador:** {m.get('Proprietario_Morador', '-')}")
                        st.write(f"**CPF:** {m.get('CPF', '-')}")
                        st.write(f"**Contato:** {m.get('Contato', '-')}")
                        st.write(f"**Vínculo:** {m.get('Vinculo', '-')}")
                        
                    with c2:
                        st.success("**🚘 Dados do Veículo**")
                        st.write(f"**Modelo:** {m.get('Modelo', '-')}")
                        st.write(f"**Marca:** {marca_dinamica}")
                        st.write(f"**Placa:** {m.get('Placa', '-')}")
                        st.write(f"**Cor:** {m.get('Cor', '-')}")
                        st.write(f"**Comprimento:** {comp_dinamico} m")
                        st.markdown("---")
                        st.write(f"**Setor:** {setor_dinamico}")
                        
                    with c3:
                        st.warning("**🏷️ Vaga & TAG**")
                        tag_val = str(m.get('TAG', '')).strip()
                        data_val = str(m.get('Data', '')).strip()
                        alug_val = str(m.get('Alugada', '')).strip()

                        st.write(f"**TAG:** {tag_val if tag_val and tag_val != 'nan' else '-'}")
                        st.write(f"**Data:** {data_val if data_val and data_val != 'nan' else '-'}")
                        st.write(f"**Vaga Alugada?:** {alug_val if alug_val and alug_val != 'nan' else '-'}")
                        
                        if alug_val.upper() == "SIM":
                            st.markdown("---")
                            st.write(f"**Locador:** {m.get('De_Quem', '-')}")
                            st.write(f"**CPF Locador:** {m.get('Locatario_CPF', '-')}")
                            st.write(f"**Contato Locador:** {m.get('Locatario_Contato', '-')}")

                        obs_val = str(m.get('Observacao', '')).strip()
                        if obs_val and obs_val not in ["-", "nan", "None"]:
                            st.markdown("---")
                            st.write(f"**Observação:** {obs_val}")
            
            st.divider()
            with st.expander("📊 Clique para ver a Tabela Completa de Moradores"):
                st.dataframe(df_vagas, use_container_width=True)

        # --- SUB-ABA 2: NOVO CADASTRO ---
        with sub_tab2:
            st.subheader("➕ ADICIONAR NOVO MORADOR / VAGA")
            
            modelos_disp = sorted(df_veic[col_mod].dropna().astype(str).unique().tolist())

            col1, col2, col3 = st.columns(3)
            with col1:
                bloco = st.text_input("Bloco", key="add_bloco")
                apto = st.text_input("Apto", key="add_apto")
                morador = st.text_input("Proprietário / Morador", key="add_morador")
                cpf = st.text_input("CPF", key="add_cpf")
            with col2:
                contato = st.text_input("Contato", key="add_contato")
                vinculo = st.selectbox("Vínculo", ["PROPRIETARIO(A)", "INQUILINO(A)", "MORADOR(A)"], key="add_vinc")
                modelo_sel = st.selectbox("Modelo do Veículo (Busca da Base)", [""] + modelos_disp, key="add_mod")
                placa = st.text_input("Placa", key="add_placa")
            with col3:
                cor = st.text_input("Cor", key="add_cor")
                tag = st.text_input("TAG", key="add_tag")
                alugada = st.selectbox("Vaga Alugada?", ["NÃO", "SIM"], key="add_alugada")
                local_manual = st.text_input("Local / Setor (Deixe em branco para calcular automático)", key="add_local")

            locatario_nome = ""
            locatario_cpf = ""
            locatario_contato = ""
            
            if alugada == "SIM":
                st.divider()
                st.markdown("#### 🔑 Informações do Locador (Quem está alugando a vaga)")
                cl1, cl2, cl3 = st.columns(3)
                locatario_nome = cl1.text_input("Nome do Locador", key="add_loc_nome")
                locatario_cpf = cl2.text_input("CPF do Locador", key="add_loc_cpf")
                locatario_contato = cl3.text_input("Contato do Locador", key="add_loc_contato")

            st.divider()
            observacao = st.text_area("Observação (Informações adicionais do cadastro)", key="add_obs")

            if st.button("Salvar Novo Cadastro", type="primary", key="btn_add_vaga"):
                marca_auto = ""
                comp_auto = ""
                setor_calc = ""
                
                if modelo_sel:
                    info = df_veic[df_veic[col_mod].astype(str) == modelo_sel]
                    if not info.empty:
                        marca_auto = str(info[col_mar].values[0]) if pd.notna(info[col_mar].values[0]) else ""
                        comp_auto = str(info[col_com].values[0]) if pd.notna(info[col_com].values[0]) else ""
                        peso_val = info[col_pes].values[0] if pd.notna(info[col_pes].values[0]) else 1000
                        setor_calc = calcular_setor_vaga(comp_auto, peso_val)

                local_final = local_manual if local_manual != "" else setor_calc

                novo = {
                    "Bloco": str(bloco), "Apto": str(apto), "Proprietario_Morador": str(morador),
                    "CPF": str(cpf), "Contato": str(contato), "Vinculo": str(vinculo),
                    "Modelo": str(modelo_sel), "Marca": marca_auto, "Placa": str(placa),
                    "Cor": str(cor), "Comprimento": comp_auto, "TAG": str(tag),
                    "Data": datetime.now().strftime("%d/%m/%Y"), "Alugada": str(alugada),
                    "De_Quem": str(locatario_nome), "Locatario_CPF": str(locatario_cpf),
                    "Locatario_Contato": str(locatario_contato), "Observacao": str(observacao),
                    "Local": str(local_final)
                }
                df_vagas = pd.concat([df_vagas, pd.DataFrame([novo])], ignore_index=True)
                save_data(df_veic, df_vagas)
                st.success("Cadastro salvo com sucesso!")
                st.rerun()

        # --- SUB-ABA 3: EDITAR ---
        with sub_tab3:
            st.subheader("✏️ EDITAR CADASTRO EXISTENTE")
            if not df_vagas.empty:
                opcoes_dict_edit = {
                    f"Linha {idx} | Bloco {row['Bloco']} - Apto {row['Apto']} | {row['Proprietario_Morador']}": idx
                    for idx, row in df_vagas.iterrows()
                }
                
                reg_para_editar = st.selectbox("Selecione o cadastro para EDITAR:", [""] + list(opcoes_dict_edit.keys()), key="sb_editar")
                
                if reg_para_editar and reg_para_editar in opcoes_dict_edit:
                    idx_edit = opcoes_dict_edit[reg_para_editar]
                    dados_atuais = df_vagas.loc[idx_edit]
                    
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        e_bloco = st.text_input("Bloco", value=str(dados_atuais.get("Bloco", "")), key="ed_bloco")
                        e_apto = st.text_input("Apto", value=str(dados_atuais.get("Apto", "")), key="ed_apto")
                        e_morador = st.text_input("Proprietário / Morador", value=str(dados_atuais.get("Proprietario_Morador", "")), key="ed_morador")
                        e_cpf = st.text_input("CPF", value=str(dados_atuais.get("CPF", "")), key="ed_cpf")
                    with c2:
                        e_contato = st.text_input("Contato", value=str(dados_atuais.get("Contato", "")), key="ed_contato")
                        e_vinculo = st.text_input("Vínculo", value=str(dados_atuais.get("Vinculo", "")), key="ed_vinc")
                        
                        modelos_disp = sorted(df_veic[col_mod].dropna().astype(str).unique().tolist())
                        mod_atual = str(dados_atuais.get("Modelo", ""))
                        idx_mod = modelos_disp.index(mod_atual) + 1 if mod_atual in modelos_disp else 0
                        e_modelo = st.selectbox("Modelo do Veículo", [""] + modelos_disp, index=idx_mod, key="ed_mod")
                        
                        e_placa = st.text_input("Placa", value=str(dados_atuais.get("Placa", "")), key="ed_placa")
                    with c3:
                        e_cor = st.text_input("Cor", value=str(dados_atuais.get("Cor", "")), key="ed_cor")
                        e_tag = st.text_input("TAG", value=str(dados_atuais.get("TAG", "")), key="ed_tag")
                        
                        alugada_atual = str(dados_atuais.get("Alugada", "")).strip().upper()
                        opts_alugada = ["NÃO", "SIM"]
                        idx_alugada = opts_alugada.index(alugada_atual) if alugada_atual in opts_alugada else 0
                        e_alugada = st.selectbox("Vaga Alugada?", opts_alugada, index=idx_alugada, key="ed_alugada")
                        
                        e_local = st.text_input("Local / Setor", value=str(dados_atuais.get("Local", "")), key="ed_local")

                    e_loc_nome = ""
                    e_loc_cpf = ""
                    e_loc_contato = ""
                    
                    if e_alugada == "SIM":
                        st.divider()
                        st.markdown("#### 🔑 Informações do Locador (Quem está alugando a vaga)")
                        ce1, ce2, ce3 = st.columns(3)
                        e_loc_nome = ce1.text_input("Nome do Locador", value=str(dados_atuais.get("De_Quem", "")), key="ed_loc_nome")
                        e_loc_cpf = ce2.text_input("CPF do Locador", value=str(dados_atuais.get("Locatario_CPF", "")), key="ed_loc_cpf")
                        e_loc_contato = ce3.text_input("Contato do Locador", value=str(dados_atuais.get("Locatario_Contato", "")), key="ed_loc_contato")

                    st.divider()
                    e_observacao = st.text_area("Observação (Informações adicionais do cadastro)", value=str(dados_atuais.get("Observacao", "")), key="ed_obs")

                    if st.button("Atualizar Informações", type="primary", key="btn_upd_vaga"):
                        info = df_veic[df_veic[col_mod].astype(str) == e_modelo]
                        
                        marca_auto = str(info[col_mar].values[0]) if not info.empty and pd.notna(info[col_mar].values[0]) else str(dados_atuais.get("Marca", ""))
                        comp_auto = str(info[col_com].values[0]) if not info.empty and pd.notna(info[col_com].values[0]) else str(dados_atuais.get("Comprimento", ""))
                        peso_auto = info[col_pes].values[0] if not info.empty and pd.notna(info[col_pes].values[0]) else 1000
                        
                        setor_calc = calcular_setor_vaga(comp_auto, peso_auto)
                        local_final = e_local if e_local != "" else setor_calc
                        
                        df_vagas.loc[idx_edit, "Bloco"] = str(e_bloco)
                        df_vagas.loc[idx_edit, "Apto"] = str(e_apto)
                        df_vagas.loc[idx_edit, "Proprietario_Morador"] = str(e_morador)
                        df_vagas.loc[idx_edit, "CPF"] = str(e_cpf)
                        df_vagas.loc[idx_edit, "Contato"] = str(e_contato)
                        df_vagas.loc[idx_edit, "Vinculo"] = str(e_vinculo)
                        df_vagas.loc[idx_edit, "Modelo"] = str(e_modelo)
                        df_vagas.loc[idx_edit, "Marca"] = marca_auto
                        df_vagas.loc[idx_edit, "Placa"] = str(e_placa)
                        df_vagas.loc[idx_edit, "Cor"] = str(e_cor)
                        df_vagas.loc[idx_edit, "Comprimento"] = comp_auto
                        df_vagas.loc[idx_edit, "TAG"] = str(e_tag)
                        df_vagas.loc[idx_edit, "Alugada"] = str(e_alugada)
                        df