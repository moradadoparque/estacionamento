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
    
    # Converte tudo para texto limpando valores nulos/NaN
    df_vagas = df_vagas.fillna("").astype(str).replace(["nan", "None", "NaN"], "")
    
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
                
                selecionado = st.selectbox("Pesquise ou selecione um morador:", [""] + list(opcoes_dict.keys()), key="sb_consultar")
                
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
                    
                    # Trata exibições para não exibir 'nan' ou vazio
                    tag_val = str(m.get('TAG', '')).strip()
                    tag_val = tag_val if tag_val and tag_val != "nan" else "-"
                    
                    data_val = str(m.get('Data', '')).strip()
                    data_val = data_val if data_val and data_val != "nan" else "-"
                    
                    alugada_val = str(m.get('Alugada', '')).strip()
                    alugada_val = alugada_val if alugada_val and alugada_val != "nan" else "-"
                    
                    with c1:
                        st.info("**🏠 Unidade / Identificação**")
                        st.write(f"**Bloco:** {m.get('Bloco', '-') or '-'}")
                        st.write(f"**Apto:** {m.get('Apto', '-') or '-'}")
                        st.write(f"**Proprietário / Morador:** {m.get('Proprietario_Morador', '-') or '-'}")
                        st.write(f"**CPF:** {m.get('CPF', '-') or '-'}")
                        st.write(f"**Contato:** {m.get('Contato', '-') or '-'}")
                        st.write(f"**Vínculo:** {m.get('Vinculo', '-') or '-'}")
                        
                    with c2:
                        st.success("**🚘 Dados do Veículo**")
                        st.write(f"**Modelo:** {m.get('Modelo', '-') or '-'}")
                        st.write(f"**Marca:** {marca_dinamica or '-'}")
                        st.write(f"**Placa:** {m.get('Placa', '-') or '-'}")
                        st.write(f"**Cor:** {m.get('Cor', '-') or '-'}")
                        st.write(f"**Comprimento:** {comp_dinamico} m")
                        st.markdown("---")
                        st.write(f"**Setor:** {setor_dinamico or '-'}")
                        
                    with c3:
                        st.warning("**🏷️ Vaga & TAG**")
                        st.write(f"**TAG:** {tag_val}")
                        st.write(f"**Data:** {data_val}")
                        st.write(f"**Vaga Alugada?:** {alugada_val}")
                        
                        if alugada_val.upper() == "SIM":
                            st.markdown("---")
                            st.write(f"**Locatário:** {m.get('De_Quem', '-') or '-'}")
                            st.write(f"**CPF Locatário:** {m.get('Locatario_CPF', '-') or '-'}")
                            st.write(f"**Contato Locatário:** {m.get('Locatario_Contato', '-') or '-'}")

                        # Trata observação garantindo conversão segura para string
                        obs_val = str(m.get('Observacao', '') or '').strip()
                        if obs_val and obs_val.lower() not in ["nan", "none", "-"]:
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
                st.markdown("#### 🔑 Informações do Locatário (Quem está alugando a vaga)")
                cl1, cl2, cl3 = st.columns(3)
                locatario_nome = cl1.text_input("Nome do Locatário", key="add_loc_nome")
                locatario_cpf = cl2.text_input("CPF do Locatário", key="add_loc_cpf")
                locatario_contato = cl3.text_input("Contato do Locatário", key="add_loc_contato")

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
                        st.markdown("#### 🔑 Informações do Locatário (Quem está alugando a vaga)")
                        ce1, ce2, ce3 = st.columns(3)
                        e_loc_nome = ce1.text_input("Nome do Locatário", value=str(dados_atuais.get("De_Quem", "")), key="ed_loc_nome")
                        e_loc_cpf = ce2.text_input("CPF do Locatário", value=str(dados_atuais.get("Locatario_CPF", "")), key="ed_loc_cpf")
                        e_loc_contato = ce3.text_input("Contato do Locatário", value=str(dados_atuais.get("Locatario_Contato", "")), key="ed_loc_contato")

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
                        df_vagas.loc[idx_edit, "De_Quem"] = str(e_loc_nome)
                        df_vagas.loc[idx_edit, "Locatario_CPF"] = str(e_loc_cpf)
                        df_vagas.loc[idx_edit, "Locatario_Contato"] = str(e_loc_contato)
                        df_vagas.loc[idx_edit, "Observacao"] = str(e_observacao)
                        df_vagas.loc[idx_edit, "Local"] = str(local_final)
                        
                        save_data(df_veic, df_vagas)
                        st.success("Cadastro atualizado com sucesso!")
                        st.rerun()

        # --- SUB-ABA 4: EXCLUIR ---
        with sub_tab4:
            st.subheader("🗑️ EXCLUIR CADASTRO")
            st.warning("⚠️ **ATENÇÃO:** A exclusão é permanente e removerá o morador da base de dados.")
            
            if not df_vagas.empty:
                opcoes_dict_del = {
                    f"Bloco {row['Bloco']} - Apto {row['Apto']} | {row['Proprietario_Morador']} | Placa: {row['Placa']}": idx
                    for idx, row in df_vagas.iterrows()
                }
                
                reg_para_excluir = st.selectbox("Selecione o cadastro para EXCLUIR:", [""] + list(opcoes_dict_del.keys()), key="sb_excluir")
                
                if reg_para_excluir and reg_para_excluir in opcoes_dict_del:
                    idx_del = opcoes_dict_del[reg_para_excluir]
                    m_del = df_vagas.loc[idx_del]
                    
                    st.error(f"Você está prestes a excluir: **Bloco {m_del['Bloco']} - Apto {m_del['Apto']} | {m_del['Proprietario_Morador']}**")
                    
                    if st.button("🔴 Confirmar Exclusão Permanente", type="primary"):
                        df_vagas = df_vagas.drop(index=idx_del).reset_index(drop=True)
                        save_data(df_veic, df_vagas)
                        st.success("Cadastro excluído com sucesso!")
                        st.rerun()

    # =========================================================================
    # ---------------- ABA 2: CADASTRO DE VEÍCULOS ----------------------------
    # =========================================================================
    with tab_veic:
        sub_veic1, sub_veic2, sub_veic3, sub_veic4 = st.tabs(["🔍 CONSULTAR", "➕ NOVO CADASTRO", "✏️ EDITAR", "🗑️ EXCLUIR"])

        # --- SUB-ABA VEÍCULOS 1: CONSULTAR ---
        with sub_veic1:
            st.subheader("🔍 CONSULTAR VEÍCULO DA BASE")
            if not df_veic.empty:
                dict_veic = {
                    f"{row[col_mod]} | Marca: {row[col_mar]} | Comprimento: {row[col_com]}m": idx
                    for idx, row in df_veic.iterrows()
                }
                veic_sel = st.selectbox("Pesquise ou selecione um modelo de veículo:", [""] + list(dict_veic.keys()), key="sb_consultar_veic")
                
                if veic_sel and veic_sel in dict_veic:
                    idx_v = dict_veic[veic_sel]
                    v_row = df_veic.loc[idx_v]
                    
                    val_largura = v_row[col_lar] if col_lar in v_row and pd.notna(v_row[col_lar]) else "-"
                    
                    st.divider()
                    st.markdown(f"### 🚘 Modelo: **{v_row[col_mod]}**")
                    c_v1, c_v2 = st.columns(2)
                    with c_v1:
                        st.info("**Ficha do Veículo**")
                        st.write(f"**Modelo:** {v_row[col_mod]}")
                        st.write(f"**Marca:** {v_row[col_mar]}")
                    with c_v2:
                        st.success("**Especificações Técnicas**")
                        st.write(f"**Peso Aproximado:** {v_row[col_pes]} kg")
                        st.write(f"**Largura:** {val_largura} m")
                        st.write(f"**Comprimento:** {v_row[col_com]} m")
                        setor_calc_v = calcular_setor_vaga(v_row[col_com], v_row[col_pes])
                        st.write(f"**Setor:** {setor_calc_v}")

            st.divider()
            with st.expander("📊 Clique para ver a Tabela Completa de Veículos"):
                st.dataframe(df_veic, use_container_width=True)

        # --- SUB-ABA VEÍCULOS 2: NOVO CADASTRO ---
        with sub_veic2:
            st.subheader("➕ CADASTRAR NOVO MODELO DE VEÍCULO")
            
            c1, c2, c3, c4, c5 = st.columns(5)
            v_modelo = c1.text_input("Modelo (ex: ONIX)", key="nv_mod")
            v_marca = c2.text_input("Marca (ex: CHEVROLET)", key="nv_mar")
            v_peso = c3.number_input("Peso (kg)", value=1000, step=50, key="nv_pes")
            v_larg = c4.number_input("Largura (m)", value=1.70, step=0.01, format="%.2f", key="nv_lar")
            v_comp = c5.number_input("Comprimento (m)", value=4.00, step=0.01, format="%.2f", key="nv_com")
            
            setor_calculado_novo = calcular_setor_vaga(v_comp, v_peso)
            st.text_input("Setor (Calculado Automático)", value=setor_calculado_novo, disabled=True, key="nv_setor")
            
            if st.button("Cadastrar Veículo", type="primary", key="btn_add_veic_auto"):
                if v_modelo.strip() == "":
                    st.error("Por favor, informe o Modelo do veículo.")
                else:
                    novo_veic = {
                        col_mod: v_modelo.strip().upper(),
                        col_mar: v_marca.strip().upper(),
                        col_pes: v_peso,
                        col_lar: v_larg,
                        col_com: v_comp
                    }
                    df_veic = pd.concat([df_veic, pd.DataFrame([novo_veic])], ignore_index=True)
                    save_data(df_veic, df_vagas)
                    st.success(f"Veículo '{v_modelo.upper()}' cadastrado com sucesso no setor {setor_calculado_novo}!")
                    st.rerun()

        # --- SUB-ABA VEÍCULOS 3: EDITAR ---
        with sub_veic3:
            st.subheader("✏️ EDITAR MODELO DE VEÍCULO")
            if not df_veic.empty:
                dict_veic_edit = {
                    f"Linha {idx} | {row[col_mod]} ({row[col_mar]})": idx
                    for idx, row in df_veic.iterrows()
                }
                veic_edit_sel = st.selectbox("Selecione o veículo para EDITAR:", [""] + list(dict_veic_edit.keys()), key="sb_edit_veic")
                
                if veic_edit_sel and veic_edit_sel in dict_veic_edit:
                    idx_ve_edit = dict_veic_edit[veic_edit_sel]
                    v_edit = df_veic.loc[idx_ve_edit]
                    
                    ce1, ce2, ce3, ce4, ce5 = st.columns(5)
                    e_v_mod = ce1.text_input("Modelo", value=str(v_edit[col_mod]), key="ev_mod")
                    e_v_mar = ce2.text_input("Marca", value=str(v_edit[col_mar]), key="ev_mar")
                    
                    val_peso = float(v_edit[col_pes]) if pd.notna(v_edit[col_pes]) else 1000.0
                    val_larg = float(v_edit[col_lar]) if col_lar in v_edit and pd.notna(v_edit[col_lar]) else 1.70
                    val_comp = float(v_edit[col_com]) if pd.notna(v_edit[col_com]) else 4.00
                    
                    e_v_peso = ce3.number_input("Peso (kg)", value=val_peso, step=50.0, key="ev_pes")
                    e_v_larg = ce4.number_input("Largura (m)", value=val_larg, step=0.01, format="%.2f", key="ev_lar")
                    e_v_comp = ce5.number_input("Comprimento (m)", value=val_comp, step=0.01, format="%.2f", key="ev_com")
                    
                    setor_calculado_edit = calcular_setor_vaga(e_v_comp, e_v_peso)
                    st.text_input("Setor (Calculado Automático)", value=setor_calculado_edit, disabled=True, key="ev_setor")
                    
                    if st.button("Atualizar Veículo e Sincronizar Moradores", type="primary", key="btn_upd_veic_auto"):
                        novo_mod = e_v_mod.strip().upper()
                        novo_mar = e_v_mar.strip().upper()
                        novo_comp = e_v_comp
                        novo_peso = e_v_peso
                        novo_larg = e_v_larg
                        novo_setor = setor_calculado_edit
                        
                        mod_antigo = str(v_edit[col_mod]).strip().upper()
                        
                        df_veic.loc[idx_ve_edit, col_mod] = novo_mod
                        df_veic.loc[idx_ve_edit, col_mar] = novo_mar
                        df_veic.loc[idx_ve_edit, col_pes] = novo_peso
                        df_veic.loc[idx_ve_edit, col_lar] = novo_larg
                        df_veic.loc[idx_ve_edit, col_com] = novo_comp
                        
                        mask_m = df_vagas["Modelo"].astype(str).str.strip().str.upper() == mod_antigo
                        if mask_m.any():
                            df_vagas.loc[mask_m, "Modelo"] = novo_mod
                            df_vagas.loc[mask_m, "Marca"] = novo_mar
                            df_vagas.loc[mask_m, "Comprimento"] = str(novo_comp)
                            df_vagas.loc[mask_m, "Local"] = novo_setor
                        
                        save_data(df_veic, df_vagas)
                        st.success("Dados do veículo e moradores vinculados atualizados com sucesso!")
                        st.rerun()

        # --- SUB-ABA VEÍCULOS 4: EXCLUIR ---
        with sub_veic4:
            st.subheader("🗑️ EXCLUIR MODELO DE VEÍCULO")
            st.warning("⚠️ **ATENÇÃO:** A exclusão removerá o modelo da base de consulta de veículos.")
            
            if not df_veic.empty:
                dict_veic_del = {
                    f"Linha {idx} | {row[col_mod]} | Marca: {row[col_mar]}": idx
                    for idx, row in df_veic.iterrows()
                }
                veic_del_sel = st.selectbox("Selecione o veículo para EXCLUIR:", [""] + list(dict_veic_del.keys()), key="sb_del_veic")
                
                if veic_del_sel and veic_del_sel in dict_veic_del:
                    idx_ve_del = dict_veic_del[veic_del_sel]
                    v_del = df_veic.loc[idx_ve_del]
                    
                    st.divider()
                    st.error(f"Você está prestes a excluir o modelo: **{v_del[col_mod]} ({v_del[col_mar]})**")
                    
                    if st.button("🔴 Confirmar Exclusão de Veículo", type="primary"):
                        df_veic = df_veic.drop(index=idx_ve_del).reset_index(drop=True)
                        save_data(df_veic, df_vagas)
                        st.success("Modelo de veículo excluído com sucesso!")
                        st.rerun()
