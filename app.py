import streamlit as st
import pandas as pd
import datetime
import os
import re
import tempfile
import shutil
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Conversor de Escala SAP", layout="centered")

st.title("✈️ Conversor de Escala - Rampa BSB")
st.markdown("""
Esta ferramenta transforma a planilha de escala operacional em arquivos TSV para importação no SAP.

**Nota:** O arquivo `Legenda.xlsx` deve estar na mesma pasta deste programa (no GitHub).
""")

# --- CONFIGURAÇÃO DO ARQUIVO FIXO ---
ARQUIVO_LEGENDA_FIXO = 'Legenda.xlsx'

# --- FUNÇÕES UTILITÁRIAS ---

def limpar_nome_arquivo(nome):
    return re.sub(r'[<>:"/\\|?*]', '', str(nome))

def carregar_legenda(caminho_arquivo):
    """Lê a legenda de um arquivo local fixo."""
    if not os.path.exists(caminho_arquivo):
        return None, f"❌ ERRO: O arquivo '{caminho_arquivo}' não foi encontrado na pasta do programa."
        
    try:
        df_legenda = pd.read_excel(caminho_arquivo)
        df_legenda['HORARIO ROSTER'] = df_legenda['HORARIO ROSTER'].astype(str).str.strip()
        df_legenda['CODIGO SAP'] = df_legenda['CODIGO SAP'].astype(str).str.strip()
        return pd.Series(df_legenda['CODIGO SAP'].values, index=df_legenda['HORARIO ROSTER']).to_dict(), None
    except Exception as e:
        return None, f"Erro ao ler o arquivo de legenda: {e}"

def processar_dados(df_escala, dicionario_legenda, mes_ano, temp_dir, formato_saida):
    mapa_excecoes = {
        'FR': 'FOLG', 'FRD': 'FOLG', 'FA': 'FAGR', 'FPA': 'FOLG', 
        'EP': 'FOLG', 'T': 'FOLG', 'FE': 'FOLG', 'CIPA': 'FOLG'
    }
    codigos_ignorar = ['DSR', 'ATESTADO', 'LICENÇA', 'FERIAS']
    
    arquivos_gerados = 0
    colaboradores_processados = 0
    log_erros = []
    
    # Lista acumuladora para o modo consolidado
    dados_todos_consolidado = []

    # Identificar colunas (Lógica robusta)
    colunas = list(df_escala.columns)
    indice_dia_1 = -1
    for i, col in enumerate(colunas):
        if isinstance(col, datetime.datetime):
            if col.day == 1:
                indice_dia_1 = i
                break
        else:
            col_str = str(col).split('.')[0].strip()
            try:
                if int(col_str) == 1:
                    indice_dia_1 = i
                    break
            except:
                pass
    
    if indice_dia_1 == -1:
        return 0, 0, ["ERRO CRÍTICO: Não foi possível encontrar a coluna do dia '1'."]

    # Processamento
    barra_progresso = st.progress(0)
    total_linhas = len(df_escala)

    for index, row in df_escala.iterrows():
        if index % 5 == 0:
            barra_progresso.progress(min(index / total_linhas, 1.0))

        try:
            bp = row['BP']
            if 'NOME COMPLETO' in df_escala.columns:
                nome = row['NOME COMPLETO']
            elif 'NOME' in df_escala.columns:
                nome = row['NOME']
            else:
                nome = f'Colaborador_{index}'
            horario_padrao = str(row['HORÁRIO']).strip()
        except:
            continue

        if pd.isna(bp) or str(bp).strip() == '':
            continue
            
        try:
            bp_int = int(bp)
        except:
            bp_int = str(bp).strip()

        dados_colaborador = []
        ultimo_dia_processado = 0
        
        for i in range(indice_dia_1, len(colunas)):
            col_header = colunas[i]
            dia_num = 0
            
            try:
                if isinstance(col_header, datetime.datetime):
                    dia_num = col_header.day
                else:
                    col_str = str(col_header).split('.')[0].strip()
                    if col_str.isdigit():
                        dia_num = int(col_str)
                    else:
                        break 
                
                if dia_num > 31: break
                if dia_num < ultimo_dia_processado: break
                ultimo_dia_processado = dia_num

            except:
                break

            valor_celula = row[colunas[i]]
            codigo_final = None
            valor_str = ""
            if not pd.isna(valor_celula):
                valor_str = str(valor_celula).strip().upper()

            if valor_str == "":
                codigo_final = dicionario_legenda.get(horario_padrao)
            elif valor_str in mapa_excecoes:
                codigo_final = mapa_excecoes[valor_str]
            elif any(ign in valor_str for ign in codigos_ignorar):
                continue
            else:
                codigo_final = dicionario_legenda.get(valor_str)

            if codigo_final:
                try:
                    mes_ano_limpo = mes_ano.replace('/', '')
                    data_saida = f"{dia_num:02d}{mes_ano_limpo}"
                    datetime.datetime.strptime(f"{dia_num:02d}/{mes_ano}", "%d/%m/%Y") 
                    dados_colaborador.append([bp_int, codigo_final, data_saida, "02"])
                except:
                    pass

        # Lógica de Salvamento dependendo do formato escolhido
        if dados_colaborador:
            colaboradores_processados += 1
            
            if formato_saida == "Arquivos Individuais (ZIP)":
                # Salva um arquivo por pessoa
                nome_limpo = limpar_nome_arquivo(nome).strip()
                nome_arquivo = f"{bp_int}_{nome_limpo}.tsv"
                caminho_completo = os.path.join(temp_dir, nome_arquivo)
                
                df_saida = pd.DataFrame(dados_colaborador)
                df_saida.to_csv(caminho_completo, sep='\t', header=False, index=False)
                arquivos_gerados += 1
            else:
                # Acumula na lista geral
                dados_todos_consolidado.extend(dados_colaborador)
        else:
            if horario_padrao not in dicionario_legenda:
                 log_erros.append(f"BP {bp_int} ({nome}): Sem dados gerados. Horário '{horario_padrao}' não achado na legenda.")

    # Se for consolidado, salva o arquivo único no final
    if formato_saida == "Arquivo Único (Consolidado)" and dados_todos_consolidado:
        nome_arquivo_consol = f"CONSOLIDADO_IMPORTACAO_{mes_ano.replace('/','')}.tsv"
        caminho_completo = os.path.join(temp_dir, nome_arquivo_consol)
        df_saida = pd.DataFrame(dados_todos_consolidado)
        df_saida.to_csv(caminho_completo, sep='\t', header=False, index=False)
        arquivos_gerados = 1 # Gerou 1 arquivo mestre

    barra_progresso.progress(100)
    return arquivos_gerados, colaboradores_processados, log_erros

# --- INTERFACE DO USUÁRIO ---

if not os.path.exists(ARQUIVO_LEGENDA_FIXO):
    st.error(f"⚠️ Atenção: O arquivo '{ARQUIVO_LEGENDA_FIXO}' não foi encontrado na pasta.")
    st.info("Por favor, suba o arquivo Legenda.xlsx no GitHub.")
    st.stop()

st.subheader("1. Configuração")
col_conf1, col_conf2 = st.columns(2)

with col_conf1:
    mes_ano = st.text_input("Mês/Ano de Referência", value="12/2024", help="Formato: mm/aaaa")

with col_conf2:
    formato_saida = st.radio(
        "Formato de Saída:",
        ("Arquivos Individuais (ZIP)", "Arquivo Único (Consolidado)")
    )

st.subheader("2. Upload da Escala")
f_escala = st.file_uploader("Carregar arquivo de Escala (.xlsx)", type=['xlsx'])

st.markdown("---")

if st.button("🚀 Processar Arquivos", type="primary"):
    if not f_escala:
        st.warning("Por favor, carregue o arquivo de escala.")
    else:
        with st.spinner("Carregando Legenda e processando Escala..."):
            # 1. Carregar Legenda
            dicionario_legenda, erro_legenda = carregar_legenda(ARQUIVO_LEGENDA_FIXO)
            if erro_legenda:
                st.error(erro_legenda)
                st.stop()
            
            # 2. Carregar Escala (Busca inteligente de cabeçalho)
            try:
                df_temp = pd.read_excel(f_escala, header=None, nrows=20)
                indice_cabecalho_bp = -1
                for idx, row in df_temp.iterrows():
                    for val in row:
                        if str(val).strip().upper() == 'BP':
                            indice_cabecalho_bp = idx
                            break
                    if indice_cabecalho_bp != -1: break
                
                if indice_cabecalho_bp == -1:
                    st.error("Não encontrei a coluna 'BP' nas primeiras 20 linhas.")
                    st.stop()

                f_escala.seek(0)
                df_escala = pd.read_excel(f_escala, header=indice_cabecalho_bp)

                tem_dia_1 = False
                for col in df_escala.columns:
                    col_str = str(col).split('.')[0].strip()
                    if isinstance(col, datetime.datetime) and col.day == 1: tem_dia_1 = True
                    elif col_str == '1': tem_dia_1 = True
                    else:
                         try:
                             if int(col_str) == 1: tem_dia_1 = True
                         except: pass
                
                if not tem_dia_1 and indice_cabecalho_bp > 0:
                    linha_dias = df_temp.iloc[indice_cabecalho_bp - 1]
                    novas_colunas = []
                    for i, col_orig in enumerate(df_escala.columns):
                        val_cima = linha_dias.iloc[i]
                        val_final = col_orig
                        
                        if isinstance(val_cima, datetime.datetime): 
                            val_final = val_cima
                        else:
                            val_str_cima = str(val_cima).split('.')[0].strip()
                            if val_str_cima.isdigit(): 
                                val_check = int(val_str_cima)
                                if 1 <= val_check <= 31: val_final = val_check
                        
                        novas_colunas.append(val_final)
                    df_escala.columns = novas_colunas

            except Exception as e:
                st.error(f"Erro ao ler escala: {e}")
                st.stop()

            # 3. Processar
            with tempfile.TemporaryDirectory() as tmpdirname:
                # Passamos a nova variável 'formato_saida' para a função
                qtd_arquivos, qtd_colaboradores, erros = processar_dados(df_escala, dicionario_legenda, mes_ano, tmpdirname, formato_saida)
                
                if qtd_arquivos > 0:
                    st.balloons()
                    
                    if formato_saida == "Arquivos Individuais (ZIP)":
                        st.success(f"Sucesso! {qtd_arquivos} arquivos gerados para {qtd_colaboradores} colaboradores.")
                        shutil.make_archive(os.path.join(tmpdirname, 'arquivos_importacao'), 'zip', tmpdirname)
                        
                        with open(os.path.join(tmpdirname, 'arquivos_importacao.zip'), "rb") as f:
                            st.download_button(
                                label="📥 Baixar Arquivos (ZIP)",
                                data=f,
                                file_name=f"importacao_sap_{mes_ano.replace('/','')}.zip",
                                mime="application/zip"
                            )
                    else:
                        # Modo Consolidado: Baixa apenas o arquivo TSV único
                        st.success(f"Sucesso! Arquivo consolidado gerado contendo {qtd_colaboradores} colaboradores.")
                        nome_consol = f"CONSOLIDADO_IMPORTACAO_{mes_ano.replace('/','')}.tsv"
                        caminho_consol = os.path.join(tmpdirname, nome_consol)
                        
                        with open(caminho_consol, "rb") as f:
                            st.download_button(
                                label="📥 Baixar Arquivo Único (TSV)",
                                data=f,
                                file_name=nome_consol,
                                mime="text/tab-separated-values"
                            )

                else:
                    st.warning("Nenhum arquivo foi gerado.")

                if erros:
                    with st.expander("Ver Logs / Avisos"):
                        for e in erros:
                            st.write(e)

