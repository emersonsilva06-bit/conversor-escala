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

**Nota:** O arquivo `Legenda.xlsx` deve estar na mesma pasta deste programa.
""")

# --- CONFIGURAÇÃO DO ARQUIVO FIXO ---
ARQUIVO_LEGENDA_FIXO = 'Legenda.xlsx'
ABAS_ALVO = ['GH_EQUIPES', 'GH_OPERADOR', 'GH_CENTRAL', 'GH_SUPORTES']
TAMANHO_LOTE = 20  # Quantidade máxima de BPs por arquivo consolidado

# --- FUNÇÕES UTILITÁRIAS ---

def limpar_nome_arquivo(nome):
    return re.sub(r'[<>:"/\\|?*]', '', str(nome))

def extrair_id_google_sheets(url):
    """Extrai o ID da planilha de um link do Google Sheets."""
    padrao = r"/d/([a-zA-Z0-9-_]+)"
    match = re.search(padrao, url)
    if match:
        return match.group(1)
    return None

def carregar_legenda(caminho_arquivo):
    """Lê a legenda de um arquivo local fixo."""
    if not os.path.exists(caminho_arquivo):
        return None, f"❌ ERRO: O arquivo '{caminho_arquivo}' não foi encontrado na pasta do programa."
        
    try:
        df_legenda = pd.read_excel(caminho_arquivo)
        # Normalização TOTAL: String, Strip e Upper
        df_legenda['HORARIO ROSTER'] = df_legenda['HORARIO ROSTER'].astype(str).str.strip().str.upper()
        df_legenda['CODIGO SAP'] = df_legenda['CODIGO SAP'].astype(str).str.strip()
        return pd.Series(df_legenda['CODIGO SAP'].values, index=df_legenda['HORARIO ROSTER']).to_dict(), None
    except Exception as e:
        return None, f"Erro ao ler o arquivo de legenda: {e}"

def buscar_cabecalho_inteligente(df_temp):
    """Procura a linha onde está o cabeçalho 'BP'."""
    indice_cabecalho_bp = -1
    for idx, row in df_temp.iterrows():
        for val in row:
            if str(val).strip().upper() == 'BP':
                indice_cabecalho_bp = idx
                return indice_cabecalho_bp
    return -1

def corrigir_colunas_datas(df, df_raw_header, indice_cabecalho):
    """Ajusta nomes de colunas usando a linha superior se necessário."""
    tem_dia_1 = False
    for col in df.columns:
        col_str = str(col).split('.')[0].strip()
        if isinstance(col, datetime.datetime) and col.day == 1: tem_dia_1 = True
        elif col_str == '1': tem_dia_1 = True
        else:
             try:
                 if int(col_str) == 1: tem_dia_1 = True
             except: pass
    
    if not tem_dia_1 and indice_cabecalho > 0:
        linha_dias = df_raw_header.iloc[indice_cabecalho - 1]
        novas_colunas = []
        for i, col_orig in enumerate(df.columns):
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
        df.columns = novas_colunas
    return df

def processar_dados(df_escala, dicionario_legenda, mes_ano, temp_dir, formato_saida, nome_aba_grupo=""):
    # --- FILTRO DE STATUS ---
    colunas_map = {str(c).strip().upper(): c for c in df_escala.columns}
    if 'STATUS' in colunas_map:
        col_status_real = colunas_map['STATUS']
        df_escala = df_escala[df_escala[col_status_real].astype(str).str.strip().str.upper() == 'ATIVO']

    # Mapa de exceções
    mapa_excecoes = {
        'FR': 'FOLG', 'FRD': 'FOLG', 'FA': 'FAGR', 'FPA': 'FOLG', 
        'EP': 'FOLG', 'T': 'FOLG', 'FE': 'FOLG', 'CIPA': 'FOLG',
        'INSS': 'FOLG', 'AUD': 'FOLG'
    }
    codigos_ignorar = ['DSR', 'ATESTADO', 'LICENÇA', 'FERIAS']
    
    arquivos_gerados = 0
    colaboradores_processados = 0
    log_erros = []
    
    # Variáveis para controle de lotes (Consolidado)
    dados_lote_atual = []
    bps_no_lote = 0
    contador_partes = 1

    # Identificar colunas
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
        return 0, 0, [f"ERRO CRÍTICO na aba '{nome_aba_grupo}': Não foi possível encontrar a coluna do dia '1'."]

    # Processamento
    total_linhas = len(df_escala)

    for index, row in df_escala.iterrows():
        try:
            bp = row['BP']
            if 'NOME COMPLETO' in df_escala.columns:
                nome = row['NOME COMPLETO']
            elif 'NOME' in df_escala.columns:
                nome = row['NOME']
            else:
                nome = f'Colaborador_{index}'
            horario_padrao = str(row['HORÁRIO']).strip().upper()
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
        dias_com_erro = []
        
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
                        continue 
                
                if dia_num > 31: break
                if dia_num < ultimo_dia_processado: break
                ultimo_dia_processado = dia_num

            except:
                continue

            valor_celula = row.iloc[i] 
            
            codigo_final = None
            valor_str = ""
            
            try:
                if not pd.isna(valor_celula):
                    valor_str = str(valor_celula).strip().upper()
            except:
                continue

            if valor_str == "":
                codigo_final = dicionario_legenda.get(horario_padrao)
                if not codigo_final:
                    dias_com_erro.append(f"Dia {dia_num} (Vazio -> Padrão '{horario_padrao}' não achado)")
            elif valor_str in mapa_excecoes:
                codigo_final = mapa_excecoes[valor_str]
            elif any(ign in valor_str for ign in codigos_ignorar):
                continue
            else:
                codigo_final = dicionario_legenda.get(valor_str)
                if not codigo_final:
                    dias_com_erro.append(f"Dia {dia_num} (Código '{valor_str}' não achado)")

            if codigo_final:
                try:
                    mes_ano_limpo = mes_ano.replace('/', '')
                    data_saida = f"{dia_num:02d}{mes_ano_limpo}"
                    datetime.datetime.strptime(f"{dia_num:02d}/{mes_ano}", "%d/%m/%Y") 
                    dados_colaborador.append([bp_int, codigo_final, data_saida, "02"])
                except:
                    pass

        # Lógica de Salvamento
        if dados_colaborador:
            colaboradores_processados += 1
            
            if formato_saida == "Arquivos Individuais (ZIP)":
                nome_limpo = limpar_nome_arquivo(nome).strip()
                nome_arquivo = f"{bp_int}_{nome_limpo}.tsv"
                caminho_completo = os.path.join(temp_dir, nome_arquivo)
                df_saida = pd.DataFrame(dados_colaborador)
                df_saida.to_csv(caminho_completo, sep='\t', header=False, index=False)
                arquivos_gerados += 1
            else:
                # Modo Consolidado (Lotes)
                dados_lote_atual.extend(dados_colaborador)
                bps_no_lote += 1
                
                # Se atingiu 20 BPs, salva o lote e reseta
                if bps_no_lote >= TAMANHO_LOTE:
                    sufixo = f"_{nome_aba_grupo}" if nome_aba_grupo else ""
                    nome_consol = f"CONSOLIDADO{sufixo}_PART{contador_partes}_{mes_ano.replace('/','')}.tsv"
                    caminho_consol = os.path.join(temp_dir, nome_consol)
                    
                    df_saida = pd.DataFrame(dados_lote_atual)
                    df_saida.to_csv(caminho_consol, sep='\t', header=False, index=False)
                    
                    arquivos_gerados += 1
                    contador_partes += 1
                    bps_no_lote = 0
                    dados_lote_atual = [] # Limpa memória

        if dias_com_erro:
             msg_erros = ", ".join(dias_com_erro[:3])
             if len(dias_com_erro) > 3: msg_erros += "..."
             log_erros.append(f"[{nome_aba_grupo}] BP {bp_int}: Dias faltando/pulados: {msg_erros}")
        elif not dados_colaborador:
             log_erros.append(f"[{nome_aba_grupo}] BP {bp_int}: Nenhum dia gerado.")

    # Se for consolidado, salva o que sobrou no último lote (resto da divisão)
    if formato_saida == "Arquivo Único (Consolidado)" and dados_lote_atual:
        sufixo = f"_{nome_aba_grupo}" if nome_aba_grupo else ""
        nome_consol = f"CONSOLIDADO{sufixo}_PART{contador_partes}_{mes_ano.replace('/','')}.tsv"
        caminho_consol = os.path.join(temp_dir, nome_consol)
        
        df_saida = pd.DataFrame(dados_lote_atual)
        df_saida.to_csv(caminho_consol, sep='\t', header=False, index=False)
        arquivos_gerados += 1

    return arquivos_gerados, colaboradores_processados, log_erros

# --- INTERFACE DO USUÁRIO ---

if not os.path.exists(ARQUIVO_LEGENDA_FIXO):
    st.error(f"⚠️ Atenção: O arquivo '{ARQUIVO_LEGENDA_FIXO}' não foi encontrado na pasta.")
    st.info("Por favor, suba o arquivo Legenda.xlsx no GitHub.")
    st.stop()

st.subheader("1. Configuração")
col_conf1, col_conf2 = st.columns(2)

with col_conf1:
    mes_ano = st.text_input("Mês/Ano de Referência", value="01/2026", help="Formato: mm/aaaa")

with col_conf2:
    formato_saida = st.radio(
        "Formato de Saída:",
        ("Arquivos Individuais (ZIP)", "Arquivo Único (Consolidado)")
    )
    if formato_saida == "Arquivo Único (Consolidado)":
        st.caption("ℹ️ Serão gerados arquivos com no máximo 20 colaboradores cada.")

st.subheader("2. Fonte de Dados")
fonte_dados = st.radio("Escolha como carregar a escala:", ("Upload de Arquivo (.xlsx)", "Link Google Sheets"))

df_final_dict = {}
precisa_buscar_header = False 

if fonte_dados == "Upload de Arquivo (.xlsx)":
    f_escala = st.file_uploader("Carregar arquivo de Escala", type=['xlsx'])
else:
    url_gsheets = st.text_input("Cole o Link do Google Sheets aqui:")
    st.info(f"O sistema irá buscar automaticamente as abas: {', '.join(ABAS_ALVO)}")

st.markdown("---")

if st.button("🚀 Processar Arquivos", type="primary"):
    pode_processar = False
    
    if fonte_dados == "Upload de Arquivo (.xlsx)" and f_escala:
        pode_processar = True
    elif fonte_dados == "Link Google Sheets" and url_gsheets:
        pode_processar = True
    else:
        st.warning("Por favor, carregue o arquivo ou insira o link.")

    if pode_processar:
        with st.spinner("Carregando dados e processando..."):
            # 1. Carregar Legenda
            dicionario_legenda, erro_legenda = carregar_legenda(ARQUIVO_LEGENDA_FIXO)
            if erro_legenda:
                st.error(erro_legenda)
                st.stop()
            
            # 2. Preparar DataFrames
            dict_dfs_para_processar = {}

            if fonte_dados == "Link Google Sheets":
                sheet_id = extrair_id_google_sheets(url_gsheets)
                if not sheet_id:
                    st.error("Link do Google Sheets inválido.")
                    st.stop()
                
                url_export = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
                try:
                    dfs_google = pd.read_excel(url_export, sheet_name=None)
                    abas_encontradas = []
                    for aba in ABAS_ALVO:
                        if aba in dfs_google:
                            dict_dfs_para_processar[aba] = dfs_google[aba]
                            abas_encontradas.append(aba)
                    
                    if not abas_encontradas:
                        st.error(f"Não encontrei nenhuma das abas obrigatórias: {ABAS_ALVO}")
                        st.stop()
                    else:
                        st.success(f"Abas encontradas: {', '.join(abas_encontradas)}")
                    precisa_buscar_header = True

                except Exception as e:
                    st.error(f"Erro ao ler Google Sheets: {e}")
                    st.stop()
            else:
                try:
                    xls = pd.ExcelFile(f_escala)
                    abas_presentes = [aba for aba in ABAS_ALVO if aba in xls.sheet_names]
                    
                    if abas_presentes:
                        st.success(f"Abas encontradas no arquivo: {', '.join(abas_presentes)}")
                        for aba in abas_presentes:
                            dict_dfs_para_processar[aba] = pd.read_excel(f_escala, sheet_name=aba)
                        precisa_buscar_header = True
                    else:
                        df_temp = pd.read_excel(f_escala, header=None, nrows=20)
                        idx_header = buscar_cabecalho_inteligente(df_temp)
                        
                        if idx_header == -1:
                            st.error("Não encontrei a coluna 'BP' nas primeiras 20 linhas.")
                            st.stop()
                        
                        f_escala.seek(0)
                        df_loaded = pd.read_excel(f_escala, header=idx_header)
                        df_loaded = corrigir_colunas_datas(df_loaded, df_temp, idx_header)
                        dict_dfs_para_processar["Arquivo_Upload"] = df_loaded
                        precisa_buscar_header = False

                except Exception as e:
                    st.error(f"Erro ao ler arquivo de upload: {e}")
                    st.stop()

            # 3. Loop de Processamento
            with tempfile.TemporaryDirectory() as tmpdirname:
                total_arquivos = 0
                total_colab = 0
                todos_erros = []

                progresso_abas = st.progress(0)
                total_abas = len(dict_dfs_para_processar)
                
                for i, (nome_aba, df_atual) in enumerate(dict_dfs_para_processar.items()):
                    if precisa_buscar_header:
                        try:
                            idx_bp_virtual = -1
                            if 'BP' not in [str(c).upper().strip() for c in df_atual.columns]:
                                for idx, row in df_atual.head(20).iterrows():
                                    if any(str(v).strip().upper() == 'BP' for v in row):
                                        idx_bp_virtual = idx
                                        break
                                
                                if idx_bp_virtual != -1:
                                    new_header = df_atual.iloc[idx_bp_virtual]
                                    df_atual = df_atual[idx_bp_virtual+1:].reset_index(drop=True)
                                    df_atual.columns = new_header
                        except:
                            pass

                    qtd_arq, qtd_col, erros_aba = processar_dados(df_atual, dicionario_legenda, mes_ano, tmpdirname, formato_saida, nome_aba_grupo=nome_aba)
                    
                    total_arquivos += qtd_arq
                    total_colab += qtd_col
                    todos_erros.extend(erros_aba)
                    
                    progresso_abas.progress((i + 1) / total_abas)

                # Finalização e Download
                if total_arquivos > 0:
                    nome_zip = f"importacao_sap_{mes_ano.replace('/','')}.zip"
                    caminho_zip = os.path.join(tmpdirname, 'arquivos_importacao')
                    
                    # Como agora o consolidado gera MÚLTIPLOS arquivos (Parts),
                    # A melhor forma de entregar é SEMPRE via ZIP.
                    shutil.make_archive(caminho_zip, 'zip', tmpdirname)
                    
                    with open(caminho_zip + ".zip", "rb") as f:
                        st.balloons()
                        tipo_saida = "arquivos consolidados (em partes)" if formato_saida == "Arquivo Único (Consolidado)" else "arquivos individuais"
                        st.success(f"Sucesso! Gerados {total_arquivos} {tipo_saida}.")
                            
                        st.download_button(
                            label="📥 Baixar Resultados (ZIP)",
                            data=f,
                            file_name=nome_zip,
                            mime="application/zip"
                        )
                else:
                    st.warning("Nenhum arquivo foi gerado.")

                if todos_erros:
                    st.warning("Alguns dias ou colaboradores foram pulados. Veja os detalhes abaixo:")
                    with st.expander("Ver Logs de Erros e Dias Pulados"):
                        for e in todos_erros:
                            st.write(e)





