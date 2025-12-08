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
ABAS_ALVO = ['GH_EQUIPES', 'GH_OPERADOR', 'GH_CENTRAL', 'GH_SUPORTES']

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
    # Mapa de exceções (Tudo em MAIÚSCULO para garantir)
    mapa_excecoes = {
        'FR': 'FOLG', 'FRD': 'FOLG', 'FA': 'FAGR', 'FPA': 'FOLG', 
        'EP': 'FOLG', 'T': 'FOLG', 'FE': 'FOLG', 'CIPA': 'FOLG',
        'INSS': 'FOLG', 'AUD': 'FOLG'
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
            # Normaliza horário padrão para UPPER
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
        dias_com_erro = [] # Para reportar quais dias falharam
        
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
                        # Se não for número, CONTINUA (pula coluna 'Total' ou vazia) em vez de quebrar
                        continue 
                
                # Se o dia for absurdamente grande, para
                if dia_num > 31: break
                
                # Se o dia "voltou" (ex: estava no 31 e veio 1), acabou o mês
                if dia_num < ultimo_dia_processado: break
                
                ultimo_dia_processado = dia_num

            except:
                continue

            # CORREÇÃO AQUI: Usar iloc para pegar pelo índice da coluna, evitando erro de nomes duplicados
            valor_celula = row.iloc[i] 
            
            codigo_final = None
            valor_str = ""
            
            # Verificação segura de NA
            try:
                if not pd.isna(valor_celula):
                    valor_str = str(valor_celula).strip().upper()
            except:
                # Se der erro na verificação (ex: array numpy), ignora
                continue

            # Lógica de Decisão
            if valor_str == "":
                codigo_final = dicionario_legenda.get(horario_padrao)
                if not codigo_final:
                    dias_com_erro.append(f"Dia {dia_num} (Vazio -> Padrão '{horario_padrao}' não achado)")
            elif valor_str in mapa_excecoes:
                codigo_final = mapa_excecoes[valor_str]
            elif any(ign in valor_str for ign in codigos_ignorar):
                continue # Ignora intencionalmente
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
                dados_todos_consolidado.extend(dados_colaborador)
        
        # Loga erros se houve dias faltando
        if dias_com_erro:
             # Limita a 3 erros por pessoa para não poluir
             msg_erros = ", ".join(dias_com_erro[:3])
             if len(dias_com_erro) > 3: msg_erros += "..."
             log_erros.append(f"[{nome_aba_grupo}] BP {bp_int}: Dias faltando/pulados: {msg_erros}")
        elif not dados_colaborador:
             log_erros.append(f"[{nome_aba_grupo}] BP {bp_int}: Nenhum dia gerado.")

    # Se for consolidado, salva o arquivo único desta aba no final
    if formato_saida == "Arquivo Único (Consolidado)" and dados_todos_consolidado:
        sufixo = f"_{nome_aba_grupo}" if nome_aba_grupo else ""
        nome_arquivo_consol = f"CONSOLIDADO{sufixo}_{mes_ano.replace('/','')}.tsv"
        
        caminho_completo = os.path.join(temp_dir, nome_arquivo_consol)
        df_saida = pd.DataFrame(dados_todos_consolidado)
        df_saida.to_csv(caminho_completo, sep='\t', header=False, index=False)
        arquivos_gerados = 1 

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

st.subheader("2. Fonte de Dados")
fonte_dados = st.radio("Escolha como carregar a escala:", ("Upload de Arquivo (.xlsx)", "Link Google Sheets"))

df_final_dict = {} # Dicionário para guardar {NomeAba: DataFrame}
modo_google = False

if fonte_dados == "Upload de Arquivo (.xlsx)":
    f_escala = st.file_uploader("Carregar arquivo de Escala", type=['xlsx'])
    if f_escala:
        df_final_dict["Geral"] = f_escala # Guarda o arquivo para processamento padrão

else:
    modo_google = True
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
            
            # 2. Preparar DataFrames (Upload vs Google)
            dict_dfs_para_processar = {}

            if modo_google:
                sheet_id = extrair_id_google_sheets(url_gsheets)
                if not sheet_id:
                    st.error("Link do Google Sheets inválido. Certifique-se de copiar o link completo.")
                    st.stop()
                
                url_export = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
                try:
                    # Lê todas as abas do Google Sheets
                    dfs_google = pd.read_excel(url_export, sheet_name=None)
                    
                    # Filtra apenas as abas desejadas
                    abas_encontradas = []
                    for aba in ABAS_ALVO:
                        if aba in dfs_google:
                            dict_dfs_para_processar[aba] = dfs_google[aba]
                            abas_encontradas.append(aba)
                    
                    if not abas_encontradas:
                        st.error(f"Não encontrei nenhuma das abas obrigatórias no link fornecido. Abas procuradas: {ABAS_ALVO}")
                        st.stop()
                    else:
                        st.success(f"Abas encontradas: {', '.join(abas_encontradas)}")

                except Exception as e:
                    st.error(f"Erro ao ler Google Sheets. Verifique se a planilha está compartilhada como 'Qualquer pessoa com o link'. Detalhe: {e}")
                    st.stop()
            else:
                # Modo Upload (Arquivo Único / Aba Padrão)
                try:
                    # Lê para achar cabeçalho
                    df_temp = pd.read_excel(f_escala, header=None, nrows=20)
                    idx_header = buscar_cabecalho_inteligente(df_temp)
                    
                    if idx_header == -1:
                        st.error("Não encontrei a coluna 'BP' nas primeiras 20 linhas do arquivo enviado.")
                        st.stop()
                    
                    f_escala.seek(0)
                    df_loaded = pd.read_excel(f_escala, header=idx_header)
                    # Aplica correção de colunas
                    df_loaded = corrigir_colunas_datas(df_loaded, df_temp, idx_header)
                    
                    dict_dfs_para_processar["Arquivo_Upload"] = df_loaded
                except Exception as e:
                    st.error(f"Erro ao ler arquivo de upload: {e}")
                    st.stop()

            # 3. Loop de Processamento (Processa cada aba encontrada)
            with tempfile.TemporaryDirectory() as tmpdirname:
                total_arquivos = 0
                total_colab = 0
                todos_erros = []

                # Barra de progresso geral
                progresso_abas = st.progress(0)
                total_abas = len(dict_dfs_para_processar)
                
                for i, (nome_aba, df_atual) in enumerate(dict_dfs_para_processar.items()):
                    # Se for modo Google, tenta ajuste fino de cabeçalho
                    if modo_google:
                        try:
                            idx_bp_virtual = -1
                            if 'BP' not in [str(c).upper().strip() for c in df_atual.columns]:
                                for idx, row in df_atual.head(20).iterrows():
                                    if any(str(v).strip().upper() == 'BP' for v in row):
                                        idx_bp_virtual = idx
                                        break
                                
                                if idx_bp_virtual != -1:
                                    new_header = df_atual.iloc[idx_bp_virtual]
                                    df_atual = df_atual[idx_bp_virtual+1:]
                                    df_atual.columns = new_header
                        except:
                            pass

                    # Processa a aba
                    qtd_arq, qtd_col, erros_aba = processar_dados(df_atual, dicionario_legenda, mes_ano, tmpdirname, formato_saida, nome_aba_grupo=nome_aba)
                    
                    total_arquivos += qtd_arq
                    total_colab += qtd_col
                    todos_erros.extend(erros_aba)
                    
                    progresso_abas.progress((i + 1) / total_abas)

                # Finalização e Download
                if total_arquivos > 0:
                    nome_zip = f"importacao_sap_{mes_ano.replace('/','')}.zip"
                    caminho_zip = os.path.join(tmpdirname, 'arquivos_importacao')
                    shutil.make_archive(caminho_zip, 'zip', tmpdirname)
                    
                    with open(caminho_zip + ".zip", "rb") as f:
                        st.balloons()
                        if formato_saida == "Arquivo Único (Consolidado)":
                            st.success(f"Sucesso! Gerados {total_arquivos} arquivos consolidados.")
                        else:
                            st.success(f"Sucesso! Gerados {total_arquivos} arquivos individuais.")
                            
                        st.download_button(
                            label="📥 Baixar Resultados (ZIP)",
                            data=f,
                            file_name=nome_zip,
                            mime="application/zip"
                        )
                else:
                    st.warning("Nenhum arquivo foi gerado em nenhuma das abas.")

                if todos_erros:
                    st.warning("Alguns dias ou colaboradores foram pulados. Veja os detalhes abaixo:")
                    with st.expander("Ver Logs de Erros e Dias Pulados"):
                        for e in todos_erros:
                            st.write(e)


