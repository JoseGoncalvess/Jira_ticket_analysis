import xml.etree.ElementTree as ET
import os
import pandas as pd
from datetime import datetime





def convert_to_date(date_string) -> str:
    if not isinstance(date_string, str):
        return date_string
    try:
        formato_original = "%a, %d %b %Y %H:%M:%S %z"
        objeto_data = datetime.strptime(date_string, formato_original)
        formato_desejado = "%d/%m/%Y"
        return objeto_data.strftime(formato_desejado)
    except (ValueError, TypeError):
        return date_string  
    

def limpar_cto(valor_celula):
    # Verificação de segurança: se a célula estiver vazia ou não for texto, retorna como está
    if not isinstance(valor_celula, str):
        return valor_celula

    # 1. Separa por vírgula
    itens = valor_celula.split(',')

    # 2. Filtra: Mantém o item SE "CTO" NÃO estiver nele (usando lower() para pegar cto, CTO, Cto)
    # O .strip() remove espaços extras que possam ter sobrado
    itens_filtrados = [item.strip() for item in itens if "cto" not in item.lower()]

    # 3. Junta de volta com vírgula
    return ",".join(itens_filtrados)



def processar_arquivo_xml(caminho_do_arquivo, dados_agregados, log_callback):
   
    file = os.path.basename(caminho_do_arquivo)
    log_callback(f"---> Processando arquivo: {file}")
    try:
        tree = ET.parse(caminho_do_arquivo)
        root = tree.getroot()
    except ET.ParseError as e:
        log_callback(f"     [ERRO] Arquivo XML mal formatado: {e}")
        return

    lista_de_itens = root.findall('.//item')
    
  
    chamados_neste_arquivo = []
    for item in lista_de_itens:
        titulo_node = item.find('title')
        data_crated_node = item.find('created')
        link_node = item.find('link')
        key_node = item.find('key')
        status_node = item.find('status')
        response_node = item.find('assignee')
        
        titulo = titulo_node.text.strip() if titulo_node is not None and titulo_node.text else "Título não encontrado"
        dataCreated = data_crated_node.text if data_crated_node is not None else ""
        link = link_node.text if link_node is not None else ""
        key = key_node.text if key_node is not None else "Key não encontrado"
        status = status_node.text if status_node is not None else "Status não encontrado"
        responsavel = response_node.text if response_node is not None else "Responsavel não encontrado não encontrado"

        chamados_vinculados = []
        cto_value=''
        issuelinks_node = item.find('issuelinks')
        if issuelinks_node is not None:
            for link_node in issuelinks_node.findall('.//issuekey'):
                if link_node.text.__contains__("CTO"):
                        cto_value = link_node.text
                if link_node.text:
                    chamados_vinculados.append(link_node.text)
        # chamados_neste_arquivo.append({
        #     "Tipo":key.split("-")[0],
        #     'Titulo': titulo,
        #     'Vinculados': chamados_vinculados,
        #     'Criado em': convert_to_date(dataCreated),
        #     'Link': link
        # })

        chamados_neste_arquivo.append({
            'Criação': convert_to_date(dataCreated),
            'status':status,
            "Chave":key,
            "Alteração de Status":"",
            'Vinculado (AVB/MOB/FLUX)':", ".join(chamados_vinculados),
            "Responsável":responsavel,
            "CTO":cto_value,
            'Link': link
        })



    log_callback(f"     Encontrados {len(chamados_neste_arquivo)} chamados no arquivo {file}.")
    
    
    for tiket in chamados_neste_arquivo:
        # list_of_vinc = tiket["Vinculados"]
        list_of_vinc = tiket["Vinculado (AVB/MOB/FLUX)"]
        
        for key in dados_agregados.keys():
            if key in list_of_vinc:
                dados_agregados[key].append(tiket)


def criar_planilhas_por_empresa(dados,log_callback,caminho_saida):
   
    log_callback("\n--- Iniciando a criação das planilhas finais ---")
 
    pasta_relatorios = caminho_saida if caminho_saida else "Relatorios_Jira" 
    
    pasta_relatorios = f"{pasta_relatorios}//Relatorios_Jira"

    if not os.path.exists(pasta_relatorios):
        os.makedirs(pasta_relatorios)
        log_callback(f"Pasta '{pasta_relatorios}' criada.")

    for id_empresa, lista_de_chamados in dados.items():
        if not lista_de_chamados:
            log_callback(f"- ID {id_empresa}: Não possui chamados. Pulando...")
            continue
            
        log_callback(f"- ID {id_empresa}: Encontrados {len(lista_de_chamados)} chamados. Gerando planilha...")
        df = pd.DataFrame(lista_de_chamados)
        

        df['Vinculado (AVB/MOB/FLUX)'] = df['Vinculado (AVB/MOB/FLUX)'].apply(limpar_cto)
        ### Garante que as colunas existam antes de manipulá-las
        # if 'Vinculados' in df.columns:
        #     df['Vinculados'] = df['Vinculados'].apply(lambda x: ', '.join(x) if isinstance(x, list) else '')

        nome_arquivo = f"Relatorio_{id_empresa}.xlsx"
        caminho_completo = os.path.join(pasta_relatorios, nome_arquivo)
        df.to_excel(caminho_completo, index=False)
        log_callback(f"  -> Planilha '{caminho_completo}' criada com sucesso!")

    log_callback("\nProcesso finalizado.")