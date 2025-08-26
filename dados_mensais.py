""" Este módulo contém as funções que extram e as que processam e registram os dados mensais de cada usina de cada site monitorado."""

import xlrd
from playwright.async_api import Page
import logging
import asyncio
from pathlib import Path
from typing import Optional
import json
from config import *

logger = logging.getLogger('Dados mensais')

logger.setLevel(logging.INFO)

file_handler = logging.FileHandler(Path(CAMINHO_PASTA_LOGS, 'dados_mensais.log'), mode='a', encoding='utf-8')

file_formatter = logging.Formatter(FORMATACAO_LOGGING)
file_handler.setFormatter(file_formatter)

logger.addHandler(file_handler)



async def extrair_dados_mensais_solis(pagina_usina: Page, nome_usina: str) -> tuple:
    logger.info(f'Extraindo os dados mensais da usina {nome_usina}')

    await pagina_usina.get_by_role('button', name='Mês').click()
    await pagina_usina.get_by_role("button", name="").click()

    await asyncio.sleep(1.2)

    infos_mes = await pagina_usina.locator('div.feature-content div.grid-connected-box > div.grid-connected-item').all_inner_texts()

    rendimentos_totais = await pagina_usina.locator('div.electrical-info-item').all_inner_texts()

    return infos_mes, rendimentos_totais



def processar_dados_mensais_solis(dados_extraidos: tuple[list[str], list[str]], nome_usina: str):
    pasta_falhas = Path(CAMINHO_PASTA_PRINTS, 'Solis', 'Falhas')

    contador_falhas = sum(1 for item in pasta_falhas.iterdir() if item.is_file() and item.name.startswith(f'falha {nome_usina}'))

    dados_processados = {
        'Usina': nome_usina,
        'Interferências': contador_falhas
    }

    dados_mes, dados_totais = dados_extraidos

    for dado in dados_mes:
        dado_formatado = dado.split('\n')

        descricao = dado_formatado[0]
        valor = dado_formatado[-1]

        dados_processados[descricao] = valor

    for dado in dados_totais:
        dado_formatado = dado.split('\n\n')

        descricao = dado_formatado[0]

        if descricao == 'Rendimento mensal':
            continue

        valor = dado_formatado[1]

        dados_processados[descricao] = valor

    caminho_arquivo = Path(CAMINHO_PASTA_DADOS_MENSAIS, 'Solis', f'dados das usinas Solis mês {AGORA.month-1}.json')

    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo_json:
            dados_atuais = json.load(arquivo_json)

    except FileNotFoundError:
        dados_atuais = [] # Se o arquivo não existir, começa com uma lista vazia

    dados_atuais.append(dados_processados)

    with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_json:
        json.dump(dados_atuais, arquivo_json, ensure_ascii=False, indent=4)



async def extrair_dados_mensais_solplanet(pagina: Page, nome_usina: str):
    pass



def processar_dados_mensais_solplanet(dados_extraidos: dict, nome_usina: str):
    pass



async def extrair_dados_mensais_sungrow(pagina_usina: Page, nome_usina: str) -> tuple:
    logger.info(f'Extraindo os dados mensais de geração e receita da usina Sungrow - {nome_usina}')

    await pagina_usina.get_by_role('tab', name='Mensal').click()
    await asyncio.sleep(2) # delay para esperar as animações terminarem

    await pagina_usina.locator('div.date-select-pannel > span.iconfont.icon-a-G2_Leftarrow_20').click()
    await asyncio.sleep(1.5)

    info_geracao = await pagina_usina.locator('div.indicator-area').inner_text()

    await pagina_usina.get_by_role('tab', name='Vida útil').click()
    await asyncio.sleep(2)

    info_totais = await pagina_usina.locator('div.indicator-area').inner_text()

    return info_geracao, info_totais



def processar_dados_mensais_sungrow(dados_extraidos: tuple[str, str], nome_usina: str):
    pasta_falhas = Path(CAMINHO_PASTA_PRINTS, 'Sungrow', 'Falhas')

    contador_falhas = sum(1 for item in pasta_falhas.iterdir() if item.is_file() and item.name.startswith(f'falha {nome_usina}'))

    dados_processados = {
        'Usina': nome_usina,
        'Interferências': contador_falhas
    }

    infos_mes, infos_totais = dados_extraidos

    dados_mes = infos_mes.split('\n')

    qtd_geracao_mes = dados_mes[4]
    unidade_geracao_mes = dados_mes[5]

    dados_processados['Rendimento mensal'] = f'{qtd_geracao_mes} {unidade_geracao_mes}'

    dados_totais = infos_totais.split('\n')

    qtd_geracao_total = dados_totais[4]
    unidade_geracao_total = dados_totais[5]

    dados_processados['Rendimento total'] = f'{qtd_geracao_total} {unidade_geracao_total}'

    caminho_arquivo = Path(CAMINHO_PASTA_DADOS_MENSAIS, 'Sungrow', f'dados das usinas Sungrow mês {AGORA.month-1}.json')

    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo_json:
            dados_atuais = json.load(arquivo_json)

    except FileNotFoundError:
        dados_atuais = [] # Se o arquivo não existir, começa com uma lista vazia

    dados_atuais.append(dados_processados)

    with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_json:
        json.dump(dados_atuais, arquivo_json, ensure_ascii=False, indent=4)




async def extrair_dados_mensais_phb(pagina_usina: Page, nome_usina: str) -> tuple:
    logger.info(f'Extraindo dados mensais da usina {nome_usina}')

    mes = f'0{DATA_ATUAL.month -1}' if DATA_ATUAL.month < 10 else str(DATA_ATUAL.month -1)
    data_procurada = f'{mes}.{DATA_ATUAL.year}'

    geracao_total = await pagina_usina.locator('div.kpi-item.kpi-power.total-power ').filter(has_text='Geração Total').text_content()

    await pagina_usina.get_by_text('Geração de Energia&Renda').click()

    await pagina_usina.get_by_text('Mês').click()

    await pagina_usina.wait_for_load_state('networkidle')

    await pagina_usina.locator('div.goodwe-station-charts__export.fr').click()
    await asyncio.sleep(3)

    linhas = await pagina_usina.locator('table.el-table__body tbody > tr').all()

    geracao_mes = ''
    for linha in linhas:
        celulas = await linha.locator('td').all()

        data = await celulas[0].text_content()

        if data == data_procurada:
            geracao_mes = await celulas[-2].text_content()

            return geracao_mes, geracao_total



def processar_dados_mensais_phb(dados_extraidos: tuple[str, str], nome_usina: str):

    dados_mes, dados_totais = dados_extraidos

    geracao_total = dados_totais.replace('Geração Total ', '').strip()

    dados_processados = {
        'Usina': nome_usina,
        'Rendimento mensal': f'{dados_mes} kWh',
        'Rendimento total': geracao_total
    }

    caminho_arquivo = Path(CAMINHO_PASTA_DADOS_MENSAIS, 'PHB', f'dados das usinas PHB mês {AGORA.month-1}.json')

    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo_json:
            dados_atuais = json.load(arquivo_json)

    except FileNotFoundError:
        dados_atuais = [] # Se o arquivo não existir, começa com uma lista vazia

    dados_atuais.append(dados_processados)

    with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_json:
        json.dump(dados_atuais, arquivo_json, ensure_ascii=False, indent=4)



async def extrair_dados_mensais_growatt(pagina_usina: Page, nome_usina: str):
    logger.info(f'Iniciando a extração dos dados mensais da usina {nome_usina}')

    await pagina_usina.locator('ul.dateSelectUl1 > li').filter(has_text='Month').click()

    await pagina_usina.get_by_role('button', name='Export').first.click()
    await asyncio.sleep(1)

    await pagina_usina.get_by_placeholder('Please select the month').clear()

    await pagina_usina.get_by_placeholder('Please select the month').fill(f'{DATA_ATUAL.year}-{DATA_ATUAL.month-1}')
    await pagina_usina.get_by_text('Export data').click()

    async with pagina_usina.context.expect_page() as nova_pag:
        await pagina_usina.locator('span.all-bottom-btn0').filter(has_text='Export').click(force=True)

    pagina_download = await nova_pag.value

    await pagina_download.wait_for_load_state('networkidle')

    async with pagina_download.expect_download() as download_info:
        await pagina_download.get_by_role('button', name='Enviar mesmo assim').click(force=True, timeout=6000)

    download = await download_info.value

    await pagina_download.close()

    failure = await download.failure()

    if failure:
        logger.error(f"ERRO: O download do arquivo com as informações mensais falhou: {failure}")

    else:
        await download.save_as(Path(CAMINHO_PASTA_DADOS_MENSAIS, 'Growatt', f'{nome_usina} mês {DATA_ATUAL.month-1}.xls'))
        logger.info('Download concluído com sucesso')



def processar_dados_mensais_growatt(nome_usina: str):
    """ Faz o processamento dos dados brutos do mês que foram extraídos pela função extrair_dados_mensais_growatt e os insere em um txt informativo.
    
    Args:
        nome_usina (str): o nome da usina para sua identificação e seleção do arquivo correto.
    
    """
    logger.info('Processanado os dados mensais...')

    workbook = xlrd.open_workbook(Path(CAMINHO_PASTA_DADOS_MENSAIS, 'Growatt', f'{nome_usina} mês {DATA_ATUAL.month-1}.xls'), formatting_info=True)

    sheet = workbook.sheet_by_index(0)

    geracao_mes = sheet.cell_value(rowx=9, colx=5)
    geracao_total = sheet.cell_value(rowx=10, colx=5)
    ganho_mensal = sheet.cell_value(rowx=11, colx=5)
    ganhos_totais = sheet.cell_value(rowx=12, colx=5)

    dados_processados = {
        'Usina': nome_usina,
        'Rendimento mensal': geracao_mes,
        'Rendimento Total': geracao_total,
        'Ganho mensal': ganho_mensal,
        'Ganho total': ganhos_totais
    }

    caminho_arquivo = Path(CAMINHO_PASTA_DADOS_MENSAIS, 'Growatt', f'dados das usinas Growatt mês {AGORA.month-1}.json')

    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo_json:
            dados_atuais = json.load(arquivo_json)

    except FileNotFoundError:
        dados_atuais = [] # Se o arquivo não existir, começa com uma lista vazia

    dados_atuais.append(dados_processados)

    with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_json:
        json.dump(dados_atuais, arquivo_json, ensure_ascii=False, indent=4)



async def extrair_dados_mensais_shine(pagina_usina: Page, nome_usina: str) -> str:
    """ Extrai dados de geração no gráfico do site Shine Monitor do mês anterior ao mês atual e os insere em um json informativo.

    Args:
        pagina_usina (Page): a página da usina que está sendo monitorada no momento. Deve já estar logada e na tela inicial.

        nome_usina (str): o nome da usina para bucar o arquivo txt onde as informações serão inseridas.

        geracao_total (str | float): a informação de geração total até o momento. Como essa informação precisa ser extraída em outra aba da página que não está no escopo dessa função esse valor deve ser passado como argumento para ser inserido no json de informações mensais.

    """
    logger.info(f'Extraindo dados mensais da usina Shine - {nome_usina}')

    data_procurada = f'{DATA_ATUAL.year}-0{DATA_ATUAL.month-1}' if DATA_ATUAL.month < 10 else f'{DATA_ATUAL.year}-{DATA_ATUAL.month-1}'

    await pagina_usina.get_by_role('link', name='Energia Ano').click()

    box = await pagina_usina.locator('div#yearContainer canvas').last.bounding_box()

    if not box:
        logger.error('Não foi possível extrair as coordenadas do canvas')
        return

    canvas_width = box['width']
    canvas_height = box['height']
    canvas_x = box['x']
    canvas_y = box['y']

    passo = 89 # Tamanho de cada barra do gráfico

    for x in range(0, int(canvas_width), passo):
        await pagina_usina.mouse.move(canvas_x + x, canvas_y + canvas_height / 2)
        await asyncio.sleep(0.3)

        tooltip = pagina_usina.locator('div#yearContainer div.echarts-tooltip.zr-element')

        if await tooltip.is_visible():
            texto_tooltip = await tooltip.text_content()

            if texto_tooltip.startswith(data_procurada):

                texto_split = texto_tooltip.split(':')
                geracao_mes = texto_split[-1]

                return geracao_mes



def processar_dados_mensais_shine(geracao_mensal: str, nome_usina: str, geracao_total: Optional[str | float] = None):
    pasta_falhas = Path(CAMINHO_PASTA_PRINTS, 'Shine', 'Falhas')

    contador_falhas = sum(1 for item in pasta_falhas.iterdir() if item.is_file() and item.name.startswith(f'falha {nome_usina}'))

    dados_processados = {
        'Usina': nome_usina,
        'Interferências': contador_falhas,
        'Rendimento mensal': geracao_mensal.strip(),
        'Rendimento total': geracao_total                  
    }

    caminho_arquivo = Path(CAMINHO_PASTA_DADOS_MENSAIS, 'Shine', f'dados das usinas Shine mês {AGORA.month-1}.json')

    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo_json:
            dados_atuais = json.load(arquivo_json)

    except FileNotFoundError:
        dados_atuais = [] # Se o arquivo não existir, começa com uma lista vazia

    dados_atuais.append(dados_processados)

    with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_json:
        json.dump(dados_atuais, arquivo_json, ensure_ascii=False, indent=4) 
