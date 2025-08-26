""" Módulo com as funções que realizam o processo diário do monitoramento.

Inclui as funções que fazem o login e tiram os prints (visão geral, inversores, e gráfico de inversores caso haja) de cada usina de cada site.
Tambem inclui a função que manda email informativo em caso de algo fora do normal ser detectado em alguma usina."""

from playwright.async_api import Browser, Page, expect
from config import *
import yagmail
import asyncio
from typing import Literal, Optional
import random
import logging
import sys
from dados_mensais import *
from config import *


logger = logging.getLogger('Monitoramento')

logger.setLevel(logging.INFO)

file_handler = logging.FileHandler(Path(CAMINHO_PASTA_LOGS, 'monitoramento.log'), mode='a', encoding='utf-8')

file_formatter = logging.Formatter(FORMATACAO_LOGGING)
file_handler.setFormatter(file_formatter)

logger.addHandler(file_handler)

# Configurando o logger que escreve as informações no terminal 
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.ERROR)

console_formatter = logging.Formatter('%(levelname)s: %(message)s')
console_handler.setFormatter(console_formatter)

logger.addHandler(console_handler)



def enviar_email(
    config_do_email: Literal['inversor_offline', 'historico_de_falhas', 'erro_no_codigo', ],
    site: Optional[Literal['Solis', 'Sungrow', 'Solplanet', 'Shine', 'Growatt', 'PHB']] = None, 
    usina: Optional[str] = None, 
    qtd_inversores: Optional[int] = None, 
    erro_capturado: Optional[str] = None,
    onde_ocorreu_erro: Optional[str] = None,
    tipo_da_falha: Optional[Literal['pendente', 'resolvida', 'aviso']] = 'não especificado'
):
    """ Envia um email de aviso para os destinatários, o conteúdo depende da configuração escolhida.
    
    As configurações do email incluem 3 opções: inversor_offline, erro_no_código, histórico_de_falhas.

    A primeira monta a mensagem de aviso de inversores offline a partir dos parâmetros de site, usina e qtd_inversores.
    A segunda informa que houve uma exceção inesperada, montando a mensagem a partir dos parâmetros de erro_capturado e onde_ocorreu_erro.
    A terceira deve ser utilizada após a leitura do histórico de falhas da usina para montar a mensagem com base nos parâmetros de site, usina, falha_identificada, codigo_falha, momento_falha e tipo_falha.
    
    Args:
        config_do_email: deve ser uma das 3 opções disponíveis (inversor_offline, erro_no_código, historico_de_falhas), serve para customizar o que será enviado no email.

        site (str): o nome do site da usina monitorada. usado na primeira e segunda configuração.

        usina (str): nome da usina, usada na primeira e segunda configuração

        qtd_inversores (int): número de inversores offline, usado na primeira configuração.

        erro_capturado (Exception): a exceção que algum trecho de código levantou.

        onde_ocorreu_erro (str): uma breve descrição de onde o erro foi levantado, que será usada para montar o corpo do email.

        tipo_falha (str): informação sobre o tipo da falha que pode ser pendente, resolvida ou por padrão 'não especificado'.
        
    Raises:
        FileNotFoundError: erro levantado caso o caminho para os prints dos inversores (que é anexado junto ao email para maior detalhamento) de determinada usina não for encontrado. Só se aplica na configuração 'inversores_offline'.

        ValueError: erro levantado caso a configuração especificada não seja igual a nenhuma das aceitas (inversores_offline, historico_de_falhas, erro_no_codigo)

    """ 
    if config_do_email == 'inversor_offline':
        assunto = 'Inversor(es) offline'

        corpo_email = f'Aviso! Foi verificado que na usina {site} {usina} há {qtd_inversores} inversores que não estão online.\nMomento da verificação: {DATA_ATUAL} às {HORARIO_ATUAL.hour}:{HORARIO_ATUAL.minute}'

        try:

            if usina == 'PHB':
                # É preciso montar os caminhos para as screenshots dos inversores PHB devido a forma como eles são exibidos (em carrossel de imagens)
                caminho_print_inversores = [Path(CAMINHO_PASTA_PRINTS, site, f'{usina} inversor {n}.png') for n in range(1, 5)]

            else:
                caminho_print_inversores = Path(CAMINHO_PASTA_PRINTS, site, f'{usina} - inversores.png')

            anexo = caminho_print_inversores

        except FileNotFoundError:
            logger.error(f'Não foi possível encontrar o caminho para o print dos inversores da usina {site} - {usina}')


    elif config_do_email == 'historico_de_falhas':
        assunto = 'Falha encontrada no histórico da usina'

        corpo_email = f'Aviso! Foi encontra uma falha no histórico da usina {site} - {usina}.\nFalha {tipo_da_falha}\nMomento da ocorrêcia: {AGORA}'

        anexo = Path(CAMINHO_PASTA_PRINTS, site, 'Falhas', f'falha {usina} - {AGORA.date()}.png').resolve()


    elif config_do_email == 'erro_no_codigo':
        assunto = 'Erro durante a execução do código'

        corpo_email = f'Aviso! O código do monitoramento apresentou o erro abaixo às {HORARIO_ATUAL.hour}:{HORARIO_ATUAL.minute}\n {erro_capturado}\n\nO erro ocorreu durante a execução do(a): {onde_ocorreu_erro}.'


    else:
        logger.error(f'Tentativa de envio de email com configuração inválida (Config: {config_do_email})')
        raise ValueError


    logger.info('Enviando email...')
    try:
        yag = yagmail.SMTP(user=remetente, password=senha_de_app, host=host, port=porta)

        if config_do_email == 'erro_no_codigo':
            yag.send(
                to=destinatario,
                subject=assunto,
                contents=corpo_email,      
            )

        else:
            yag.send(
                to=destinatario,
                subject=assunto,
                contents=corpo_email,
                attachments=anexo
            )

    except Exception as e:
        logger.error(f'Erro durante o envio de email: {e}')

    else:
        logger.info(f'Email enviado com sucesso para o destinatário {destinatario}')



async def resolver_captcha_solplanet(pagina: Page) -> bool:
    """ Resolve o captcha do site Solplanet para poder concluir o login.
    
    Args:
        pagina (Page): A página inicial da solplanet

        locator_arrastavel (Locator): o locator que deve ser arrastado para resolver o captcha

    Raises:
        CoordenadasCaptchaError: esse erro será levantado caso alguma das operações de extração das coordenadas do canvas do captcha ou do locator arrastável falhe.

    Returns:
        bool: retorna True em caso de sucesso na resolução do captcha, False caso contrário.
    
    """
    logger.info('Iniciando resolução do captcha')

    locator_arrastavel = pagina.locator('div.slider-button')
    await locator_arrastavel.wait_for(state='visible')

    box_locator = await locator_arrastavel.bounding_box(timeout=3000)

    if not box_locator:
        logger.error('Não foi possível extrair as coordenadas do seletor arrastável')
        return False


    div_captcha = pagina.locator('div.ant-modal-body')
    await div_captcha.wait_for(state='visible')

    canvas = pagina.locator('div.image-container > canvas.canvas').first
    canvas_box = await canvas.bounding_box(timeout=3000)

    if not canvas_box:
        logger.error('Não foi possível extrair as coordenadas do canvas do captcha')
        return False

    canvas_style = await canvas.get_attribute('style', timeout=3000) 
    canvas_style_left = int(canvas_style.split(';')[0].replace('left: ', '').replace('px', '').strip())

    try:
        meio_vertical_seletor = box_locator['y'] + box_locator['height'] / 2
        meio_horizontal_seletor = box_locator['x'] + box_locator['width'] / 2

        meio_canvas = canvas_box['x'] + canvas_box['width'] / 2

        await pagina.mouse.move(meio_horizontal_seletor, meio_vertical_seletor)
        await asyncio.sleep(0.7)

        await pagina.mouse.down(button='left')

        await pagina.mouse.move(meio_canvas, (meio_vertical_seletor + random.randint(-2, 2)), steps=random.randint(9, 16))

        locator_style = await locator_arrastavel.get_attribute('style')
        locator_style_left = int(locator_style.split(':')[-1].replace('px;', '').strip())

        logger.info(f'Atributo left do canvas: {canvas_style_left} | atributo left do locator: {locator_style_left}')

        if canvas_style_left != locator_style_left:
            diferenca_pixels = canvas_style_left - locator_style_left

            logger.info(f'Diferença: {diferenca_pixels}. Corrigindo o posicionamento...')

            await pagina.mouse.move(meio_canvas + diferenca_pixels, meio_vertical_seletor)

        await asyncio.sleep(0.7)
        await pagina.mouse.up(button='left')

        await expect(pagina).to_have_url('https://internation-pro-cloud.solplanet.net/plant-center/plant-overview-all/plant-overview', timeout=7000)
        return True

    except (TimeoutError, AssertionError):
        return False

    except Exception as e:
        logger.error(f'Erro inesperado ao resolver o captcha Solplanet: {e}')
        enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro='recaptcha Solplanet')



async def gerenciar_tentativas_captcha_solplanet(pagina_login: Page) -> bool:
    """ Faz o gerenciamento das chamadas da função resolver_captcha_solplanet durante algumas tentativas.
    
    A função irá verificar se o captcha foi resolvido e caso não tenha sido irá tentar até no máximo 7 vezes. Em caso de falha o captcha será reiniciado para uma nova imagem. Em caso de sucesso mostra tambem o número de tentativas gastas.

    Args:
        pagina_login (Page): a página de login da Solplanet que será passado para a função de resolução do captcha.

        locator_arrastavel (locator): o locator que possa ser movido pelo mouse para completar o desafio.

    Returns:
        bool: retorna True em caso de sucesso (a função resolver_captcha_solplanet tambem retorne True), False caso contrário (máximo de tentativas atingido).

    """
    tentativas = 6
    tentativa_atual = 0

    while tentativa_atual <= tentativas:
        tentativa_atual += 1
        sucesso = await resolver_captcha_solplanet(pagina_login)

        if sucesso:
            logger.info(f'Sucesso! Captcha resolvido em {tentativa_atual} tentativas!')
            return True

        else:
            if tentativa_atual == tentativas:
                    logger.critical('Máximo de tentativas do captcha Solplanet')
                    enviar_email(config_do_email='erro_no_codigo', erro_capturado='CaptchaError', onde_ocorreu_erro='captcha Solplanet')
                    return False

            logger.warning(f'Tentativa {tentativa_atual}/{tentativas} falhou')
            await asyncio.sleep(1.3)

            await pagina_login.locator('span.reload-tips').filter(has_text='Refresh and re-verify').click()            



async def analisar_status_inversores_solis(pagina: Page, nome_usina: str):
    """ Lê os status dos inversores de determinada usina do site Solis, caso algum não esteja online enviará uma notificação por email.
    
    Args:
        pagina (Page): a página da usina atual

        nome_usina (str): o nome da usina é usado para identificá-la caso alguma problema seja encontrado.

    """
    logger.info(f'Iniciando análise dos status dos inversores da usina Solis - {nome_usina}')

    infos_inversores = await pagina.locator('tbody > tr').all()

    contador = 0
    for inversor in infos_inversores:
        info_inversor = await inversor.text_content()

        info_inversor = info_inversor.replace('  ', ' ').lower().strip().split(' ')
        status = info_inversor[0]

        if status != 'on-line':
            logger.warning(f'O inversor de SN número {info_inversor[2]} da usina Solis - {nome_usina} não está online! Status atual: {status}')
            contador += 1

    logger.info(f'Análise dos status dos inversores da usina Solis - {nome_usina} concluída')

    if contador == 0:
        logger.info(f'Todos os inversores estão online!\n')

    else:
        logger.warning(f'Há {contador} inversores offline.')
        enviar_email(config_do_email='inversor_offline', site='Solis', usina=nome_usina, qtd_inversores=contador)



async def analisar_status_inversores_solplanet(pagina: Page, nome_usina: str):
    """ Lê os status dos inversores de determinada usina do site Solplanet, caso algum não esteja online enviará uma notificação por email.
    
    Args:
        pagina (Page): a página da usina atual

        nome_usina (str): o nome da usina é usado para identificá-la caso alguma problema seja encontrado.

    """ 
    logger.info(f'Iniciando análise dos status dos inversores da usina Solplanet - {nome_usina}')

    contador = 0

    inversores = await pagina.locator('#rc-tabs-1-panel-item-1 div.ant-collapse.ant-collapse-icon-position-start.ant-collapse-ghost').all()

    for inversor in inversores:
        texto_linha = await inversor.locator('tr').last.inner_text()

        if 'normal' not in texto_linha.lower():
            logger.warning(f'Há um inversor offline na usina Solplanet - {nome_usina}')
            contador += 1

    logger.info(f'Análise do status dos inversores  da usina Solplanet - {nome_usina} concluída')

    if contador == 0:
        logger.info(f'Todos os inversores da usina Solplanet - {nome_usina} estão online')

    else:
        enviar_email(config_do_email='inversor_offline', site='Solplanet', usina=nome_usina, qtd_inversores=contador)



async def analisar_status_inversores_sungrow(pagina: Page, nome_usina: str): 
    """ Lê os status dos inversores de determinada usina do site Sungrow, caso algum não esteja online enviará uma notificação por email.
    
    Args:
        pagina (Page): a página da usina atual

        nome_usina (str): o nome da usina é usado para identificá-la caso alguma problema seja encontrado.

    """ 
    logger.info(f'Iniciando análise dos status dos inversores da usina Sungrow {nome_usina}')

    div_cards = await pagina.locator('div.card-container div.container').all()

    contador = 0
    for card in div_cards:
        status = await card.locator('div.isc-tag').text_content()

        if status.lower().replace(' ', '') != 'normal':
            logger.warning(f'Há um inversor que não está online usina Sungrow {nome_usina}!  Status: {status}')
            contador += 1

    logger.info(f'Análise dos status dos inversores da usina Sungrow {nome_usina} concluída')

    if contador == 0:
        logger.info('Todos os inversores estão online!')

    else:
        logger.warning(f'Há {contador} inversores offline')
        enviar_email(config_do_email='inversor_offline', site='Sungrow', usina=nome_usina, qtd_inversores=contador)



async def analisar_status_inversores_phb(pagina: Page, nome_usina: str):
    """ Lê os status dos inversores de determinada usina do site PHB, caso algum não esteja online enviará uma notificação por email..
    
    Args:
        pagina (Page): a página da usina atual

        nome_usina (str): o nome da usina usado para indetificá-la em caso de alguma falha. 

    """
    logger.info(f'Iniciando análise dos status dos inversores da usina Sungrow - {nome_usina}')

    status = await pagina.locator('div.device-status').all()

    contador = 0
    for item in status:
        texto_status = await item.inner_text()

        if texto_status.lower().strip() not in ('trabalhando', 'working'):
            contador += 1
            logger.warning(f'Há um inversor que não está online na usina PHB - {nome_usina}. Status: {item}')

    logger.info(f'Análise do status dos inversores PHB - {nome_usina} concluído')

    if contador == 0:
        logger.info('Todos os inversores estão online')

    else:
        logger.warning(f'Há {contador} inversores offline')
        enviar_email(config_do_email='inversor_offline', site='PHB', usina=nome_usina, qtd_inversores=contador)



async def analisar_status_inversores_growatt(pagina: Page, nome_usina: str):
    """ Lê os status dos inversores de determinada usina do site Growatt, caso algum não esteja online enviará uma notificação por email.
    
    Args:
        pagina (Page): a página da usina atual

        nome_usina (str): o nome da usina usado para indetificá-la em caso de alguma falha. 

    """
    logger.info(f'Iniciando análise dos status dos inversores da usina Growatt - {nome_usina}')

    tabela = await pagina.locator('tbody#inverterRefreshData > tr').all()
    contador = 0

    for linha in tabela:
        texto_linha = await linha.text_content()

        if 'online' not in texto_linha.lower():
            contador += 1
            logger.warning(f'Há um inversor que não está online na usina Growatt - {nome_usina}')

    logger.info(f'Análise dos status dos inversores da usina Growatt - {nome_usina} concluída')

    if contador == 0:
        logger.info(f'Todos os inversores estão online!')

    else:
        logger.warning(f'Há {contador} inversores offline')
        enviar_email(config_do_email='inversor_offline', site='Growatt', usina=nome_usina, qtd_inversores=contador)



async def analisar_status_inversores_shine(pagina: Page, nome_usina):
    """ Lê os status dos inversores de determinada usina do site ShineMonitor (Renovigi), caso algum não esteja online enviará uma notificação por email.
    
    Args:
        pagina (Page): a página da usina atual

        nome_usina (str): o nome da usina usado para indetificá-la em caso de alguma falha. 

    """
    logger.info('Iniciando análise dos status dos inversores')

    div_inversores = await pagina.locator('div#basicInfo div.basic_box_bottom').all()

    contador = 0
    for inversor in div_inversores:
        status_inversor = await inversor.inner_text()

        if 'normal' not in status_inversor.lower().strip():
            contador += 1
            logger.warning(f'Há um inversor que não está online na usina Shine - {nome_usina}. Status: {status_inversor}')

    logger.info(f'Análise dos status dos inversores da usina Shine - {nome_usina} concluída')

    if contador == 0:
        logger.info('Todos os inversores estão online')

    else:
        logger.warning(f'Há {contador} inversores que não estão online')
        enviar_email(config_do_email='inversor_offline', site='Shine', usina=nome_usina, qtd_inversores=contador)



async def analisar_historico_de_falhas_solis(pagina: Page, nome_usina: str):
    """ Lê o histórico de falhas de determinada usina do site Solis, e caso existam falhas irá enviar um email de aviso.
    
    Args:
        pagina (Page): página da usina atual.

        nome_usina (str): o nome da usina para sua identificação caso haja falhas.

    """
    logger.info('Iniciando leitura do histórico de falhas')

    try:
        await pagina.locator('a').filter(has_text='Alarme').click()
        await pagina.wait_for_load_state('networkidle')
        await asyncio.sleep(2.5)

        if await pagina.locator('div.no-data-content').count() > 0:
            logger.info(f'Sem falhas pendentes no histórico da usina {nome_usina}')

        else:
            logger.warning(f'Falha encontrada na usina Solis - {nome_usina}')

            await pagina.locator('div.gl-table-box').screenshot(
                type='png',
                path=Path(CAMINHO_PASTA_PRINTS, 'Solis', 'Falhas', f'falha {nome_usina} - {AGORA.date()}.png')
            )

            enviar_email(
                config_do_email='historico_de_falhas', 
                site='Solis', 
                usina=nome_usina, 
                tipo_da_falha='pendente', 
            )

    except Exception as e:
        logger.error(f'Erro ao ler o histórico de falhas da usina Solis - {nome_usina}: {e}')
        enviar_email('erro_no_codigo', 'Solis', erro_capturado=e, onde_ocorreu_erro='análise do histórico de falhas Solis')         



async def analisar_historico_falhas_solplanet(pagina: Page, nome_usina: str):
    """ Lê o histórico de falhas de determinada usina do site Solplanet, e caso existam falhas irá enviar um email de aviso.
    
    Args:
        pagina (Page): página da usina atual.

        nome_usina (str): o nome da usina para sua identificação caso haja falhas.

    """
    logger.info(f'Iniciando análise do histórico de falhas da usina Solplanet - {nome_usina}')

    await pagina.get_by_role('tab', name='Fault information').click()
    await pagina.wait_for_load_state('networkidle')
    await asyncio.sleep(2.5)

    if await pagina.locator('div.ant-empty-description').count() > 0:
        logger.info(f'Sem avisos pendentes na usina Solplanet - {nome_usina}')

    else:
        logger.warning(f'Falha encontrada na usina Solplanet - {nome_usina}')

        await pagina.locator('div#rc-tabs-2-panel-plantDetailError').screenshot(
            type='png', 
            path=Path(CAMINHO_PASTA_PRINTS, 'Solplanet', 'Falhas', f'falha {nome_usina} - {AGORA.date()}.png').resolve()
            )

        enviar_email(
            config_do_email='historico_de_falhas', 
            site='Solplanet', 
            usina=nome_usina,
            tipo_da_falha='aviso'
        )


async def analisar_historico_de_falhas_sungrow(pagina: Page, nome_usina: str):
    """ Analisa o histórico de falhas de determinada usina do site Sungrow, e caso existam falhas irá enviar um email de aviso.
    
    A função lê o histórico de falhas da usina para o dia atual, pressupondo que a página passada como argumento já está na aba da usina.
    Caso não esteja irá falhar. 
    Caso algo esteja fora do normal, chama outra função que enviará um aviso explicando onde e qual o tipo da falha.

    Args:
        pagina (Page): a página web da usina, considerando que a página já está na aba de 'Dispositivos' para fazer a análise a partir daí.

        nome_usina (str): o nome da usina que está sendo analisada, será usado em caso de localização de alguma falha nos inversores.

    """
    await pagina.locator('span.menu-item-text').filter(has_text='Falha').click()
    await pagina.wait_for_load_state('networkidle')
    await asyncio.sleep(2.5)

    if await pagina.locator('div.empty-container').count() > 0:
        logger.info(f'Sem falhas pendentes no histórico para a usina Sungrow - {nome_usina}')

    else:
        try:
            logger.warning(f'Falha encontrada na usina Sungrow - {nome_usina}')

            await pagina.locator('div#plant-detail-overview-mount-loading-node').screenshot(
                type='png', 
                path=Path(CAMINHO_PASTA_PRINTS, 'Sungrow', 'Falhas', f'falha {nome_usina} - {AGORA.date()}.png')
            )

            enviar_email(
                config_do_email='historico_de_falhas', 
                site='Sungrow', 
                usina=nome_usina,  
                tipo_da_falha='pendente',
            )

        except Exception as e:
            logger.error(f'Erro ao ler o histórico de falhas da usina Sungrow - {nome_usina}: {e}')
            enviar_email('erro_no_codigo', 'Solis', erro_capturado=e, onde_ocorreu_erro='análise do histórico de falhas Solis')



async def analisar_historico_de_falhas_shine(pagina: Page, nome_usina: str):
    """ Lê o histórico de falhas de determinada usina do site Shine, e caso existam falhas irá enviar um email de aviso.
    
    Args:
        pagina (Page): página da usina atual.

        nome_usina (str): o nome da usina para sua identificação caso haja falhas.

    """
    logger.info(f'Analisando histórico de falhas da usina Shine {nome_usina}')

    await pagina.get_by_role('link', name='Alerta').click()
    await pagina.wait_for_load_state('networkidle')
    await asyncio.sleep(2.5)

    tabela_falhas = pagina.locator('tbody#pltWarnsTbody')
    await tabela_falhas.wait_for(state='visible')

    if await tabela_falhas.get_by_role('cell', name='No alarm for equipment').count() > 0:
        logger.info(f'Sem falhas pendentes no histórico da usina {nome_usina}')

    else:
        logger.warning(f'Falha encontrada na usina Shine - {nome_usina}')

        await pagina.locator('div#plantAlarm').screenshot(
            type='png', 
            path=Path(CAMINHO_PASTA_PRINTS, 'Shine', 'Falhas', f'falha {nome_usina} - {AGORA.date()}.png').resolve()
        )

        enviar_email(
           'historico_de_falhas', 
            tipo_da_falha='pendente',
            site='Shine', 
            usina=nome_usina
        )



async def monitoramento_solis(browser: Browser, lista_usinas: list, semaforo: asyncio.Semaphore):
    """ Realiza o monitoramento das usinas do site SolisCloud.
    
    Args:
        browser (Browser): a instância do navegador que será utilizado

        lista_usinas (list): a lista contendo o nome das usinas que serão monitoradas no site.

    """
    async with semaforo:
        logger.info('Iniciando monitoramento Solis...')
        info_solis = sites['Solis']

        async with await browser.new_context(viewport=VIEWPORT_PADRAO) as context:
            pagina_inicial = await context.new_page()

            await pagina_inicial.goto(info_solis['url'])

            print(f'Página inicial Solis aberta')

            try:
                await pagina_inicial.get_by_role('textbox', name='Username/Email').fill(info_solis['login'])
                await pagina_inicial.get_by_role('textbox', name='Palavra-passe').fill(info_solis['senha'])

                await pagina_inicial.locator('label.el-checkbox.el-checkbox--default.el-tooltip__trigger').click()

                await pagina_inicial.get_by_role('button', name="Login").click()

                await pagina_inicial.wait_for_load_state('domcontentloaded')

            except Exception as e:
                logger.critical(f'Erro durante o login da Solis: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro='login do site Solis')

            else:
                logger.info('Login na Solis realizado com sucesso, monitorando as usinas...')

            try:
                for usina in lista_usinas:
                    try:
                        async with pagina_inicial.context.expect_page() as nova_pag:
                            await pagina_inicial.locator('div.station-name', has_text=usina).first.click()

                    except Exception:
                        logger.error(f'Não foi possível encontrar a usina {usina}, continuando para a próxima...')
                        continue

                    pag_usina = await nova_pag.value

                    await pag_usina.wait_for_load_state('networkidle')

                    await pag_usina.screenshot(
                        full_page=True, 
                        type='png', 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Solis', f'{usina} - visão geral.png')
                    )

                    if DATA_ATUAL.day == 1:
                        dados_extraidos = await extrair_dados_mensais_solis(pag_usina, usina)
                        processar_dados_mensais_solis(dados_extraidos, usina)

                    await pag_usina.locator('a').filter(has_text='Dispositivo').click()
                    await asyncio.sleep(2)

                    area_inversores = pag_usina.locator('div#equipment.equipment')
                    await area_inversores.wait_for(state='visible')

                    await asyncio.sleep(1.5)

                    await area_inversores.screenshot(
                        type='png', 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Solis', f'{usina} - inversores.png')
                    )

                    await analisar_status_inversores_solis(pag_usina, usina)
                    await asyncio.sleep(2)

                    await analisar_historico_de_falhas_solis(pag_usina, usina)

                    await pag_usina.close()

            except Exception as e:
                logger.error(f'Erro inesperado durante o monitoramento Solis: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro=f'monitoramento da usina {usina}')

        logger.info('Monitoramento Solis concluído')



async def monitoramento_solplanet(browser: Browser, lista_usinas: list, semaforo: asyncio.Semaphore):
    """ Realiza o monitoramento das usinas do site SoltPlanet.
    
    Args:
        browser (Browser): a instância do navegador que será utilizado

        lista_usinas (list): a lista contendo o nome das usinas que serão monitoradas no site. 

    """
    async with semaforo:
        logger.info('Iniciando monitoramento SoltPlanet...')

        info_soltplanet = sites['Solplanet']

        async with await browser.new_context(viewport=VIEWPORT_PADRAO) as contexto:
            pag_inicial = await contexto.new_page()

            await pag_inicial.goto(info_soltplanet['url'])

            print(f'Página inicial Solplanet aberta')

            try:
                await pag_inicial.get_by_placeholder('Please enter your email address or phone number').fill(info_soltplanet['login'])
                await pag_inicial.get_by_placeholder('Please enter your password').fill(info_soltplanet['senha'])

                await pag_inicial.get_by_role('checkbox').check()

                await pag_inicial.get_by_role('button', name="login").click()
                await asyncio.sleep(0.5)

            except Exception as e:
                logger.critical(f'Erro durante o login da SolPlanet: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro='login do site Solplanet')

            else:
                logger.info('Login realizado, resolvendo o recaptcha...')

            captcha_resolvido = await gerenciar_tentativas_captcha_solplanet(pag_inicial)
            await asyncio.sleep(1)

            if not captcha_resolvido:
                return # o erro já é registrado dentro da função que gerencia as tentativas por isso não é preciso registrar de novo

            try:
                for usina in lista_usinas:
                    async with pag_inicial.context.expect_page() as nova_pag:
                        await pag_inicial.get_by_text(usina).click()

                    pag_usina = await nova_pag.value

                    imagem = pag_usina.get_by_role('img', name='avatar').last
                    await imagem.wait_for(state='visible')

                    await asyncio.sleep(8.5)

                    limitador = pag_usina.locator('div.ant-card-head-title').filter(has_text='Energy flow diagram')

                    area_limite = await limitador.bounding_box()
                    limite_altura = area_limite['y'] - 20 # <- reduzindo 20px para não pegar a borda desse locator

                    await pag_usina.screenshot(
                        type='png', 
                        clip={'x': 0, 'y': 0, 'width': 1920, 'height': limite_altura},
                        path=Path(CAMINHO_PASTA_PRINTS, 'Solplanet', f'{usina} - visão geral.png')
                    )

                    grafico = pag_usina.locator('div#rc-tabs-0-panel-power')
                    await asyncio.sleep(1)

                    await grafico.screenshot(
                        type='png', 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Solplanet', f'{usina} - gráfico.png')
                    )

                    area_inversores = pag_usina.locator('#rc-tabs-1-panel-item-1')
                    await asyncio.sleep(1)

                    await area_inversores.screenshot(
                        type='png', 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Solplanet', f'{usina} - inversores.png')
                    )

                    await analisar_status_inversores_solplanet(pag_usina, usina)
                    await asyncio.sleep(2)

                    await analisar_historico_falhas_solplanet(pag_usina, usina)

                    await pag_usina.close()

            except Exception as e:
                logger.error(f'Erro inesperado durante o monitoramento da usina Solplanet - {usina}: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro=f'monitoramento da usina {usina}')

        logger.info('Monitoramento SoltPlanet concluído com sucesso')



async def monitoramento_sungrow(browser: Browser, lista_usinas: list, semaforo: asyncio.Semaphore):
    """ Realiza o monitoramento das usinas do site ISolarCloud (Sungrow).
    
    Args:
        browser (Browser): a instância do navegador que será utilizado

        lista_usinas (list): a lista contendo o nome das usinas que serão monitoradas no site. 

    """
    async with semaforo:
        logger.info('Iniciando monitoramento Sungrow...')
        info_sungrow = sites['Sungrow']

        async with await browser.new_context(viewport=VIEWPORT_PADRAO) as context:
            pag_inicial = await context.new_page()

            await pag_inicial.goto(info_sungrow['url'])

            print(f'Página inicial Sungrow aberta')

            try:
                await pag_inicial.get_by_placeholder('Conta').fill(info_sungrow['login'])
                await pag_inicial.get_by_placeholder('Senha').fill(info_sungrow['senha'])
                await pag_inicial.get_by_role('button').filter(has_text='Entrar').click()

                await pag_inicial.locator('div.menu-item').filter(has_text='Estação de energia').click()

            except Exception as e:
                logger.critical(f'Erro durante o login da Sungrow: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro='login no site Sungrow')

            else:
                logger.info('Login na Sungrow realizado com sucesso, monitorando as usinas...')

            try:
                for usina in lista_usinas:
                    await pag_inicial.wait_for_load_state('domcontentloaded')

                    await pag_inicial.locator('div.plant-name').filter(has_text=usina).click()

                    await pag_inicial.wait_for_load_state('networkidle')
                    await asyncio.sleep(4.5)

                    await pag_inicial.screenshot(
                        type='png', 
                        full_page=False, 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Sungrow', f'{usina} - visão geral.png')
                    )

                    canvas = pag_inicial.locator('canvas')
                    await asyncio.sleep(2)

                    await canvas.screenshot(
                        type='png', 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Sungrow', f'{usina} - gráfico.png')
                    )

                    if DATA_ATUAL.day == 1:
                        dados_do_mes = await extrair_dados_mensais_sungrow(pag_inicial, usina)
                        processar_dados_mensais_sungrow(dados_do_mes, usina)

                    await pag_inicial.locator('span.menu-item-text').filter(has_text='Dispositivos').click()
                    await pag_inicial.wait_for_load_state('networkidle')

                    area_inversores = pag_inicial.locator('div.card-container')
                    await area_inversores.wait_for(state='visible')
                    await asyncio.sleep(1)

                    await area_inversores.screenshot(
                        type='png', 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Sungrow', f'{usina} - inversores.png')
                    )

                    await analisar_status_inversores_sungrow(pag_inicial, usina)
                    await asyncio.sleep(2)

                    await analisar_historico_de_falhas_sungrow(pag_inicial, usina)

                    await pag_inicial.get_by_text('Estação de energia').nth(1).click()

            except Exception as e:
                logger.error(f'Erro inesperado durante o monitoramento da usina Sungrow - {usina}: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro=f'monitoramento da usina {e}')

        logger.info('Monitoramento Sungrow concluído com sucesso!')



async def monitoramento_growatt(browser: Browser, lista_usinas: list, semaforo: asyncio.Semaphore):
    """ Realiza o monitoramento das usinas do site Growatt.
    
    Args:
        browser (Browser): a instância do navegador que será utilizado

        lista_usinas (list): a lista contendo o nome das usinas que serão monitoradas no site. 

    """
    async with semaforo:
        logger.info('Iniciando monitoramento Growatt...')
        info_growatt = sites['Growatt']

        async with await browser.new_context(viewport=VIEWPORT_PADRAO) as context:
            pag_inicial = await context.new_page()

            await pag_inicial.goto(info_growatt['url'])

            print(f'Página inicial Growatt aberta')

            try:
                await pag_inicial.get_by_placeholder('Usuário').fill(info_growatt['login'])
                await pag_inicial.get_by_placeholder('Senha').fill(info_growatt['senha'])
                await pag_inicial.get_by_role('button', name='Entrar').click()

            except Exception as e:
                logger.critical(f'Erro durante o Login da Growatt: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro='login no site Growatt')

            else:
                logger.info('Login na Growatt realizado com sucesso, monitorando as usinas...')

            try:
                for usina in lista_usinas:
                    async with pag_inicial.context.expect_page() as nova_pag:
                        await pag_inicial.locator('tbody#tbl_data_plant td.plantName').filter(has_text=usina).click(click_count=2, delay=120)

                    pag_usina = await nova_pag.value

                    await pag_usina.wait_for_load_state('networkidle')
                    await asyncio.sleep(2)

                    area_limite = await pag_usina.locator('span').filter(has_text='Device List').bounding_box()
                    limite_altura = area_limite['y'] - 20 # <- reduzindo 20px para não pegar a borda desse locator

                    await pag_usina.screenshot(
                        type='png', 
                        clip={'x': 0, 'y': 0, 'width': 1920, 'height': limite_altura}, 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Growatt', f'{usina} - visão geral.png')
                    )

                    inversores = pag_usina.locator('tbody#inverterRefreshData')
                    await inversores.screenshot(
                        type='png', 
                        path=Path(CAMINHO_PASTA_PRINTS, 'Growatt', f'{usina} - inversores.png')
                    )

                    await analisar_status_inversores_growatt(pag_usina, usina)

                    if DATA_ATUAL.day == 1:
                        await extrair_dados_mensais_growatt(pag_usina, usina)
                        processar_dados_mensais_growatt(usina)

                    await pag_usina.close()

            except Exception as e:
                logger.error(f'Erro inesperado durante o monitoramento da usina Growatt - {usina}: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro=f'monitoramento da usina {usina}')

        logger.info('Monitoramento Growatt concluído com sucesso!')



async def monitoramento_phb(browser: Browser, lista_usinas: list, semaforo: asyncio.Semaphore):
    """ Realiza o monitoramento das usinas do site Solar Portal (PHB).
    
    Args:
        browser (Browser): a instância do navegador que será utilizado

        lista_usinas (list): a lista contendo o nome das usinas que serão monitoradas no site. 

    """
    async with semaforo:
        logger.info('Iniciando monitoramento PHB...')
        info_phb = sites['PHB']

        async with await browser.new_context(viewport=VIEWPORT_PADRAO) as context:
            pag_inicial = await context.new_page()

            await pag_inicial.goto(info_phb['url'])

            print(f'Página inicial PHB aberta')

            for usina in lista_usinas:
                try:
                    await pag_inicial.get_by_role("textbox", name="Endereço de e-mail").fill(info_phb['login_'+ usina], timeout=5000) 
                    await pag_inicial.get_by_role('textbox', name='Por favor, digite sua senha').fill(info_phb['senha_' + usina], timeout=5000)

                    await pag_inicial.locator('input#readStatement').check()

                    await pag_inicial.get_by_role('button', name='Login').click()

                except Exception:
                    try:
                        # Em caso de não haver correspondencia dos locators devido a linguagem, o código tenta executar com os nomes em ingles
                        await pag_inicial.get_by_role("textbox", name="Email Address").fill(info_phb['login_'+ usina], timeout=5000) 
                        await pag_inicial.get_by_role('textbox', name='Please enter your password').fill(info_phb['senha_' + usina], timeout=5000)

                        await pag_inicial.locator('input#readStatement').check()
                        await pag_inicial.get_by_role('button', name='Log In').click()

                    except Exception as e:
                        logger.critical(f'Erro durante o login PHB: {e}')                             
                        enviar_email(config_do_email='erro_no_codigo', site='PHB', erro_capturado=e, onde_ocorreu_erro='login no site PHB')
                        return

                    else:
                        logger.info(f'Login na PHB realizado com sucesso, monitorando as usinas...')

                await pag_inicial.wait_for_load_state('networkidle')

                try:
                    grafico = pag_inicial.locator('canvas').last
                    await grafico.wait_for(state='visible', timeout=15000)
                    await asyncio.sleep(2) # Aguardando o gráfico de geração estar visível e acabar as animações

                    div_inversores = pag_inicial.locator('div.row.foot-row')
                    area_inversores = await div_inversores.bounding_box()

                    await pag_inicial.screenshot(
                        type='png', 
                        clip= {'x': 0,'y': 0,'width': 1920,'height': area_inversores['y']}, 
                        path=Path(CAMINHO_PASTA_PRINTS, 'PHB', f'{usina} - visão geral.png')
                    )

                    await div_inversores.wait_for(state='attached')

                    for n in range(1, 5):
                        await asyncio.sleep(0.8)

                        await div_inversores.screenshot(
                            type='png', 
                            path=Path(CAMINHO_PASTA_PRINTS, 'PHB', f'{usina} - inversor {n}.png')
                        )

                        await div_inversores.hover()

                        await pag_inicial.locator('div#data_carousel i.el-icon-arrow-right').click(force=True)
                        await asyncio.sleep(0.8)

                    await analisar_status_inversores_phb(pag_inicial, usina)

                    if DATA_ATUAL.day == 1:
                        dados = await extrair_dados_mensais_phb(pag_inicial, usina)
                        processar_dados_mensais_phb(dados, usina)

                    await pag_inicial.get_by_role('link', name='Sair').click()

                    botao_confirmar = pag_inicial.get_by_role('button', name='Cofirmar')

                    await expect(botao_confirmar).to_be_visible()
                    await botao_confirmar.click()

                except Exception as e:
                    logger.error(f'Erro durante o monitoramento da usina PHB {usina}: {e}')              
                    enviar_email(config_do_email='erro_no_codigo', erro_capturado=e, site='PHB', usina=usina, onde_ocorreu_erro=f'monitoramento da usina {usina}')

        logger.info('Monitoramento PHB concluído com sucesso!')



async def monitoramento_shine(browser: Browser, lista_usinas: list, semaforo: asyncio.Semaphore):
    """ Realiza o monitoramento das usinas do site ShineMonitor.
    
    Args:
        browser (Browser): a instância do navegador que será utilizado

        lista_usinas (list): a lista contendo o nome das usinas que serão monitoradas no site.

    """
    async with semaforo:
        logger.info('Iniciando monitoramento Shine...')
        info_shine = sites['Shine']

        async with await browser.new_context(viewport=VIEWPORT_PADRAO, ignore_https_errors=True) as context:
            pag_inicial = await context.new_page()

            await pag_inicial.goto(info_shine['url'])

            print(f'Página inicial Shine aberta!')

            try:
                await pag_inicial.get_by_placeholder('Digite o nome do usuário').fill(info_shine['login'])
                await pag_inicial.get_by_placeholder('Por favor, digite sua senha').fill(info_shine['senha'])
                await pag_inicial.locator('div#loginbtn').filter(has_text='Login').click()

                await pag_inicial.wait_for_load_state('networkidle')

            except Exception as e:
                logger.critical(f'Erro durante o login da Shine: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro='login no site Shine')

            else:
                logger.info('Login na Shine realizado com sucesso, monitorando as usinas...')

            try:
                await pag_inicial.wait_for_load_state('networkidle')
                await asyncio.sleep(1)

                await pag_inicial.screenshot(
                    type='png', 
                    full_page=True, 
                    path=Path(CAMINHO_PASTA_PRINTS, 'Shine', 'UFV - Faz Fundão - visão geral.png')
                )

                geracao_total = await pag_inicial.locator('strong#stats03').text_content()

                await pag_inicial.get_by_text('Visão Geral da Geração de Energia').click()
                await pag_inicial.get_by_role('link', name='Energia Mês').click()

                grafico = pag_inicial.locator('div#MonthContainer')

                await grafico.wait_for(state='visible')
                await asyncio.sleep(2)

                await grafico.screenshot(
                    type='png', 
                    path=Path(CAMINHO_PASTA_PRINTS, 'Shine', 'UFV - Faz Fundão - inversores.png')
                )

                await analisar_status_inversores_shine(pag_inicial, 'UFV - Faz Fundão')
                await asyncio.sleep(2)

                if DATA_ATUAL.day == 1:
                    dados = await extrair_dados_mensais_shine(pag_inicial, 'UFV - Faz Fundão')
                    processar_dados_mensais_shine(dados, 'UFV - Faz Fundão', geracao_total)

                await analisar_historico_de_falhas_shine(pag_inicial, 'UFV - Faz Fundão')

            except Exception as e:
                logger.error(f'Erro inesperado durante o monitoramento Shine: {e}')
                enviar_email('erro_no_codigo', erro_capturado=e, onde_ocorreu_erro=f'monitoramento da usina UFV - Faz Fundão')

        logger.info('Monitoramento Shine concluído com sucesso!')
