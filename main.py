from playwright.async_api import async_playwright
from time import perf_counter
import asyncio
import logging
import sys
from config import *
from organizacao_prints import *
from monitoramento import *
from pathlib import Path

logger = logging.getLogger('Main')

# Configurando o logger que manda os registros para um arquivo .log
logger.setLevel(logging.INFO)

file_handler = logging.FileHandler(os.path.join(CAMINHO_PASTA_LOGS, 'main.log'), mode='a', encoding='utf-8')

file_formatter = logging.Formatter(FORMATACAO_LOGGING)
file_handler.setFormatter(file_formatter)

logger.addHandler(file_handler)

# Configurando o logger que escreve as informações no terminal 
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.ERROR) # Este handler só captura ERROR e níveis superiores.

console_formatter = logging.Formatter('%(levelname)s: %(message)s')
console_handler.setFormatter(console_formatter)

logger.addHandler(console_handler)



async def main():
    inicio = perf_counter()

    logger.info('MONITORAMENTO INICIADO')

    print('----- Monitoramento iniciado... -----')

    mapeamento_site_usinas = {
        'Solis': ['Usina 1', 'Usina 2', 'Usina 3'],

        'Sungrow': ['Usina 4', 'Usina 5'],

        'Growatt': ['Usina 6', 'Usina 7'],

        'PHB': ['Usina 8, Usina 9'],

        'Solplanet': ['Usina 10', 'Usina 11'],

        'Shine': ['Usina 12']
    }

    semaphore = asyncio.Semaphore(2)

    try:
        if DATA_ATUAL.day == 1 and HORARIO_ATUAL.hour == 6:
            for site in mapeamento_site_usinas.keys():
                for usina in mapeamento_site_usinas[site]:
                    caminho_docx = Path(CAMINHO_PASTA_RAIZ, 'Histórico de monitoramentos', site, f'{usina} - mês {DATA_ATUAL.month}.docx')

                    if not caminho_docx.exists():
                        criar_docx_monitoramentos(nome_usina=usina, site=site)


        async with async_playwright() as pw:
            chrome = await pw.chromium.launch(headless=False)

            tasks = [
                asyncio.create_task(monitoramento_solis(chrome, mapeamento_site_usinas['Solis'], semaphore)),

                asyncio.create_task(monitoramento_solplanet(chrome, mapeamento_site_usinas['Solplanet'], semaphore)),

                asyncio.create_task(monitoramento_phb(chrome, mapeamento_site_usinas['PHB'], semaphore)),

                asyncio.create_task(monitoramento_growatt(chrome, mapeamento_site_usinas['Growatt'], semaphore)),

                asyncio.create_task(monitoramento_shine(chrome, mapeamento_site_usinas['Shine'], semaphore)),

                asyncio.create_task(monitoramento_sungrow(chrome, mapeamento_site_usinas['Sungrow'], semaphore))
            ]

            await asyncio.gather(*tasks)      


        if HORARIO_ATUAL >= HORARIO_PARA_INSERIR_PRINTS:
            screenshots = organizar_screenshots(mapeamento_site_usinas)
            inserir_prints_docx(mapeamento_site_usinas, screenshots)

    except KeyboardInterrupt:
        print('Execução interrompida pelo usuário')

    except Exception as e:
        logger.critical(f'Erro inesperado capturado na função main: {e}')
        enviar_email(config_do_email='erro_no_codigo', erro_capturado=e, onde_ocorreu_erro='erro inesperado capturado na função main')

    else:
        print('\n----- Monitoramento concluído! -----')

    finally:
        fim = perf_counter()

        tempo = fim - inicio

        print(f'Tempo de execução: {tempo:.4f} segundos.')

        logger.info(f'MONITORAMENTO FINALIZADO COM SUCESSO EM {tempo:.4f} SEGUNDOS')


if __name__ == "__main__":
    asyncio.run(main())
