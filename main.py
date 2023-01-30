import os.path
import re
import locale
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import timedelta
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException

for j in range(4, 3, -1):
    # seta a linguagem padrão para o Brasil. Será utilizado para o formato de data
    locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')

    # Inicializa o executável do web driver linkado ao Chrome
    browser = webdriver.Chrome(r'.\chromedriver.exe')

    # Pega o link da página que queremos acessar
    browser.get('https://app.octadesk.com/login')
    browser.maximize_window()

    # Encontra o campo de login, importante olhar a estrutura HTML do site e encontrar o tipo para o XPATH
    input_login = browser.find_element(By.XPATH, "//input[@type='text']")

    # Pega a variável input_login e insere credencial de email nela
    input_login.send_keys('joaofernandes.kaizen@gmail.com')

    # Encontra o local da senha e coloca a senha
    input_senha = browser.find_element(By.XPATH, "//input[@type='password']")
    input_senha.send_keys('KznMUZ@123')

    # Encontra o botão de login e clica nele
    botao_login = browser.find_element(By.XPATH, "//button[@type='button']")
    botao_login.click()

    # Clica na aba de conversas Encerradas
    time.sleep(10)
    browser.switch_to.frame(0)
    time.sleep(1)
    encerradas = WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                                              "button[class='_btn_1gyyu_1 "
                                                                              "_btn-theme-primary_1gyyu_98"
                                                                              " _btn-link_1gyyu_40 "
                                                                              "ChatSidebarMenu___tabs_cRCTf']")))
    encerradas.click()

    # Clica nos filtros
    filtros = browser.find_element(By.CSS_SELECTOR, "button[class='_btn-icon_1ke0k_1 _btn-icon-theme-info_1ke0k_21']")
    filtros.click()

    # Clica no campo de data
    data_campo = browser.find_element(By.CSS_SELECTOR, "input[placeholder='Selecione uma data ou período']")
    data_campo.click()

    # Clica no botão personalizado
    data_personalizado = browser.find_element(By.CSS_SELECTOR, "button[class='_btn_1gyyu_1 _btn-theme-primary_1gyyu_98 "
                                                               "_btn-sm_1gyyu_51 _btn-outlined_1gyyu_22 "
                                                               "_btn-block_1gyyu_37']")
    data_personalizado.click()

    # Pega a data de hoje e coloca o período a ser filtrado, transformando em string para que a tag seja encontrada
    data = date.today() - timedelta(days=j)
    data_ext = data.strftime("%A, %d de %B de %Y")
    # Checa se há 2 zeros e cria a string de data corretamente, excluindo as dezenas
    if data_ext.count("0") == 2 and data_ext[-21] != "0":
        data_correta = data_ext.replace('0', '', 1)
    else:
        data_correta = data_ext

    # Clica na data para fazer a filtragem; checa se a tag da data está no mês atual ou anterior
    if data.month == date.today().month:
        # módulo parent pega a div anterior ao span, que é o "pai" dela
        data_filtro = browser.find_element(By.XPATH,
                                           "//span[@aria-label='" + data_correta + "'][@tabindex='-1']/parent::*")
    else:
        data_filtro = browser.find_element(By.XPATH, "//span[@aria-label='" + data_correta + "']/parent::*")

    data_filtro.click()
    data_filtro.click()

    # Aplica o filtro
    filtro_tag = browser.find_element(By.CSS_SELECTOR, "button[class='_btn_1gyyu_1 "
                                                       "_btn-theme-primary_1gyyu_98 _btn-lg_1gyyu_56']")
    filtro_tag.click()

    # Fecha a caixa de filtros
    fecha = browser.find_element(By.CSS_SELECTOR, "button[aria-label='close']")
    fecha.click()

    # Clica nos botões para carregar as conversas

    time.sleep(2)
    # Clica em todos os botões de carregar mais da página
    while True:
        try:
            browser.implicitly_wait(3)
            botao_carregar_mais = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                                                               "button[class='_btn_1gyyu_1"
                                                                                               " _btn-theme-primary_"
                                                                                               "1gyyu_98 "
                                                                                               "_btn-sm_1gyyu_51 "
                                                                                               "_btn-outlined_1gyyu_22 "
                                                                                               "DefaultList_btn-more"
                                                                                               "_2zcP-']")))
        except TimeoutException as error:
            break

        botao_carregar_mais.click()

    # Pega a altura da página em pixels para saber a quantidade de conversas
    origem = browser.find_element(By.XPATH, '//div[@class="ChatSidebarList_sidebar-list-container_1y6Bh"]'
                                            '[@data-cy="conversation_list"]')
    altura_pag = int(origem.get_attribute('scrollHeight'))
    scroll_origem = ScrollOrigin.from_element(origem)  # Seta a origem do scroll para o início da página
    ActionChains(browser).scroll_from_origin(scroll_origem, 0, -altura_pag).perform()  # Scrolla para o topo

    # Função para análise de palavras em uma conversa com expressão regular
    def analise_vendedor_geral(mens):
        try:
            mensagem = WebDriverWait(browser, 10).until(EC.presence_of_all_elements_located((By.XPATH,
                                                                                             '//div[@class='
                                                                                             '"ChatConversationMessage'
                                                                                             '_chat-conversation-'
                                                                                             'message_3F2m0 '
                                                                                             'ChatConversationMessage'
                                                                                             '_agent-message'
                                                                                             '_2upwR"]//descendant::'
                                                                                             'span[@class='
                                                                                             '"MessageText_message'
                                                                                             '-text_1ycuI"]')))

            for i in range(len(mens)):
                for texto in mensagem:
                    teste = re.search(mens[i], texto.text)
                    if teste is not None:
                        string = 'Sim'
                        return string
            else:
                string = 'Não'
                return string
        except TimeoutException:
            string = ""
            return string

    # Peça faltante colocar 'possa achar', em orçamento colocar o span da imagem
    def analise_orcamento(imagem_check):
        # Acha orçamento pela imagem
        try:
            time.sleep(1)
            imagem_tempo = browser.find_element(By.XPATH,
                                                '//div[@class='
                                                '"ChatConversationMessage'
                                                '_chat-conversation'
                                                '-message_3F2m0 '
                                                'ChatConversationMessage'
                                                '_agent-message'
                                                '_2upwR"]//descendant::ul[@class="MessageAttachments_'
                                                'attachments-list_2yody MessageAttachments_has-message-text_63PZr"]'
                                                '//img[@alt="image.png"]//following::time')
            data_hora_de_atribuicao = PegarDataEHoraDeAtribuicaoDaConversa()
            tempo_de_orcamento = imagem_tempo.get_attribute('datetime')
            TO = CalculoDeTempoDeResposta(data_hora_de_atribuicao, tempo_de_orcamento)
            string_orcamento = 'Sim'

            return string_orcamento, TO

        except NoSuchElementException:
            try:
                imagem_tempo = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                                '//div[@class='
                                                                                                '"ChatConversation'
                                                                                                'Message '
                                                                                                '_chat-conversation'
                                                                                                '-message_3F2m0 '
                                                                                                'ChatConversation'
                                                                                                'Message '
                                                                                                '_agent-message'
                                                                                                '_2upwR"]//descendant::'
                                                                                                'ul[@class='
                                                                                                '"MessageAttachments_'
                                                                                                'attachments'
                                                                                                '-list_2yody"]// '
                                                                                                'img[@alt="image.png'
                                                                                                '"]// '
                                                                                                'following::time')))

                data_hora_de_atribuicao = PegarDataEHoraDeAtribuicaoDaConversa()
                tempo_de_orcamento = imagem_tempo.get_attribute('datetime')
                TPO = CalculoDeTempoDeResposta(data_hora_de_atribuicao, tempo_de_orcamento)
                string_orcamento = 'Sim'

                return string_orcamento, TPO
            # Acha orçamento pela conversa
            except TimeoutException:
                teve_orcamento = analise_vendedor_geral(imagem_check)
                tempo_de_orcamento = "00:00:00"

                return teve_orcamento, tempo_de_orcamento


    def anal_cli(analise):
        try:
            mens_cli = WebDriverWait(browser, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'ChatConversation'
                                                                    'Message_chat-'
                                                                    'conversation-'
                                                                    'message_3F2m0')))

            for i in range(len(analise)):
                for texto in mens_cli:
                    analise_cliente = re.search(analise[i], texto.text)
                    if analise_cliente is None:
                        pass
                    else:
                        string = 'Sim'
                        return string

            if analise_cliente is None:
                string = 'Não'
                return string
        except TimeoutException:
            string = ""
            return string


    def PegarDataEHoraDeAtribuicaoDaConversa():
        # Pega as informações do balão de conversas atribuída
        try:
            atribu = WebDriverWait(browser, 10).until(EC.presence_of_all_elements_located((By.XPATH,
                                                                                           '//li[@class="mWGFOXnPu6ieL'
                                                                                           '-jptpvj3"]'
                                                                                           '/label["@class='
                                                                                           '_2_Wz9b8BlOuWRLjxM296cl"]'
                                                                                           )))
        except TimeoutException:
            return ""

            # Checa se a conversa foi apenas finalizada, se não houve atribuição
        if len(atribu) > 1:
            ActionChains(browser).move_to_element(atribu[1]).perform()
        else:
            ActionChains(browser).move_to_element(atribu[0]).perform()

        # Acha a div escondida com as informações de horário de atribuição
        time.sleep(1)
        '''
        atribu_hidden = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@data-tippy -root=""]'
                                                      '/div[@class'
                                                      '="tippy-box" '
                                                      ']/div['
                                                      '@class="tippy '
                                                      '-content"]')))
        '''
        try:
            atribu_hidden = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tippy-box')))
            return atribu_hidden.text
        except TimeoutException:
            return ""


    def CalculoDeTempoDeResposta(atribuicao, mensagem):
        # Pega a data e hora da primeira mensagem enviada

        # Pega o tempo de atribuição de chamado e transforma tudo em segundo
        try:
            prim_resp_sec = (int(mensagem[11:13]) - 3) * 60 * 60 + int(mensagem[14:16]) * 60 + int(mensagem[17:19])

            atribu_hidden_sec = int(atribuicao[11:13]) * 60 * 60 + int(atribuicao[14:16]) * 60

            tempo_em_segundos = prim_resp_sec - atribu_hidden_sec

            # tempo em formato hora:min:seexcept:g
            tempo_final = str(timedelta(seconds=tempo_em_segundos))

            return tempo_final
        except ValueError:
            return "00:00:00"




    # Análise de TPR
    def tpr():
        atribu_hidden = PegarDataEHoraDeAtribuicaoDaConversa()
        # Pega a data e hora da primeira mensagem do vendedor
        try:
            mensagem = WebDriverWait(browser, 10).until(EC.presence_of_all_elements_located((By.XPATH,
                                                                                             '//div[@class='
                                                                                             '"ChatConversationMessage'
                                                                                             '_chat-conversation'
                                                                                             '-message_3F2m0 '
                                                                                             'ChatConversationMessage'
                                                                                             '_agent-message'
                                                                                             '_2upwR"]')))
        except TimeoutException:
            return ""

        # Pega a primeira mensagem enviada pelo vendedor através de checagem por regular expresion
        i = 0
        while True:
            # Checa o caso em que não foi atribuída a nenhum vendedor e para no último range de mensagens
            if len(mensagem) == i:
                i -= 1
                break

            teste = re.match('Vinicius Macedo|Danilo Castro|Claudio Henrique|Juvenal Junior|Giselly Ramos',
                             mensagem[i].text)
            if teste is None:
                i += 1
            else:
                break

        prim_resp = mensagem[i].get_attribute('time')

        tpr = CalculoDeTempoDeResposta(atribu_hidden, prim_resp)

        return tpr


    # Clica em cada uma das conversas e as analisa
    pixel = 0
    dados = []
    while pixel <= altura_pag - 77:
        insere_dados = []
        # Pega o nome do cliente
        print('--------------------------------')
        browser.implicitly_wait(3)
        nome = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                '//div[@class="vue-recycle'
                                                                                '-scroller__item-view"]'
                                                                                '[@style="transform: translateY(' + str(
                                                                                    pixel) + 'px);"]//descendant::p')))
        insere_dados.append(nome.text)

        # Pega o nome do vendedor
        try:
            vendedor = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                        '//div[@class="vue-recycle'
                                                                                        '-scroller__item-view"]'
                                                                                        '[@style="transform: '
                                                                                        'translateY(' + str(
                                                                                            pixel) + 'px);"]'
                                                                                                     '//descendant::span'
                                                                                                     '[@class="DefaultChat'
                                                                                                     'Body___agent-name'
                                                                                                     '_GBLB1"]')))
            print(f'Nome: {nome.text}, vendedor: {vendedor.text}')
            insere_dados.append(vendedor.text)
        except TimeoutException as error:
            vendedor = ""
            print(f'Nome: {nome.text}, vendedor: {vendedor}')
            insere_dados.append(vendedor)

        # Pega a conversa
        chat = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                            '//div'
                                                                            '[@class="vue-recycle-scroller__item-view"]'
                                                                            '[@style="transform: translateY(' + str(
                                                                                pixel) + 'px);"]/descendant::div[1]')))

        # Clica na conversa
        ActionChains(browser).move_to_element(chat).click().perform()

        # pega o número de telefone
        time.sleep(2)
        numero = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                              '//div'
                                                                              '[@class="_list-item_ufped_1 '
                                                                              '_expansion-item_k6il9_1 '
                                                                              '_expansion-item--sm_k6il9_10 '
                                                                              '_expansion-item--no-padding_k6il9_13 '
                                                                              'UserInformations_profile__card--min-height'
                                                                              '_2ItlU"]/descendant::span')))

        print(f'Telefone: {numero.text[4:]}')
        insere_dados.append(numero.text[4:])

        # Pega a data de saída
        data_out = WebDriverWait(browser, 10).until(EC.presence_of_all_elements_located((By.XPATH,
                                                                                         '//li[@class="mWGFOXnPu6ieL'
                                                                                         '-jptpvj3"]'
                                                                                         '/label["@class='
                                                                                         '_2_Wz9b8BlOuWRLjxM296cl"]')))

        time.sleep(0.5)
        if len(data_out) > 2:
            ActionChains(browser).move_to_element(data_out[2]).perform()
        if len(data_out) <= 2:
            ActionChains(browser).move_to_element(data_out[0]).perform()

        browser.implicitly_wait(3)
        try:
            atribu_hidden_data = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@data-tippy'
                                                          '-root=""]/'
                                                          'div[@class='
                                                          '"tippy-box"]'
                                                          '/div[@class'
                                                          '="tippy-content"]'
                                                )))

            data_saida = atribu_hidden_data.text
        except TimeoutException:
            data_saida = "01/01/1900"

        # Scrolla o chat até em cima

        altura_anterior = 0
        while True:
            origem_chat = browser.find_element(By.CLASS_NAME, 'MessageListContainer_discussion_3509h')
            scroll_origem_chat = ScrollOrigin.from_element(
                origem_chat)  # Seta a origem do scroll para o início da página
            altura_pag_chat = int(origem_chat.get_attribute('scrollHeight'))
            ActionChains(browser).scroll_from_origin(scroll_origem_chat, 0,
                                                     -altura_pag_chat).perform()  # Scrolla para o
            # topo
            time.sleep(1)
            if altura_anterior == altura_pag_chat:
                break
            altura_anterior = altura_pag_chat

        # Pega a data de entrada
        data_ent = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'ChatDiscussionForm'
                                                                                                   'Information_chat-'
                                                                                                   'information-'
                                                                                                   'container_35NM3')))
        print(f'Data de entrada: {data_ent.text[17:]}')
        insere_dados.append(data_ent.text[17:])

        print(f'Data de saída: {data_saida}')
        insere_dados.append(data_saida)

        # Analises dentro das conversas
        orcamento = ["Segue seu orçamento!"]

        peca_faltante = ["não temos a peça em nosso estoque", "possa encontrar", "possa achar",
                         "não temos esse item em nosso estoque",
                         "não inclui esse item", "Ideal Peças", "Sucesso Auto peças", "catálogo de peças não inclui"]

        follow_up = ["prosseguir", "continuar"]

        achou_caro = ["caro", "mais em conta", "desconto"]

        foi_na_loja = ["fui buscar", "na loja", "busquei"]

        reserva = ['\d{11}|\d{3}\.\d{3}\.\d{3}-\d{2}|\d{3} \d{3} \d{3}-\d{2}|\d{3}\.\d{3}\.\d{3}\.\d{2}|\d{9}-\d{2}'
                   '|\d{3} \d{3} \d{3} \d{2}']

        # Análise de peça faltante
        peca_faltante_guarda = analise_vendedor_geral(peca_faltante)
        print(f'Peça faltante? {peca_faltante_guarda}')
        insere_dados.append(peca_faltante_guarda)

        # Análise de orçamento e tempo de orçamento
        orcamento_guarda, tempo_orcamento = analise_orcamento(orcamento)
        print(f'Teve Orçamento? {orcamento_guarda}')
        if len(tempo_orcamento) > 8:
            print(f'TO: {tempo_orcamento[8:]}')
            insere_dados.append(tempo_orcamento[8:])
        else:
            print(f'TO: {tempo_orcamento}')
            insere_dados.append(tempo_orcamento)
        insere_dados.append(orcamento_guarda)

        # Analisa se houve pedido reservado
        reserva_guarda = anal_cli(reserva)
        print(f'Pedido reservado? {reserva_guarda}')
        insere_dados.append(reserva_guarda)

        # Analisa se o cliente achou caro
        achou_caro_guarda = anal_cli(achou_caro)
        print(f'Achou caro? {achou_caro_guarda}')
        insere_dados.append(achou_caro_guarda)

        # Analisa Follow Up
        follow_up_guarda = analise_vendedor_geral(follow_up)
        print(f'Teve Follow Up? {follow_up_guarda}')
        insere_dados.append(follow_up_guarda)

        # Analisa se foi buscar
        analise_busca_guarda = anal_cli(foi_na_loja)
        print(f'Foi na loja? {analise_busca_guarda}')
        insere_dados.append(analise_busca_guarda)

        tpr_guarda = tpr()
        # Analisa o TPR
        if len(tpr_guarda) > 8:
            print(f'TPR: {tpr_guarda[8:]}')
            insere_dados.append(tpr_guarda[8:])
        else:
            print(f'TPR: {tpr_guarda}')
            insere_dados.append(tpr_guarda)

        dados.append(insere_dados)
        print(dados)

        # Aumenta a quantidade de pixels do tamanho da conversa do height da página para pegar a conversa anterio
        pixel += 77


    class InfoPlanilha:
        def __init__(self, caminho_planilha):
            self.caminho_planilha = caminho_planilha

        def ColocaDados(self, dados):
            # If modifying these scopes, delete the file token.json.
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

            creds = None

            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'client_secret.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())

            try:
                service = build('sheets', 'v4', credentials=creds)

                # Call the Sheets API
                # For para atualizar os dados da planilha de dados do report HH
                sheet = service.spreadsheets()

                altura_planilha = sheet.values().get(spreadsheetId=self.caminho_planilha,
                                                     range='Dados!A:A').execute().get('values', [])
                # Insere os dados na planilha de dados
                sheet.values().update(spreadsheetId=self.caminho_planilha,
                                      range='Dados!A' + str(len(altura_planilha) + 1),
                                      valueInputOption='USER_ENTERED', body={"values": dados
                                                                             }).execute()

            except HttpError as err:
                print(err)


    caminho_planilha = InfoPlanilha('1mNoCy6g9WQ1J5J4nj-u5JB1Zn2BJpBaInrEEpr_kh6E')

    caminho_planilha.ColocaDados(dados)

    browser.quit()
