import os
import time
import re
import numpy as np
import pandas as pd
import pyautogui
import pyperclip
from pywinauto import Application
import easyocr

class RoboAnexo:
    def __init__(self):
        """Initialize the RoboAnexo class with necessary attributes."""
        # Coordenadas de botões e campos na interface
        self.BOTAO_VOLTAR = (1274, 400)
        self.BOTAO_SETA_SALVAR = (1477, 968)
        self.CAMPO_SALVAR_COMO = (1608, 932)
        self.CONSULTA_PROCESSOS = (180, 360)
        self.CAMPO_PESQUISA = (348, 361)
        self.BOTAO_PESQUISAR = (502, 610)
        self.BOTAO_MENU = (120, 181)
        self.BOTAO_SALVAR = (1564, 953)
        self.BOTAO_FECHAR = (1656, 967)

        # Diretório base no Desktop onde a pasta mãe será criada
        self.desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
        self.anexos_dir = os.path.join(self.desktop_dir, 'anexos')
        os.makedirs(self.anexos_dir, exist_ok=True)

        # Inicializar o leitor OCR
        self.reader = easyocr.Reader(['pt'])

    def processar_linha(self, numero_processo):
        """Processa uma linha de número de processo clicando e pesquisando na interface."""
        try:
            pyautogui.click(self.CONSULTA_PROCESSOS)
            time.sleep(2.5)

            numero_processo_str = str(int(float(numero_processo)))
            print(f"Digitando GCPJ: {numero_processo_str}")

            pyperclip.copy(numero_processo_str)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1.5)

            pyautogui.click(self.BOTAO_PESQUISAR)
            time.sleep(1.5)
        except Exception as e:
            print(f"Erro ao processar linha: {str(e)}")

    def extrair_informacoes_da_tabela(self):
        """Extrai todas as informações da tabela (Nome Documento, Tipo de Anexo, Arquivo) usando OCR."""
        try:
            # Capturar a tabela inteira
            tabela_x1, tabela_y1 = 25, 85
            tabela_x2, tabela_y2 = 950, 200

            screenshot_tabela = pyautogui.screenshot(region=(tabela_x1, tabela_y1, tabela_x2-tabela_x1, tabela_y2-tabela_y1))
            screenshot_np = np.array(screenshot_tabela)

            # Usar OCR para detectar todo o texto na tabela
            resultados_ocr = self.reader.readtext(screenshot_np, detail=1)

            # Criar uma lista para armazenar as linhas da tabela
            linhas = []

            # Ordenar resultados por coordenada Y para agrupar por linhas
            resultados_ocr.sort(key=lambda x: x[0][0][1])

            # Valor Y da linha anterior para comparação
            last_y = -1
            linha_atual = []

            # Agrupar resultados em linhas
            for res in resultados_ocr:
                bbox = res[0]  # Bounding box
                texto = res[1]  # Texto detectado

                # Calcular a coordenada Y média do item
                y_medio = (bbox[0][1] + bbox[2][1]) / 2

                # Se este item estiver em uma nova linha (Y mudou significativamente)
                if last_y != -1 and abs(y_medio - last_y) > 15:  # Ajustável
                    # Ordenar itens da linha atual por X
                    linha_atual.sort(key=lambda x: x[0][0][0])
                    # Adicionar à lista de linhas
                    linhas.append([item[1] for item in linha_atual])
                    # Iniciar nova linha
                    linha_atual = []

                linha_atual.append(res)
                last_y = y_medio

            # Adicionar a última linha
            if linha_atual:
                linha_atual.sort(key=lambda x: x[0][0][0])
                linhas.append([item[1] for item in linha_atual])

            # Ignorar a primeira linha (cabeçalho da tabela)
            linhas = linhas[1:]

            # Extrair informações estruturadas de cada linha
            resultados = []
            for linha in linhas:
                if len(linha) >= 3:  # Deve ter pelo menos Nome Documento, Tipo Anexo e Arquivo
                    resultados.append({
                        "nome_documento": linha[0],
                        "tipo_anexo": linha[1],
                        "arquivo": linha[2],
                        "extensao": linha[2].split(".")[-1].lower() if "." in linha[2] else ""
                    })

            return resultados

        except Exception as e:
            print(f"Erro ao extrair informações da tabela: {str(e)}")
            return []

    def obter_info_via_ocr_campo(self, x1, y1, x2, y2):
        """Captura texto de uma região específica via OCR."""
        try:
            screenshot = pyautogui.screenshot(region=(x1, y1, x2-x1, y2-y1))
            screenshot_np = np.array(screenshot)
            resultados = self.reader.readtext(screenshot_np)
            texto = ' '.join([res[1] for res in resultados])
            return texto.strip()
        except Exception as e:
            print(f"Erro ao obter texto via OCR: {str(e)}")
            return ""

    def capturar_texto_radio_button(self, radio_button, window):
        """Captura o texto associado ao radio button usando OCR na área ao redor do radio button."""
        try:
            # Obter a posição do radio button
            rect = radio_button.rectangle()

            # Tentar localizar a coluna de referência usando a imagem COLUNA.png
            try:
                coluna_pos = pyautogui.locateOnScreen('baixar-anexos-assets/COLUNA.png', confidence=0.7)
                if coluna_pos:
                    print(f"Imagem COLUNA.png encontrada")
                    # Ajustar a área de captura com base na posição da coluna
                    x1 = int(coluna_pos.left)
                    x2 = int(x1 + 300)  # Largura suficiente para capturar o texto completo
                    y1 = int(rect.top - 5)
                    y2 = int(rect.bottom + 5)
                    print(f"Área de captura: ({x1}, {y1}, {x2-x1}, {y2-y1})")
                else:
                    print("Imagem COLUNA.png não encontrada, usando posição padrão")
                    x1 = rect.right + 5
                    y1 = rect.top - 5
                    x2 = x1 + 300
                    y2 = rect.bottom + 5
            except Exception as img_err:
                print(f"Erro ao localizar imagem COLUNA.png: {img_err}")
                x1 = rect.right + 5
                y1 = rect.top - 5
                x2 = x1 + 300
                y2 = rect.bottom + 5

            # Verificar se as coordenadas são válidas
            if x1 < 0 or y1 < 0 or x2 <= x1 or y2 <= y1:
                print("Coordenadas inválidas, ajustando...")
                x1 = max(0, rect.right + 5)
                y1 = max(0, rect.top - 5)
                x2 = x1 + 300
                y2 = y1 + 30

            # Garantir que a largura e altura sejam positivas
            width = max(1, x2 - x1)
            height = max(1, y2 - y1)

            # Capture o texto usando OCR
            try:
                screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
                screenshot_np = np.array(screenshot)

                # Usar EasyOCR para extrair o texto
                resultados = self.reader.readtext(screenshot_np)

                # Processar os resultados do OCR
                if resultados:
                    # Ordenar os resultados da esquerda para a direita
                    resultados.sort(key=lambda x: x[0][0][0])
                    # Concatenar todos os textos encontrados
                    texto = ' '.join([res[1].strip() for res in resultados])
                    print(f"Texto extraído via OCR: '{texto}'")
                else:
                    texto = ""
                    print("Nenhum texto encontrado via OCR")
            except Exception as ocr_err:
                print(f"Erro ao capturar screenshot ou processar OCR: {ocr_err}")
                texto = ""

            # Se OCR falhar ou texto estiver vazio, tente métodos alternativos
            if not texto:
                # Tente obter o nome diretamente do elemento
                texto = radio_button.element_info.name.strip()
                print(f"Texto obtido do elemento: '{texto}'")

                # Se ainda estiver vazio, tente os textos dos elementos vizinhos
                if not texto:
                    # Tente encontrar elementos de texto próximos (como labels)
                    text_elements = list(window.descendants(control_type="Text"))
                    for element in text_elements:
                        elem_rect = element.rectangle()
                        # Verificar se o elemento de texto está próximo ao radio button
                        if (abs(elem_rect.left - rect.right) < 100 and 
                            abs(elem_rect.top - rect.top) < 20):
                            texto = element.element_info.name.strip()
                            print(f"Texto obtido de elemento vizinho: '{texto}'")
                            if texto:
                                break

            # Limpar o texto de caracteres especiais e informações desnecessárias
            if texto:
                # Remover indicadores de radio button como "○" ou "●"
                texto = re.sub(r'[○●⚪⚫]', '', texto).strip()
                # Remover informações extras como "(selected)" ou "[x]"
                texto = re.sub(r'\(selected\)|\[x\]', '', texto).strip()

                # Verificar se o texto contém palavras-chave
                texto_upper = texto.upper()

                # Verificar se é um CIP (mesmo que contenha INVEST)
                if "CIP" in texto_upper:
                    print(f"Documento CIP detectado: '{texto}'")
                    return "CIP"

                # Verificar se é um IMPF
                elif "IMPF" in texto_upper:
                    print(f"Documento IMPF detectado: '{texto}'")
                    return "IMPF"

                # Verificar se é um CONTRATO
                elif "CONTRATO" in texto_upper:
                    print(f"Documento CONTRATO detectado: '{texto}'")
                    return "CONTRATO"

                # Verificar se é um CALCULO
                elif any(termo in texto_upper for termo in ["CALCULO", "CÁLCULO", "PLANILHA"]):
                    print(f"Documento CALCULO detectado: '{texto}'")
                    return "CALCULO"

                # Verificar se é um LOG
                elif "LOG" in texto_upper:
                    print(f"Documento LOG detectado: '{texto}'")
                    return "LOG"

                # Verificar se é um DOCUMENTO
                elif "DOCUMENTO" in texto_upper:
                    print(f"DOCUMENTO detectado: '{texto}'")
                    return "DOCUMENTO"

                # Verificar se contém "INVEST" e "CIP" (mesmo separados)
                elif "INVEST" in texto_upper and "CIP" in texto_upper:
                    print(f"Documento CIP (INVEST) detectado: '{texto}'")
                    return "CIP"

                # Verificar se é um CIP de INVESTIGAÇÃO
                elif "INVEST" in texto_upper:
                    print(f"Possível documento CIP (INVEST) detectado: '{texto}'")
                    # Verificar se há outros elementos que indicam que é um CIP
                    try:
                        # Capturar uma área maior para verificar se há menção a CIP
                        area_maior_x1 = max(0, x1 - 100)
                        area_maior_x2 = x2 + 100
                        area_maior_y1 = max(0, y1 - 20)
                        area_maior_y2 = y2 + 20

                        screenshot_maior = pyautogui.screenshot(region=(
                            area_maior_x1, 
                            area_maior_y1, 
                            area_maior_x2 - area_maior_x1, 
                            area_maior_y2 - area_maior_y1
                        ))
                        screenshot_maior_np = np.array(screenshot_maior)

                        resultados_maior = self.reader.readtext(screenshot_maior_np)
                        texto_maior = ' '.join([res[1].strip() for res in resultados_maior])

                        if "CIP" in texto_maior.upper():
                            print(f"Confirmado CIP em área maior: '{texto_maior}'")
                            return "CIP"
                    except:
                        pass

            print(f"Texto final após limpeza: '{texto}'")
            return texto if texto else f"Documento_{id(radio_button)}"

        except Exception as e:
            print(f"Erro ao capturar texto do radio button: {str(e)}")
            return f"Documento_{id(radio_button)}"

    def determinar_extensao_e_nome(self, nome_documento):
        """
        Determina a extensão do arquivo e, se necessário, um novo nome com base no nome do documento
        Retorna uma tupla: (nome_final, extensao)
        """
        # Converter para maiúsculas para comparação
        nome_upper = nome_documento.upper()

        # Verificar se o nome contém "CIP" (case insensitive)
        if re.search(r'CIP', nome_upper, re.IGNORECASE):
            print(f"Documento CIP detectado. Usando nome 'CIP' e extensão .xlsx")
            return ("CIP", "xlsx")

        # Verificar se o nome contém "IMPF" (case insensitive)
        elif re.search(r'IMPF', nome_upper, re.IGNORECASE):
            print(f"Documento IMPF detectado. Usando nome 'IMPF' e extensão .pdf")
            return ("IMPF", "pdf")

        # Verificar se o nome contém "CONTRATO" (case insensitive)
        elif re.search(r'CONTRATO', nome_upper, re.IGNORECASE):
            # Se o nome contiver informações adicionais, usar o nome completo
            if len(nome_documento.split()) > 1:
                print(f"Documento CONTRATO com informações adicionais detectado. Usando nome completo e extensão .pdf")
                return (nome_documento, "pdf")
            else:
                print(f"Documento CONTRATO simples detectado. Usando nome 'CONTRATO' e extensão .pdf")
                return ("CONTRATO", "pdf")

        # Verificar se o nome contém "CALCULO" (case insensitive)
        elif re.search(r'CALCULO', nome_upper, re.IGNORECASE):
            print(f"Documento CALCULO detectado. Renomeando para 'PLANILHA DE DÉBITO' e usando extensão .pdf")
            return ("PLANILHA DE DÉBITO", "pdf")

        # Verificar se o nome contém "LOG" (case insensitive)
        elif re.search(r'LOG', nome_upper, re.IGNORECASE):
            print(f"Documento LOG detectado. Usando nome 'LOG' e extensão .pdf")
            return ("LOG", "pdf")

        # Verificar se o nome contém "DOCUMENTO" (case insensitive)
        elif re.search(r'DOCUMENTO', nome_upper, re.IGNORECASE):
            print(f"DOCUMENTO detectado. Usando nome 'DOCUMENTO' e extensão .pdf")
            return ("DOCUMENTO", "pdf")

        # Para todos os outros casos, usar o nome original e PDF como padrão
        else:
            print(f"Tipo de documento não específico: '{nome_documento}'. Usando extensão .pdf")
            return (nome_documento, "pdf")

    def capturar_informacoes_linha_vertical(self, radio_button, window):
        """Captura as informações completas da linha vertical associada ao radio button."""
        try:
            # Obter a posição do radio button
            rect = radio_button.rectangle()

            # Tentar localizar a coluna de referência usando a imagem COLUNA.png
            try:
                coluna_pos = pyautogui.locateOnScreen('baixar-anexos-assets/COLUNA.png', confidence=0.7)
                if coluna_pos:
                    print(f"Imagem COLUNA.png encontrada")
                    # Ajustar a área de captura com base na posição da coluna
                    x1 = int(coluna_pos.left)
                    x2 = int(x1 + 1200)  # Largura para capturar todas as colunas
                    y1 = int(rect.top - 10)  # Aumentar a altura para capturar mais texto acima
                    y2 = int(rect.bottom + 10)  # Aumentar a altura para capturar mais texto abaixo
                    print(f"Área de captura: ({x1}, {y1}, {x2-x1}, {y2-y1})")
                else:
                    print("Imagem COLUNA.png não encontrada, usando posição padrão")
                    x1 = rect.right + 5
                    y1 = rect.top - 10
                    x2 = x1 + 1200
                    y2 = rect.bottom + 10
            except Exception as img_err:
                print(f"Erro ao localizar imagem COLUNA.png: {img_err}")
                x1 = rect.right + 5
                y1 = rect.top - 10
                x2 = x1 + 1200
                y2 = rect.bottom + 10

            # Verificar se as coordenadas são válidas
            if x1 < 0 or y1 < 0 or x2 <= x1 or y2 <= y1:
                print("Coordenadas inválidas, ajustando...")
                x1 = max(0, rect.right + 5)
                y1 = max(0, rect.top - 10)
                x2 = x1 + 1200
                y2 = y1 + 40  # Aumentar para capturar mais texto verticalmente

            # Garantir que a largura e altura sejam positivas
            width = max(1, x2 - x1)
            height = max(1, y2 - y1)

            # Capture o texto usando OCR
            try:
                screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
                screenshot_np = np.array(screenshot)

                # Salvar a imagem para debug (temporário)
                try:
                    screenshot.save('debug_captura.png')
                    print("Imagem de debug salva como 'debug_captura.png'")
                except:
                    print("Não foi possível salvar a imagem de debug")

                # Usar EasyOCR para extrair o texto com mais detalhes
                resultados = self.reader.readtext(screenshot_np, detail=1)

                # Processar os resultados do OCR
                if resultados:
                    print(f"Encontrados {len(resultados)} blocos de texto")

                    # Exibir todos os resultados com suas coordenadas para debug
                    for i, res in enumerate(resultados):
                        bbox = res[0]  # Coordenadas [topo-esq, topo-dir, inf-dir, inf-esq]
                        texto = res[1]  # Texto detectado
                        conf = res[2]   # Confiança

                        # Calcular posição média X e Y
                        x_medio = sum([p[0] for p in bbox]) / 4
                        y_medio = sum([p[1] for p in bbox]) / 4

                        print(f"Bloco {i}: '{texto}' (X={x_medio:.1f}, Y={y_medio:.1f}, Conf={conf:.2f})")

                    # Agrupar os textos por coluna usando análise de posicionamento X
                    # Primeiro, ordenar todos os blocos por posição X
                    resultados_ordenados_x = sorted(resultados, key=lambda res: sum([p[0] for p in res[0]]) / 4)

                    # Identificar separações de colunas (com base em lacunas na posição X)
                    posicoes_x = [sum([p[0] for p in res[0]]) / 4 for res in resultados_ordenados_x]

                    colunas = []
                    coluna_atual = [resultados_ordenados_x[0]]

                    for i in range(1, len(resultados_ordenados_x)):
                        # Se houver uma distância significativa no eixo X, considerar uma nova coluna
                        if posicoes_x[i] - posicoes_x[i-1] > 50:  # Ajustável conforme necessário
                            # Adicionar coluna anterior
                            colunas.append(coluna_atual)
                            # Iniciar nova coluna
                            coluna_atual = [resultados_ordenados_x[i]]
                        else:
                            coluna_atual.append(resultados_ordenados_x[i])

                    # Adicionar a última coluna
                    if coluna_atual:
                        colunas.append(coluna_atual)

                    print(f"Detectadas {len(colunas)} possíveis colunas")

                    # Para cada coluna, ordenar os blocos por posição Y e juntar os textos
                    textos_por_coluna = []
                    for i, coluna in enumerate(colunas):
                        # Ordenar blocos dentro da coluna por posição Y
                        blocos_ordenados = sorted(coluna, key=lambda res: sum([p[1] for p in res[0]]) / 4)
                        # Extrair e juntar os textos
                        texto_coluna = " ".join([bloco[1] for bloco in blocos_ordenados])
                        textos_por_coluna.append(texto_coluna)
                        print(f"Texto da coluna {i+1}: '{texto_coluna}'")

                    # Atribuir as colunas encontradas
                    nome_documento = ""
                    tipo_anexo = ""
                    nome_arquivo = ""
                    extensao = ""

                    # Verificar se temos pelo menos 1 coluna (Nome Documento)
                    if len(textos_por_coluna) >= 1:
                        nome_documento = textos_por_coluna[0].strip()
                        print(f"Nome do documento (coluna 1): '{nome_documento}'")

                    # Verificar se temos pelo menos 2 colunas (Tipo de Anexo)
                    if len(textos_por_coluna) >= 2:
                        tipo_anexo = textos_por_coluna[1].strip()
                        print(f"Tipo de anexo (coluna 2): '{tipo_anexo}'")

                    # Verificar se temos pelo menos 3 colunas (Arquivo)
                    if len(textos_por_coluna) >= 3:
                        nome_arquivo = textos_por_coluna[2].strip()
                        print(f"Nome do arquivo (coluna 3): '{nome_arquivo}'")

                        # Extrair a extensão
                        if "." in nome_arquivo:
                            extensao = nome_arquivo.split(".")[-1].lower()
                        else:
                            # Tentar extrair extensão sem ponto
                            for ext in ["pdf", "xlsx", "xls", "doc", "docx"]:
                                if ext.upper() in nome_arquivo.upper():
                                    extensao = ext
                                    break
                        print(f"Extensão detectada: '{extensao}'")

                    # Se não conseguimos detectar as colunas corretamente, tentar o método anterior
                    if not nome_documento or not extensao:
                        print("Usando método alternativo para detectar colunas")

                        # Extrair todos os textos encontrados
                        textos = [res[1].strip() for res in resultados]
                        print(f"Textos extraídos via OCR: {textos}")

                        # Identificar o Arquivo (terceira coluna) - buscar por padrões de extensão
                        idx_arquivo = -1
                        for i, texto in enumerate(textos):
                            if re.search(r'\.(PDF|XLSX|XLS|DOC|DOCX)', texto.upper()):
                                idx_arquivo = i
                                nome_arquivo = texto
                                # Extrair a extensão
                                if "." in texto:
                                    extensao = texto.split(".")[-1].lower()
                                print(f"Arquivo detectado na posição {i}: '{texto}' (extensão: {extensao})")
                                break

                        # Identificar o Tipo de Anexo (segunda coluna)
                        idx_tipo_anexo = -1
                        for i, texto in enumerate(textos):
                            if ("DOCUMENTO LOCALIZADO" in texto.upper() or 
                                "INVESTIGAÇÃO" in texto.upper() or 
                                "INVESTIGACAO" in texto.upper()):
                                idx_tipo_anexo = i
                                print(f"Tipo de anexo detectado na posição {i}: '{texto}'")
                                break

                        # Se encontramos o tipo de anexo, os textos antes dele são o nome do documento
                        if idx_tipo_anexo > 0 and (not nome_documento or "DOCUMENTO LOCALIZADO" in nome_documento.upper()):
                            nome_documento = " ".join(textos[:idx_tipo_anexo])
                            print(f"Nome do documento (método alternativo): '{nome_documento}'")

                    # Limpar o nome do documento de múltiplos espaços e caracteres problemáticos
                    if nome_documento:
                        nome_documento = re.sub(r'\s+', ' ', nome_documento).strip()
                        print(f"Nome do documento após limpeza: '{nome_documento}'")

                    # Tratamentos especiais baseados no conteúdo do nome do documento
                    if "EMPF" in nome_documento.upper() and "CONTRATO" in nome_documento.upper():
                        print(f"Detectado documento EMPF CONTRATO: '{nome_documento}'")
                        if not extensao:
                            extensao = "pdf"

                    # Se detectamos "CIP" no nome e não temos extensão, assumir XLSX
                    if "CIP" in nome_documento.upper() and not extensao:
                        extensao = "xlsx"
                        print(f"Documento CIP detectado sem extensão. Assumindo .xlsx")

                    # Se ainda não temos extensão, verificar o nome do documento
                    if not extensao:
                        # Tentar inferir a extensão com base no nome do documento
                        if any(termo in nome_documento.upper() for termo in ["RECIBO", "REGULAMENTO", "CONTRATO"]):
                            extensao = "pdf"
                            print(f"Inferindo extensão PDF com base no nome do documento")

                    # Verificar se o nome do documento contém parte do tipo de anexo (erro comum)
                    if tipo_anexo and nome_documento and "DOCUMENTO LOCALIZADO" in nome_documento.upper():
                        # Remover a parte que pertence ao tipo de anexo
                        nome_documento = re.sub(r'DOCUMEN[TI]O LOCALIZAD[OU].*', '', nome_documento, flags=re.IGNORECASE).strip()
                        print(f"Nome do documento corrigido (removida parte do tipo anexo): '{nome_documento}'")

                    return {
                        "nome_documento": nome_documento,
                        "nome_arquivo": nome_arquivo,
                        "extensao": extensao
                    }
                else:
                    print("Nenhum texto encontrado via OCR")
                    return {
                        "nome_documento": f"Documento_{id(radio_button)}",
                        "nome_arquivo": "",
                        "extensao": ""
                    }

            except Exception as ocr_err:
                print(f"Erro ao capturar screenshot ou processar OCR: {ocr_err}")
                return {
                    "nome_documento": f"Documento_{id(radio_button)}",
                    "nome_arquivo": "",
                    "extensao": ""
                }

        except Exception as e:
            print(f"Erro ao capturar informações da linha: {str(e)}")
            return {
                "nome_documento": f"Documento_{id(radio_button)}",
                "nome_arquivo": "",
                "extensao": ""
            }

    def capturar_nome_radio_button(self, radio_button):
        """Captura o nome do documento clicando três vezes à direita do radio button e copiando o texto selecionado."""
        try:
            # Obter a posição do radio button
            rect = radio_button.rectangle()

            # Calcular a posição para clicar (à direita do radio button)
            click_x = rect.right + 50  # 50 pixels à direita do radio button
            click_y = rect.top + (rect.bottom - rect.top) / 2  # No meio do radio button verticalmente

            # Limpar a área de transferência
            pyperclip.copy('')

            # Clicar três vezes para selecionar todo o texto
            pyautogui.click(click_x, click_y, clicks=3, interval=0.1)
            time.sleep(0.5)

            # Copiar o texto selecionado
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)

            # Obter o texto da área de transferência
            texto = pyperclip.paste().strip()

            if texto:
                print(f"Texto capturado pelo método de seleção: '{texto}'")
                return texto
            else:
                print("Não foi possível capturar texto pelo método de seleção")
                return None

        except Exception as e:
            print(f"Erro ao capturar nome do radio button: {str(e)}")
            return None

    def interagir_com_checkboxes(self, gcpj_number):
        """Interage com checkboxes na interface para processar documentos."""
        try:
            time.sleep(2)

            app = Application(backend="uia").connect(title_re="GCPJ - Gestão e Controle de Processos Jurídicos.*")
            window = app.window(title_re="GCPJ - Gestão e Controle de Processos Jurídicos.*")
            window.set_focus()

            # Acessar a seção de anexos
            for _ in range(5):
                window.type_keys("{PGDN}")
                time.sleep(0.2)

            anexo_btn = window.child_window(title="anexo", control_type="Button")
            if anexo_btn.exists():
                anexo_btn.click()
                time.sleep(2)
            else:
                print("Botão 'anexo' não encontrado.")
                return

            # Voltar para o topo da tela
            for _ in range(5):
                window.type_keys("{PGUP}")
                time.sleep(0.2)

            # Criar pasta para o processo atual se não existir
            new_folder_path = os.path.join(self.anexos_dir, str(int(gcpj_number)))
            os.makedirs(new_folder_path, exist_ok=True)

            # Controle de páginas
            pagina_atual = 1
            documentos_processados = set()  # Conjunto para armazenar nomes de documentos já processados
            documentos_pagina_atual = set()  # Conjunto para armazenar nomes de documentos da página atual
            max_paginas = 10  # Limite máximo de páginas para evitar loop infinito
            ultima_pagina_repetida = False  # Flag para controlar repetição da última página

            while pagina_atual <= max_paginas:
                print(f"\n{'='*50}")
                print(f"Processando página {pagina_atual}")
                print(f"{'='*50}")

                # Limpar conjunto de documentos da página atual
                documentos_pagina_atual.clear()

                # Obter lista de radio buttons na página atual
                tentativas = 0
                radio_buttons = []

                while not radio_buttons and tentativas < 3:
                    radio_buttons = list(window.descendants(control_type="RadioButton"))
                    if not radio_buttons:
                        print(f"Tentativa {tentativas+1}: Nenhum radio button encontrado. Tentando novamente...")
                        time.sleep(1)
                        tentativas += 1

                if not radio_buttons:
                    print("Não foi possível encontrar radio buttons na página atual.")
                    if pagina_atual > 1:
                        print("Fim da navegação pelas páginas.")
                        break
                    else:
                        print("Não foram encontrados documentos para download.")
                        return

                print(f"Encontrados {len(radio_buttons)} radio buttons na página {pagina_atual}.")

                # Processar cada radio button
                for i, radio_button in enumerate(radio_buttons):
                    try:
                        print(f"\n{'='*30}")
                        print(f"Processando radio button {i+1}/{len(radio_buttons)} (página {pagina_atual})")
                        print(f"{'='*30}")

                        # Clicar no radio button para selecioná-lo
                        print(f"Clicando no radio button {i+1}")
                        radio_button.click()
                        time.sleep(0.5)

                        # Tentar capturar o nome do documento usando o método de 3 cliques
                        nome_documento = self.capturar_nome_radio_button(radio_button)

                        # Se não conseguiu capturar pelo método de 3 cliques, usar OCR
                        if not nome_documento:
                            print("Método de 3 cliques falhou, usando OCR...")
                            info = self.capturar_informacoes_linha_vertical(radio_button, window)
                            nome_documento = info["nome_documento"]
                            nome_arquivo = info["nome_arquivo"]
                            extensao = info["extensao"]
                        else:
                            # Se conseguiu capturar pelo método de 3 cliques, usar OCR apenas para extensão
                            info = self.capturar_informacoes_linha_vertical(radio_button, window)
                            extensao = info["extensao"]
                            nome_arquivo = info["nome_arquivo"]

                        # Verificar se a extensão é válida (apenas PDF ou XLSX)
                        if extensao not in ["pdf", "xlsx"]:
                            print(f"Arquivo ignorado: extensão '{extensao}' não é PDF ou XLSX")
                            continue

                        # Adicionar o nome do documento ao conjunto da página atual
                        documentos_pagina_atual.add(nome_documento)

                        # Verificar se já processamos esse documento
                        if nome_documento in documentos_processados:
                            print(f"Documento '{nome_documento}' já foi processado anteriormente. Pulando.")
                            continue

                        print(f"Nome do documento: '{nome_documento}'")
                        print(f"Extensão: '{extensao}'")

                        # Definir o nome final do arquivo
                        nome_final = nome_documento

                        # Verificar se é um dos tipos de documentos específicos que precisam ser renomeados
                        nome_upper = nome_documento.upper()

                        # Verificar se é um CIP
                        if "CIP" in nome_upper and not any(termo in nome_upper for termo in ["RECIBO", "REGULAMENTO"]):
                            print(f"Documento CIP identificado para renomeação.")
                            nome_final = "CIP"
                            extensao = "xlsx"

                        # Verificar se é um IMPF
                        elif "IMPF" in nome_upper and not any(termo in nome_upper for termo in ["RECIBO", "REGULAMENTO"]):
                            print(f"Documento IMPF identificado para renomeação.")
                            nome_final = "IMPF"
                            extensao = "pdf"

                        # Verificar se é um CALCULO ou PLANILHA DE DÉBITO
                        elif any(termo in nome_upper for termo in ["CALCULO", "CÁLCULO", "PLANILHA DE DÉBITO"]) and not any(termo in nome_upper for termo in ["RECIBO", "REGULAMENTO"]):
                            print(f"Documento CALCULO identificado para renomeação.")
                            nome_final = "PLANILHA DE DÉBITO"
                            extensao = "pdf"

                        # Verificar se é um LOG
                        elif "LOG" in nome_upper and not any(termo in nome_upper for termo in ["RECIBO", "REGULAMENTO"]):
                            print(f"Documento LOG identificado para renomeação.")
                            nome_final = "LOG"
                            extensao = "pdf"

                        # Verificar se é um DOCUMENTO genérico
                        elif "DOCUMENTO" in nome_upper and len(nome_documento.split()) < 3 and not any(termo in nome_upper for termo in ["RECIBO", "REGULAMENTO"]):
                            print(f"DOCUMENTO genérico identificado para renomeação.")
                            nome_final = "DOCUMENTO"
                            extensao = "pdf"

                        print(f"Nome final: '{nome_final}', Extensão: '{extensao}'")

                        # Construir o nome do arquivo com extensão
                        nome_arquivo = f"{nome_final}.{extensao}"

                        # Remover caracteres inválidos para nome de arquivo
                        nome_arquivo = re.sub(r'[\\/*?:"<>|]', "_", nome_arquivo)

                        # Caminho completo para o arquivo
                        caminho_completo = os.path.join(new_folder_path, nome_arquivo)

                        # Verificar se o arquivo já existe (adicionar numeração se existir)
                        contador = 1
                        nome_base, ext = os.path.splitext(nome_arquivo)
                        while os.path.exists(caminho_completo):
                            nome_arquivo = f"{nome_base}_{contador}{ext}"
                            caminho_completo = os.path.join(new_folder_path, nome_arquivo)
                            contador += 1

                        print(f"Caminho completo do arquivo: {caminho_completo}")

                        # Tentar 3 vezes para garantir o clique no botão visualizar
                        tentativa_visualizar = 0
                        while tentativa_visualizar < 3:
                            try:
                                # Clicar no botão visualizar
                                visualizar_btn = window.child_window(title="visualizar arquivo", control_type="Button")
                                if visualizar_btn.exists():
                                    visualizar_btn.click()
                                    print("Clicou no botão visualizar arquivo")
                                    time.sleep(5)
                                    break  # Saiu com sucesso
                                else:
                                    print(f"Tentativa {tentativa_visualizar+1}: Botão 'visualizar arquivo' não encontrado")
                                    if tentativa_visualizar == 2:  # Última tentativa, pular para o próximo item
                                        print("Botão 'visualizar arquivo' não encontrado após 3 tentativas. Pulando.")
                                        break
                            except Exception as e:
                                print(f"Erro ao clicar no botão visualizar (tentativa {tentativa_visualizar+1}): {e}")

                            tentativa_visualizar += 1
                            time.sleep(1)

                        # Se não conseguiu clicar no botão visualizar após 3 tentativas, pular para próximo
                        if tentativa_visualizar >= 3:
                            continue

                        # Salvar o arquivo
                        try:
                            print(f"Salvando arquivo como '{nome_arquivo}' (extensão: {extensao})...")

                            # Tentar salvar o arquivo, com tentativas em caso de falha
                            tentativas_salvar = 0
                            while tentativas_salvar < 3:
                                try:
                                    pyautogui.click(self.BOTAO_SETA_SALVAR)
                                    time.sleep(2.5)

                                    pyautogui.click(self.CAMPO_SALVAR_COMO)
                                    time.sleep(2.5)

                                    # Limpar o campo de texto antes de colar
                                    pyautogui.hotkey('ctrl', 'a')
                                    time.sleep(2)

                                    pyperclip.copy(caminho_completo)  # Usar pyperclip para evitar problemas com caracteres especiais
                                    pyautogui.hotkey('ctrl', 'v')
                                    time.sleep(2)
                                    pyautogui.press('enter')
                                    time.sleep(5)  # Tempo para download

                                    # Se existir botão de sobrescrever, clicar nele
                                    try:
                                        overwrite_btn = pyautogui.locateOnScreen('baixar-anexos-assets/sobrescrever.png', confidence=0.8)
                                        if overwrite_btn:
                                            pyautogui.click(overwrite_btn)
                                            print("Clicou no botão sobrescrever")
                                            time.sleep(2)
                                    except:
                                        pass

                                    # Verificar se o arquivo foi salvo com sucesso
                                    if os.path.exists(caminho_completo):
                                        print(f"Arquivo salvo com sucesso: {caminho_completo}")
                                        # Adicionar o documento ao conjunto de processados
                                        documentos_processados.add(nome_documento)
                                        break  # Sair do loop de tentativas
                                    else:
                                        # Verificar se foi salvo com outra extensão
                                        extensoes_possiveis = ["pdf", "xlsx"]
                                        arquivo_encontrado = False

                                        for ext_alt in extensoes_possiveis:
                                            if ext_alt == extensao:
                                                continue  # Pular a extensão original

                                            caminho_alt = os.path.join(new_folder_path, f"{nome_base}.{ext_alt}")

                                            if os.path.exists(caminho_alt):
                                                print(f"Arquivo salvo com extensão diferente: {caminho_alt}")

                                                # Se for um CIP e foi salvo como PDF, tentar renomear para xlsx
                                                if nome_final == "CIP" and ext_alt != "xlsx":
                                                    try:
                                                        novo_caminho = os.path.splitext(caminho_alt)[0] + ".xlsx"
                                                        os.rename(caminho_alt, novo_caminho)
                                                        print(f"Arquivo CIP renomeado para XLSX: {novo_caminho}")
                                                        caminho_completo = novo_caminho
                                                        arquivo_encontrado = True
                                                        # Adicionar o documento ao conjunto de processados
                                                        documentos_processados.add(nome_documento)
                                                        break
                                                    except Exception as rename_err:
                                                        print(f"Erro ao renomear arquivo CIP para XLSX: {rename_err}")
                                                        # Manter com a extensão encontrada
                                                        caminho_completo = caminho_alt
                                                        arquivo_encontrado = True
                                                        # Adicionar o documento ao conjunto de processados
                                                        documentos_processados.add(nome_documento)
                                                        break
                                                else:
                                                    # Manter com a extensão encontrada
                                                    print(f"Mantendo arquivo com extensão {ext_alt}")
                                                    caminho_completo = caminho_alt
                                                    arquivo_encontrado = True
                                                    # Adicionar o documento ao conjunto de processados
                                                    documentos_processados.add(nome_documento)
                                                    break

                                        if arquivo_encontrado:
                                            break  # Sair do loop de tentativas
                                        elif tentativas_salvar < 2:  # Tentar novamente se não encontrou
                                            print(f"Tentativa {tentativas_salvar+1}: Arquivo não encontrado após salvar")
                                            tentativas_salvar += 1
                                        else:
                                            print("Não foi possível salvar o arquivo após 3 tentativas")
                                            break

                                except Exception as e:
                                    print(f"Erro ao salvar (tentativa {tentativas_salvar+1}): {e}")
                                    tentativas_salvar += 1

                                    # Tentar fechar diálogos abertos
                                    pyautogui.press('esc')
                                    time.sleep(0.5)

                            # Fechar visualizador de documento
                            pyautogui.click(self.BOTAO_FECHAR)
                            time.sleep(1)

                            # Voltar para a lista
                            pyautogui.click(self.BOTAO_VOLTAR)
                            time.sleep(1)

                            print(f"Documento '{nome_documento}' processado e salvo como '{os.path.basename(caminho_completo)}'")

                        except Exception as e:
                            print(f"Erro ao salvar documento: {e}")
                            # Tentar fechar possíveis diálogos abertos e voltar à tela principal
                            try:
                                pyautogui.press('esc')
                                time.sleep(0.5)
                                pyautogui.click(self.BOTAO_FECHAR)
                                time.sleep(0.5)
                                pyautogui.click(self.BOTAO_VOLTAR)
                                time.sleep(1)
                            except:
                                pass

                    except Exception as e:
                        print(f"Erro ao processar radio button {i+1}: {str(e)}")
                        # Tentar voltar para a lista de anexos
                        try:
                            pyautogui.press('esc')
                            time.sleep(0.5)
                            pyautogui.click(self.BOTAO_VOLTAR)
                            time.sleep(1)
                        except:
                            pass

                # Verificar se todos os documentos desta página já foram processados antes
                if documentos_pagina_atual and documentos_pagina_atual.issubset(documentos_processados - documentos_pagina_atual):
                    print("Todos os documentos desta página já foram processados anteriormente.")
                    break

                # Verificar se a página atual é igual à anterior (repetição)
                if pagina_atual > 1:
                    # Capturar screenshot da página atual
                    screenshot_atual = pyautogui.screenshot()
                    screenshot_atual.save(os.path.join(self.anexos_dir, f"pagina_{pagina_atual}.png"))

                    # Comparar com a página anterior
                    if os.path.exists(os.path.join(self.anexos_dir, f"pagina_{pagina_atual-1}.png")):
                        # Aqui você pode implementar uma comparação mais robusta das imagens
                        # Por enquanto, vamos usar uma verificação simples de tamanho
                        if os.path.getsize(os.path.join(self.anexos_dir, f"pagina_{pagina_atual}.png")) == os.path.getsize(os.path.join(self.anexos_dir, f"pagina_{pagina_atual-1}.png")):
                            print("Página atual é igual à anterior. Fim da navegação.")
                            break

                # Ir para a próxima página - TENTAR LOCALIZAR E CLICAR NA SETINHA.PNG
                print("Tentando avançar para a próxima página...")

                # Tentar localizar a setinha.png
                tentativas_setinha = 0
                setinha_encontrada = False

                while tentativas_setinha < 3 and not setinha_encontrada:
                    try:
                        # Tentar encontrar a setinha.png com diferentes níveis de confiança
                        confiancas = [0.8, 0.7, 0.6]
                        for confianca in confiancas:
                            try:
                                print(f"Tentando localizar setinha.png com confiança {confianca}")
                                setinha_pos = pyautogui.locateOnScreen('baixar-anexos-assets/setinha.png', confidence=confianca)
                                if setinha_pos:
                                    print(f"Setinha.png encontrada em {setinha_pos} com confiança {confianca}")
                                    pyautogui.click(setinha_pos)
                                    print("Clicou na setinha para avançar")
                                    time.sleep(2)
                                    setinha_encontrada = True
                                    break
                            except Exception as e:
                                print(f"Erro ao tentar localizar setinha.png com confiança {confianca}: {e}")

                        # Se não encontrou com nenhuma confiança, tente tirar screenshot e salvar para debug
                        if not setinha_encontrada:
                            # Capturar toda a tela para debug
                            try:
                                screenshot = pyautogui.screenshot()
                                debug_path = os.path.join(self.anexos_dir, f"debug_pagina_{pagina_atual}.png")
                                screenshot.save(debug_path)
                                print(f"Screenshot de debug salvo em {debug_path}")
                            except:
                                print("Não foi possível salvar screenshot de debug")

                    except Exception as e:
                        print(f"Erro na tentativa {tentativas_setinha+1} de localizar setinha.png: {e}")

                    if not setinha_encontrada:
                        tentativas_setinha += 1
                        time.sleep(1)

                # Se a setinha foi encontrada, avançamos para a próxima página
                if setinha_encontrada:
                    pagina_atual += 1
                    # Esperar que a página carregue
                    time.sleep(3)

                    # Rolar para o topo da página para garantir que vemos todos os elementos
                    for _ in range(3):
                        window.type_keys("{PGUP}")
                        time.sleep(0.2)
                else:
                    # Se não conseguimos encontrar a setinha após várias tentativas
                    print("Não foi possível localizar a setinha.png após várias tentativas")

                    # Verificar se temos novos documentos nesta página comparando com os já processados
                    if documentos_pagina_atual and not documentos_pagina_atual - documentos_processados:
                        print("Não há novos documentos nesta página. Assumindo que chegamos ao fim.")
                        break

                    # Tentar usar PageDown como último recurso
                    print("Tentando usar PageDown como último recurso")
                    window.type_keys("{PGDN}")
                    time.sleep(2)

                    # Verificar se houve mudança nos radio buttons
                    novos_radio_buttons = list(window.descendants(control_type="RadioButton"))
                    if len(novos_radio_buttons) > 0 and novos_radio_buttons != radio_buttons:
                        print("Página alterada com PageDown")
                        pagina_atual += 1
                    else:
                        print("Não foi possível avançar para a próxima página. Fim da navegação.")
                        break

            print(f"\nProcessamento concluído. Total de {len(documentos_processados)} documentos processados em {pagina_atual} páginas.")

        except Exception as e:
            print(f"Erro ao executar automação: {str(e)}")

    def main(self):
        """Função principal para iniciar o processamento dos documentos."""
        try:
            # Verificar se o arquivo existe
            if not os.path.exists('BaseDeDados.xlsx'):
                print("ERRO: Arquivo 'BaseDeDados.xlsx' não encontrado!")
                return

            df = pd.read_excel('BaseDeDados.xlsx')

            if 'GCPJ' not in df.columns:
                print("ERRO: Coluna 'GCPJ' não encontrada no arquivo BaseDeDados.xlsx!")
                print(f"Colunas disponíveis: {df.columns.tolist()}")
                return

            print("Iniciando processamento em 3 segundos...")
            time.sleep(3)

            for index, row in df.iterrows():
                numero_processo = row['GCPJ']
                print(f"\n=== Processando número: {numero_processo} ===")

                self.processar_linha(numero_processo)
                time.sleep(1)

                self.interagir_com_checkboxes(numero_processo)
                time.sleep(1)

                pyautogui.click(self.BOTAO_MENU)
                time.sleep(2)

            print("\nProcessamento concluído. PDFs e XLSXs salvos na pasta 'anexos' no Desktop.")

        except Exception as e:
            print(f"Erro no processamento: {str(e)}")