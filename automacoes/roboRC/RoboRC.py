import pandas as pd
import pyautogui
import time
import pyperclip
import os
import sys
import numpy as np
import easyocr
from PIL import ImageGrab
from datetime import datetime

# COORDENADAS
CAMPO_GCPJ = (338, 313) # ONDE IRÁ DIGITAR CÓDIGO GCPJ
CAMPO_REFERENCIA = (529, 509) # ONDE IRÁ CLICAR E DIGITAR "CI-PROCESSO SEM MOVIMENTACAO"


# Região aproximada onde o botão salvar deve estar (para limitar a busca e acelerar)
# Definida com base na parte inferior da tela
REGIAO_BOTAO_SALVAR = (300, 850, 550, 950)

# Inicializa o leitor OCR (apenas uma vez para economizar recursos)
reader = None

def log_mensagem(mensagem, arquivo_log="automacao_log.txt"):
    """Registra mensagens com timestamp no console e em um arquivo de log"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mensagem_log = f"[{timestamp}] {mensagem}"
    print(mensagem_log)
    
    # Também salva em um arquivo de log
    with open(arquivo_log, "a", encoding="utf-8") as f:
        f.write(mensagem_log + "\n")

def encontrar_botao_por_texto(texto_botao, regiao=None, confianca=0.6):
    """Localiza um botão pelo texto usando OCR e retorna suas coordenadas"""
    global reader
    
    # Inicializa o leitor OCR se ainda não foi inicializado
    if reader is None:
        log_mensagem("Inicializando o leitor OCR...")
        reader = easyocr.Reader(['pt'], gpu=False)
    
    log_mensagem(f"Procurando botão com texto '{texto_botao}'...")
    
    # Captura a tela ou região específica
    if regiao is None:
        # Captura toda a tela
        screenshot = ImageGrab.grab()
    else:
        # Captura apenas a região especificada (x1, y1, x2, y2)
        screenshot = ImageGrab.grab(bbox=regiao)
    
    # Converte para formato numpy (necessário para o easyocr)
    img_np = np.array(screenshot)
    
    # Realiza o reconhecimento de texto com configurações otimizadas para velocidade
    resultados = reader.readtext(img_np, detail=1, paragraph=False, decoder='greedy', beamWidth=1, batch_size=1)
    
    # Procura pelo texto do botão
    for (bbox, texto, prob) in resultados:
        if texto_botao.lower() in texto.lower() and prob >= confianca:
            # Calcula o centro do botão
            (top_left, top_right, bottom_right, bottom_left) = bbox
            centro_x = int((top_left[0] + bottom_right[0]) / 2)
            centro_y = int((top_left[1] + bottom_right[1]) / 2)
            
            # Ajusta as coordenadas se estiver usando uma região específica
            if regiao is not None:
                centro_x += regiao[0]
                centro_y += regiao[1]
            
            log_mensagem(f"Botão '{texto_botao}' encontrado com confiança {prob:.2f}! Coordenadas: ({centro_x}, {centro_y})")
            return centro_x, centro_y, texto
    
    log_mensagem(f"Botão com texto '{texto_botao}' não encontrado na tela.")
    return None

def verificar_mudanca_tela(antes, depois, limiar=0.9):
    """Verifica se houve mudança significativa na tela após clicar no botão"""
    try:
        # Converte para escala de cinza para comparação mais rápida
        antes_gray = np.mean(antes, axis=2) if len(antes.shape) > 2 else antes
        depois_gray = np.mean(depois, axis=2) if len(depois.shape) > 2 else depois
        
        # Calcula a diferença entre as imagens
        diff = np.abs(antes_gray - depois_gray)
        
        # Calcula a média das diferenças
        media_diff = np.mean(diff)
        
        # Se a diferença média for maior que o limiar, houve mudança na tela
        mudanca = media_diff > limiar
        
        log_mensagem(f"Diferença detectada na tela: {media_diff:.4f} (limiar: {limiar})")
        log_mensagem(f"Verificação de mudança na tela: {'Detectada' if mudanca else 'Não detectada'}")
        
        # Se não detectou mudança, espera um pouco e tenta novamente
        if not mudanca:
            time.sleep(2)
            # Captura a tela novamente
            nova_depois = np.array(ImageGrab.grab())
            diff = np.abs(antes_gray - np.mean(nova_depois, axis=2))
            media_diff = np.mean(diff)
            mudanca = media_diff > limiar
            log_mensagem(f"Segunda verificação - Diferença: {media_diff:.4f} (limiar: {limiar})")
            log_mensagem(f"Segunda verificação de mudança: {'Detectada' if mudanca else 'Não detectada'}")
        
        return mudanca
    except Exception as e:
        log_mensagem(f"Erro ao verificar mudança na tela: {str(e)}")
        # Em caso de erro, retorna True para não interromper o processo
        return True

def clicar_botao_por_texto(texto_botao, regiao=None, confianca=0.6, timeout=5, verificar_clique=True):
    """Localiza um botão pelo texto e clica nele"""
    inicio = time.time()
    while time.time() - inicio < timeout:
        # Tenta encontrar o botão
        resultado = encontrar_botao_por_texto(texto_botao, regiao, confianca)
        if resultado:
            x, y, texto_encontrado = resultado
            log_mensagem(f"Clicando no botão '{texto_encontrado}' nas coordenadas ({x}, {y})")
            pyautogui.click(x, y)
            time.sleep(2)  # Espera após o clique
            return True
        
        time.sleep(1)  # Espera entre tentativas
    
    log_mensagem(f"Timeout: Não foi possível encontrar o botão '{texto_botao}' após {timeout} segundos")
    return False

def ler_planilha():
    """Função para ler a planilha Excel com os casos"""
    try:
        # Caminho da planilha Base.xlsx
        caminho_atual = os.path.dirname(os.path.abspath(__file__))
        caminho_planilha = os.path.join(caminho_atual, 'Base.xlsx')
        log_mensagem(f"Tentando ler a planilha: {caminho_planilha}")
        
        df = pd.read_excel(caminho_planilha)
        log_mensagem(f"Planilha lida com sucesso. {len(df)} casos encontrados.")
        return df
    except FileNotFoundError:
        log_mensagem(f"Erro: Arquivo 'Base.xlsx' não encontrado em {caminho_planilha}")
        return None
    except Exception as e:
        log_mensagem(f"Erro ao ler a planilha: {str(e)}")
        return None

def verificar_campo_referencia(regiao, max_tentativas=3):
    """Verifica se o campo referência contém o texto correto"""
    for tentativa in range(max_tentativas):
        # Captura a região do campo
        screenshot = ImageGrab.grab(bbox=regiao)
        img_np = np.array(screenshot)
        
        # Usa OCR para ler o texto
        global reader
        if reader is None:
            reader = easyocr.Reader(['pt'], gpu=False)
        
        resultados = reader.readtext(img_np, detail=0)
        texto_encontrado = ' '.join(resultados).strip()
        
        log_mensagem(f"Verificando campo referência - Tentativa {tentativa + 1}")
        log_mensagem(f"Texto encontrado: '{texto_encontrado}'")
        
        if "CI-PROCESSO SEM MOVIMENTACAO" in texto_encontrado:
            log_mensagem("Campo referência preenchido corretamente!")
            return True
        
        log_mensagem("Campo referência não está preenchido corretamente. Aguardando...")
        time.sleep(2)
    
    return False

def processar_casos():
    """Função principal para processar cada caso da planilha"""
    df = ler_planilha()
    if df is None:
        log_mensagem("Não foi possível ler a planilha. Encerrando automação.")
        return
    
    # Cria um arquivo de log para os GCPJs processados
    with open("gcpjs_processados.txt", "a", encoding="utf-8") as f:
        f.write(f"\n--- INÍCIO DA EXECUÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
    
    # Itera sobre cada linha da planilha
    total_casos = len(df)
    casos_processados = 0
    casos_com_erro = 0
    
    log_mensagem(f"Iniciando processamento de {total_casos} casos...")
    
    for index, row in df.iterrows():
        try:
            gcpj = str(row['CÓDIGO GCPJ'])
            log_mensagem(f"Processando caso {index+1}/{total_casos} - GCPJ: {gcpj}")
            
            # Copia e cola o código GCPJ
            pyautogui.click(CAMPO_GCPJ[0], CAMPO_GCPJ[1])
            pyperclip.copy(gcpj)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(10)
            
            # Clica no botão pesquisar
            pyautogui.press('enter')
            time.sleep(15)
            
            # Pressiona a seta para baixo 211 vezes para encontrar a opção
            log_mensagem("Pressionando seta para baixo 211 vezes para encontrar 'CI-PROCESSO SEM MOVIMENTACAO'...")
            for i in range(211):
                pyautogui.press('down')
                if i % 50 == 0 and i > 0:  # Log a cada 50 pressionamentos
                    log_mensagem(f"Pressionado {i} vezes...")
            
            log_mensagem("Concluído o pressionamento de 211 vezes da seta para baixo")
            time.sleep(5)  # Espera após pressionar as setas
            
            
            # Verifica se o campo referência foi preenchido corretamente
            regiao_referencia = (CAMPO_REFERENCIA[0]-250, CAMPO_REFERENCIA[1]-10, 
                               CAMPO_REFERENCIA[0]+350, CAMPO_REFERENCIA[1]+25)
            
            if not verificar_campo_referencia(regiao_referencia):
                # Tenta mais uma vez caso a primeira verificação falhe
                time.sleep(2)
                if not verificar_campo_referencia(regiao_referencia):
                    log_mensagem("ERRO CRÍTICO: Não foi possível confirmar o preenchimento correto do campo referência após várias tentativas")
                    with open("gcpjs_processados.txt", "a", encoding="utf-8") as f:
                        f.write(f"GCPJ: {gcpj} - ERRO: Campo referência não foi preenchido corretamente\n")
                    return  # Interrompe toda a execução do programa
            
            # Cola a descrição
            pyautogui.press('tab')
            time.sleep(8)
            pyperclip.copy("Sem Movimentação processual no período")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(8)
            
            # Clica no botão salvar
            pyautogui.press('tab')
            time.sleep(8)
            pyautogui.press('enter')
            
            time.sleep(10)
            
            # Aperta enter
            pyautogui.press('enter')
            time.sleep(5)
            
            # Registra o sucesso
            log_mensagem(f"Caso GCPJ {gcpj} processado com sucesso")
            with open("gcpjs_processados.txt", "a", encoding="utf-8") as f:
                f.write(f"GCPJ: {gcpj} - CONCLUÍDO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            
            casos_processados += 1
            time.sleep(8)  # Pequena pausa antes de passar para o próximo caso
            
        except Exception as e:
            log_mensagem(f"Erro ao processar caso {index+1} (GCPJ: {gcpj}): {str(e)}")
            with open("gcpjs_processados.txt", "a", encoding="utf-8") as f:
                f.write(f"GCPJ: {gcpj} - ERRO: {str(e)}\n")
            casos_com_erro += 1
            continue
    
    # Registra o resumo da execução
    log_mensagem(f"Processamento concluído. Total: {total_casos}, Sucesso: {casos_processados}, Erros: {casos_com_erro}")
    with open("gcpjs_processados.txt", "a", encoding="utf-8") as f:
        f.write(f"\n--- FIM DA EXECUÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        f.write(f"Total: {total_casos}, Sucesso: {casos_processados}, Erros: {casos_com_erro}\n\n")

if __name__ == "__main__":
    # Verifica se as bibliotecas necessárias estão instaladas
    try:
        import easyocr
        import numpy as np
        from PIL import ImageGrab
    except ImportError:
        print("ERRO: Bibliotecas necessárias não estão instaladas.")
        print("Execute os seguintes comandos para instalar:")
        print("pip install easyocr numpy pillow")
        sys.exit(1)
    
    # Pequena pausa inicial para dar tempo de posicionar a janela
    log_mensagem("Iniciando automação em 5 segundos...")
    log_mensagem("ATENÇÃO: Posicione a janela corretamente antes de iniciar")
    time.sleep(10)
    
    processar_casos()
    
    log_mensagem("Automação concluída!")