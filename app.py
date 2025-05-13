import os
import customtkinter as ctk
from PIL import Image
from automacoes.baixarAnexos.RoboAnexo import RoboAnexo
import automacoes.roboRC.RoboRC as roboRC
import automacoes.upload.api as uploadAPI
from automacoes.teste.temp import Temporario
import automacoes.verificador.script as verificador


class AutomacaoApp:
    def __init__(self):
        """Initialize the automation application with UI setup and theme configuration."""
        # Configuração do tema
        self.tema_atual = "dark"  # Variável para controlar o tema
        ctk.set_appearance_mode(self.tema_atual)
        ctk.set_default_color_theme("blue")

        # Criar janela principal
        self.janela = ctk.CTk()
        self.janela.title("Central de Automação - Barros ADV")
        self.janela.geometry("900x600")

        # Inicializa componentes de interface
        self.setup_ui()

    def setup_ui(self):
        # Define os componentes de interface
        self.create_side_frame()
        self.create_main_frame()
        self.criar_botoes()
        self.create_log_area()
        self.mostrar_home()

    def create_side_frame(self):
        """Create the side frame with theme toggle and automation buttons."""
        self.frame_lateral = ctk.CTkFrame(self.janela, width=250)
        self.frame_lateral.pack(side="left", fill="y", padx=10, pady=10)

        # Título lateral
        self.label_titulo = ctk.CTkLabel(
            self.frame_lateral, 
            text="Painel de Automações", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.label_titulo.pack(pady=20)

        # Botão para alternar tema
        self.btn_tema = ctk.CTkButton(
            self.frame_lateral,
            text="🌙 Tema Escuro" if self.tema_atual == "dark" else "☀️ Tema Claro",
            command=self.alternar_tema,
            width=220,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#444",
            hover_color="#333"
        )
        self.btn_tema.pack(pady=8)

    def create_main_frame(self):
        """Create the main frame with scrollable content and image display."""
        self.frame_principal = ctk.CTkScrollableFrame(self.janela)
        self.frame_principal.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.frame_imagem = ctk.CTkFrame(
            self.frame_principal,
            fg_color="transparent",
            height=150
        )
        self.frame_imagem.pack(pady=(20, 10))

        self.carregar_logo()

    def carregar_logo(self):
        try:
            diretorio_atual = os.path.dirname(os.path.abspath(__file__))
            caminho_light = os.path.join(diretorio_atual, "assets/logos/logo-azul.png")
            caminho_dark = os.path.join(diretorio_atual, "assets/logos/logo-branca.png")

            # Carrega e redimensiona a imagem
            imagem = ctk.CTkImage(
                light_image=Image.open(caminho_light),
                dark_image=Image.open(caminho_dark),
                size=(180, 100)
            )

            # Cria label com a imagem
            self.label_imagem = ctk.CTkLabel(
                self.frame_imagem,
                text="",
                image=imagem
            )
            self.label_imagem.pack()

        except Exception as e:
            print(f"Erro detalhado ao carregar imagem: {str(e)}")
            self.label_imagem = ctk.CTkLabel(
                self.frame_imagem,
                text=f"Erro ao carregar imagem:\n{str(e)}",
                font=ctk.CTkFont(size=12),
                wraplength=200
            )
            self.label_imagem.pack()

    def criar_botoes(self):
        # Botão Home
        btn_home = ctk.CTkButton(
            self.frame_lateral,
            text="Home",
            command=self.mostrar_home,
            width=220,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2d5a27",
            hover_color="#1f3d1b"
        )
        btn_home.pack(pady=8)

        # Separador
        separador = ctk.CTkFrame(self.frame_lateral, height=2, fg_color="gray50")
        separador.pack(fill="x", padx=10, pady=10)

        # Frame para os botões de automação
        frame_botoes = ctk.CTkFrame(self.frame_lateral, fg_color="transparent")
        frame_botoes.pack(fill="x", expand=False, padx=5)

        # Botões de automação
        automacoes = [
            "Baixar Anexos",
            "Upload",
            "Cadastro DRC",
            "Verificador",
        ]

        for automacao in automacoes:
            btn = ctk.CTkButton(
                frame_botoes,
                text=automacao,
                command=lambda x=automacao: self.mostrar_automacao(x),
                width=220,
                height=40,
                font=ctk.CTkFont(size=14)
            )
            btn.pack(pady=8)

    def create_log_area(self):
        """Create the log area for displaying automation details and progress."""
        # Frame para a área de conteúdo
        self.frame_conteudo = ctk.CTkFrame(self.frame_principal)
        self.frame_conteudo.pack(fill="both", expand=True, padx=10, pady=10)

        # Frame para título da automação
        self.frame_titulo = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        self.frame_titulo.pack(fill="x", padx=20, pady=(10, 5))

        # Label para título (inicialmente vazio)
        self.label_titulo_automacao = ctk.CTkLabel(
            self.frame_titulo,
            text="",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.label_titulo_automacao.pack(pady=(10, 5))

        # Frame para descrição
        self.frame_descricao = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        self.frame_descricao.pack(fill="x", padx=20, pady=5)

        # Labels para descrição
        self.label_subtitulo_desc = ctk.CTkLabel(
            self.frame_descricao,
            text="",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.label_subtitulo_desc.pack(anchor="w")

        self.label_descricao = ctk.CTkLabel(
            self.frame_descricao,
            text="",
            font=ctk.CTkFont(size=14),
            justify="left",
            wraplength=600
        )
        self.label_descricao.pack(anchor="w", pady=5)

        # Frame para parâmetros
        self.frame_parametros = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        self.frame_parametros.pack(fill="x", padx=20, pady=5)

        # Labels para parâmetros
        self.label_subtitulo_param = ctk.CTkLabel(
            self.frame_parametros,
            text="",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.label_subtitulo_param.pack(anchor="w")

        self.label_parametros = ctk.CTkLabel(
            self.frame_parametros,
            text="",
            font=ctk.CTkFont(size=14),
            justify="left",
            wraplength=600
        )
        self.label_parametros.pack(anchor="w", pady=5)

        # Frame para controles (botão e barra de progresso)
        self.frame_controles = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        self.frame_controles.pack(fill="x", padx=20, pady=20)

        self.entrada_label = ctk.CTkLabel(
            self.frame_controles,
            text="Entre com o caminho do diretório dos anexos:",
            font=ctk.CTkFont(size=14, weight="bold"),
            justify="center",
            wraplength=600
        )

        self.entrada = ctk.CTkEntry(self.frame_controles, width=300)
        self.entrada.pack_forget()

        # Botão de execução (inicialmente escondido)
        self.btn_executar = ctk.CTkButton(
            self.frame_controles,
            text="Executar Automação",
            command=lambda: None,  # Será atualizado quando selecionar uma automação
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=8
        )

        # Barra de progresso (inicialmente escondida)
        self.barra_progresso = ctk.CTkProgressBar(
            self.frame_controles,
            width=300,
            height=15,
            corner_radius=5,
            progress_color="#4da6ff"
        )

        # Esconde os controles inicialmente
        self.btn_executar.pack_forget()
        self.barra_progresso.pack_forget()

    def alternar_tema(self):
        """Toggle between light and dark themes."""
        self.tema_atual = "light" if self.tema_atual == "dark" else "dark"
        ctk.set_appearance_mode(self.tema_atual)

        # Atualiza o texto do botão
        novo_texto = "🌙Tema Escuro" if self.tema_atual == "dark" else "☀️Tema Claro"
        self.btn_tema.configure(text=novo_texto)

    def iniciar(self):
        """Start the main application loop."""
        self.janela.mainloop()

    def obter_entrada(self):
        caminho = self.entrada.get()
        return caminho

    def mostrar_automacao(self, nome_automacao):
        # Atualiza o título
        self.label_titulo_automacao.configure(text=nome_automacao)
        
        # Atualiza a descrição
        self.label_subtitulo_desc.configure(text="Descrição")
        self.label_descricao.configure(text=self.obter_descricao(nome_automacao))
        
        # Atualiza os parâmetros
        self.label_subtitulo_param.configure(text="Parâmetros")
        self.label_parametros.configure(text=self.obter_parametros(nome_automacao))
        
        # Configura e mostra o botão de execução
        self.btn_executar.configure(
            command=lambda: self.executar_automacao(nome_automacao)
        )

        if nome_automacao != 'Upload':
            self.entrada.grid_forget()
            self.entrada_label.grid_forget()
        else:
            self.entrada_label.grid(row=1,column=1,pady=0)
            self.entrada.grid(row=2,column=1,pady= 10)
            self.entrada.delete(0, ctk.END)
        
        self.btn_executar.grid(row=3,column=1,pady=5)
        
        # Mostra e reseta a barra de progresso
        self.barra_progresso.grid(row=4,column=1,pady=5)
        self.barra_progresso.set(0)

    def obter_descricao(self, nome_automacao):
        descricoes = {
            "Baixar Anexos": """Esta automação é responsável por realizar o download de novos anexos.
            
• Baixa os novos anexos via GCPJ
• Organiza os anexos em subpastas dentro do diretório 'Anexos' na sua área de trabalho""",
            
            "Cadastro DRC": """Automatiza o processo de envio de relatórios diários.
            
• Coleta dados do sistema principal
• Formata as informações em planilhas
• Envia e-mails automáticos para a lista de destinatários""",
            
            "Upload" : """Essa Automação é responsável Upload de novos anexos.

• Renomear os anexos para o formato aceito no sistema
• Subir os novo anexos no sistema Octopus""",

            "Verificador" : """Essa automação é responsável por verificar anexos não cadastrados no sistema."""
        }
        return descricoes.get(nome_automacao, "Descrição não disponível")

    def obter_parametros(self, nome_automacao):
        parametros = {
            "Baixar Anexos": """• Entrada: Planilha excel com os processos GCPJ que deverão ser processados
• Formato de saída: Documentos organizados em subdiretórios na pasta 'Anexos' """,
            
            "Cadastro DRC": """• Lista de destinatários
• Horário de execução
• Tipo de relatório""",

            "Upload" : """• Entrada: o Caminho do diretório onde há anexos que deseja subir.
•Saída: Novos sub diretórios organizados em 'FEITOS', 'VAZIOS' caso a subpasta estiver vazia e 'NÃO ENCONTRADOS' caso o processo não esteja no sistema""",

            "Verificador" : """• Entrada: Planilha excel com os processos GCPJ que deverão ser verificados
• Formato de saída: Nova planilha excel na sua Área de Trabalho com os casos que não possuem pasta no sistema. """

        }
        return parametros.get(nome_automacao, "Sem parâmetros disponíveis")

    def executar_automacao(self, nome_automacao):
        # Inicia o progresso
        self.barra_progresso.set(0)

        # Criar uma instância de outra classe
        instancia_robo_anexo = RoboAnexo()
        instancia_temporario = Temporario()

        # Dicionário de automações com funções associadas
        automacoes = {
            "Baixar Anexos": lambda: instancia_robo_anexo.main(),
            "Upload" : lambda: uploadAPI.main(self.obter_entrada()),
            "Cadastro DRC": lambda: instancia_temporario.main(),
            "Verificador" : verificador.main,
        }
        
        # Obter a função de automação com base no nome
        func_automacao = automacoes.get(nome_automacao)

        if func_automacao != None:
            try:
                func_automacao()  # Chama a função associada
                print(f"Processo '{nome_automacao}' finalizado.")
            except Exception as e:
                print(f"Erro ao executar automação '{nome_automacao}': {str(e)}")
        else:
            print("Automação não encontrada")

        # Simula conclusão
        self.barra_progresso.set(1)
        return "Automação concluída."

    def mostrar_home(self):
        # Esconde os controles de automação
        self.btn_executar.grid_forget()
        self.barra_progresso.grid_forget()
        self.entrada.grid_forget()
        self.entrada_label.grid_forget()
        
        # Atualiza o título com emoji e cor personalizada
        self.label_titulo_automacao.configure(
            text="⭐ Central de Automações BARROS",
            font=ctk.CTkFont(size=28, weight="bold"),
            wraplength=600
        )  
        
        # Atualiza a descrição
        self.label_subtitulo_desc.configure(
            text="✨ Principais Recursos",
            font=ctk.CTkFont(size=18, weight="bold"),
            wraplength=600
        )
        
        self.label_descricao.configure(
            text="""Automações de tarefas diárias dos escritório da Barros & Advogados:

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚡ Automações Disponíveis:


• ⬇️ Baixar Anexo GCPJ
   Faça download de forma automática de documentos do sistema GCPJ

• 📈 Cadastro DRC
   Geração automática de relatórios


💡 Dica: Selecione uma automação no menu lateral para começar!""",
            font=ctk.CTkFont(size=13),  # Reduzido tamanho da fonte
            wraplength=550,
            justify="left"
        )
        
        # Atualiza os parâmetros com estatísticas
        self.label_subtitulo_param.configure(
            text="📊 Painel de Controle",
            font=ctk.CTkFont(size=18, weight="bold"),
            wraplength=600
        )
               
        
        self.label_parametros.configure(
            text=f"""⚙️ Informações do Sistema:
            
• 📌 Versão: 1.0.0 Beta

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━""",
            font=ctk.CTkFont(size=13),  # Reduzido tamanho da fonte
            wraplength=550,
            justify="left"
        )

if __name__ == "__main__":
    app = AutomacaoApp()
    app.iniciar()