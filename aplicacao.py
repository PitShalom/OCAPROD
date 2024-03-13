
import os
from tkinter import Tk, filedialog, ttk, Button, Label, messagebox, Entry, StringVar, PhotoImage,Canvas
import pandas as pd
import os
from PyPDF2 import PdfFileWriter
from tqdm import tqdm
from datetime import datetime
from ttkthemes import ThemedStyle
from funcoes_pdf import (preencher_nr01, preencher_nr06, preencher_nr18, preencher_nr35,preencher_nr10comp,
                         preencher_fichaEPI,preencher_fichaEPI_adm_geral,preencher_fichaEPI_adm_obra,preencher_fichaEPI_almoxarife,preencher_fichaEPI_civil,preencher_fichaEPI_hidra,preencher_fichaEPI_solda,
                         preencher_CA, preencher_nr05, preencher_nr10basic,
                         preencher_nr11,preencher_nr12, preencher_nr17,
                         preencher_nr18_pemt, preencher_nr20_infla, preencher_nr20_brigada,
                         preencher_nr33, preencher_nr34, preencher_nr34_adm,
                         preencher_nr34_obs_quente, preencher_cracha, preencher_OS_adm_geral,
                         preencher_OS_aumoxarifado, preencher_OS_obras_civil, preencher_OS_adm_obra,
                         preencher_OS_obras_eletricas, preencher_OS_obras_hidraulicas, preencher_OS_soldador,
                         preencher_prova_NR06,preencher_prova_NR10,preencher_prova_NR11,preencher_prova_NR12,preencher_prova_NR17,
                         preencher_prova_NR18,preencher_prova_NR18pta,preencher_prova_NR20_infla,preencher_prova_NR33,preencher_prova_NR34,preencher_prova_NR35,
                         preencher_treino_NR01)

class Aplicacao:
    def __init__(self, root):
        self.root = root
        self.root.title("Preenchimento Automático de PDF")
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "FrontEnd", "imgOca.png")
        self.root = root
        self.root.title("Preenchimento Automático de PDF")

        # Configuração da imagem de fundo
        script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, "FrontEnd", "imgOca.png")

        if os.path.isfile(image_path):
            self.background_image = PhotoImage(file=image_path)
            self.canvas = Canvas(root, width=self.background_image.width(), height=self.background_image.height())
            self.canvas.pack()

            # Adiciona a imagem de fundo ao canvas
            self.canvas.create_image(0, 0, anchor="nw", image=self.background_image)

        # Verifica se o ícone existe antes de tentar configurá-lo
        if os.path.isfile(icon_path):
            icon = PhotoImage(file=icon_path)
            self.root.iconphoto(True, icon)

        style = ThemedStyle(self.root)
        style.set_theme("radiance")

        self.diretorio_modelos_pdf = r'C:\progOca\modCert'
        self.modelo_nr01 = 'nr01.pdf'
        self.modelo_nr05 = 'nr05.pdf'
        self.modelo_nr06 = 'nr06.pdf'
        self.modelo_10basic = 'nr10_basic.pdf'
        self.modelo_10comp = 'nr10_comp.pdf'
        self.modelo_11 = 'nr11.pdf'
        self.modelo_12 = 'nr12.pdf'
        self.modelo_17 = 'nr17.pdf'
        self.modelo_nr18 = 'nr18.pdf'
        self.modelo_nr18_pemt = 'nr18_pemt.pdf'
        self.modelo_nr20_infla = 'nr20_infla.pdf'
        self.modelo_nr20_brigada = 'nr20_brigada.pdf'
        self.modelo_nr33 = 'nr33.pdf'
        self.modelo_nr34 = 'nr34.pdf'
        self.modelo_nr34_adm = 'nr34_adm.pdf'
        self.modelo_nr34_obs_quente = 'nr34_obs_trab_quente.pdf'
        self.modelo_nr20_brigada = 'nr20_brigada.pdf'
        self.modelo_nr35 = 'nr35.pdf'
        self.modelo_OS_adm_geral = 'O.S - GHE ADM. GERAL.pdf'
        self.modelo_OS_adm_de_obra = 'O.S - GHE ADM DE OBRA.pdf'
        self.modelo_OS_aumoxarifado = 'O.S - GHE ALMOXARIFADO.pdf'
        self.modelo_OS_obra_civil = 'O.S - GHE OBRAS CIVIL.pdf'
        self.modelo_OS_obra_eletrica = 'O.S - GHE OBRAS ELÉTRICA.pdf'
        self.modelo_OS_obra_hidraulica = 'O.S - GHE OBRAS HIDRÁULICA.pdf'
        self.modelo_OS_soldador = 'O.S - GHE OBRAS SOLDA.pdf'
        self.modelo_CA = 'C.A.pdf'
        self.modelo_fichaEPI = 'F.EPI.pdf'
        self.modelo_epi_adm_geral = 'F.EPI - GHE 1 - ADM GERAL.pdf'
        self.modelo_epi_adm_obra = 'F.EPI - GHE 2 - ADM DE OBRA.pdf'
        self.modelo_epi_almoxarife = 'F.EPI - GHE 3 - ALMOXARIFE.pdf'
        self.modelo_epi_civil = 'F.EPI - GHE 5 - CIVIL.pdf'
        self.modelo_epi_eletrica = 'F.EPI - GHE 6 - ELÉTRICA.pdf'
        self.modelo_epi_hidra = 'F.EPI - GHE 7 - HIDRÁULICA.pdf'
        self.modelo_epi_solda = 'F.EPI - GHE 8 - SOLDA.pdf'
        self.modelo_provaNR06 = 'NR06_PROVA.pdf'
        self.modelo_provaNR10 = 'NR10 - PROVA.pdf'
        self.modelo_provaNR11 = 'NR11 - PROVA.pdf'
        self.modelo_provaNR12 = 'NR12 - PROVA.pdf'
        self.modelo_provaNR17 = 'NR17 - PROVA.pdf'
        self.modelo_provaNR18 = 'NR18 - PROVA.pdf'
        self.modelo_provaNR18_PTA = 'NR18 - PTA - PROVA.pdf'
        self.modelo_provaNR20 = 'NR20 - PROVA.pdf'
        self.modelo_provaNR33 = 'NR33 - PROVA.pdf'
        self.modelo_provaNR34 = 'NR34 - PROVA.pdf'
        self.modelo_provaNR35 = 'NR35 - PROVA.pdf'
        self.modelo_treinoNR01 = 'RELATÓRIO DE TREINAMENTO - NR01.pdf'
        self.modelo_treinoNR06 = 'RELATÓRIO DE TREINAMENTO - NR06.pdf'
        self.modelo_treinoNR10 = 'RELATÓRIO DE TREINAMENTO - NR10.pdf'
        self.modelo_treinoNR10c = 'RELATÓRIO DE TREINAMENTO - NR10C.pdf'
        self.modelo_treinoNR11 = 'RELATÓRIO DE TREINAMENTO - NR11.pdf'
        self.modelo_treinoNR12 = 'RELATÓRIO DE TREINAMENTO - NR12.pdf'
        self.modelo_treinoNR17 = 'RELATÓRIO DE TREINAMENTO - NR17.pdf'
        self.modelo_treinoNR18civil = 'RELATÓRIO DE TREINAMENTO - NR18 CIVIL.pdf'
        self.modelo_treinoNR18pemt = 'RELATÓRIO DE TREINAMENTO - NR18 PEMT.pdf'
        self.modelo_treinoNR20 = 'RELATÓRIO DE TREINAMENTO - NR20.pdf'
        self.modelo_treinoNR33 = 'RELATÓRIO DE TREINAMENTO - NR33.pdf'
        self.modelo_treinoNR34adm = 'RELATÓRIO DE TREINAMENTO - NR34 - 1.admissional.pdf'
        self.modelo_treinoNR34obs = 'RELATÓRIO DE TREINAMENTO - NR34 - 2.observador.pdf'
        self.modelo_treinoNR34bas = 'RELATÓRIO DE TREINAMENTO - NR34 - 2.observador.pdf'
        self.modelo_treinoNR35 = 'RELATÓRIO DE TREINAMENTO - NR35.pdf'
        

        self.modelo_cracha = 'CRACHA_PLANEM.pdf'

        frame = ttk.Frame(root, style="TFrame")
        frame.pack(padx=10, pady=10)

        ttk.Label(frame, text="Selecione o arquivo Excel:", style="TLabel").grid(row=0, column=0, padx=10, pady=10)
        self.entry_excel = ttk.Entry(frame, state="readonly", width=50, textvariable=StringVar(), style="TEntry")
        self.entry_excel.grid(row=0, column=1, padx=10, pady=10)
        ttk.Button(frame, text="Selecionar Excel", command=self.selecionar_excel, style="TButton").grid(row=0, column=2, padx=10, pady=10)

        self.progress_bar = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.grid(row=2, column=1, pady=10)

        self.progress_label = ttk.Label(frame, text="", style="TLabel")
        self.progress_label.grid(row=3, column=1, pady=10)

        ttk.Button(frame, text="Preencher e Salvar NRs", command=self.preencher_e_salvar_nr, style="TButton").grid(row=4, column=1, pady=20)

    def selecionar_excel(self):
        self.caminho_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel")
        self.entry_excel.config(state="normal")
        self.entry_excel.delete(0, "end")
        self.entry_excel.insert(0, self.caminho_excel)
        self.entry_excel.config(state="readonly")

    def preencher_e_salvar_nr(self):
        if not self.caminho_excel:
            self.progress_label.config(text="Erro: Selecione o arquivo Excel.")
            self.root.update()
            return

        self.progress_label.config(text="Aguarde, preenchendo e salvando PDFs...")
        self.progress_bar["value"] = 0
        self.root.update()

        planilha = pd.read_excel(self.caminho_excel)
        total_rows = len(planilha)

        required_columns = ['NOME', 'CPF', 'FUNÇÃO', 'DATA_NR18', 'DATA_NR35','DATA_NR01','DATA_NR06',
                            'NOME_SUPERINTENDENTE_OBRA','N_REGISTRO_TST','CPF_SUPERINTENDENTE','NOME_SUPERINTENDENTE_OBRA',
                            'HABILITAÇÃO_SUPERINTENDENTE','Nº_REGISTRO_SUPERINTENDENTE','REGISTRO_MATRICULA_EMPREGADO','NOME_TST']
        missing_columns = [col for col in required_columns if col not in planilha.columns]

        if missing_columns:
            messagebox.showerror("Erro", f"As seguintes colunas estão ausentes no arquivo Excel: {', '.join(missing_columns)}")
            return

        for index, linha in planilha.iterrows():
            try:
                nome = linha['NOME']
                nome_obra = linha['NOME_OBRA']
                cpf = linha['CPF']
                funcao = linha['FUNÇÃO']
                nomeTecRep = linha['NOME_SUPERINTENDENTE_OBRA']
                n_superInt = linha['Nº_REGISTRO_SUPERINTENDENTE']
                Hab_SupInt = linha ['HABILITAÇÃO_SUPERINTENDENTE']
                cpf_superInt = linha ['CPF_SUPERINTENDENTE']
                registro_empregado_epi = linha ['REGISTRO_MATRICULA_EMPREGADO']
                nome_TST = linha ['NOME_TST']
                n_registroTST = linha['N_REGISTRO_TST']
                #---------------------------------------------
                data_aso = linha['DATA_ASO']
                dataNR01 = linha['DATA_NR01']
                dataNR05 = linha['DATA_NR05']
                dataNR06 = linha['DATA_NR06']
                dataNR10_basica = linha['DATA_NR10_basica']
                dataNR10_complementar = linha['DATA_NR10_complementar']
                dataNR11 = linha['DATA_NR11']
                dataNR12 = linha['DATA_NR12']
                dataNR17 = linha['DATA_NR17']
                dataNR18 = linha['DATA_NR18']
                dataNR18_pemt = linha['DATA_NR18_pta']
                dataNR20_inflamaveis = linha['DATA_20_inflamaveis']
                dataNR20_brigada = linha['DATA_NR20_brigada']
                dataNR33 = linha['DATA_NR33']
                dataNR34 = linha['DATA_NR34_basico']
                dataNR34_adm = linha['DATA_NR34_adimissional']
                dataNR34_obs_quente = linha['DATA_NR34_obs_quente']
                dataNR35 = linha['DATA_NR35']
                

                # Caminho da pasta do colaborador
                nome_colaborador = '_'.join(nome.split())
                colaborador_folder = os.path.join(self.diretorio_modelos_pdf, 'pdfBaixados', nome_colaborador)
                if not os.path.exists(colaborador_folder):
                    os.makedirs(colaborador_folder)

                data_atual = datetime.now().strftime('%d-%m-%y')

                output_path_nr01 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR01.pdf')
                output_path_nr05 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR05.pdf')
                output_path_nr06 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR06.pdf')
                output_path_nr10basic = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR10basica.pdf')
                output_path_nr10comp = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR10complementar.pdf')
                output_path_nr11 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR11.pdf')
                output_path_nr12 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR12.pdf')
                output_path_nr17 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR17.pdf')
                output_path_nr18 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR18.pdf')
                output_path_nr18_pemt = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR18_pemt.pdf')
                output_path_nr20_infla = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR20_brigada.pdf')
                output_path_nr20_brigada = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR20_inflamaveis.pdf')
                output_path_nr33 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR33.pdf')
                output_path_nr34 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR34.pdf')
                output_path_nr34_obs_quente = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR34_obs_quente.pdf')
                output_path_nr34_adm = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR34_adm.pdf')
                output_path_nr35 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_NR35.pdf')
                output_path_OS_adm_geral = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_OS_adm_geral.pdf')
                output_path_OS_adm_obra = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_OS_adm_obra.pdf')
                output_path_OS_aumoxarifado = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_OS_aumoxarifado.pdf')
                output_path_OS_obras_civil = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_OS_obras_civil.pdf')
                output_path_OS_obra_eletrica = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_OS_obras_eletricas.pdf')
                output_path_OS_obra_hidraulica = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_OS_obras_hidraulica.pdf')
                output_path_OS_soldador = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_OS_soldador.pdf')
              #-------------
                output_path_CA = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_C.A.pdf')
              #-----------
                output_path_fichaEPI = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_fichaEPI.pdf')
                output_path_epi_adm_geral = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_fichaEPI_Adm_Geral.pdf')
                output_path_epi_adm_obra = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_fichaEPI_Adm_Obra.pdf')
                output_path_epi_almoxarife = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_fichaEPI_Almoxarife.pdf')
                output_path_epi_civil = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_fichaEPI_Civil.pdf')
                output_path_epi_hidra = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_fichaEPI_Hidraulica.pdf')
                output_path_epi_solda = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_fichaEPI_solda.pdf')
                output_path_prova_NR06 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR06.pdf')
                output_path_prova_NR10 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR10.pdf')
                output_path_prova_NR11 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR11.pdf')
                output_path_prova_NR12 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR12.pdf')
                output_path_prova_NR17 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR17.pdf')
                output_path_prova_NR18 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR18.pdf')
                output_path_prova_NR18pta = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR18pta.pdf')
                output_path_prova_NR20infla = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR20_Inflamaveis.pdf')
                output_path_prova_NR33 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR33.pdf')
                output_path_prova_NR34 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR34.pdf')
                output_path_prova_NR35 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Prova_NR35.pdf')
                output_path_treino_NR01 = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_RELATÓRIO DE TREINAMENTO - NR01.pdf')
              #---------
                output_path_cracha = os.path.join(colaborador_folder, f'{nome_colaborador}_{data_atual}_Cracha.pdf')


                preencher_nr01(nome, cpf, funcao, dataNR01, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr01), output_path_nr01, incluir_funcao=False)
                preencher_nr05(nome, cpf, funcao, dataNR05, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr05), output_path_nr05, incluir_funcao=False)
                preencher_nr06(nome, cpf, funcao, dataNR06, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr06), output_path_nr06, incluir_funcao=False)
                preencher_nr10basic(nome, cpf, funcao, dataNR10_basica, nomeTecRep,Hab_SupInt, n_superInt, os.path.join(self.diretorio_modelos_pdf, self.modelo_10basic), output_path_nr10basic, incluir_funcao=False)
                preencher_nr10comp(nome, cpf, funcao, dataNR10_complementar, nomeTecRep,Hab_SupInt, n_superInt, os.path.join(self.diretorio_modelos_pdf, self.modelo_10comp), output_path_nr10comp, incluir_funcao=False)
                preencher_nr11(nome, cpf, funcao, dataNR11, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_11), output_path_nr11, incluir_funcao=False)
                preencher_nr12(nome, cpf, funcao, dataNR12, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_12), output_path_nr12, incluir_funcao=False)
                preencher_nr17(nome, cpf, funcao, dataNR17, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_17), output_path_nr17, incluir_funcao=False)
                preencher_nr18(nome, cpf, funcao, dataNR18, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr18), output_path_nr18, incluir_funcao=False)
                preencher_nr18_pemt(nome, cpf, funcao, dataNR18_pemt, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr18_pemt), output_path_nr18_pemt, incluir_funcao=False)
                preencher_nr20_brigada(nome, cpf, funcao, dataNR20_inflamaveis, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr20_brigada), output_path_nr20_brigada, incluir_funcao=False)
                preencher_nr20_infla(nome, cpf, funcao, dataNR20_brigada, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr20_infla), output_path_nr20_infla, incluir_funcao=False)
                preencher_nr33(nome, cpf, funcao, dataNR33, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr33), output_path_nr33, incluir_funcao=False)
                preencher_nr34(nome, cpf, funcao, dataNR34, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr34), output_path_nr34, incluir_funcao=False)
                preencher_nr34_adm(nome, cpf, funcao, dataNR34_adm, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr34_adm), output_path_nr34_adm, incluir_funcao=False)
                preencher_nr34_obs_quente(nome, cpf, funcao, dataNR34_obs_quente, nome_TST,Hab_SupInt, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr34_obs_quente), output_path_nr34_obs_quente, incluir_funcao=False)
                preencher_nr35(nome, cpf, funcao, dataNR35,Hab_SupInt, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_nr35), output_path_nr35, incluir_funcao=False)
               #------
                preencher_OS_adm_geral(nome, cpf, funcao, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_OS_adm_geral), output_path_OS_adm_geral, incluir_funcao=True)
                preencher_OS_adm_obra(nome, cpf, funcao, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_OS_adm_de_obra), output_path_OS_adm_obra, incluir_funcao=True)
                preencher_OS_aumoxarifado(nome, cpf, funcao, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_OS_aumoxarifado), output_path_OS_aumoxarifado, incluir_funcao=True)
                preencher_OS_obras_civil(nome, cpf, funcao, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_OS_obra_civil), output_path_OS_obras_civil, incluir_funcao=True)
                preencher_OS_obras_eletricas(nome, cpf, funcao, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_OS_obra_eletrica), output_path_OS_obra_eletrica, incluir_funcao=True)
                preencher_OS_obras_hidraulicas(nome, cpf, funcao, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_OS_obra_hidraulica), output_path_OS_obra_hidraulica, incluir_funcao=True)
                preencher_OS_soldador(nome, cpf, funcao, nome_TST, n_registroTST, os.path.join(self.diretorio_modelos_pdf, self.modelo_OS_soldador), output_path_OS_soldador, incluir_funcao=True)
               #------
                
                preencher_fichaEPI(nome,funcao, registro_empregado_epi, os.path.join(self.diretorio_modelos_pdf, self.modelo_fichaEPI), output_path_fichaEPI, incluir_funcao=True)
                preencher_fichaEPI_adm_geral(nome,funcao, registro_empregado_epi, os.path.join(self.diretorio_modelos_pdf, self.modelo_epi_adm_geral),output_path_epi_adm_geral, incluir_funcao=True)
                preencher_fichaEPI_adm_obra(nome,funcao, registro_empregado_epi, os.path.join(self.diretorio_modelos_pdf, self.modelo_epi_adm_obra), output_path_epi_adm_obra, incluir_funcao=True)
                preencher_fichaEPI_almoxarife(nome,funcao, registro_empregado_epi, os.path.join(self.diretorio_modelos_pdf, self.modelo_epi_almoxarife), output_path_epi_almoxarife, incluir_funcao=True)
                preencher_fichaEPI_civil(nome,funcao, registro_empregado_epi, os.path.join(self.diretorio_modelos_pdf, self.modelo_epi_civil), output_path_epi_civil, incluir_funcao=True)
                preencher_fichaEPI_hidra(nome,funcao, registro_empregado_epi, os.path.join(self.diretorio_modelos_pdf, self.modelo_epi_hidra), output_path_epi_hidra, incluir_funcao=True)
                preencher_fichaEPI_solda(nome,funcao, registro_empregado_epi, os.path.join(self.diretorio_modelos_pdf, self.modelo_epi_solda), output_path_epi_solda, incluir_funcao=True)

               #-------
                
                preencher_prova_NR06(nome, funcao,dataNR06, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR06), output_path_prova_NR06, incluir_funcao=False)
                preencher_prova_NR10(nome, funcao,dataNR10_basica, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR10), output_path_prova_NR10, incluir_funcao=False)
                preencher_prova_NR11(nome, funcao,dataNR11, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR11), output_path_prova_NR11, incluir_funcao=False)
                preencher_prova_NR12(nome, funcao,dataNR12, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR12), output_path_prova_NR12, incluir_funcao=False)
                preencher_prova_NR17(nome, funcao,dataNR17, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR17), output_path_prova_NR17, incluir_funcao=False)
                preencher_prova_NR18(nome, funcao,dataNR18, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR18), output_path_prova_NR18, incluir_funcao=False)
                preencher_prova_NR18pta(nome, funcao,dataNR18_pemt, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR18_PTA), output_path_prova_NR18pta, incluir_funcao=False)
                preencher_prova_NR20_infla(nome, funcao,dataNR20_inflamaveis, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR20), output_path_prova_NR20infla, incluir_funcao=False)
                preencher_prova_NR33(nome, funcao,dataNR33, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR33), output_path_prova_NR33, incluir_funcao=False)
                preencher_prova_NR34(nome, funcao,dataNR34, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR34), output_path_prova_NR34, incluir_funcao=False)
                preencher_prova_NR35(nome, funcao,dataNR35, os.path.join(self.diretorio_modelos_pdf, self.modelo_provaNR35), output_path_prova_NR35, incluir_funcao=False)

               #-------
                preencher_treino_NR01(nome,nome_TST, funcao,dataNR01, os.path.join(self.diretorio_modelos_pdf, self.modelo_treinoNR01), output_path_treino_NR01, incluir_funcao=False)

               #-------
                preencher_CA(nome, cpf, funcao, Hab_SupInt, n_superInt, cpf_superInt, nomeTecRep, os.path.join(self.diretorio_modelos_pdf, self.modelo_CA), output_path_CA, incluir_funcao=True)
                preencher_cracha(nome,nome_obra,funcao,data_aso,dataNR06,dataNR05,dataNR18,dataNR35,dataNR12,dataNR01,dataNR10_basica,dataNR10_complementar,dataNR11,dataNR18_pemt,dataNR20_inflamaveis,dataNR20_brigada,dataNR33,dataNR34,dataNR34_adm,dataNR34_obs_quente,dataNR17, os.path.join(self.diretorio_modelos_pdf, self.modelo_cracha), output_path_cracha, incluir_funcao=True)
                progress_value = (index + 1) / total_rows * 100
                self.progress_bar["value"] = progress_value
                self.root.update_idletasks()
            except Exception as e:
                if "File is not open for writing" in str(e):
                    messagebox.showerror("Erro", "Feche o PDF aberto antes de continuar.")
                    self.root.update()
                    return
                else:
                    messagebox.showerror("Erro", f"FECHE O PDF ABERTO: {str(e)}")
                    print({str(e)})
                    return
        self.progress_bar["value"] = 100
        self.progress_label.config(text="Concluído, preenchimento e salvamento concluídos com sucesso!")
        self.root.update()
        

    
if __name__ == "__main__":
    app = Aplicacao(Tk())
    app.root.mainloop()
