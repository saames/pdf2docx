from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from pdf2docx import Converter

class Tela:
    def __init__(self, master):
        self._janela = master
        self._janela.geometry('320x120')
        self._janela.title('PDF para DOCX')

        self._lbl_titulo = tk.Label(self._janela, text='PDF para DOCX', bg="#014078", fg='#FFFFFF', font=('Arial', 14, 'bold'))
        self._lbl_titulo.pack(fill='x')
        self._btn_selecionarArquivos = tk.Button(self._janela, text='Selecionar Arquivos', command=self.selecionar_arquivos, bg="#014078", fg='#FFFFFF', width=18, font=('Arial', 10, 'bold'))
        self._btn_selecionarArquivos.pack(pady='30')

        self.centraliza(self._janela)

    def selecionar_arquivos(self):
        self._caminhoArquivosPDF = filedialog.askopenfilenames()
        if len(self._caminhoArquivosPDF)==0:
            raise Exception("No files selected")
        for arquivo in self._caminhoArquivosPDF:
            if not '.pdf' in arquivo:
                messagebox.showerror('ERRO', 'Selecione apenas arquivos em PDF.')
                raise Exception("Incorrect File Format")
        
        #Cria uma nova janela para confirmar a seleção de arquivos
        self._janela_confirmar = tk.Toplevel(self._janela)
        
        colunas = ['nomeArq', 'dataModificacaoArq','tamanhoArq']
        self._tvw_tabelaArquivosSelecionados = ttk.Treeview(self._janela_confirmar, height=10, columns=colunas, show='headings')
        #Configura o cabaçalho das colunas
        self._tvw_tabelaArquivosSelecionados.heading('nomeArq', text='Nome do arquivo')
        self._tvw_tabelaArquivosSelecionados.heading('dataModificacaoArq', text='Data de modificação')
        self._tvw_tabelaArquivosSelecionados.heading('tamanhoArq', text='Tamanho')
        #Ajusta tamanho da coluna
        self._tvw_tabelaArquivosSelecionados.column('nomeArq', width=300, minwidth=150)
        self._tvw_tabelaArquivosSelecionados.column('dataModificacaoArq', width=120, minwidth=120)
        self._tvw_tabelaArquivosSelecionados.column('tamanhoArq', width=100, minwidth=50)

        #Preencher a tabela
        for arquivo in self._caminhoArquivosPDF:
            dataMod = datetime.fromtimestamp(os.path.getctime(arquivo)).strftime('%d/%m/%Y %H:%M')
            tamanhoMB = (os.path.getsize(arquivo))
            if tamanhoMB >= 1048576: # (1 MB em bytes)
                tamanhoMB = f'{(tamanhoMB/1024)/1024:.2f} MB'
            else:
                tamanhoMB = f'{tamanhoMB/1024:.2f} KB'
            
            self._tvw_tabelaArquivosSelecionados.insert('', 'end', values=(os.path.basename(arquivo), dataMod, tamanhoMB))
        
        self._tvw_tabelaArquivosSelecionados.grid(row=0, column=0)
        self._scb_barraTabela = ttk.Scrollbar(self._janela_confirmar, command=self._tvw_tabelaArquivosSelecionados.yview)
        self._scb_barraTabela.grid(row=0, column=1, sticky='ns')
        self._tvw_tabelaArquivosSelecionados.configure(yscrollcommand=self._scb_barraTabela.set)

        self._ftm_botoes = tk.Frame(self._janela_confirmar)
        self._ftm_botoes.grid(row=1, column=0, sticky='EW')

        self._btn_excluir = tk.Button(self._ftm_botoes, text='Remover da lista', command=self.excluir, bg="#B80F0A", fg="#FFFFFF", width=17, font=('Arial', 10, 'bold'))
        self._btn_excluir.pack(side='left', padx=5, pady=5)
        self._btn_converter = tk.Button(self._ftm_botoes, text='Converter arquivos', command=self.converter_arquivo, bg="#014078", fg='#FFFFFF', width=17, font=('Arial', 10, 'bold'))
        self._btn_converter.pack(side='right', padx=5, pady=5)

        self.centraliza(self._janela_confirmar)
        
    def excluir(self):
        itens = self._tvw_tabelaArquivosSelecionados.selection()
        if len(itens) > 0:
            resposta = messagebox.askquestion('Confirmar', 'Tem certeza da remoção?', parent=self._janela_confirmar)
            _temp_caminhoArquivosPDF = list(self._caminhoArquivosPDF)
            if resposta == 'yes':
                for item in itens:
                    nomeArq = self._tvw_tabelaArquivosSelecionados.item(item, 'values')[0]
                    for caminho in _temp_caminhoArquivosPDF:
                        if nomeArq in caminho:
                            _temp_caminhoArquivosPDF.remove(caminho)
                    self._tvw_tabelaArquivosSelecionados.delete(item)
            self._caminhoArquivosPDF = tuple(_temp_caminhoArquivosPDF)
        else:
            messagebox.showwarning('Aviso', 'Escolha pelo menos um arquivo para remover da lista de conversão.', parent=self._janela_confirmar)

    def converter_arquivo(self):
        for arquivo in self._caminhoArquivosPDF:
            temp_caminho = arquivo.split('/')
            temp_caminho.pop()
            caminho = '/'.join(temp_caminho)
            arquivoDOCX = caminho + '/' + os.path.splitext(os.path.basename(arquivo))[0] + '.docx'
            self._converter = Converter(arquivo)
            self._converter.convert(arquivoDOCX)
            self._converter.close()
        
        self._janela_confirmar.destroy()
        messagebox.showinfo('Status da conversão', 'Arquivo(s) convertido(s) com sucesso!')

    def centraliza(self, master): # Centraliza a janela quando gerada
        largura_monitor = master.winfo_screenwidth()
        altura_monitor = master.winfo_screenheight()
        master.update_idletasks()
        largura_janela = master.winfo_width()
        altura_janela = master.winfo_height()
        x = largura_monitor // 2 - largura_janela // 2
        y = altura_monitor //2 - altura_janela // 2
        master.geometry(f'{largura_janela}x{altura_janela}+{x}+{y}')

gui = tk.Tk()
Tela(gui)
gui.mainloop()