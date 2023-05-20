import os
from docx import Document
import tkinter as tk
from tkinter import messagebox, filedialog

class Aplicativo:
    def __init__(self, master):
        self.master = master
        self.master.title("Atualização em lote de documentos do Word")

        # Crie um rótulo para as instruções
        self.label = tk.Label(self.master, text="Digite as informações abaixo:")
        self.label.pack()

        # Crie uma entrada de texto para o nome
        self.nome_label = tk.Label(self.master, text="Nome:")
        self.nome_label.pack()
        self.nome_entry = tk.Entry(self.master)
        self.nome_entry.pack()

        # Crie uma entrada de texto para o endereço
        self.endereco_label = tk.Label(self.master, text="Endereço:")
        self.endereco_label.pack()
        self.endereco_entry = tk.Entry(self.master)
        self.endereco_entry.pack()

        # Crie uma entrada de texto para o CEP
        self.cep_label = tk.Label(self.master, text="CEP:")
        self.cep_label.pack()
        self.cep_entry = tk.Entry(self.master)
        self.cep_entry.pack()

        # Crie um botão para selecionar a pasta de documentos
        self.selecionar_pasta_button = tk.Button(self.master, text="Selecionar pasta", command=self.selecionar_pasta)
        self.selecionar_pasta_button.pack()

        # Crie um botão para criar os documentos
        self.criar_button = tk.Button(self.master, text="Criar documentos", state="disabled", command=self.criar_documentos)
        self.criar_button.pack()

    def selecionar_pasta(self):
        # Abra uma janela de seleção de diretório e obtenha o caminho selecionado
        diretorio = filedialog.askdirectory()

        # Verifique se um diretório foi selecionado
        if diretorio:
            self.diretorio = diretorio

            # Ative o botão "Criar documentos"
            self.criar_button["state"] = "normal"
        else:
            self.diretorio = None

    def criar_documentos(self):
        # Obtenha as informações digitadas pelo usuário
        novo_nome = self.nome_entry.get()
        novo_endereco = self.endereco_entry.get()
        novo_cep = self.cep_entry.get()

        # Itere através de cada arquivo no diretório selecionado
        for arquivo in os.listdir(self.diretorio):
            if arquivo.endswith(".docx") or arquivo.endswith(".doc"):
                caminho_arquivo = os.path.join(self.diretorio, arquivo)
                doc = Document(caminho_arquivo)

                # Substitua as informações antigas pelas novas em cada parágrafo do documento
                for paragrafo in doc.paragraphs:
                    texto = paragrafo.text
                    novo_texto = texto.replace("[nome]", novo_nome)
                    novo_texto = novo_texto.replace("[endereço]", novo_endereco)
                    novo_texto = novo_texto.replace("[cep]", novo_cep)
                    paragrafo.text = novo_texto

                # Salve o documento atualizado com o mesmo nome do original
                doc.save(caminho_arquivo)

        # Exiba uma mensagem informando que a atualização foi concluída
        tk.messagebox.showinfo("Atualização concluída", "Os documentos foram atualizados com sucesso!")
root = tk.Tk()
app = Aplicativo(root)
root.mainloop()