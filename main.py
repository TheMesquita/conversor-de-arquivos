from tkinter import Tk, Button, Label
from tkinter.filedialog import askopenfilename, askdirectory
from pdf2docx import Converter

#estrutura principal do código para conversão (orientado á objetos)
class PDFtoWordConverter:
    def __init__(self):
        self.pdf_file_path = None
        self.docx_file_path = None

    def select_pdf_file(self):
        self.pdf_file_path = askopenfilename(filetypes=[("Arquivo PDF", "*.pdf")])
        if self.pdf_file_path:
            self.pdf_label.config(text=self.pdf_file_path)
        else:
            self.pdf_label.config(text="Nenhum arquivo PDF selecionado.")

    def select_destination_directory(self):
        self.docx_file_path = askdirectory()
        if self.docx_file_path:
            self.docx_label.config(text=self.docx_file_path)
        else:
            self.docx_label.config(text="Nenhum local de destino selecionado.")

    def convert(self):
        if self.pdf_file_path and self.docx_file_path:
            cv = Converter(self.pdf_file_path)
            cv.convert(f"{self.docx_file_path}/arquivo.docx") #definido o nome padrão do arquivo convertido como 'arquivo.docx' 
            cv.close()
            self.status_label.config(text="Conversão concluída com sucesso!")
        else:
            self.status_label.config(text="Selecione o arquivo PDF e o local de destino.")

#criação da parte gráfica, adcionando tela e botoões para a seleçãp do arquivo desejado
    def tela(self):
        root = Tk()
        root.title("PDFtoWORD")

        pdf_button = Button(root, text="Selecionar PDF", font = "arial 12 bold", padx = 15, pady = 15, border = 5, command=self.select_pdf_file)
        pdf_button.pack()

        self.pdf_label = Label(root, text="Nenhum arquivo PDF selecionado.", font = "arial 12", padx = 15, pady = 15)
        self.pdf_label.pack()

        docx_button = Button(root, text="Selecione o local de destino", font = "arial 12 bold", padx = 15, pady = 15, border = 5, command=self.select_destination_directory)
        docx_button.pack()

        self.docx_label = Label(root, text="Nenhum local de destino selecionado.", font = "arial 12", padx = 15, pady = 15)
        self.docx_label.pack()

        convert_button = Button(root, text="Converter", font = "arial 12 bold", padx = 15, pady = 15, border = 5, command=self.convert)
        convert_button.pack()

        self.status_label = Label(root, text="")
        self.status_label.pack()

        root.mainloop()


converter = PDFtoWordConverter()
converter.tela()

#Leonardo Mesquita - 2023 - ADS


