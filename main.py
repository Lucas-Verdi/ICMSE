from tkinter  import *
import lxml.etree as ET
from tkinter.filedialog import askopenfilenames
import getpass
import xlwings as xl
from threading import Thread


user = getpass.getuser()
arquivos = []
results = []
numeronota = []
ipinota = []
dataemi = []
razaosocial = []
valornota = []

def start():
    a = Th(1)
    a.start()


def ler_icms(xml_file):
    print("Analisando arquivo XML {}".format(xml_file))

    tree = ET.parse(xml_file)
    root = tree.getroot()

    namespaces = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

    final = root.xpath(".//nfe:ICMSTot/nfe:vICMS/text()", namespaces=namespaces)
    numero = root.xpath(".//nfe:nNF/text()", namespaces=namespaces)
    ipi = root.xpath(".//nfe:ICMSTot/nfe:vIPI/text()", namespaces=namespaces)
    data = root.xpath(".//nfe:dhEmi/text()", namespaces=namespaces)
    razao = root.xpath(".//nfe:emit/nfe:xNome/text()", namespaces=namespaces)
    valortot = root.xpath(".//nfe:ICMSTot/nfe:vNF/text()", namespaces=namespaces)



    numeronota.append(numero)
    ipinota.append(ipi)
    dataemi.append(data)
    razaosocial.append(razao)
    valornota.append(valortot)

    return results.append(final)


class Th(Thread):

    def __init__(self, num):
        Thread.__init__(self)
        self.num = num

    def run(self):
        global arquivos


        arqtemp = askopenfilenames(multiple=True)
        arquivos.extend(arqtemp)


        for i in range(0, len(arquivos)):
            ler_icms(arquivos[i])

        app = xl.App()
        workbook = app.books.add()
        sheet = workbook.sheets.active
        sheet.range('A1').value = "DATA"
        sheet.range('B1').value = "NOTA"
        sheet.range('C1').value = "RAZAO SOCIAL"
        sheet.range('D1').value = "ICMS"
        sheet.range('E1').value = "IPI"
        sheet.range('F1').value = "TOTAL NF"

        sheet.range('A2').value = dataemi
        sheet.range('B2').value = numeronota
        sheet.range('C2').value = razaosocial
        sheet.range('D2').value = results
        sheet.range('E2').value = ipinota
        sheet.range('F2').value = valornota

        arquivos.clear()
        results.clear()
        numeronota.clear()
        ipinota.clear()
        dataemi.clear()
        razaosocial.clear()
        valornota.clear()


#INTERFACE
janela = Tk()
janela.title("ICMSE")
janela.geometry("230x100")
label1 = Label(janela, text="Olá {}!".format(user), font="Arial 10 bold", justify=CENTER)
label1.grid(column=0, row=0, padx=2, pady=2)
bt1 = Button(janela, text="INSERIR XML", font="Arial 10 bold", justify=CENTER)
bt1.grid(column=0, row=1, padx=70, pady=30)
bt1.bind("<Button>", lambda e: start())
janela.mainloop()
