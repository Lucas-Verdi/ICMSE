from tkinter  import *
import lxml.etree as ET
from tkinter.filedialog import askopenfilenames
import getpass
import xlwings as xl


user = getpass.getuser()
arquivos = []
results = []
numeronota = []
ipinota = []


def ler_icms(xml_file):
    print("Analisando arquivo XML {}".format(xml_file))

    tree = ET.parse(xml_file)
    root = tree.getroot()

    namespaces = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

    final = root.xpath(".//nfe:ICMSTot/nfe:vICMS/text()", namespaces=namespaces)
    numero = root.xpath(".//nfe:nNF/text()", namespaces=namespaces)
    ipi = root.xpath(".//nfe:ICMSTot/nfe:vIPI/text()", namespaces=namespaces)

    numeronota.append(numero)
    ipinota.append(ipi)

    return results.append(final)


def main():
    global arquivos


    arqtemp = askopenfilenames(multiple=True)
    arquivos.extend(arqtemp)


    for i in range(0, len(arquivos)):
        ler_icms(arquivos[i])

    app = xl.App()
    workbook = app.books.add()
    sheet = workbook.sheets.active
    sheet.range('A1').value = "NOTA"
    sheet.range('B1').value = "ICMS"
    sheet.range('C1').value = "IPI"

    sheet.range('A2').value = numeronota
    sheet.range('B2').value = results
    sheet.range('C2').value = ipinota


#INTERFACE
janela = Tk()
janela.title("ICMSE")
janela.geometry("230x100")
label1 = Label(janela, text="Olá {}!".format(user), font="Arial 10 bold", justify=CENTER)
label1.grid(column=0, row=0, padx=2, pady=2)
bt1 = Button(janela, text="INSERIR XML", font="Arial 10 bold", justify=CENTER)
bt1.grid(column=0, row=1, padx=70, pady=30)
bt1.bind("<Button>", lambda e: main())
janela.mainloop()
