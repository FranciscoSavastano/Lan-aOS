from tkinter import *
from tkinter import ttk
from tkcalendar import DateEntry

root = Tk()

class app(): 
    def __init__(self):
        self.root = root
        self.window()
        self.subwindow()
        self.botoes()
        root.mainloop()        
        
    
    def window(self): #cria a janela principal
        self.root.title('Lança OS')
        self.root.configure(background='#ffffff')
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
    def subwindow(self): #cria as subjanelas para os widgets, e os insere
        self.frame_1 = Frame()
        self.frame_1 = Label(
            bg='#cac3ba', bd=2, highlightbackground='#000000', highlightthickness=1)
        self.frame_1.place(relx=0.04, rely=0.05, relwidth=0.9, relheight=0.9)
        self.frame_botoes = Frame()
        self.frame_botoes.place(relx=0.04, rely= 0.05, relwidth=0.9, relheight=0.9)
        entrytext = Label(
            text='Codigo do parceiro: \n\n\nNumero da Nota: \n\n\nData de emissão: \n\n\nNumero da OS: \n\n\nTipo de operação: \n\n\n\n\nProblema:\n\n\n\n\nNumero da NS:\n\n\n\n Consultor(a):', font=('verdana', 11))
        entrytext.place(relx=0.06, rely=0.134)
    def botoes(self):
        global entryCodPar, entryOs, entryNF, entryEmissao, entryNS, entryConsul, entryTroca, entryProblema
        entryCodPar = Entry(self.frame_botoes)
        entryCodPar.place(relx= 0.4, rely= 0.1)
        entryNF = Entry(self.frame_botoes)
        entryNF.place(relx=0.4, rely = 0.185)
        entryEmissao = DateEntry(self.frame_botoes, locale = "pt_BR")
        entryEmissao.place(relx= 0.4, rely= 0.27)        
        entryOs = Entry(self.frame_botoes)
        entryOs.place(relx= 0.4, rely= 0.36)
        entryTroca = ttk.Combobox(        state="readonly",
        values=["Troca defeito", "Em Teste", "Compra Errada", "Outros"]
        )
        entryTroca.place(relx= 0.4, rely= 0.45)
        entryProblema = Text(self.frame_botoes, height= 4, width= 35)
        entryProblema.place(relx= 0.4, rely= 0.5)
        entryNS = Text(self.frame_botoes, height= 4, width= 35)
        entryNS.place(relx= 0.4, rely= 0.65)
        entryConsul = Entry(self.frame_botoes)
        entryConsul.place(relx= 0.4, rely= 0.84)
        enviar = Button(text = ("Enviar"), command= self.Oswin)
        enviar.place(relx= 0.4, rely= 0.9)
    def Oswin(self):
        CodPar = entryCodPar.get()
        Nf = entryNF.get()
        Emissao = entryEmissao.get()
        Os = entryOs.get()
        Troca = entryTroca.get()
        Problema = entryProblema.get("1.0", END)
        Ns = entryNS.get("1.0", END)
        win = Toplevel()
        win.geometry('400x300')
        win.resizable(False, False)
        res_1 = Frame(win)
        res_1.place(relx=0, rely=0, relheight=1, relwidth=1)
        lab_1 = Label(res_1, bg='#d3d3d3',
                      text = f"Codigo do parceiro : {CodPar}\n Numero da Nota: {Nf}\n Emissão da Nota: {Emissao}\n Numero da OS: {Os} \n Tipo de operação: {Troca}\n Problema: {Problema} \n NS do produto: {Ns}")
        lab_1.place(relheight=1, relwidth=1)
        writetxt = Button(res_1, text = ("Salvar como texto"), command= self.Oswin)
        writetxt.place(relx= 0.2, rely= 0.9)
        writeexcel = Button(res_1,text = ("Salvar em excel"), command= self.Oswin)
        writeexcel.place(relx= 0.5, rely= 0.9)

app()