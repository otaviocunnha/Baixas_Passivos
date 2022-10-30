from botcity.core import DesktopBot
import pandas as pd
from botcity.maestro import *
import datetime

planilha = pd.read_excel(r'C:\Users\botcity1.admin\Desktop\Contas_Pagar\baixas_passivo\Vencidos.xlsx')

class Bot(DesktopBot):
    def action(self, execution=None):
       
       
        self.browse(r"C:\Users\botcity1.admin\Desktop\Protheus.lnk")
  

       
        if not self.find( "ok", matching=0.97, waiting_time=90000):
            self.not_found("ok")
        self.click()

       #usuraio e senha
        if not self.find( "usuario", matching=0.97, waiting_time=90000):
            self.not_found("usuario")
        self.click_relative(48, 33)
       ##usuario
        self.paste('')
        self.tab()
       ##senha 
        self.paste('')
        if not self.find( "entra", matching=0.97, waiting_time=10000):
            self.not_found("entra")
        self.click()
        #preparar o ambiente 
        self.wait(15000) 
        self.tab()
        self.tab()
        self.tab()
        self.paste('06')
        ##modulo financeiro
        if not self.find( "entrar_1", matching=0.97, waiting_time=10000):
            self.not_found("entrar_1")
        self.wait(4000)
        self.click()
        #entrar na rotina
        if not self.find( "Favoritos", matching=0.97, waiting_time=900000):
            self.not_found("Favoritos")
        self.click()
        if not self.find( "funcoes", matching=0.97, waiting_time=10000):
            self.not_found("funcoes")
        self.click()
        #data
        self.wait(10000)
        self.paste('31/10/2022')
        self.wait(5000)
        self.tab()
        self.paste('00601000')
        
        if not self.find( "confirmar", matching=0.97, waiting_time=10000):
            self.not_found("confirmar")
        self.click()
        #parametros filtro
        if not self.find( "filtro", matching=0.97, waiting_time=900000):
            self.not_found("filtro")
        self.click()
        if not self.find( "coluna", matching=0.97, waiting_time=90000):
            self.not_found("coluna")
        self.click()
        if not self.find( "scroll", matching=0.97, waiting_time=10000):
            self.not_found("scroll")
        self.triple_click()
        self.wait(5000)
        if not self.find( "fornecedor", matching=0.97, waiting_time=10000):
            self.not_found("fornecedor")
        self.wait(3000)
        self.click()
        self.wait(3000)
        self.space()
        for fornecedor in planilha["Cliente"]:
            if not self.find( "procurar", matching=0.97, waiting_time=10000):
                self.not_found("procurar")
            self.click_relative(84, 9)
            self.paste(f'{fornecedor}'.zfill(6))
            if not self.find( "lupa", matching=0.97, waiting_time=90000):
                self.not_found("lupa")
            self.click()
            while True:
                if not self.find( "Outras_acoes", matching=0.97, waiting_time=10000):
                    self.not_found("Outras_acoes")
                self.click()
                while not self.find( "Baixa_manual", matching=0.97, waiting_time=10000):
                    if not self.find( "Outras_acoes", matching=0.97, waiting_time=10000):
                        self.not_found("Outras_acoes")
                    self.click()   
                if not self.find( "baixa_manual", matching=0.97, waiting_time=10000):
                    self.not_found("baixa_manual")
                    self.wait(2000)
                self.click()
                if not self.find( "baixar", matching=0.97, waiting_time=10000):
                    self.not_found("baixar")
                self.click()
                if self.find( "fechar", matching=0.97, waiting_time=10000):
                    self.click()
                    self.wait(10000)
                    break
                else:
                    if not self.find( "esperar", matching=0.97, waiting_time=10000):
                        self.not_found("esperar")
                    if not self.find( "normal", matching=0.97, waiting_time=30000):
                        self.not_found("normal")                
                    self.wait(3000)
                    self.click()
                    if not self.find( "dacao", matching=0.97, waiting_time=20000):
                        self.not_found("dacao")
                    self.click()
                    self.wait(5000)
                    if not self.find( "salvar", matching=0.97, waiting_time=25000):
                        self.not_found("salvar")
                    self.click()              
                  


    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
