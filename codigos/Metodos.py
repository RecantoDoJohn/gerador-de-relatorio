import openpyxl
import Planilha
import meses

class operacoes:
    def __init__(self, nome_planilha: str):
        self.planilha = openpyxl.load_workbook(nome_planilha)
        self.nomeda_planilha = nome_planilha
        self.tabela = self.planilha.active

    def organizar(self, metodo):
        lista = []
        resultado_lista = []
        for linha in self.tabela.iter_rows(min_row=2):
            for i in range(1, 5):
                lista.append(linha[i].value)
            
            match metodo:
                case 0:
                    resultado_lista.append(self.media(lista))
                case 1:
                    resultado_lista.append(self.mediana(lista))
                case 2:
                    resultado_lista.append(f'{self.variancia(lista):.2f}')
                case 3:
                    resultado_lista.append(f'{self.desvio_padrao(lista):.2f}')
                case _:
                    return "ainda n tem"
        Planilha.atualizar_planilha( self.nomeda_planilha, meses.metodos[metodo], resultado_lista)
        return f'Adcionado {meses.metodos[metodo]} a planilha'
    
    def media(self, lista: list):
        tamanho = len(lista)
        total = 0
        for item in lista:
            total += item
        return (total//tamanho)
    
    def mediana(self, lista: list):
        tamanho = len(lista)
        lista_row = sorted(lista)
        if tamanho % 2 == 0:
            return (lista_row[(tamanho // 2) - 1] + lista_row[tamanho // 2]) / 2
        else:
            return lista_row[tamanho // 2]
        
    def variancia(self, lista):
        media = sum(lista) / len(lista)
        soma_diferencas = sum((x - media) ** 2 for x in lista)
        variancia = soma_diferencas / len(lista)
        
        return variancia
    
    def desvio_padrao(self, lista):
        vari = self.variancia(lista)
        return vari ** (1/2)

        

if __name__ == '__main__':
   operacoes('possivel_planilha.xlsx').organizar(0)