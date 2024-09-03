import openpyxl

def atualizar_planilha(planilha, metodo, lista):
    wb = openpyxl.load_workbook(planilha)

    pagina = wb.active

    for letra in 'ghijklmnopq':
        if pagina[f'{letra}1'].value == None:
            pagina[f'{letra}1'] = metodo

            i = 2
            for item in lista:
                pagina[f"{letra}{i}"] = item
                i += 1
            break
    wb.save(planilha)