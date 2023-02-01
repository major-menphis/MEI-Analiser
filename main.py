import pandas as pd
import os
from time import sleep
import PyPDF2


# diretory define
def diretory_define():
    path = "./"
    os.chdir(path)


# read txt in pdf
def read_pdf_file():
    diretory_define()
    list_data = []
    for file in os.listdir():
        if file.endswith('.pdf'):
            data_dict = dict()
            empresa = ''
            cnpj = ''
            print(f'ANALISANDO: {file}')
            pdf_file = open(f'{file}', 'rb')
            read_pdf = PyPDF2.PdfReader(pdf_file)
            # number_of_pages = len(read_pdf.pages)
            page = read_pdf.pages[0]
            page_content = page.extract_text()
            parsed = ''.join(page_content)
            lines = parsed.splitlines()
            for index, value in enumerate(lines):
                if value.strip() == 'Nome Empresarial':
                    empresa = lines[index+1]
                elif value.strip() == 'CNPJ':
                    cnpj = lines[index+1]
                elif value.strip() == 'CNPJ Data de Abertura':
                    cnpj = lines[index+1][:18]
                data_dict['Empresa'] = empresa
                data_dict['CNPJ'] = cnpj
            list_data.append(data_dict)
    if list_data != '':
        return list_data
    else:
        return False


# saving report
def save_report():
    print('Processando...')
    direc = os.getcwd()
    pasta = direc.split("\\")[-1]
    diretory_define()
    list_data = read_pdf_file()
    if list_data != []:
        df = pd.DataFrame(list_data)
        writer = pd.ExcelWriter(f'{pasta}.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name=f'{pasta}', header=True, index=False)
        workbook = writer.book
        worksheet = writer.sheets[pasta]
        # add custom formats
        fmt_centered = workbook.add_format({
            'valign': 'vcenter',
            'align': 'center'
        })
        fmt_text = workbook.add_format({
            'text_wrap': True
        })
        fmt_header = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#5DADE2',
            'font_color': '#FFFFFF',
            'border': 1
        })
        # set zoom
        worksheet.set_zoom(100)
        # apply custom formats to the cells
        worksheet.set_column("A:A", 50, fmt_centered)
        worksheet.set_column("B:B", 20, fmt_text)
        # apply format header
        for col, value in enumerate(df.columns.values):
            worksheet.write(0, col, value, fmt_header)
        # save
        writer.close()
        print('Arquivo gerado com sucesso!')
        sleep(2)
    else:
        print('Não há certificados do MEI para gerar o relatório.')
        sleep(3)
        return


if __name__ == "__main__":
    save_report()
