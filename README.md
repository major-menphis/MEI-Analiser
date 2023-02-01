# Gerador de relatórios de nomes de empresas e cnpj's vinculados

Segue a mesma idéia do projeto Gerador de Relatórios para Excel no link https://github.com/major-menphis/gerador-relatorios-excel.
Os diferenciais estão na parte do arquivo que é lido, neste projeto ele lê apenas arquivos em pdf encontrados com a extensão .pdf
Analisa o nome das empresas de acordo com o Certificado do MEI emitido pela RFB e retorna uma planilha com nome da empresa e cnpj.
O projeto foi criado apenas pra reduzir a dificuldade de um colega em criar esse relatório em grande escala, podendo ser acrescido de mais funcionalidades e retornar mais informações caso nescessário.

Obs: Código simples, sem o devido tratamento de erros.

Comando para gerar o executável: `pyinstaller --onefile --name MEI-Analiser --icon=icon.ico main.py`