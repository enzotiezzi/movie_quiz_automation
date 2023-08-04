import os
import xlsxwriter

# Número
# Nome do filme
# Dica 1
# Dica 2
# Dica 3
# Dica 4
# Dica 5
# Genero
# Onde assistir 1
# Onde assistir 2
# Ja utilizado
# Data

path = input("Qual o caminho para a pasta com os filmes?")

workbook = xlsxwriter.Workbook('MovieQuiz.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "Número")
worksheet.write(0, 1, "Nome do filme")
worksheet.write(0, 2, "Dica 1")
worksheet.write(0, 3, "Dica 2")
worksheet.write(0, 4, "Dica 3")
worksheet.write(0, 5, "Dica 4")
worksheet.write(0, 6, "Dica 5")
worksheet.write(0, 7, "Genero")
worksheet.write(0, 8, "Onde assistir 1")
worksheet.write(0, 9, "Onde assistir 2")
worksheet.write(0, 10, "Ja utilizado")
worksheet.write(0, 11, "Data")

row = 1

for folder in os.listdir(path):
    column = 2
    print("Filme atual " + folder)
    
    worksheet.write(row, 0, row)
    worksheet.write(row, 1, folder
                    )
    folderPath = os.path.join(path, folder)

    if(os.path.isdir(folderPath)):
        for file in os.listdir(folderPath):
            print("Arquivo " + file)

            if("movie" not in file and not file.startswith('.')):
                worksheet.write(row, column, file.split('.')[0])
                column +=1
        row+=1

workbook.close()