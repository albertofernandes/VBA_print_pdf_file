# VBA_print_pdf_file

Create a excel file. Insert a table and name the Sheet -> "database"
Create another Sheet, and design the page you want to print. Name the sheet -> "design_page"

Then, you will have 2 types of variables:
1) Variables which you will use to write on your "design_page" and give your printed file a filename
  For each variable, create a Dim [write_variable] As Long
2) File location variables
  For your output folder location and sheet name
  Name the variable path_pdf with your output file location
  Name the variable printing_sheet with your "design_page" sheet name
  Name the variable database_sheet with your "database" sheet name
    

Sub Print_PDF()

'Create variables which you will use to write on your "design_page" and give your printed file a filename
Dim example_1 As Long
Dim example_2 As Long
Dim example_3 As Long

'Create variables for output folder location, and sheet names
Dim path_pdf As String
Dim database_sheet As String
Dim printing_sheet As String

Dim cells_number As String

'Change de column number for each variable you will use from your database
example_1 = 5 'column 5
example_2 = 22 'column 22
example_3 = 23 'column 23

'Database sheet name
database_sheet = "Database Sheet Name"

'Printing Desing Sheet Name
printing_sheet = "Printing Design Sheet name"

'Output path folder
path_pdf = "C:\Users\xpto\"


'Number of rows on your database, used to fix the number of pdf created, by counting number of full cells on column x. 
'Change Range("x:x"), where x is you column
cells_number = Worksheets(sep_base_dados).Range("x:x").Cells.SpecialCells(xlCellTypeConstants).Count

For i = 2 To cells_number
  
  'For each cell you want to fill, change 
  Worksheets(printing_sheet).Cells(7, 2) = "João Manuel de Barros Figueiredo Cruz, Autoridade de Saúde de Braga do Aces Cávado 1 – Braga, nos termos do artigo 5.º do Decreto-Lei n.º 82/2009, de 2 de abril, alterado pelo Decreto-Lei n.º 135/2013, de 4 de outubro, determino o isolamento profilático de " & Worksheets(sep_base_dados).Cells(i, coluna_nome) & ", portador do BI / CC n.º " & Worksheets(sep_base_dados).Cells(i, coluna_cc) & ", com validade até " & Worksheets(sep_base_dados).Cells(i, coluna_validade_cc) & ", com o número de identificação de segurança social " & Worksheets(sep_base_dados).Cells(i, coluna_niss) & ", pelo período de " & Worksheets(sep_base_dados).Cells(i, coluna_inicio_isolamento) & " a " & Worksheets(sep_base_dados).Cells(i, coluna_fim_isolamento) & ", por motivo de perigo de contágio e como medida de contenção de COVID-19.----------------------------------"
        Worksheets(sep_folha_impressao).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        destino_pdf & Worksheets(sep_base_dados).Cells(i, 24) & "@@" & Worksheets(sep_base_dados).Cells(i, 20) & "@@" & Worksheets(sep_base_dados).Cells(i, 13) & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        From:=1, To:=1, OpenAfterPublish:=False
    
Next i

Debug.Print ("Fim")

End Sub
