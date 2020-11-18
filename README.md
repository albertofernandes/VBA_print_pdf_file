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
  
  'For each cell you want to fill, change Cells(x, y), where x is the row of the cell and y its column,
  ' and add the text you want on that cell, and change Worksheets(sep_base_dados).Cells(i, example_1), where example_1 os the column where the data is
  Worksheets(printing_sheet).Cells(x, y) = "This is a test" & Worksheets(database_sheet).Cells(i, example_1) & "Ending Test"
  
  'Then, a pdf will be printed, to path pdf. You can name it with text and the cells at columns you chose.
  Worksheets(printing_sheet).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
  path_pdf & Worksheets(sep_base_dadosdatabase_sheet).Cells(i, example_2) & "@@" & Worksheets(sep_base_dados).Cells(i, example_3) & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        From:=1, To:=1, OpenAfterPublish:=False
    
Next i

'At the end, you will see a message End
Debug.Print ("End")

End Sub
