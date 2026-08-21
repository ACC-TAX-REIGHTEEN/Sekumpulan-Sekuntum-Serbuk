Sub GabungFileExcelKeSheet()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSumber As Workbook
    Dim wbTujuan As Workbook
    Dim namaSheet As String
    
    folderPath = "E:\DataExcel\" 
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbTujuan = ThisWorkbook
    fileName = Dir(folderPath & "*.xls*")
    
    Do While fileName <> ""

        If fileName <> wbTujuan.Name Then
            Set wbSumber = Workbooks.Open(folderPath & fileName)
            
            namaSheet = Left(fileName, InStrRev(fileName, ".") - 1)
            namaSheet = Left(namaSheet, 31) 
            
            wbSumber.Sheets(1).Copy After:=wbTujuan.Sheets(wbTujuan.Sheets.Count)
            wbTujuan.Sheets(wbTujuan.Sheets.Count).Name = namaSheet
            
            wbSumber.Close SaveChanges:=False
        End If
        
        fileName = Dir
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Selesai! Seluruh file berhasil digabungkan ke sheet terpisah.", vbInformation, "Sukses"
End Sub
