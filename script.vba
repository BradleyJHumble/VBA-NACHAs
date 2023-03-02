Sub GenerateNachaFile()
    
    ' Set up variables for the NACHA file
    Dim FileHeaderRecord As String
    Dim BatchHeaderRecord As String
    Dim EntryDetailRecord As String
    Dim BatchControlRecord As String
    Dim FileControlRecord As String
    Dim NachaFile As String
    
    ' Define the header information for the NACHA file
    FileHeaderRecord = "101 123456789 1234567892001011431A094101YourBankName                    CompanyName                     "
    BatchHeaderRecord = "5200YOURCOMPANYNAME             1234567892001011431PPD   0000000000000001 123456789YOURBANKNAME                    "
    BatchControlRecord = "8200000001001234567890000000000000000000000000000000000000000000000000000000000000123456789YOURBANKNAME                    "
    FileControlRecord = "9000001000001000000012000000000000000000000000000000000000000000000000000000000000123456789                                            "
    
    ' Set up variables for the data in the Excel worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the data in the worksheet and generate the entry detail records
    Dim i As Long
    For i = 2 To LastRow
        Dim RoutingNumber As String
        RoutingNumber = ws.Cells(i, 1).Value
        Dim AccountNumber As String
        AccountNumber = ws.Cells(i, 2).Value
        Dim Amount As Double
        Amount = ws.Cells(i, 3).Value
        Dim IndividualName As String
        IndividualName = ws.Cells(i, 4).Value
        Dim EntryDetailRecordLine As String
        EntryDetailRecordLine = "625" & RoutingNumber & AccountNumber & Format(Amount, "000000000000.00") & " " & IndividualName & "                                "
        EntryDetailRecord = EntryDetailRecord & EntryDetailRecordLine
    Next i
    
    ' Compile the NACHA file
    NachaFile = FileHeaderRecord & vbCrLf & BatchHeaderRecord & vbCrLf & EntryDetailRecord & BatchControlRecord & vbCrLf & FileControlRecord
    
    ' Save the NACHA file
    Dim FilePath As String
    FilePath = "C:\NACHA\NACHA_FILE.TXT"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(FilePath, True)
    ts.Write NachaFile
    ts.Close
    
    ' Display a message box indicating that the NACHA file has been generated
    MsgBox "The NACHA file has been generated and saved to " & FilePath
    
End Sub
