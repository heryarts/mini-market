Attribute VB_Name = "ModulUtama"
Public DbCon As New ADODB.Connection
Public RsFind As New ADODB.Recordset
Public SQL, ConDb, Periode As String
Public Trans         As New Convert


'---------------------------------------------------------------------------------------
' Procedure : bukaDatabase
' DateTime  : 11/21/2012 09:43
' Author    : Admin
' Purpose   : fungsi untuk membuka koneksi ke database
'---------------------------------------------------------------------------------------

Public Sub bukaDatabase()
   On Error GoTo bukaDatabase_Error
App.Title = "Aplikasi Minimarket"
With DbCon
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=matahari;Initial Catalog=minimarket;server=."
    adocon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & pass & ";Initial Catalog=[" & DbName & "];server=" & GetSetting(App.Title, "startup", "server", "(local)")
    ConDb = .ConnectionString
    .Open
End With

   On Error GoTo 0
   Exit Sub

bukaDatabase_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bukaDatabase of Module ModuleUtama"

End Sub

Function FormatTgl(ddate As Date) As String
    FormatTgl = Format(ddate, "mm/dd/yyyy")
End Function

Public Sub Enter(ByVal Key As Integer, Optional ByRef XX As Object)
     If Key = 13 Then SendKeys "{tab}"
     If Key = 38 Then
         If XX Is Nothing Then
            SendKeys "+{tab}"
            Exit Sub
         End If
     ElseIf Key = 40 Then
         If XX Is Nothing Then
            SendKeys "{tab}"
            Exit Sub
         End If
     End If
End Sub
