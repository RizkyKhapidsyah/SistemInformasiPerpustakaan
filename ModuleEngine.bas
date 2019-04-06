Attribute VB_Name = "Module1"
Option Explicit

Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public Kalimat As String
Global Objek As Control
Public X As Integer

Public Sub NyambunggUtama()
If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PerpustakaanDwi"
End Sub

Public Sub PusatError()
    MsgBox "Maaf, ada Kesalahan internal program, silahkan re-start program ini." & vbCrLf & vbCrLf & _
           "Error Code : " & Err.Number & vbCrLf & _
           "Deskripsi  : " & Err.Description, vbCritical + vbOKOnly, "MainSystem : Error"
End Sub
