Attribute VB_Name = "mdlBuscador"
Option Compare Database
Option Explicit

Public Sub Buscar(NTextBox As TextBox, NListBox As ListBox, NTabla As String, ParamArray NCamposWhere() As Variant)
    On Error GoTo ManipulaError
    
    Dim NCampo As Variant, SQL As String

    For Each NCampo In NCamposWhere
        SQL = SQL & "[" & NCampo & "] Like '*" & Replace(NTextBox.Text, "'", "''") & "*' OR "
    Next NCampo
    
    NListBox.RowSource = "SELECT * FROM [" & NTabla & "] WHERE " & Mid(SQL, 1, Len(SQL) - 3)
    Exit Sub
    
ManipulaError:
    MsgBox Err.Description, vbCritical, "Avíso"
End Sub


