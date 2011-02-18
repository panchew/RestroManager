Attribute VB_Name = "mdlPortada"
Option Explicit
Dim strSQL As String

Public Sub frmPortada_ResizeWindow()
  With frmPortada
      .Left = 0
      .Top = 0
      .Width = Screen.Width
      .Height = Screen.Height '- 500
      .fraContainer.Left = (.Width - .fraContainer.Width) / 2
  End With
End Sub

Public Sub frmPortada_CargarDatosEmpresa()
  Dim rsEmpresa As New ADODB.Recordset
  
  With frmPortada
    strSQL = "SELECT * FROM Empresa"
    rsEmpresa.Open strSQL, connADO, adOpenKeyset, adLockOptimistic, adCmdText
    If rsEmpresa.EOF = False Then
      If rsEmpresa!Nombre <> "" Then
        strNombreEmpresa = rsEmpresa!Nombre
        .Caption = .Caption & " - " & rsEmpresa!Nombre
      End If
      If rsEmpresa!Logotipo <> "" Then
        .imgLogotipo.Picture = LoadPicture(App.Path & "/images/" & rsEmpresa!Logotipo)
      End If
      If rsEmpresa!ColorFondo <> "" Then
        .BackColor = rsEmpresa!ColorFondo
        .fraContainer.BackColor = rsEmpresa!ColorFondo
      End If
    Else
        strNombreEmpresa = ""
    End If
    rsEmpresa.Close
    Set rsEmpresa = Nothing
  End With
  
End Sub
