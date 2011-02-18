Attribute VB_Name = "mdlLogin"
Option Explicit
Dim strSQL As String

Public Sub OpenDatabase()
  Dim strRuta As String
  Dim strArchivo As String
  
On Error GoTo ErrError
  'Obtener información de archivo de almacenamiento de datos
  With frmLogin
    .ocxINI.Archivo = App.Path & "\options.ini"
    .ocxINI.Seccion = "DATOS"
    .ocxINI.LLave = "RUTA"
    strRuta = .ocxINI.LeeIni
    .ocxINI.LLave = "ARCHIVO"
    strArchivo = .ocxINI.LeeIni
    strRutaConn = App.Path & strRuta & strArchivo
    'Set strRuta = Nothing
    'Set strArchivo = Nothing
    
Abrir:
    If strRutaConn <> "" Then
      connADO.Provider = "Microsoft.Jet.OLEDB.4.0"
      connADO.Properties("Data Source") = strRutaConn
      connADO.Properties("Jet OLEDB:Database Password") = "restro"
      connADO.Open
      commADO.ActiveConnection = connADO
      GoTo Fin
    Else
      Err.Number = 32755
      GoTo ErrError
    End If
      
ErrError:
    Select Case Err.Number
      Case 0
          GoTo Fin
      Case -2147467259
        'MsgBox "No se ha encontrado el archivo de base de datos; favor de elegir uno", vbCritical + vbOKOnly, "Error"
        .cdgFileOpen.DefaultExt = "mdb"
        .cdgFileOpen.Filter = "Base de datos|*.mdb; Todos los archivos|*.*"
        .cdgFileOpen.ShowOpen
        strRutaConn = .cdgFileOpen.FileName
        'Set strRutaConn = Nothing
        GoTo Abrir
      Case 32755, -2147217843
        MsgBox "No se ha elegido ningún archivo de datos; las opciones no estarán disponibles", vbCritical + vbOKOnly, "Error"
        Exit Sub
      Case Else
        MsgBox "Se ha presentado el siguiente error: " & vbCrLf & "(" & Err.Number & ")" & Err.Description
    End Select
  End With
  
Fin:
      
End Sub

Public Sub VerificarUsuario()
  Dim rsUsuarios As New adodb.Recordset

  With frmLogin
    strSQL = "SELECT * FROM tblUsuarios WHERE Nombre = '" & .txtUsuario.Text & "'"
    rsUsuarios.Open strSQL, connADO, adOpenKeyset, adLockOptimistic, adCmdText
    If rsUsuarios.EOF = False Then
      If rsUsuarios!Password = .txtContraseña.Text Then
        idUsuario = rsUsuarios!idUsuario
        rsUsuarios.Close
        Set rsUsuarios = Nothing
        ObtenerCorteActual
        frmPortada.Show
        Unload frmLogin
      Else
        MsgBox "Contraseña incorrecta"
      End If
    Else
        MsgBox "Usuario no econtrado"
    End If
  End With
  
End Sub

Public Sub ObtenerCorteActual()
  Dim rsCorteCaja As New adodb.Recordset
  
  strSQL = "SELECT * FROM tblCorteCaja WHERE FechaCorte is NULL"
  rsCorteCaja.Open strSQL, connADO, adOpenKeyset, adLockOptimistic, adCmdText
  If rsCorteCaja.EOF = False Then
        strIdCorteActual = rsCorteCaja!IdCorte
  End If
  rsCorteCaja.Close
  Set rsCorteCaja = Nothing
  
End Sub

