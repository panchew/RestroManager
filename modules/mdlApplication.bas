Attribute VB_Name = "mdlApplication"
Option Explicit

'DB connection objects
Public connADO As New ADODB.Connection
Public commADO As New ADODB.Command
'DB path folder
Public strRutaConn As String
'Usuario en sesión
Public strIdUsuario As String
'Corte de caja actual
Public strIdCorteActual As String

