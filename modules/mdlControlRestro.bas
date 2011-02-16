Attribute VB_Name = "mdlControlRestro"
Option Explicit

Public strRuta As String
Public strArchivo As String
Public strIdUsuario As String

'Public strBaseConn As String
Public strRutaConn As String

Public connADO As New ADODB.Connection
Public commADO As New ADODB.Command

Public strSQL As String
Public strEmpresa As String

Public strZona As String
Public strMesa As String

Public intIndex As Integer ' índices de mesas
Public strIdMesa As String

Public strComanda As String
Public sngDescuento As Single

Public blnImpreso As Boolean
Public varObservaciones As Variant

Public strIdPedido As String
Public strIdCorteActual As String
Public blnBrowse As Boolean

Public strPosX As String
Public strPosY As String

Public strAgregarEditarCliente As String
Public strFormCliente As String

Public Sub TransferirDatos()

      ' Transferir catInsumos
      
      ' Transferir catPermisos
      
      ' Transferir Clientes
      
      ' Transferir Cobros
      
      ' Transferir corte caja
      
      ' Transferir facturas
      
      ' Transferir observaciones
      
      ' Transferir pedidos
      
      ' Transferir perfiles
      
      ' Transferir permisos
      
      ' Transferir productos
      
      ' Transferir propinas
      
      ' Transferir recetas
      
      ' Transferir usuarios
      
      ' Transferir zonas
      
End Sub

