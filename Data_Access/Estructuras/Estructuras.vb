Public Class Estructuras
    Public Structure strRespuesta
        Public codigo As String
        Public mensaje As String
        Public url As String
        Public color As String
        Public objeto As Object
    End Structure

    Public Structure strLogin
        Public accion As Integer
        Public nombreUsuario As String
        Public pwd As String
        Public accesoHash As String
        Public arrConsultas() As TipoConsultas
    End Structure

    Public Structure TipoConsultas
        Public idConsulta As Integer
        Public nombreConsulta As String
    End Structure

    Public Structure strCliente
        Public idCliente As Integer
        Public nombreCliente As String
        Public telefono As String
        Public direccion As String
        Public email As String
        Public codUsuario As String
        Public pwd As String
        Public idUsrMod As Integer
        Public esAdministrador As Boolean
        Public idTipoConsulta As Integer
        Public activo As Boolean
    End Structure

    Public Structure strRecuperacion
        Public numCel As String
        Public pwd As String
    End Structure
    Public Structure strSelNombresCorp
        Public idCliente As Integer
        Public nombreCliente As String
    End Structure

    Public Structure strRolesUsuarios
        Public esAdmin As Boolean
        Public bloqueado As Boolean
        Public activo As Boolean
        Public idUsr As Integer
        Public idUsrMod As Integer
        Public fallidos As Integer
        Public accion As Integer
        Public idEmpresa As Integer
    End Structure

    Public Structure strOTP
        Public codUsr As String
        Public otp As String
        Public celular As String
        Public pwd As String
    End Structure

    Public Structure strEmpresas
        Public accion As Integer
        Public nombre As String
        Public activo As Boolean
        Public idEmpresa As Integer
        Public idUsrMod As Integer
    End Structure

End Class
