Imports System.Configuration
Imports Data_Access.Estructuras
Imports MySql.Data.MySqlClient

Public Class clData
    Dim objFuncionalidadess As New RVLFuncionesUtilidades.FuncionesUtilidades
    Public _conexLocal = objFuncionalidadess.DesencriptarAES(ConfigurationManager.AppSettings("_cadenaConexion"))
    Public _maximoIntentos = ConfigurationManager.AppSettings("_maximoIntentos")

    Public Function Login(ByRef entrada As strLogin) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)
        Dim objFunciones As New RVLFuncionesUtilidades.FuncionesUtilidades



        Try
            Dim comm As New MySqlCommand("portal_spLogin", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()

            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("accion", entrada.accion)
                comm.Parameters.AddWithValue("usrName", entrada.nombreUsuario)
                comm.Parameters.AddWithValue("_hash", entrada.accesoHash)
                comm.Parameters.AddWithValue("codSalida", MySqlDbType.Int32).Direction = ParameterDirection.Output

                da.SelectCommand = comm
                da.SelectCommand.Connection = con
                da.Fill(dtResult)
                con.Close()

                If dtResult.Rows.Count > 0 Then
                    Dim pwdBd As String = dtResult.Rows(0).Item("pwd").ToString()
                    Dim pwdCl As String = entrada.pwd
                    Dim auxPwd = objFunciones.DesencriptarAES(pwdBd).objeto
                    Dim bloqueado As Boolean = dtResult(0).Item("bloqueado").ToString()

                    If Not bloqueado Then
                        If (auxPwd = pwdCl) Or (entrada.accion = 4) Then
                            'entrada.accion = 2
                            'Dim _actualizacionHash = ActualizarHash(entrada)

                            retorno.codigo = 0
                            retorno.mensaje = "Login correcto"

                            Select Case entrada.accion
                                Case 1
                                    ReDim entrada.arrConsultas(dtResult.Rows.Count - 1)

                                    retorno.url = dtResult(0).Item("url").ToString()
                                    retorno.objeto = dtResult(0).Item("idUsuario").ToString() & "|" & dtResult(0).Item("reqCambioPwd").ToString() & "|" &
                                                     dtResult(0).Item("nombreCliente").ToString() & "|" & dtResult(0).Item("esAdmin").ToString()

                                    Dim reset = ResetFallidos(dtResult(0).Item("idUsuario").ToString())

                                    'dr("Material") = nombre columna dt

                                    Dim x As Integer = 0
                                    For Each dr As DataRow In dtResult.Rows
                                        entrada.arrConsultas(x).idConsulta = dr("idTipoConsulta")
                                        entrada.arrConsultas(x).nombreConsulta = dr("nombreTipoConsulta")

                                        x = x + 1

                                        retorno.objeto = retorno.objeto & "|"
                                    Next


                                Case 4
                                    dtResult.Rows(0).Item("pwd") = auxPwd
                                    retorno.objeto = objFuncionalidadess.DtToJSON(dtResult)
                            End Select

                        Else
                            Dim fallidos As New strRolesUsuarios

                            fallidos.accion = 1
                            fallidos.fallidos = dtResult(0).Item("intentosFallidos").ToString() + 1
                            fallidos.bloqueado = If(fallidos.fallidos = _maximoIntentos, True, False)
                            fallidos.idUsr = dtResult(0).Item("idUsuario").ToString()
                            fallidos.idUsrMod = dtResult(0).Item("idUsuario").ToString()

                            RegistrarIntentosFallidos(fallidos)

                            retorno.codigo = 1
                            retorno.mensaje = "Usuario o contraseña incorrectos"
                        End If
                    Else
                        retorno.codigo = 2
                        retorno.mensaje = "Cuenta bloqueada. Póngase en contacto con el Administrador para el desbloqueo, o proceda al cambio de esta en la opción ""Olvidó su contraseña"""
                    End If

                Else
                    retorno.codigo = 2
                    retorno.mensaje = "Usuario o contraseña incorrectos"
                End If

            Else
                retorno.codigo = 10
                retorno.mensaje = "Ocurrió un error al abrir la conexión para el sistema de logueo"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 11
            retorno.mensaje = "Error crítico en el sistema de logueo: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function ResetFallidos(ByVal _idUsuario As String) As strRespuesta
        Dim Query As String = "update segurinet.portal_usuario set intentosFallidos = 0, idUsrMod = " & _idUsuario & ", fechaMod = now() where idUsuario = " & _idUsuario & ";"

        Dim esValido As Boolean = False
        Dim retorno As New strRespuesta
        Dim recuperacion As New strRecuperacion

        Using Conn As New MySqlConnection(_conexLocal.objeto)
            Using Comm As New MySqlCommand()
                With Comm
                    .Connection = Conn
                    .CommandText = Query
                    .CommandType = CommandType.Text
                End With
                Try
                    Conn.Open()
                    Dim Reader As MySqlDataReader = Comm.ExecuteReader
                    Conn.Close()

                    retorno.codigo = 0
                    retorno.mensaje = "OK"

                Catch ex As Exception
                    Conn.Close()
                    retorno.codigo = 18
                    retorno.mensaje = "Error crítico al resetear fallidos: " & ex.Message
                End Try
            End Using
        End Using

        Return retorno

    End Function

    Public Function ActualizarHash(ByVal entrada As strLogin) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spLogin", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()

            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("accion", entrada.accion)
                comm.Parameters.AddWithValue("usrName", entrada.nombreUsuario)
                comm.Parameters.AddWithValue("_hash", entrada.accesoHash)
                comm.Parameters.AddWithValue("codSalida", MySqlDbType.Int32).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("codSalida").Value.ToString()
                retorno.mensaje = If(retorno.codigo = 0, "Actualización de hash exitosa", "Error al actualizar hash para el usuario")
            Else
                retorno.codigo = 11
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la actualización del hash"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 12
            retorno.mensaje = "Error crítico en la actualización del hash del usuario: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function ValidarHash(ByVal entrada As strLogin) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spLogin", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()

            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("accion", entrada.accion)
                comm.Parameters.AddWithValue("usrName", entrada.nombreUsuario)
                comm.Parameters.AddWithValue("_hash", entrada.accesoHash)
                comm.Parameters.AddWithValue("codSalida", MySqlDbType.Int32).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("codSalida").Value.ToString()
                retorno.mensaje = If(retorno.codigo = 0, "Hash válido", "Hash Incorrecto")
            Else
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la validación del Hash"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 13
            retorno.mensaje = "Error crítico en la validación del Hash: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function RegistrarCliente(ByVal entrada As strCliente) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spRegistrarCliente", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable
            Dim auxPwd = objFuncionalidadess.EncriptarAES(entrada.pwd)

            If auxPwd.codigo = 0 Then
                con.Open()

                If con.State = ConnectionState.Open Then

                    comm.CommandType = CommandType.StoredProcedure
                    comm.Parameters.AddWithValue("_idEmpresa", entrada.idCliente)
                    comm.Parameters.AddWithValue("_nombreCliente", entrada.nombreCliente)
                    comm.Parameters.AddWithValue("_telefono", entrada.telefono)
                    comm.Parameters.AddWithValue("_email", entrada.email)
                    comm.Parameters.AddWithValue("_codUsuario", entrada.codUsuario)
                    comm.Parameters.AddWithValue("_pwd", auxPwd.objeto)
                    comm.Parameters.AddWithValue("_idUsrMod", entrada.idUsrMod)
                    comm.Parameters.AddWithValue("_esAdmin", entrada.esAdministrador)
                    comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                    comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output

                    comm.ExecuteNonQuery()
                    con.Close()

                    retorno.codigo = comm.Parameters("_cod").Value.ToString()
                    retorno.mensaje = comm.Parameters("_msj").Value.ToString()
                Else
                    retorno.codigo = 14
                    retorno.mensaje = "Ocurrió un error al abrir la conexión para el registro del cliente"
                End If
            Else
                retorno.codigo = 15
                retorno.mensaje = "Ocurrió un error al encritar la contraseña: " & auxPwd.errCatch
            End If


        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 16
            retorno.mensaje = "Error crítico en el registro del cliente: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function ObtenerNumCel(ByVal codUsuario As String) As strRespuesta
        Dim Query As String = "select cli.telefono,usu.pwd 
                                from segurinet.portal_cliente cli 
                                inner join segurinet.portal_usuario usu on usu.idCliente = cli.idCliente and usu.activo = 1 
                                where usu.codUsuario = '" & codUsuario & "' 
                                and cli.activo = 1 limit 1;"

        Dim esValido As Boolean = False
        Dim retorno As New strRespuesta
        Dim recuperacion As New strRecuperacion

        Using Conn As New MySqlConnection(_conexLocal.objeto)
            Using Comm As New MySqlCommand()
                With Comm
                    .Connection = Conn
                    .CommandText = Query
                    .CommandType = CommandType.Text
                    '.Parameters.AddWithValue(Operator1, Value1)
                End With
                Try
                    Conn.Open()
                    Dim Reader As MySqlDataReader = Comm.ExecuteReader
                    While Reader.Read() OrElse (Reader.NextResult And Reader.Read)
                        recuperacion.numCel = Reader.GetString(0)
                        recuperacion.pwd = objFuncionalidadess.DesencriptarAES(Reader.GetString(1)).objeto
                        esValido = True
                    End While

                    Conn.Close()

                    If esValido Then
                        retorno.codigo = 0
                        retorno.mensaje = "Consulta exitosa"
                        retorno.objeto = recuperacion
                    Else
                        retorno.codigo = 17
                        retorno.mensaje = codUsuario & " no produjo resultdos"
                    End If

                Catch ex As Exception
                    Conn.Close()
                    retorno.codigo = 18
                    retorno.mensaje = "Error al correr query de restauración de contraseña: " & ex.Message
                End Try
            End Using
        End Using

        Return retorno

    End Function

    Public Function GenerarOTP(ByVal _codUsr As String) As strRespuesta
        Dim retorno As New strRespuesta
        Dim otp As New strOTP
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spGenerarOTP", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()

            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_codUsr", _codUsr)
                comm.Parameters.AddWithValue("_otp", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_numCel", String.Empty).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("_cod").Value.ToString()
                retorno.mensaje = comm.Parameters("_msj").Value.ToString()

                If retorno.codigo = 0 Then
                    otp.otp = comm.Parameters("_otp").Value.ToString()
                    otp.celular = comm.Parameters("_numCel").Value.ToString()
                    retorno.objeto = otp
                End If

            Else
                retorno.codigo = 14
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la generación de OTP"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 16
            retorno.mensaje = "Error crítico al generar OTP: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function ValidarOTP(ByVal entrada As strOTP) As strRespuesta
        Dim retorno As New strRespuesta
        Dim otp As New strOTP
        Dim pwdUsr As String = String.Empty
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spValidarOTP", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()

            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_codUsr", entrada.codUsr)
                comm.Parameters.AddWithValue("_otp", entrada.otp)
                comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_pwdUsr", String.Empty).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("_cod").Value.ToString()
                retorno.mensaje = comm.Parameters("_msj").Value.ToString()

                If retorno.codigo = 0 Then
                    pwdUsr = comm.Parameters("_pwdUsr").Value.ToString()
                    retorno.objeto = objFuncionalidadess.DesencriptarAES(pwdUsr).objeto
                End If
            Else
                retorno.codigo = 14
                retorno.mensaje = "Ocurrió un error al abrir la conexión en la validación de OTP"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 16
            retorno.mensaje = "Error crítico al validar OTP: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function RegistrarNombreCorp(ByVal entrada As strCliente) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spRegistrarNombreCorporativo", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable
            Dim auxPwd = objFuncionalidadess.EncriptarAES(entrada.pwd)

            con.Open()
            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_nombreEmpresa", entrada.nombreCliente)
                comm.Parameters.AddWithValue("_idUsrMod", entrada.idUsrMod)
                comm.Parameters.AddWithValue("_idTipoConsulta", entrada.idTipoConsulta)
                comm.Parameters.AddWithValue("_idEmpresa", entrada.idCliente)
                comm.Parameters.AddWithValue("_activo", entrada.activo)
                comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_auxIdEmpresa", MySqlDbType.Int32).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("_cod").Value.ToString()
                retorno.mensaje = comm.Parameters("_msj").Value.ToString()
                retorno.objeto = comm.Parameters("_auxIdEmpresa").Value.ToString()
            Else
                retorno.codigo = 18
                retorno.mensaje = "Ocurrió un error al abrir la conexión para el registro del Nombre Corporativo"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 19
            retorno.mensaje = "Error crítico en el registro del Nombre Corporativo: " & ex.Message()
        End Try

        Return retorno
    End Function
    Public Function ObtenerSelNomCorp(ByVal accion As Integer) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spGetSelect", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()
            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("accion", accion)

                da.SelectCommand = comm
                da.SelectCommand.Connection = con
                da.Fill(dtResult)
                con.Close()

                If dtResult.Rows.Count > 0 Then
                    retorno.codigo = 0
                    retorno.mensaje = "OK"
                    retorno.objeto = objFuncionalidadess.DtToJSON(dtResult).objeto
                End If
            Else
                retorno.codigo = 20
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la carga del select"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 21
            retorno.mensaje = "Error crítico para la carga del select: " & ex.Message()
        End Try

        Return retorno

    End Function

    Public Function ActualizarPwdUsuario(ByVal entrada As strCliente) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spCambiarPwd", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable
            Dim auxPwd = objFuncionalidadess.EncriptarAES(entrada.pwd)

            con.Open()
            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_pwd", auxPwd.objeto)
                comm.Parameters.AddWithValue("_IdUsrMod", entrada.idUsrMod)
                comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("_cod").Value.ToString()
                retorno.mensaje = comm.Parameters("_msj").Value.ToString()
            Else
                retorno.codigo = 22
                retorno.mensaje = "Ocurrió un error al abrir la conexión para el cambio de contraseña"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 23
            retorno.mensaje = "Error crítico en el cambio de contraseña: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function ActualizarInformacionUsuario(ByVal entrada As strCliente) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spActualizarInfUsu", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable
            Dim auxPwd = If(Not IsNothing(entrada.pwd), objFuncionalidadess.EncriptarAES(entrada.pwd), Nothing)

            con.Open()
            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_nombreCliente", entrada.nombreCliente)
                comm.Parameters.AddWithValue("_telefono", entrada.telefono)
                comm.Parameters.AddWithValue("_email", entrada.email)
                comm.Parameters.AddWithValue("_idUsrMod", entrada.idUsrMod)
                comm.Parameters.AddWithValue("_pwd", auxPwd.objeto)
                comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("_cod").Value.ToString()
                retorno.mensaje = comm.Parameters("_msj").Value.ToString()
            Else
                retorno.codigo = 24
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la actualiación del perfil"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 25
            retorno.mensaje = "Error crítico en la actualiación del perfil: " & ex.Message()
        End Try

        Return retorno
    End Function
    Public Function CargarUsuarios(ByVal idUsu As Integer) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spConsultaUsuarios", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()
            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_idUsr", idUsu)

                da.SelectCommand = comm
                da.SelectCommand.Connection = con
                da.Fill(dtResult)
                con.Close()

                If dtResult.Rows.Count > 0 Then
                    retorno.codigo = 0
                    retorno.mensaje = "OK"
                    retorno.objeto = objFuncionalidadess.DtToJSON(dtResult).objeto
                End If
            Else
                retorno.codigo = 26
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la carga de usuarios"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 27
            retorno.mensaje = "Error crítico para la carga de usuarios: " & ex.Message()
        End Try

        Return retorno

    End Function

    Public Function AdministrarRolesUsuario(ByVal entrada As strRolesUsuarios) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spAdministrarRoles", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()
            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_esAdmin", entrada.esAdmin)
                comm.Parameters.AddWithValue("_bloqueado", entrada.bloqueado)
                comm.Parameters.AddWithValue("_activo", entrada.activo)
                comm.Parameters.AddWithValue("_idUsuario", entrada.idUsr)
                comm.Parameters.AddWithValue("_idUsrMod", entrada.idUsrMod)
                comm.Parameters.AddWithValue("_idEmpresa", entrada.idEmpresa)
                comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output

                comm.ExecuteNonQuery()
                con.Close()

                retorno.codigo = comm.Parameters("_cod").Value.ToString()
                retorno.mensaje = comm.Parameters("_msj").Value.ToString()
            Else
                retorno.codigo = 28
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la actualiación de roles"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 29
            retorno.mensaje = "Error crítico en la actualiación de roles: " & ex.Message()
        End Try

        Return retorno
    End Function

    Public Function RegistrarIntentosFallidos(ByVal entrada As strRolesUsuarios) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spRegistrarFallidos", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()

            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_accion", entrada.accion)
                comm.Parameters.AddWithValue("_intentosFallidos", entrada.fallidos)
                comm.Parameters.AddWithValue("_bloqueado", entrada.bloqueado)
                comm.Parameters.AddWithValue("_idUsuario", entrada.idUsr)
                comm.Parameters.AddWithValue("_idUsrMod", entrada.idUsrMod)

                comm.ExecuteNonQuery()
                con.Close()
            Else
                retorno.codigo = 30
                retorno.mensaje = "Ocurrió un error al abrir la conexión para el registro de fallidos"
            End If


        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 31
            retorno.mensaje = "Error crítico en el registro de fallidos: " & ex.Message()
        End Try

        Return retorno
    End Function
    Public Function CRUDempresas(ByVal entrada As strEmpresas) As strRespuesta
        Dim retorno As New strRespuesta
        Dim con As New MySqlConnection(_conexLocal.objeto)

        Try
            Dim comm As New MySqlCommand("portal_spActualizarEmpresas", con)
            Dim da As New MySqlDataAdapter
            Dim dtResult As New DataTable

            con.Open()

            If con.State = ConnectionState.Open Then

                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.AddWithValue("_accion", entrada.accion)
                comm.Parameters.AddWithValue("_nombreEmpresa", entrada.nombre)
                comm.Parameters.AddWithValue("_activo", entrada.activo)
                comm.Parameters.AddWithValue("_idEmpresa", entrada.idEmpresa)
                comm.Parameters.AddWithValue("_idUsrMod", entrada.idUsrMod)
                comm.Parameters.AddWithValue("_cod", MySqlDbType.Int32).Direction = ParameterDirection.Output
                comm.Parameters.AddWithValue("_msj", String.Empty).Direction = ParameterDirection.Output

                da.SelectCommand = comm
                da.SelectCommand.Connection = con
                da.Fill(dtResult)
                con.Close()

                Select Case entrada.accion
                    Case 1
                        If dtResult.Rows.Count > 0 Then

                            retorno.codigo = 0
                            retorno.mensaje = "Operación exitosa"
                            retorno.objeto = objFuncionalidadess.DtToJSON(dtResult)

                        Else
                            retorno.codigo = 32
                            retorno.mensaje = "No se encontraron resultados"
                        End If
                    Case 2
                        retorno.codigo = comm.Parameters("_cod").Value.ToString()
                        retorno.mensaje = comm.Parameters("_msj").Value.ToString()
                End Select

            Else
                retorno.codigo = 33
                retorno.mensaje = "Ocurrió un error al abrir la conexión para la consulta de empresas"
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            retorno.codigo = 34
            retorno.mensaje = "Error crítico en el sistema de logueo: " & ex.Message()
        End Try

        Return retorno
    End Function
End Class
