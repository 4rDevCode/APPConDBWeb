'*******************************************************************************************************************************************
'*                                                                                                                                         *
'*                                                        JEFRINSSON J F CALDERON                                                          *
'*                                               doblejota4r@gmail.com | jefrinssonjfcalderon@gmail.com                                    *
'*                                                                                                                                         *
'*******************************************************************************************************************************************


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'              IMPORTACIONES USADAS
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Imports System.Data.SqlClient
Imports System.Data

Public Class cls_Mantenimiento

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              ATRIBUTOS
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private _Id As Integer
    Private _Nombre As String
    Private _Apellidos As String
    Private _Correo As String
    Private _Telefono As String
    Private _Estado As String

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              CONSTRUCTORES
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Sub New()
    End Sub

    Public Sub New(oId As Integer, oNombre As String, oApellidos As String, oCorreo As String, oTelefono As String, oEstado As String)
        _Id = oId
        _Nombre = oNombre
        _Apellidos = oApellidos
        _Correo = oCorreo
        _Telefono = oTelefono
        _Estado = oEstado
    End Sub

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              FUNCIONES
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Public Function UlitmoId() As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("SELECT TOP 1 ID FROM MANTENIMIENTO ORDER BY ID DESC", LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    Public Function CapturarIdWeb(Id As String) As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("SELECT ID_WEB FROM MANTENIMIENTO WHERE ID = " + Id, LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    Public Function ListaMantenimientoSinIdWEB() As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("SELECT * FROM MANTENIMIENTO WHERE ID_WEB IS NULL", LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    Public Function ListaMantenimientoConIdWEB() As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("SELECT * FROM MANTENIMIENTO WHERE ID_WEB IS NOT NULL", LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    Public Function ListaMantenimientoWEB() As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_WEB()
            Dim DA As New SqlDataAdapter("SELECT * FROM MANTENIMIENTO", WEB_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_WEB()
        Return DT
    End Function
    Public Function Filtrar(ByVal Var As String) As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("SELECT ID, APELLIDOS + ', ' + NOMBRE AS DATOS, ESTADO, ID_WEB FROM MANTENIMIENTO " +
                                         "WHERE ((APELLIDOS LIKE '%" + Var + "%') OR (NOMBRE LIKE '%" + Var + "%')) OR  " +
                                         "(APELLIDOS + ', ' + NOMBRE LIKE '%" + Var + "%')", LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    Public Function Buscar(ByVal Id As String) As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("SELECT * FROM MANTENIMIENTO WHERE ID = " + Id, LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    'Public Function ObtenerFechaDeUltimaSincronizacion() As DataTable
    '    Dim DT As New DataTable
    '    Try
    '        Call CONECTAR_LOCAL()
    '        Dim DA As New SqlDataAdapter("SELECT CAST(FECHA_HORA AS datetime2) FROM ULTIMA_SINCRONISACION", LOCAL_CON)
    '        DA.Fill(DT)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    '    Call DESCONECTAR_LOCAL()
    '    Return DT
    'End Function
    Public Function ObtenerFechaDeUltimaSincronizacion() As String

        Dim Msj As String = ""
        Dim sql As String = _
            "DECLARE @FECHA AS VARCHAR(23) = '' " +
            "IF(EXISTS(SELECT * FROM ULTIMA_SINCRONIZACION)) " +
            "	SET @FECHA = (SELECT CAST(DATEADD(MINUTE, -10, FECHA_HORA) AS datetime2) FROM ULTIMA_SINCRONIZACION) " +
            "SELECT @FECHA "
        Dim cmd As New SqlCommand(sql, LOCAL_CON)
        Try
            CONECTAR_LOCAL()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_LOCAL()
        Return Msj
    End Function

    Public Function ListaMantenimientoParaSincronizarUpdateLocal() As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("IF(EXISTS(SELECT * FROM ULTIMA_SINCRONIZACION)) " +
                                         "  SELECT * FROM MANTENIMIENTO WHERE FECHA_UPDATE > (SELECT FECHA_HORA FROM ULTIMA_SINCRONIZACION) AND ID_WEB IS NOT NULL " +
                                         "ELSE " +
                                         "	SELECT * FROM MANTENIMIENTO WHERE FECHA_UPDATE IS NOT NULL  AND ID_WEB IS NOT NULL", LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    Public Function ListaMantenimientoParaSincronizarUpdateWeb(Fecha As String) As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_WEB()
            Dim DA As New SqlDataAdapter("IF('" + Fecha + "' <> '') " +
                                        "  SELECT * FROM MANTENIMIENTO WHERE FECHA_UPDATE >= '" + Fecha + "' " +
                                        "ELSE " +
                                        "	SELECT * FROM MANTENIMIENTO WHERE FECHA_UPDATE IS NOT NULL ", WEB_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_WEB()
        Return DT
    End Function
    Public Function BuscarPorIdWeb(ByVal Id As String) As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_LOCAL()
            Dim DA As New SqlDataAdapter("SELECT * FROM MANTENIMIENTO WHERE ID_WEB = " + Id, LOCAL_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_LOCAL()
        Return DT
    End Function
    Public Function BuscarPorIdEnLaWeb(ByVal Id As String) As DataTable
        Dim DT As New DataTable
        Try
            Call CONECTAR_WEB()
            Dim DA As New SqlDataAdapter("SELECT * FROM MANTENIMIENTO WHERE ID = " + Id, WEB_CON)
            DA.Fill(DT)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call DESCONECTAR_WEB()
        Return DT
    End Function
    Public Function Transaccion(ByVal T As String) As String
        Dim MSJ As String
        Try
            Call CONECTAR_LOCAL()
            Dim CMD As New SqlCommand("sp_T_Mantenimiento", LOCAL_CON)
            CMD.CommandType = 4
            With CMD.Parameters
                .AddWithValue("@T", T)
                .AddWithValue("@Id", _Id)
                .AddWithValue("@Nombre", _Nombre)
                .AddWithValue("@Apellidos", _Apellidos)
                .AddWithValue("@Correo", _Correo)
                .AddWithValue("@TLF", _Telefono)
                .AddWithValue("@Estado", _Estado)
                .Add("@Msj", SqlDbType.VarChar, 300).Direction = 2
            End With
            CMD.ExecuteNonQuery()
            MSJ = CMD.Parameters("@Msj").Value
        Catch ex As Exception
            Throw ex
        End Try
        Call DESCONECTAR_LOCAL()
        Return MSJ
    End Function
    Public Function UpdateLocalDesdeWeb(ID As String) As String

        Dim Msj As String = ""
        Dim sql As String = _
            "UPDATE MANTENIMIENTO SET NOMBRE = '" + _Nombre + "', APELLIDOS = '" + _Apellidos + "', CORREO = '" + _Correo + "', TLF = '" + _Telefono + "', " +
            "ESTADO = '" + _Estado + "', FECHA_UPDATE = GETDATE() where ID_WEB = " + ID
        Dim cmd As New SqlCommand(sql, LOCAL_CON)
        Try
            CONECTAR_LOCAL()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_LOCAL()
        Return Msj
    End Function
    Public Function InsertWeb() As String

        Dim Msj As String = ""
        Dim sql As String = _
         "INSERT MANTENIMIENTO VALUES('" + _Nombre + "', '" + _Apellidos + "', '" + _Correo + "', '" + _Telefono + "', '" + _Estado + "', NULL) " +
         "SELECT @@IDENTITY"
        Dim cmd As New SqlCommand(sql, WEB_CON)
        Try
            CONECTAR_WEB()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_WEB()
        Return Msj
    End Function
    Public Function UpdateWeb(oIdWEB As String) As String

        Dim Msj As String = ""
        Dim sql As String = _
            "UPDATE MANTENIMIENTO SET NOMBRE = '" + _Nombre + "', APELLIDOS = '" + _Apellidos + "', CORREO = '" + _Correo + "', TLF = '" + _Telefono + "', " +
            "   ESTADO = '" + _Estado + "', FECHA_UPDATE = (DATEADD(HOUR, 3, GETDATE())) WHERE ID = " + oIdWEB
        Dim cmd As New SqlCommand(sql, WEB_CON)
        Try
            CONECTAR_WEB()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_WEB()
        Return Msj
    End Function
    Public Function UpdateALaFechaDeUltimaSincronizacion() As String

        Dim Msj As String = ""
        Dim sql As String = _
            "IF(NOT EXISTS(SELECT * FROM ULTIMA_SINCRONIZACION)) " +
            "   INSERT ULTIMA_SINCRONIZACION VALUES(GETDATE()) " +
            "ELSE " +
            "	UPDATE ULTIMA_SINCRONIZACION SET FECHA_HORA = GETDATE() "
        Dim cmd As New SqlCommand(sql, LOCAL_CON)
        Try
            CONECTAR_LOCAL()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_LOCAL()
        Return Msj
    End Function
    Public Function DeleteWeb(oIdWEB As String) As String

        Dim Msj As String = ""
        Dim sql As String = _
            "DELETE FROM MANTENIMIENTO WHERE ID = " + oIdWEB
        Dim cmd As New SqlCommand(sql, WEB_CON)
        Try
            CONECTAR_WEB()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_WEB()
        Return Msj
    End Function
    Public Function UpdateIdWeb_Local(LOCAL As String, WEB As String) As String

        Dim Msj As String = ""
        Dim sql As String = _
         "UPDATE MANTENIMIENTO SET ID_WEB = " + WEB + " WHERE ID = " + LOCAL
        Dim cmd As New SqlCommand(sql, LOCAL_CON)
        Try
            CONECTAR_LOCAL()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_LOCAL()
        Return Msj
    End Function
    Public Function Eliminar(ByVal ID As String) As String

        Dim Msj As String = ""
        Dim sql As String = _
            "DELETE FROM MANTENIMIENTO WHERE ID = " + ID + " " + _
            "SELECT 'Se ha eliminado correctamente.' "
        Dim cmd As New SqlCommand(sql, LOCAL_CON)
        Try
            CONECTAR_LOCAL()
            Msj = Convert.ToString(cmd.ExecuteScalar())
        Catch ex As Exception
            Throw ex
        End Try
        DESCONECTAR_LOCAL()
        Return Msj
    End Function
End Class
