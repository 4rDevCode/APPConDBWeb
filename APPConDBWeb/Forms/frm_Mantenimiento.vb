'*******************************************************************************************************************************************
'*                                                                                                                                         *
'*                                                        JEFRINSSON J F CALDERON                                                          *
'*                                               doblejota4r@gmail.com | jefrinssonjfcalderon@gmail.com                                    *
'*                                                                                                                                         *
'*******************************************************************************************************************************************
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'              IMPORTACIONES USADAS
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Class frm_Mantenimiento

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              DECLARACION DE VARIABLES PRIVADAS DEL FORM
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private T As String = "" REM PARA EL TIPO DE TRANSACCION
    Private B As String = "" REM PARA PERMITIR LA BUSQUEDA
    Private CLS As New cls_Mantenimiento

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              EVENTOS DEL FORM
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Sub frm_Mantenimiento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgvLista.DataSource = CLS.Filtrar("")
        Call FormatoGrid() : Call Bloquear("SI") : cmbEstado.SelectedIndex = 0
    End Sub
    Private Sub Form1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.Close()
        End If
    End Sub

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              FUNCIONES LOCALES
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Sub Bloquear(ByVal R As String)
        If (R = "SI") Then
            txtNombre.Enabled = False : txtApellidos.Enabled = False : txtCorreo.Enabled = False : txtTelefono.Enabled = False : cmbEstado.Enabled = False
        Else
            txtNombre.Enabled = True : txtApellidos.Enabled = True : txtCorreo.Enabled = True : txtTelefono.Enabled = True : cmbEstado.Enabled = True
        End If
    End Sub
    Private Sub FormatoGrid()
        Dim Ancho = dgvLista.Width
        dgvLista.Columns(0).Width = ((Ancho * 10) / 100)
        dgvLista.Columns(1).Width = ((Ancho * 56) / 100)
        dgvLista.Columns(2).Width = ((Ancho * 15) / 100)
        dgvLista.Columns(3).Width = ((Ancho * 14) / 100)

        dgvLista.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvLista.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvLista.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        dgvLista.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvLista.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvLista.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvLista.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        With dgvLista
            .RowsDefaultCellStyle.BackColor = Color.White
            .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
        End With
    End Sub
    Private Sub Limpiar()
        txtNombre.Clear() : txtApellidos.Clear() : txtCorreo.Clear() : txtTelefono.Clear() : cmbEstado.SelectedIndex = 0
    End Sub
    Private Sub Nuevo()
        Bloquear("NO") : txtNombre.Focus() : btnSave.Enabled = True : btnCancel.Enabled = True : btnSearch.Enabled = False : btnEdit.Enabled = False
        btnNew.Enabled = False : dgvLista.Enabled = False
    End Sub
    Private Sub Cancelar()
        T = "" : B = "" : btnNew.Enabled = True : btnSearch.Enabled = True : btnSave.Enabled = False : btnEdit.Enabled = False : btnCancel.Enabled = False
        Call Limpiar() : dgvLista.DataSource = CLS.Filtrar("") : Bloquear("SI") : btnDelete.Enabled = False : dgvLista.Enabled = True
        P1.BackColor = Color.Transparent : P2.BackColor = Color.Transparent
    End Sub
    Private Sub Buscar()
        Limpiar()
        Bloquear("SI")
        btnSearch.Enabled = False
        btnSave.Enabled = False
        btnCancel.Enabled = True
        btnEdit.Enabled = False
        txtNombre.Enabled = True : txtApellidos.Enabled = True : txtNombre.Focus()
        P1.BackColor = Color.Green : P2.BackColor = Color.Green
    End Sub
    Private Sub Editar()
        If (dgvLista.RowCount > 0) Then
            Try
                Dim DT As New DataTable
                DT = CLS.Buscar(CStr(dgvLista.Rows(dgvLista.CurrentRow.Index).DataBoundItem(0)))
                If (DT.Rows.Count > 0) Then
                    lblID.Text = DT.Rows(0).Item(0)
                    txtNombre.Text = DT.Rows(0).Item(1)
                    txtApellidos.Text = DT.Rows(0).Item(2)
                    txtCorreo.Text = DT.Rows(0).Item(3)
                    txtTelefono.Text = DT.Rows(0).Item(4)
                    If (DT.Rows(0).Item(5) = "A") Then
                        cmbEstado.SelectedIndex = 0
                    Else
                        cmbEstado.SelectedIndex = 1
                    End If
                    dgvLista.Enabled = False
                    T = "E"
                    btnSearch.Enabled = False : btnNew.Enabled = False : btnSave.Enabled = True : btnEdit.Enabled = False : btnCancel.Enabled = True
                    btnDelete.Enabled = True : Bloquear("NO") : txtNombre.Focus()
                Else
                    MessageBox.Show("Ups: No hay ninguno elemento para editar.", "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                MessageBox.Show("Ups error: " + ex.Message + vbCrLf + "Si el error persiste informe al proveedor por favor.",
                                "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            End Try
        Else
            MessageBox.Show("Ups: No hay nada seleccionado en la lista.", "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
        End If
    End Sub

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              EVENTO CLICK DE LOS BOTONES
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        T = "G" : Call Nuevo() : P1.BackColor = Color.Transparent : P2.BackColor = Color.Transparent
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Call Cancelar() : btnNew.Focus()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If (T <> "") Then
            Try
                Dim MSJ As String = ""
                REM REGISTRO LOCAL
                Dim REG_LOCAL As New cls_Mantenimiento(lblID.Text, txtNombre.Text, txtApellidos.Text, txtCorreo.Text, txtTelefono.Text, cmbEstado.Text.Substring(0, 1))
                MSJ = REG_LOCAL.Transaccion(T)
                dgvLista.DataSource = CLS.Filtrar("")
                If (MSJ = "Registro exitoso." Or MSJ = "Edición exitosa.") Then
                    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    Try
                        Dim IDWEB As String = ""
                        If (T = "G") Then REM REGISTRO WEB
                            Dim DT_ULTIMOID As DataTable = CLS.UlitmoId()
                            Dim REG_WEB As New cls_Mantenimiento(lblID.Text, txtNombre.Text, txtApellidos.Text, txtCorreo.Text, txtTelefono.Text, cmbEstado.Text.Substring(0, 1))
                            IDWEB = REG_WEB.InsertWeb()
                            CLS.UpdateIdWeb_Local(DT_ULTIMOID.Rows(0).Item(0), IDWEB)
                        ElseIf (T = "E") Then REM EDITAR WEB
                            Dim DT_IDWEB As DataTable = CLS.CapturarIdWeb(lblID.Text)
                            If (DT_IDWEB.Rows(0).Item(0) Is DBNull.Value = True) Then
                                REM SI NO TIENE ID WEB
                                Dim REG_WEB As New cls_Mantenimiento(lblID.Text, txtNombre.Text, txtApellidos.Text, txtCorreo.Text, txtTelefono.Text, cmbEstado.Text.Substring(0, 1))
                                IDWEB = REG_WEB.InsertWeb()
                                CLS.UpdateIdWeb_Local(lblID.Text, IDWEB)
                            Else
                                REM SI TIENE ID WEB
                                Dim MOD_WEB As New cls_Mantenimiento(lblID.Text, txtNombre.Text, txtApellidos.Text, txtCorreo.Text, txtTelefono.Text, cmbEstado.Text.Substring(0, 1))
                                MOD_WEB.UpdateWeb(DT_IDWEB.Rows(0).Item(0))
                            End If
                        End If
                    Catch ex As Exception : End Try
                    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    Call Cancelar() : btnNew.Focus()
                End If
                REM MOSTRAR MENSAJE
                MessageBox.Show(MSJ, "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Ups error: " + ex.Message + vbCrLf + "Si el error persiste informe al proveedor por favor." _
                                  , "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            End Try
        End If
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        B = "Yes" : Call Buscar()
    End Sub
    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        B = "" : Call Editar()
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        End 
    End Sub
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If (lblID.Text <> "" And lblID.Text <> "0") Then
            If MessageBox.Show("¿Esta Ud. de querer Eliminar " + txtApellidos.Text + ", " + txtNombre.Text + "  del Programa?", "Aviso", _
                                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = 6 Then
                Try
                    Dim Msj As String = "" : Dim DT_IDWEB As DataTable = CLS.CapturarIdWeb(lblID.Text)
                    Msj = CLS.Eliminar(lblID.Text)
                    If (Msj = "Se ha eliminado correctamente.") Then
                        If (DT_IDWEB.Rows(0).Item(0) Is DBNull.Value = False) Then
                            CLS.DeleteWeb(DT_IDWEB.Rows(0).Item(0))
                        End If
                        Call btnCancel_Click(Nothing, Nothing)
                    End If
                    MessageBox.Show(Msj, "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Ups: No se ha podido eliminar el Área seleccionado debido a que tiene relación " + _
                                      "con otros registros ya sean de otro ejercio de esta empresa o de otra." _
                                    , "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                End Try
            End If
        End If
    End Sub

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '              EVENTOS DEL DATAGRIDVIEW
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Sub dgvLista_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLista.MouseClick
        btnSearch.Enabled = False : btnNew.Enabled = False : btnSave.Enabled = False : btnEdit.Enabled = True : btnCancel.Enabled = True
    End Sub
    Private Sub dgvLista_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLista.MouseDoubleClick
        Call Editar()
    End Sub

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '               EVENTOS DE LAS CAJAS DE TEXTO
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Sub txtNombre_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNombre.TextChanged
        If (B = "Yes") Then
            Try
                dgvLista.DataSource = CLS.Buscar(txtNombre.Text)
            Catch ex As Exception
                MessageBox.Show("Ups error: " + ex.Message + vbCrLf + "Si el error persiste informe al proveedor por favor." _
                                  , "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            End Try
        End If
    End Sub
    Private Sub txtApellidos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtApellidos.TextChanged
        If (B = "Yes") Then
            Try
                dgvLista.DataSource = CLS.Buscar(txtApellidos.Text)
            Catch ex As Exception
                MessageBox.Show("Ups error: " + ex.Message + vbCrLf + "Si el error persiste informe al proveedor por favor." _
                                  , "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            End Try
        End If
    End Sub
    Private Sub txtNombre_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNombre.KeyPress
        If (Asc(e.KeyChar) = 13 And txtNombre.Text <> "") Then
            txtApellidos.Focus()
        End If
    End Sub
    Private Sub txtApellidos_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtApellidos.KeyPress
        If (Asc(e.KeyChar) = 13 And txtApellidos.Text <> "") Then
            txtCorreo.Focus()
        End If
    End Sub
    Private Sub txtCorreo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCorreo.KeyPress
        If (Asc(e.KeyChar) = 13) Then
            txtTelefono.Focus()
        End If
    End Sub
    Private Sub txtTelefono_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTelefono.KeyPress
        If (Asc(e.KeyChar) = 13) Then
            cmbEstado.Focus()
        End If
    End Sub
    Private Sub cmbEstado_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbEstado.KeyPress
        If (Asc(e.KeyChar) = 13) Then
            btnSave.Focus()
        End If
    End Sub

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '               EVENTOS CLICK PARA SINCRONIZAR CON EL SERVIDOR WEB
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Sub btnSincronizarServidorWeb_Click(sender As Object, e As EventArgs) Handles btnSincronizarServidorWeb.Click
        Try
            Dim DT_TEMP As New DataTable
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            '       LOCAL A WEB
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            REM INSERTS
            Dim DT_LISTA As DataTable = CLS.ListaMantenimientoSinIDWEB()
            If (DT_LISTA.Rows.Count > 0) Then
                For iList = 0 To DT_LISTA.Rows.Count - 1
                    Try
                        Dim IDWEB As String = ""
                        Dim REG_WEB As New cls_Mantenimiento(DT_LISTA.Rows(iList).Item(0), DT_LISTA.Rows(iList).Item(1), DT_LISTA.Rows(iList).Item(2),
                                                             DT_LISTA.Rows(iList).Item(3), DT_LISTA.Rows(iList).Item(4), DT_LISTA.Rows(iList).Item(5))
                        IDWEB = REG_WEB.InsertWeb()
                        CLS.UpdateIdWeb_Local(DT_LISTA.Rows(iList).Item(0), IDWEB)
                    Catch ex As Exception : End Try
                Next
            End If
            REM UPDATES
            DT_LISTA = CLS.ListaMantenimientoParaSincronizarUpdateLocal()
            If (DT_LISTA.Rows.Count > 0) Then
                For iList = 0 To DT_LISTA.Rows.Count - 1
                    Dim MOD_WEB As New cls_Mantenimiento(DT_LISTA.Rows(iList).Item(0), DT_LISTA.Rows(iList).Item(1), DT_LISTA.Rows(iList).Item(2), DT_LISTA.Rows(iList).Item(3),
                                                         DT_LISTA.Rows(iList).Item(4), DT_LISTA.Rows(iList).Item(5))
                    MOD_WEB.UpdateWeb(DT_LISTA.Rows(iList).Item(6))
                Next
            End If
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            '       WEB A LOCAL
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            REM INSERTS
            DT_LISTA = CLS.ListaMantenimientoWEB()
            If (DT_LISTA.Rows.Count > 0) Then
                For iList = 0 To DT_LISTA.Rows.Count - 1
                    DT_TEMP = CLS.BuscarPorIdWeb(DT_LISTA.Rows(iList).Item(0))
                    If (DT_TEMP.Rows.Count <= 0) Then
                        Dim REG_LOCAL As New cls_Mantenimiento("0", DT_LISTA.Rows(iList).Item(1), DT_LISTA.Rows(iList).Item(2), DT_LISTA.Rows(iList).Item(3),
                                                               DT_LISTA.Rows(iList).Item(4), DT_LISTA.Rows(iList).Item(5))
                        REG_LOCAL.Transaccion("G")
                        Dim DT_ULTIMOID As DataTable = CLS.UlitmoId()
                        CLS.UpdateIdWeb_Local(DT_ULTIMOID.Rows(0).Item(0), DT_LISTA.Rows(iList).Item(0))
                    End If
                Next
            End If
            REM UPDATES
            Dim FECHA As String = CLS.ObtenerFechaDeUltimaSincronizacion()
            DT_LISTA = CLS.ListaMantenimientoParaSincronizarUpdateWeb(FECHA)
            If (DT_LISTA.Rows.Count > 0) Then
                For iList = 0 To DT_LISTA.Rows.Count - 1
                    Dim REG_LOCAL As New cls_Mantenimiento(0, DT_LISTA.Rows(iList).Item(1), DT_LISTA.Rows(iList).Item(2), DT_LISTA.Rows(iList).Item(3),
                                                           DT_LISTA.Rows(iList).Item(4), DT_LISTA.Rows(iList).Item(5))
                    REG_LOCAL.UpdateLocalDesdeWeb(DT_LISTA.Rows(iList).Item(0))
                Next
            End If

            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            '       ELIMINAR LOCALMENTE 
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            DT_LISTA = CLS.ListaMantenimientoConIdWEB()
            For iList = 0 To DT_LISTA.Rows.Count - 1
                DT_TEMP = CLS.BuscarPorIdEnLaWeb(DT_LISTA.Rows(iList).Item(6))
                If (DT_TEMP.Rows.Count = 0) Then
                    Try
                        CLS.Eliminar(DT_LISTA.Rows(iList).Item(0))
                    Catch ex As Exception : End Try
                End If
            Next

            REM ACTUALIZAR FECHA DEL ULTIMO UPDATE
            CLS.UpdateALaFechaDeUltimaSincronizacion()
            REM ACTUALIZAR LISTA
            dgvLista.DataSource = CLS.Filtrar("")

            MessageBox.Show("Se ha sincronizado correctamente ;) ", "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Ups error: " + ex.Message + vbCrLf + "Si el error persiste informe al proveedor por favor." _
                                 , "Mensaje", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
        End Try
    End Sub
End Class
