Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class Panel_Principal

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Cree una nueva instancia del formulario secundario.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Conviértalo en un elemento secundario de este formulario MDI antes de mostrarlo.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Ventana " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: agregue código aquí para abrir el archivo.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: agregue código aquí para guardar el contenido actual del formulario en un archivo.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Utilice My.Computer.Clipboard.GetText() o My.Computer.Clipboard.GetData para recuperar la información del Portapapeles.
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Cierre todos los formularios secundarios del principal.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub RegistrarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistrarToolStripMenuItem.Click
        Dim CE As New Formulario_Retencion()
        'Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub Panel_Principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim consulta As String
            Dim dt1 As New DataTable
            Dim nombre_gestor As String = ""
            Dim dominio As Integer

            consulta = "SELECT Id_Dominio From USUARIOS Where Nombre_Usuario='" & nombreusuario & "'"
            Dim adaptadorSQL As New SqlDataAdapter(consulta, New SqlConnection(coneccion))
            adaptadorSQL.Fill(dt1)

            If adaptadorSQL.Fill(dt1) <> 0 Then

                Dim j As Integer

                For j = 0 To dt1.Rows.Count - 1
                    dominio = dt1.Rows(j).Item("Id_Dominio")
                Next
                If dominio = 1 Then

                Else
                    If dominio = 2 Then
                        ToolStrip3.Visible = True
                        ToolStrip1.Visible = False
                        ToolStrip2.Visible = False
                        ToolStripLabel1.Text = "                                                                                                                                                                                        "
                        ToolStripLabel2.Text = nombreusuario
                        ToolStripLabel3.Text = Today
                    Else
                        If dominio = 4 Then
                            ToolStrip1.Visible = True
                            ToolStrip3.Visible = False
                            ToolStrip2.Visible = False
                            ToolStripLabel4.Text = "                                                                                                                                                                                        "
                            ToolStripLabel6.Text = nombreusuario
                            ToolStripLabel5.Text = Today

                        Else
                            If dominio = 3 Then
                                ToolStrip2.Visible = True
                                ToolStrip3.Visible = False
                                ToolStrip1.Visible = False
                                ToolStripLabel7.Text = "                                                                                                                                                                                        "
                                ToolStripLabel9.Text = nombreusuario
                                ToolStripLabel8.Text = Today
                            End If
                        End If
                    End If
                End If


            Else
                MsgBox("El usuario o contraseña son incorrectos")
            End If


        Finally

        End Try
    End Sub

    Private Sub ConsultarToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ConsultarToolStripMenuItem2.Click
        Dim CE As New Consulta_Indivudual_Productividad()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub CambioDeClaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CambioDeClaveToolStripMenuItem.Click
        Dim CE As New Cambio_Clave()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem12.Click
        Dim CE As New Cambio_Clave()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub HallazgosPorProcesoToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem10.Click
        Dim CE As New Consolidado_Hallazgos_por_Proceso()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem11_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem11.Click
        Dim CE As New Consolidado_Hallazgos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        Dim CE As New Consolidado_Monitoreos_por_Gestion()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem9.Click
        Dim CE As New Consolidado_Monitoreos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ConsolidadoGeneralToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsolidadoGeneralToolStripMenuItem.Click
        Dim CE As New Cantidad_Monitoreos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ConsolidadoGeneralToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ConsolidadoGeneralToolStripMenuItem1.Click
        Dim CE As New Cantidad_Hallazgos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem14_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem14.Click
        Dim CE As New Formulario_Cartera()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub RegistrarToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RegistrarToolStripMenuItem1.Click
        Dim CE As New Formulario_Dilo_Claro()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub RegistrarToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles RegistrarToolStripMenuItem2.Click
        Dim CE As New Formulario_SAC()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Dim CE As New Formulario_TMK()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim CE As New Formulario_VIP()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        Dim CE As New Formulario_Webcenter()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem17_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem17.Click
        Dim CE As New Formulario_Soporte()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ConsolidadoCorreosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsolidadoCorreosToolStripMenuItem.Click
        Dim CE As New Consolidado_Correos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem31_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem31.Click
        Dim CE As New Reevaluaciones()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub
   
    Private Sub ToolStripMenuItem33_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem33.Click
        Dim CE As New Consolidado_Correos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub CargarReevaluaciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CargarReevaluaciónToolStripMenuItem.Click
        Dim CE As New Reevaluaciones()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ActualizarReevaluzaciònToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActualizarReevaluzaciònToolStripMenuItem.Click
        Dim CE As New Actualizar_Reevaluaciones()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim c1, c2, c3, c4, c5, c6 As String
        Dim dt1, dt2, dt3, dt4, dt5, dt6 As New DataTable

        c1 = "DROP TABLE TEMP_CORREOS_PARA_ENVIAR"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)

        c2 = "SELECT ALARMAS_CORREOS.* INTO TEMP_CORREOS_PARA_ENVIAR FROM ALARMAS_CORREOS"
        Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
        adaptadorsql2.Fill(dt2)

        c3 = "DELETE PENDIENDIENTE_ENVIAR"
        Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
        adaptadorsql3.Fill(dt3)

        c4 = "DROP TABLE PENDIENDIENTE_ENVIAR"
        Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
        adaptadorsql4.Fill(dt4)

        c5 = "SELECT TEMP_CORREOS_PARA_ENVIAR.ID_CORREO, TEMP_CORREOS_PARA_ENVIAR.USUARIO, TEMP_CORREOS_PARA_ENVIAR.FECHA_GESTION, TEMP_CORREOS_PARA_ENVIAR.AÑO_GESTION, TEMP_CORREOS_PARA_ENVIAR.MES_GESTION, TEMP_CORREOS_PARA_ENVIAR.DIA_GESTION, TEMP_CORREOS_PARA_ENVIAR.PROCESO, TEMP_CORREOS_PARA_ENVIAR.FRENTE, TEMP_CORREOS_PARA_ENVIAR.IMPACTO, TEMP_CORREOS_PARA_ENVIAR.RESULTADO_ACOMPAÑAMIENTO, TEMP_CORREOS_PARA_ENVIAR.NOMBRE_ASESOR, TEMP_CORREOS_PARA_ENVIAR.ALIADO, TEMP_CORREOS_PARA_ENVIAR.OPERACION, TEMP_CORREOS_PARA_ENVIAR.CUENTA, TEMP_CORREOS_PARA_ENVIAR.MOTIVO_LLAMADA, TEMP_CORREOS_PARA_ENVIAR.OPORTUNIDAD_MEJORA, TEMP_CORREOS_PARA_ENVIAR.RECOMENDACIONES, TEMP_CORREOS_PARA_ENVIAR.OBSERCACIONES, TEMP_CORREOS_PARA_ENVIAR.UBICACION_LLAMADA INTO PENDIENDIENTE_ENVIAR FROM TEMP_CORREOS_PARA_ENVIAR LEFT JOIN CORREOS_ENVIADOS ON TEMP_CORREOS_PARA_ENVIAR.ID_CORREO = CORREOS_ENVIADOS.ID_CORREO WHERE (((CORREOS_ENVIADOS.ID_CORREO) Is Null))"
        Dim adaptadorsql5 As New SqlDataAdapter(c5, New SqlConnection(coneccion))
        adaptadorsql5.Fill(dt5)

        c6 = "INSERT INTO CORREOS_ENVIADOS ( ID_CORREO, USUARIO, FECHA_GESTION, AÑO_GESTION, MES_GESTION, DIA_GESTION, PROCESO, FRENTE, IMPACTO, RESULTADO_ACOMPAÑAMIENTO, NOMBRE_ASESOR, ALIADO, OPERACION, CUENTA, MOTIVO_LLAMADA, OPORTUNIDAD_MEJORA, RECOMENDACIONES, OBSERCACIONES, UBICACION_LLAMADA )SELECT PENDIENDIENTE_ENVIAR.ID_CORREO, PENDIENDIENTE_ENVIAR.USUARIO, PENDIENDIENTE_ENVIAR.FECHA_GESTION, PENDIENDIENTE_ENVIAR.AÑO_GESTION, PENDIENDIENTE_ENVIAR.MES_GESTION, PENDIENDIENTE_ENVIAR.DIA_GESTION, PENDIENDIENTE_ENVIAR.PROCESO, PENDIENDIENTE_ENVIAR.FRENTE, PENDIENDIENTE_ENVIAR.IMPACTO, PENDIENDIENTE_ENVIAR.RESULTADO_ACOMPAÑAMIENTO, PENDIENDIENTE_ENVIAR.NOMBRE_ASESOR, PENDIENDIENTE_ENVIAR.ALIADO, PENDIENDIENTE_ENVIAR.OPERACION, PENDIENDIENTE_ENVIAR.CUENTA, PENDIENDIENTE_ENVIAR.MOTIVO_LLAMADA, PENDIENDIENTE_ENVIAR.OPORTUNIDAD_MEJORA, PENDIENDIENTE_ENVIAR.RECOMENDACIONES, PENDIENDIENTE_ENVIAR.OBSERCACIONES, PENDIENDIENTE_ENVIAR.UBICACION_LLAMADA FROM PENDIENDIENTE_ENVIAR"
        Dim adaptadorsql6 As New SqlDataAdapter(c6, New SqlConnection(coneccion))
        adaptadorsql6.Fill(dt6)

        MsgBox("Proceso Realizado")


    End Sub

    Private Sub ToolStripMenuItem20_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem20.Click
        Dim CE As New Cantidad_Monitoreos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem22_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem22.Click
        Dim CE As New Cantidad_Hallazgos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem24_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem24.Click
        Dim CE As New Consolidado_Monitoreos_por_Gestion()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem25_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem25.Click
        Dim CE As New Consolidado_Monitoreos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem28_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem28.Click
        Dim CE As New Consolidado_Hallazgos_por_Proceso()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem29_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem29.Click
        Dim CE As New Consolidado_Hallazgos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem30_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem30.Click
        Dim CE As New Consolidado_Correos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub BasePersonalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BasePersonalToolStripMenuItem.Click
        Dim CE As New Admin_Base_Personal()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub CarteraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CarteraToolStripMenuItem.Click
        Dim CE As New Actualizar_Monitoreos_Cartera()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ClaroTeAyudaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClaroTeAyudaToolStripMenuItem.Click
        Dim CE As New Actualizacion_Monitoreos_Dilo_Claro()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub RetencionToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RetencionToolStripMenuItem1.Click
        Dim CE As New Actualizar_Monitoreos_Retencion()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub SacToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SacToolStripMenuItem1.Click
        Dim CE As New Actualizacion_Monitoreos_SAC()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub SoporteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoporteToolStripMenuItem.Click
        Dim CE As New Actualizacion_Monitoreos_Soporte()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub TMKToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TMKToolStripMenuItem1.Click
        Dim CE As New Actualización_Monitoreos_TMK()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub VIPToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles VIPToolStripMenuItem1.Click
        Dim CE As New Actualizacion_Formulario_VIP()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub WebcenterToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles WebcenterToolStripMenuItem1.Click
        Dim CE As New Actualizacion_Monitoreos_Webcenter()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub CorreosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CorreosToolStripMenuItem.Click
        Dim CE As New Actualizacion_Correos()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ToolStripMenuItem34_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem34.Click
        Dim CE As New Monitoreos_CGV()
        ''Set the Parent Form of the Child window.
        CE.MdiParent = Me
        'Display the new form.
        CE.Show()
        CE.StartPosition = FormStartPosition.CenterScreen
    End Sub
End Class
