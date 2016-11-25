Imports System.Data.SqlClient

Public Class Formulario_Soporte
    Private Sub Formulario_Soporte_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ID_MACROPROCESO = 6
        Dim c1 As String
        Dim dt1 As New DataTable

        c1 = "SELECT PROCESOS.DESCRIPCION FROM PROCESOS INNER JOIN MACROPROCESOS ON PROCESOS.ID_MACROPROCESO = MACROPROCESOS.ID_MACROPROCESO WHERE (((MACROPROCESOS.ID_MACROPROCESO)=" & ID_MACROPROCESO & "))"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)

        Dim i As Integer

        For i = 0 To dt1.Rows.Count - 1
            ComboBox10.Items.Add(dt1.Rows(i).Item("DESCRIPCION"))
        Next
        Label25.Text = nombreusuario

        Dim c11 As String
        Dim dt11 As New DataTable


        c11 = "select DESCRIPCION from [dbo].[MOTIVOS_LLAMADA]"
        Dim adaptadorsql11 As New SqlDataAdapter(c11, New SqlConnection(coneccion))
        adaptadorsql11.Fill(dt11)
        Dim h As Integer

        For h = 0 To dt11.Rows.Count - 1
            ComboBox8.Items.Add(dt11.Rows(h).Item("DESCRIPCION"))
        Next

        ComboBox6.Text = "PRIMER NIVEL"
        ComboBox2.Text = "SOPORTE"
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        TextBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox14.Text = ""
        TextBox3.DataBindings.Clear()
        ComboBox4.DataBindings.Clear()
        ComboBox14.DataBindings.Clear()

        If TextBox2.Text <> "" Then

            Dim c1 As String
            Dim dt1 As New DataTable
            Dim cedula As Double = TextBox2.Text

            c1 = "SELECT * FROM BD_PERSONAL_RETENCION WHERE CEDULA = " & cedula & ""
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)

            If adaptadorsql1.Fill(dt1) <> 0 Then
                TextBox3.DataBindings.Add(New Binding("Text", dt1, "NOMBRE"))
                ComboBox4.DataBindings.Add(New Binding("Text", dt1, "ALIADO"))
                ComboBox14.DataBindings.Add(New Binding("Text", dt1, "OPERACION"))
            Else

            End If
        Else
        End If
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox10.SelectedIndexChanged
        ComboBox11.Items.Clear()
        ComboBox11.DataBindings.Clear()
        ComboBox11.SelectedIndex = -1

        Dim c1, c2 As String
        Dim dt1, dt2 As New DataTable
        Dim id_proceso As Integer = 0
        Dim nombre_proceso As String = ComboBox10.Text


        c1 = "SELECT ID_PROCESO FROM PROCESOS WHERE DESCRIPCION like '" & nombre_proceso & "'"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
        Dim j As Integer

        For j = 0 To dt1.Rows.Count - 1
            id_proceso = dt1.Rows(j).Item("ID_PROCESO")
        Next

        c2 = "SELECT SUBPROCESOS.DESCRIPCION FROM (SUBPROCESOS INNER JOIN PROCESOS ON SUBPROCESOS.ID_PROCESO = PROCESOS.ID_PROCESO) INNER JOIN MACROPROCESOS ON PROCESOS.ID_MACROPROCESO = MACROPROCESOS.ID_MACROPROCESO WHERE (((PROCESOS.DESCRIPCION)='" & nombre_proceso & "') AND ((MACROPROCESOS.ID_MACROPROCESO)=6))"
        Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
        adaptadorsql2.Fill(dt2)

        Dim i As Integer

        For i = 0 To dt2.Rows.Count - 1
            ComboBox11.Items.Add(dt2.Rows(i).Item("DESCRIPCION"))
        Next

    End Sub
    Private Sub ComboBox11_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox11.SelectedIndexChanged
        ComboBox12.Items.Clear()
        ComboBox12.DataBindings.Clear()
        ComboBox12.SelectedIndex = -1

        Dim c1, c2 As String
        Dim dt1, dt2 As New DataTable
        Dim id_subproceso As Integer = 0
        Dim nombre_subproceso As String = ComboBox11.Text
        Dim nombre_proceso As String = ComboBox10.Text

        c1 = "SELECT ID_SUBPROCESO FROM SUBPROCESOS WHERE DESCRIPCION like '" & nombre_subproceso & "'"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
        Dim j As Integer

        For j = 0 To dt1.Rows.Count - 1
            id_subproceso = dt1.Rows(j).Item("ID_SUBPROCESO")
        Next

        c2 = "SELECT SUBITEM.DESCRIPCION FROM ((SUBPROCESOS INNER JOIN PROCESOS ON SUBPROCESOS.ID_PROCESO = PROCESOS.ID_PROCESO) INNER JOIN MACROPROCESOS ON PROCESOS.ID_MACROPROCESO = MACROPROCESOS.ID_MACROPROCESO) INNER JOIN SUBITEM ON SUBPROCESOS.ID_SUBPROCESO = SUBITEM.ID_SUBPROCESO WHERE (((MACROPROCESOS.ID_MACROPROCESO)=6) AND ((PROCESOS.DESCRIPCION)='" & nombre_proceso & "') AND ((SUBPROCESOS.DESCRIPCION)='" & nombre_subproceso & "'))"
        Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
        adaptadorsql2.Fill(dt2)

        Dim i As Integer

        For i = 0 To dt2.Rows.Count - 1
            ComboBox12.Items.Add(dt2.Rows(i).Item("DESCRIPCION"))
        Next

    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Dim dato1 As String = TextBox6.Text
        Dim dato2 As String = TextBox8.Text
        Dim resultado As String

        resultado = dato1 & dato2
        TextBox9.Text = resultado
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.LostFocus
        ComboBox14.Focus()
    End Sub

    Private Sub ComboBox14_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox14.LostFocus
        ComboBox5.Focus()
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.LostFocus
        Button1.Focus()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim CUENTA As Double = TextBox1.Text
        Dim FECHA As String = Today
        Dim NOMBRE As String = TextBox3.Text
        Dim CEDULA As Double = TextBox2.Text
        Dim ALIADO As String = ComboBox4.Text
        Dim OPERACION As String = ComboBox14.Text
        Dim NOTA_FINAL As Double = TextBox7.Text
        Dim ETIQUETA As String = TextBox6.Text
        Dim USUARIOMEC As String = ""
        Dim AÑO As Integer = Now.Year
        Dim MES As Integer = Now.Month
        Dim DIA As Integer = Now.Day
        Dim TIPO_CLIENTE As String = ComboBox1.Text
        Dim GESTOR_REPORTA As String = ""
        Dim PROCESO As String = ComboBox2.Text
        Dim FRENTE As String = ComboBox3.Text
        Dim DIVISION As String = ComboBox5.Text
        Dim RESULTADO As String = ComboBox7.Text
        Dim MOTIVO_LLAMADA As String = ComboBox8.Text
        Dim RAZON_LLAMADA As String = ComboBox9.Text
        Dim INCUMPLIMIENTO_PROCESO As String = ComboBox10.Text
        Dim ESPECIFICACION As String = ComboBox11.Text
        Dim SUB_ITEM As String = ComboBox12.Text
        Dim OBSERVACIONES As String = TextBox8.Text
        Dim FASE_AFECTADA As String = ComboBox13.Text
        Dim NIVEL_SE As String = ComboBox6.Text
        Dim MARCACION_RR As String = TextBox4.Text
        Dim MARCACION_CORRECTA As String = TextBox5.Text
        Dim NOVO As String = TextBox9.Text
        Dim REGISTRO_NOVO As Double = TextBox10.Text

        Dim c1, c2, c3, c4, c5, c6 As String
        Dim dt1, dt2, dt3, dt4, dt5, dt6 As New DataTable

        c1 = "Select Usuario from USUARIOS where Nombre_Usuario like '" & nombreusuario & "'"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
        Dim i As Integer

        For i = 0 To dt1.Rows.Count - 1
            USUARIOMEC = dt1.Rows(i).Item("Usuario")
        Next
        If (MsgBox("Desea guardar el registro?", MsgBoxStyle.YesNo, "Guardar")) = MsgBoxResult.Yes Then

            If RESULTADO <> "MONITOREO" Then

                c2 = "INSERT INTO HALLAZGOS_SOPORTE (USUARIO,FECHA_HALLAZGO,AÑO_HALLAZGO,MES_HALLAZGO,DIA_HALLAZGO,CUENTA,TIPO_CLIENTE,GESTOR_REPORTA,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & nombreusuario & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
                adaptadorsql2.Fill(dt2)


                c3 = "INSERT INTO MONITOREOS_SOPORTE (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                adaptadorsql3.Fill(dt3)

                c4 = "INSERT INTO CONSOLIDADO_MONITOREOS (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                adaptadorsql4.Fill(dt4)

                c5 = "INSERT INTO CONSOLIDADO_HALLAZGOS (USUARIO,FECHA_HALLAZGO,AÑO_HALLAZGO,MES_HALLAZGO,DIA_HALLAZGO,CUENTA,TIPO_CLIENTE,GESTOR_REPORTA,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & nombreusuario & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql5 As New SqlDataAdapter(c5, New SqlConnection(coneccion))
                adaptadorsql5.Fill(dt5)


                c6 = "INSERT INTO PRODUCTIVIDAD (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql6 As New SqlDataAdapter(c6, New SqlConnection(coneccion))
                adaptadorsql6.Fill(dt6)
                MsgBox("Registro Almacenado")

                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox9.Text = ""
                TextBox10.Text = ""

                ComboBox1.Text = ""
                ComboBox2.Text = ""
                ComboBox3.Text = ""
                ComboBox4.Text = ""
                ComboBox5.Text = ""
                ComboBox6.Text = ""
                ComboBox7.Text = ""
                ComboBox8.Text = ""
                ComboBox9.Text = ""
                ComboBox10.Text = ""
                ComboBox11.Text = ""
                ComboBox12.Text = ""
                ComboBox13.Text = ""
                ComboBox14.Text = ""
                ComboBox2.Text = "SOPORTE"
                ComboBox6.Text = "PRIMER NIVEL"
            Else
                If RESULTADO = "MONITOREO" Then
                    c3 = "INSERT INTO MONITOREOS_SOPORTE (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                    Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                    adaptadorsql3.Fill(dt3)

                    c4 = "INSERT INTO CONSOLIDADO_MONITOREOS (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                    Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                    adaptadorsql4.Fill(dt4)

                    c6 = "INSERT INTO PRODUCTIVIDAD (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                    Dim adaptadorsql6 As New SqlDataAdapter(c6, New SqlConnection(coneccion))
                    adaptadorsql6.Fill(dt6)
                    MsgBox("Registro Almacenado")

                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox9.Text = ""
                    TextBox10.Text = ""

                    ComboBox1.Text = ""
                    ComboBox2.Text = ""
                    ComboBox3.Text = ""
                    ComboBox4.Text = ""
                    ComboBox5.Text = ""
                    ComboBox6.Text = ""
                    ComboBox7.Text = ""
                    ComboBox8.Text = ""
                    ComboBox9.Text = ""
                    ComboBox10.Text = ""
                    ComboBox11.Text = ""
                    ComboBox12.Text = ""
                    ComboBox13.Text = ""
                    ComboBox14.Text = ""
                    ComboBox2.Text = "SOPORTE"
                    ComboBox6.Text = "PRIMER NIVEL"

                End If

            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim CUENTA As Double = TextBox1.Text
        Dim FECHA As String = Today
        Dim NOMBRE As String = TextBox3.Text
        Dim CEDULA As Double = TextBox2.Text
        Dim ALIADO As String = ComboBox4.Text
        Dim OPERACION As String = ComboBox14.Text
        Dim NOTA_FINAL As Double = TextBox7.Text
        Dim ETIQUETA As String = TextBox6.Text
        Dim USUARIOMEC As String = ""
        Dim OBSERVACIONES As String = TextBox8.Text
        Dim PROCESOS As String = ComboBox2.Text
        Dim FRENTES As String = ComboBox3.Text
        Dim REGISTRO_NOVO As Double = TextBox10.Text

        CUENTA_CORREO = CUENTA
        FECHA_CORREO = FECHA
        NOMBRE_ASESOR = NOMBRE
        CEDULA_ASESOR = CEDULA
        ALIADO_ASESOR = ALIADO
        OPERACION_ASESOR = OPERACION
        NOTA_ASESOR = NOTA_FINAL
        ETIQUETA_LLAMADA = ETIQUETA
        OBSERVACIONES_ASESOR = OBSERVACIONES
        PROCESO_GESTION = PROCESOS
        FRENTE_GESTION = FRENTES
        REGISTRO_NOVO_GENERAL = REGISTRO_NOVO

        Formato_Correo.Show()
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If ComboBox7.Text = "MONITOREO" Then
            Button2.Hide()
        Else
            Button2.Show()
        End If
    End Sub
End Class