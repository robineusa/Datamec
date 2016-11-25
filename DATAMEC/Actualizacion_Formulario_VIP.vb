Imports System.Data.SqlClient

Public Class Actualizacion_Formulario_VIP

    Private Sub Actualizacion_Formulario_VIP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ID_MACROPROCESO = 1
        Dim c1 As String
        Dim dt1 As New DataTable

        c1 = "SELECT DESCRIPCION FROM PROCESOS WHERE ID_MACROPROCESO =" & ID_MACROPROCESO & ""
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)

        Dim i As Integer

        For i = 0 To dt1.Rows.Count - 1
            ComboBox10.Items.Add(dt1.Rows(i).Item("DESCRIPCION"))
        Next
        Dim c11 As String
        Dim dt11 As New DataTable


        c11 = "select DESCRIPCION from [dbo].[MOTIVOS_LLAMADA]"
        Dim adaptadorsql11 As New SqlDataAdapter(c11, New SqlConnection(coneccion))
        adaptadorsql11.Fill(dt11)
        Dim h As Integer

        For h = 0 To dt11.Rows.Count - 1
            ComboBox8.Items.Add(dt11.Rows(h).Item("DESCRIPCION"))
        Next

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()
        ComboBox4.DataBindings.Clear()
        ComboBox5.DataBindings.Clear()
        ComboBox6.DataBindings.Clear()
        ComboBox7.DataBindings.Clear()
        ComboBox8.DataBindings.Clear()
        ComboBox9.DataBindings.Clear()
        ComboBox10.DataBindings.Clear()
        ComboBox11.DataBindings.Clear()
        ComboBox12.DataBindings.Clear()
        ComboBox13.DataBindings.Clear()
        ComboBox14.DataBindings.Clear()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""

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


        If TextBox10.Text <> "" Then
            Dim c1 As String
            Dim dt1 As New DataTable
            Dim NOVO As Double = TextBox10.Text

            c1 = "SELECT * FROM CONSOLIDADO_MONITOREOS WHERE REGISTRO_NOVO=" & NOVO & ""
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)

            If adaptadorsql1.Fill(dt1) <> 0 Then
                TextBox1.DataBindings.Add(New Binding("Text", dt1, "CUENTA"))
                ComboBox1.DataBindings.Add(New Binding("Text", dt1, "TIPO_CLIENTE"))
                ComboBox2.DataBindings.Add(New Binding("Text", dt1, "PROCESO"))
                ComboBox3.DataBindings.Add(New Binding("Text", dt1, "FRENTE"))
                TextBox2.DataBindings.Add(New Binding("Text", dt1, "CEDULA_ASESOR"))

                TextBox3.DataBindings.Clear()
                ComboBox4.DataBindings.Clear()
                ComboBox14.DataBindings.Clear()

                TextBox3.Text = ""
                ComboBox4.Text = ""
                ComboBox14.Text = ""

                TextBox3.DataBindings.Add(New Binding("Text", dt1, "NOMBRE_ASESOR"))
                ComboBox4.DataBindings.Add(New Binding("Text", dt1, "ALIADO"))
                ComboBox14.DataBindings.Add(New Binding("Text", dt1, "OPERACION"))
                ComboBox5.DataBindings.Add(New Binding("Text", dt1, "DIVISION"))
                TextBox4.DataBindings.Add(New Binding("Text", dt1, "MARCACION_RR"))
                TextBox5.DataBindings.Add(New Binding("Text", dt1, "MARCACION_CORRECTA"))
                ComboBox6.DataBindings.Add(New Binding("Text", dt1, "NIVEL_SE"))
                ComboBox7.DataBindings.Add(New Binding("Text", dt1, "RESULTADO"))
                ComboBox8.DataBindings.Add(New Binding("Text", dt1, "MOTIVO_LLAMADA"))
                ComboBox9.DataBindings.Add(New Binding("Text", dt1, "RAZON_LLAMADA"))
                ComboBox10.DataBindings.Add(New Binding("Text", dt1, "INCUMPLIMIENTO_DEL_PROCESO"))
                ComboBox11.DataBindings.Add(New Binding("Text", dt1, "ESPECIFICACION"))
                ComboBox12.DataBindings.Add(New Binding("Text", dt1, "SUB_ITEM"))
                ComboBox13.DataBindings.Add(New Binding("Text", dt1, "FASE_AFECTADA"))
                TextBox6.DataBindings.Add(New Binding("Text", dt1, "ETIQUETA_LLAMADA"))
                TextBox7.DataBindings.Add(New Binding("Text", dt1, "NOTA_FINAL"))
                TextBox8.DataBindings.Add(New Binding("Text", dt1, "OBSERCACIONES"))
                TextBox9.DataBindings.Add(New Binding("Text", dt1, "NOVO"))
            Else

            End If
        Else
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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

                c2 = "INSERT INTO HALLAZGOS_ALTO_VALOR_VIP(USUARIO,FECHA_HALLAZGO,AÑO_HALLAZGO,MES_HALLAZGO,DIA_HALLAZGO,CUENTA,TIPO_CLIENTE,GESTOR_REPORTA,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & nombreusuario & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
                adaptadorsql2.Fill(dt2)


                c3 = "INSERT INTO MONITOREOS_ALTO_VALOR_VIP(USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                adaptadorsql3.Fill(dt3)

                c4 = "INSERT INTO CONSOLIDADO_MONITOREOS (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                adaptadorsql4.Fill(dt4)

                c5 = "INSERT INTO CONSOLIDADO_HALLAZGOS (USUARIO,FECHA_HALLAZGO,AÑO_HALLAZGO,MES_HALLAZGO,DIA_HALLAZGO,CUENTA,TIPO_CLIENTE,GESTOR_REPORTA,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & nombreusuario & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "'," & REGISTRO_NOVO & ")"
                Dim adaptadorsql5 As New SqlDataAdapter(c5, New SqlConnection(coneccion))
                adaptadorsql5.Fill(dt5)


                c6 = "INSERT INTO PRODUCTIVIDAD (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
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
            Else
                If RESULTADO = "MONITOREO" Then

                    c3 = "INSERT INTO MONITOREOS_ALTO_VALOR_VIP(USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                    Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                    adaptadorsql3.Fill(dt3)

                    c4 = "INSERT INTO CONSOLIDADO_MONITOREOS (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
                    Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                    adaptadorsql4.Fill(dt4)

                    c6 = "INSERT INTO PRODUCTIVIDAD (USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,CUENTA,TIPO_CLIENTE,PROCESO,FRENTE,CEDULA_ASESOR,NOMBRE_ASESOR,ALIADO,OPERACION,DIVISION,RESULTADO,MOTIVO_LLAMADA,RAZON_LLAMADA,INCUMPLIMIENTO_DEL_PROCESO,ESPECIFICACION,SUB_ITEM,OBSERCACIONES,FASE_AFECTADA,NIVEL_SE,MARCACION_RR,MARCACION_CORRECTA,ETIQUETA_LLAMADA,NOTA_FINAL,NOVO,REGISTRO_NOVO)VALUES('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & ", " & DIA & "," & CUENTA & ",'" & TIPO_CLIENTE & "','" & PROCESO & "','" & FRENTE & "'," & CEDULA_ASESOR & ",'" & NOMBRE & "','" & ALIADO & "','" & OPERACION & "','" & DIVISION & "','" & RESULTADO & "','" & MOTIVO_LLAMADA & "','" & RAZON_LLAMADA & "','" & INCUMPLIMIENTO_PROCESO & "','" & ESPECIFICACION & "','" & SUB_ITEM & "','" & OBSERVACIONES & "','" & FASE_AFECTADA & "','" & NIVEL_SE & "','" & MARCACION_RR & "','" & MARCACION_CORRECTA & "','" & ETIQUETA & "','" & NOTA_FINAL & "','" & NOVO & "'," & REGISTRO_NOVO & ")"
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
                End If

            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim RESULTADO As String = ComboBox7.Text
        Dim REGISTRO_NOVO As Double = TextBox10.Text

        Dim c2, c3, c4, c5, c6 As String
        Dim dt2, dt3, dt4, dt5, dt6 As New DataTable

        If (MsgBox("Desea eliminar el registro?", MsgBoxStyle.YesNo, "Eliminar")) = MsgBoxResult.Yes Then

            If RESULTADO <> "MONITOREO" Then

                c2 = "DELETE HALLAZGOS_ALTO_VALOR_VIP WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
                adaptadorsql2.Fill(dt2)


                c3 = "DELETE MONITOREOS_ALTO_VALOR_VIP WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                adaptadorsql3.Fill(dt3)

                c4 = "DELETE CONSOLIDADO_MONITOREOS WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                adaptadorsql4.Fill(dt4)

                c5 = "DELETE CONSOLIDADO_HALLAZGOS WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                Dim adaptadorsql5 As New SqlDataAdapter(c5, New SqlConnection(coneccion))
                adaptadorsql5.Fill(dt5)

                c6 = "DELETE PRODUCTIVIDAD WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                Dim adaptadorsql6 As New SqlDataAdapter(c6, New SqlConnection(coneccion))
                adaptadorsql6.Fill(dt6)

                Call Eliminar()

                MsgBox("Registro Eliminado")

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
            Else
                If RESULTADO = "MONITOREO" Then

                    c3 = "DELETE MONITOREOS_ALTO_VALOR_VIP WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                    Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                    adaptadorsql3.Fill(dt3)

                    c4 = "DELETE CONSOLIDADO_MONITOREOS WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                    Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                    adaptadorsql4.Fill(dt4)

                    c6 = "DELETE PRODUCTIVIDAD WHERE REGISTRO_NOVO =" & REGISTRO_NOVO & ""
                    Dim adaptadorsql6 As New SqlDataAdapter(c6, New SqlConnection(coneccion))
                    adaptadorsql6.Fill(dt6)

                    Call Eliminar()

                    MsgBox("Registro Eliminado")

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
        If (MsgBox("Desea actualizar el registro?", MsgBoxStyle.YesNo, "Actualizar")) = MsgBoxResult.Yes Then

            If RESULTADO <> "MONITOREO" Then

                c2 = "UPDATE HALLAZGOS_ALTO_VALOR_VIP SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
                adaptadorsql2.Fill(dt2)


                c3 = "UPDATE MONITOREOS_ALTO_VALOR_VIP SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "',ETIQUETA_LLAMADA='" & ETIQUETA & "',NOTA_FINAL=" & NOTA_FINAL & ",NOVO='" & NOVO & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                adaptadorsql3.Fill(dt3)

                c4 = "UPDATE CONSOLIDADO_MONITOREOS SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "',ETIQUETA_LLAMADA='" & ETIQUETA & "',NOTA_FINAL=" & NOTA_FINAL & ",NOVO='" & NOVO & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                adaptadorsql4.Fill(dt4)

                c5 = "UPDATE CONSOLIDADO_HALLAZGOS SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                Dim adaptadorsql5 As New SqlDataAdapter(c5, New SqlConnection(coneccion))
                adaptadorsql5.Fill(dt5)


                c6 = "UPDATE PRODUCTIVIDAD SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "',ETIQUETA_LLAMADA='" & ETIQUETA & "',NOTA_FINAL=" & NOTA_FINAL & ",NOVO='" & NOVO & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                Dim adaptadorsql6 As New SqlDataAdapter(c6, New SqlConnection(coneccion))
                adaptadorsql6.Fill(dt6)

                Call Actualizar()

                MsgBox("Registro Actualizado")

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
            Else
                If RESULTADO = "MONITOREO" Then

                    c3 = "UPDATE MONITOREOS_ALTO_VALOR_VIP SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "',ETIQUETA_LLAMADA='" & ETIQUETA & "',NOTA_FINAL=" & NOTA_FINAL & ",NOVO='" & NOVO & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                    Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
                    adaptadorsql3.Fill(dt3)

                    c4 = "UPDATE CONSOLIDADO_MONITOREOS SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "',ETIQUETA_LLAMADA='" & ETIQUETA & "',NOTA_FINAL=" & NOTA_FINAL & ",NOVO='" & NOVO & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                    Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
                    adaptadorsql4.Fill(dt4)

                    c6 = "UPDATE PRODUCTIVIDAD SET CUENTA=" & CUENTA & ", TIPO_CLIENTE='" & TIPO_CLIENTE & "', FRENTE='" & FRENTE & "',CEDULA_ASESOR=" & CEDULA & ",NOMBRE_ASESOR='" & NOMBRE & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',DIVISION='" & DIVISION & "',RESULTADO='" & RESULTADO & "',MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',RAZON_LLAMADA='" & RAZON_LLAMADA & "',INCUMPLIMIENTO_DEL_PROCESO='" & INCUMPLIMIENTO_PROCESO & "',ESPECIFICACION='" & ESPECIFICACION & "',SUB_ITEM ='" & SUB_ITEM & "',OBSERCACIONES='" & OBSERVACIONES & "',FASE_AFECTADA='" & FASE_AFECTADA & "',NIVEL_SE='" & NIVEL_SE & " ',MARCACION_RR='" & MARCACION_RR & "',MARCACION_CORRECTA='" & MARCACION_CORRECTA & "',ETIQUETA_LLAMADA='" & ETIQUETA & "',NOTA_FINAL=" & NOTA_FINAL & ",NOVO='" & NOVO & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
                    Dim adaptadorsql6 As New SqlDataAdapter(c6, New SqlConnection(coneccion))
                    adaptadorsql6.Fill(dt6)

                    Call Actualizar()

                    MsgBox("Registro Actualizado")

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
                End If

            End If
        Else
            Exit Sub
        End If
    End Sub

    Public Sub Actualizar()
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox10.Text
        Dim transaccion As String = "Actualizacion"
        Dim ejecutor As String = nombreusuario
        Dim proceso_realizado As String = "VIP"

        c1 = "INSERT INTO SEGURIDAD_BD(REGISTRO_NOVO,TRANSACCION,PROCESO,EJECUTOR,FECHA)VALUES(" & REGISTRO_NOVO & ",'" & transaccion & "','" & proceso_realizado & "','" & ejecutor & "',getdate())"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
    End Sub
    Public Sub Eliminar()
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox10.Text
        Dim transaccion As String = "Eliminacion"
        Dim ejecutor As String = nombreusuario
        Dim proceso_realizado As String = "VIP"

        c1 = "INSERT INTO SEGURIDAD_BD(REGISTRO_NOVO,TRANSACCION,PROCESO,EJECUTOR,FECHA)VALUES(" & REGISTRO_NOVO & ",'" & transaccion & "','" & proceso_realizado & "','" & ejecutor & "',getdate())"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
    End Sub

End Class