Imports System.Data.SqlClient

Public Class Actualizacion_Correos

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()

        RichTextBox1.DataBindings.Clear()
        RichTextBox2.DataBindings.Clear()
        RichTextBox3.DataBindings.Clear()
        RichTextBox4.DataBindings.Clear()
        RichTextBox5.DataBindings.Clear()

        ComboBox1.DataBindings.Clear()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        ComboBox1.Text = ""


        If TextBox6.Text <> "" Then

            Dim c1 As String
            Dim dt1 As New DataTable
            Dim REGISTRO_NOVO As Double = TextBox6.Text

            c1 = "SELECT * FROM ALARMAS_CORREOS WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)

            If adaptadorsql1.Fill(dt1) <> 0 Then

                ComboBox1.DataBindings.Add(New Binding("Text", dt1, "IMPACTO"))
                TextBox1.DataBindings.Add(New Binding("Text", dt1, "RESULTADO_ACOMPAÑAMIENTO"))
                TextBox2.DataBindings.Add(New Binding("Text", dt1, "NOMBRE_ASESOR"))
                TextBox3.DataBindings.Add(New Binding("Text", dt1, "ALIADO"))
                TextBox4.DataBindings.Add(New Binding("Text", dt1, "OPERACION"))
                TextBox5.DataBindings.Add(New Binding("Text", dt1, "CUENTA"))
                RichTextBox1.DataBindings.Add(New Binding("Text", dt1, "MOTIVO_LLAMADA"))
                RichTextBox2.DataBindings.Add(New Binding("Text", dt1, "OPORTUNIDAD_MEJORA"))
                RichTextBox3.DataBindings.Add(New Binding("Text", dt1, "RECOMENDACIONES"))
                RichTextBox4.DataBindings.Add(New Binding("Text", dt1, "OBSERCACIONES"))
                RichTextBox5.DataBindings.Add(New Binding("Text", dt1, "UBICACION_LLAMADA"))

            Else

            End If

        Else

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox6.Text
        Dim IMPACTO As String = ComboBox1.Text
        Dim ACOMPAÑAMIENTO As Double = TextBox1.Text
        Dim NOMBRE_ASESORR As String = TextBox2.Text
        Dim ALIADO As String = TextBox3.Text
        Dim OPERACION As String = TextBox4.Text
        Dim CUENTA As Double = TextBox5.Text
        Dim MOTIVO_LLAMADA As String = RichTextBox1.Text
        Dim OPORTUNIDAD_MEJORA As String = RichTextBox2.Text
        Dim RECOMENDACIONES As String = RichTextBox3.Text
        Dim OBSERCACIONES As String = RichTextBox4.Text
        Dim UBICACION_LLAMADA As String = RichTextBox5.Text

        If (MsgBox("Desea Actualizar el  Registro?", MsgBoxStyle.YesNo, "Actualizar")) = MsgBoxResult.Yes Then

            c1 = "UPDATE ALARMAS_CORREOS SET IMPACTO='" & IMPACTO & "',RESULTADO_ACOMPAÑAMIENTO=" & ACOMPAÑAMIENTO & ",NOMBRE_ASESOR='" & NOMBRE_ASESORR & "',ALIADO='" & ALIADO & "',OPERACION='" & OPERACION & "',CUENTA=" & CUENTA & ",MOTIVO_LLAMADA='" & MOTIVO_LLAMADA & "',OPORTUNIDAD_MEJORA='" & OPORTUNIDAD_MEJORA & "',RECOMENDACIONES='" & RECOMENDACIONES & "',OBSERCACIONES='" & OBSERCACIONES & "',UBICACION_LLAMADA='" & UBICACION_LLAMADA & "' WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)

            Call Actualizar()

            MsgBox("Registro Actualizado")

            TextBox1.DataBindings.Clear()
            TextBox2.DataBindings.Clear()
            TextBox3.DataBindings.Clear()
            TextBox4.DataBindings.Clear()
            TextBox5.DataBindings.Clear()

            RichTextBox1.DataBindings.Clear()
            RichTextBox2.DataBindings.Clear()
            RichTextBox3.DataBindings.Clear()
            RichTextBox4.DataBindings.Clear()
            RichTextBox5.DataBindings.Clear()

            ComboBox1.DataBindings.Clear()

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""

            RichTextBox1.Text = ""
            RichTextBox2.Text = ""
            RichTextBox3.Text = ""
            RichTextBox4.Text = ""
            RichTextBox5.Text = ""
            ComboBox1.Text = ""
            TextBox6.Text = ""


        Else
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox6.Text
        
        If (MsgBox("Desea Eliminiar el  Registro?", MsgBoxStyle.YesNo, "Eliminar")) = MsgBoxResult.Yes Then

            c1 = "DELETE ALARMAS_CORREOS WHERE REGISTRO_NOVO=" & REGISTRO_NOVO & ""
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)

            Call Eliminar()

            MsgBox("Registro Eliminado")
            TextBox1.DataBindings.Clear()
            TextBox2.DataBindings.Clear()
            TextBox3.DataBindings.Clear()
            TextBox4.DataBindings.Clear()
            TextBox5.DataBindings.Clear()

            RichTextBox1.DataBindings.Clear()
            RichTextBox2.DataBindings.Clear()
            RichTextBox3.DataBindings.Clear()
            RichTextBox4.DataBindings.Clear()
            RichTextBox5.DataBindings.Clear()

            ComboBox1.DataBindings.Clear()

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""

            RichTextBox1.Text = ""
            RichTextBox2.Text = ""
            RichTextBox3.Text = ""
            RichTextBox4.Text = ""
            RichTextBox5.Text = ""
            ComboBox1.Text = ""
            TextBox6.Text = ""

        Else
        End If
    End Sub
    Public Sub Actualizar()
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox6.Text
        Dim transaccion As String = "Actualizacion"
        Dim ejecutor As String = nombreusuario
        Dim proceso_realizado As String = "Correos"

        c1 = "INSERT INTO SEGURIDAD_BD(REGISTRO_NOVO,TRANSACCION,PROCESO,EJECUTOR,FECHA)VALUES(" & REGISTRO_NOVO & ",'" & transaccion & "','" & proceso_realizado & "','" & ejecutor & "',getdate())"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
    End Sub
    Public Sub Eliminar()
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox6.Text
        Dim transaccion As String = "Eliminacion"
        Dim ejecutor As String = nombreusuario
        Dim proceso_realizado As String = "Correos"

        c1 = "INSERT INTO SEGURIDAD_BD(REGISTRO_NOVO,TRANSACCION,PROCESO,EJECUTOR,FECHA)VALUES(" & REGISTRO_NOVO & ",'" & transaccion & "','" & proceso_realizado & "','" & ejecutor & "',getdate())"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
    End Sub

    Private Sub Actualizacion_Correos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class