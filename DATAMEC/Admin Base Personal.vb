Imports System.Data.SqlClient

Public Class Admin_Base_Personal

    Private Sub Admin_Base_Personal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim c1, c2, c3, c4 As String
        Dim dt1, dt2, dt3, dt4 As New DataTable

        c1 = "select ALIADO FROM BD_PERSONAL_RETENCION GROUP BY ALIADO"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)

        Dim i As Integer

        For i = 0 To dt1.Rows.Count - 1
            ComboBox1.Items.Add(dt1.Rows(i).Item("ALIADO"))
        Next

        c2 = "select OPERACION FROM BD_PERSONAL_RETENCION GROUP BY OPERACION"
        Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
        adaptadorsql2.Fill(dt2)

        Dim j As Integer

        For j = 0 To dt2.Rows.Count - 1
            ComboBox2.Items.Add(dt2.Rows(j).Item("OPERACION"))
        Next

        c3 = "select FRENTE FROM BD_PERSONAL_RETENCION GROUP BY FRENTE"
        Dim adaptadorsql3 As New SqlDataAdapter(c3, New SqlConnection(coneccion))
        adaptadorsql3.Fill(dt3)

        Dim k As Integer

        For k = 0 To dt3.Rows.Count - 1
            ComboBox4.Items.Add(dt3.Rows(k).Item("FRENTE"))
        Next

        c4 = "select PROCESO FROM BD_PERSONAL_RETENCION GROUP BY PROCESO"
        Dim adaptadorsql4 As New SqlDataAdapter(c4, New SqlConnection(coneccion))
        adaptadorsql4.Fill(dt4)

        Dim l As Integer

        For l = 0 To dt4.Rows.Count - 1
            ComboBox3.Items.Add(dt4.Rows(l).Item("PROCESO"))
        Next

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()
        ComboBox4.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        TextBox2.Text = ""

        If TextBox1.Text = "" Then

        Else

            Dim c1 As String
            Dim dt1 As New DataTable
            Dim CEDULA As Double = TextBox1.Text

            c1 = "SELECT * FROM BD_PERSONAL_RETENCION WHERE CEDULA=" & CEDULA & ""
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)

            If adaptadorsql1.Fill(dt1) <> 0 Then
                ComboBox1.DataBindings.Add(New Binding("Text", dt1, "ALIADO"))
                ComboBox2.DataBindings.Add(New Binding("Text", dt1, "OPERACION"))
                ComboBox3.DataBindings.Add(New Binding("Text", dt1, "PROCESO"))
                ComboBox4.DataBindings.Add(New Binding("Text", dt1, "FRENTE"))
                TextBox2.DataBindings.Add(New Binding("Text", dt1, "NOMBRE"))
            Else

            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim CEDULA2 As Double = TextBox1.Text
        Dim ALIADO2 As String = ComboBox1.Text
        Dim OPERACION2 As String = ComboBox2.Text
        Dim PROCESO2 As String = ComboBox3.Text
        Dim FRENTE2 As String = ComboBox4.Text
        Dim NOMBRE2 As String = TextBox2.Text

        c1 = "UPDATE BD_PERSONAL_RETENCION SET NOMBRE='" & NOMBRE2 & "', ALIADO ='" & ALIADO2 & "', OPERACION='" & OPERACION2 & "', FRENTE='" & FRENTE2 & "', PROCESO='" & PROCESO2 & "' WHERE CEDULA =" & CEDULA2 & ""
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)

        Call Actualizar()

        MsgBox("Registro Actualizado")
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim CEDULA1 As Double = TextBox1.Text
        Dim ALIADO1 As String = ComboBox1.Text
        Dim OPERACION1 As String = ComboBox2.Text
        Dim PROCESO1 As String = ComboBox3.Text
        Dim FRENTE1 As String = ComboBox4.Text
        Dim NOMBRE1 As String = TextBox2.Text

        c1 = "INSERT INTO BD_PERSONAL_RETENCION(CEDULA,NOMBRE,ALIADO,OPERACION,PROCESO,FRENTE)VALUES(" & CEDULA1 & ",'" & NOMBRE1 & "','" & ALIADO1 & "','" & OPERACION1 & "','" & PROCESO1 & "','" & FRENTE1 & "')"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)

        MsgBox("Registro Almacenado")
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim CEDULA3 As Double = TextBox1.Text

        c1 = "DELETE BD_PERSONAL_RETENCION WHERE CEDULA =" & CEDULA3 & ""
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)

        Call Eliminar()

        MsgBox("Registro Eliminado")
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
    End Sub
    Public Sub Actualizar()
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox1.Text
        Dim transaccion As String = "Actualizacion"
        Dim ejecutor As String = nombreusuario
        Dim proceso_realizado As String = "Base Personal"

        c1 = "INSERT INTO SEGURIDAD_BD(REGISTRO_NOVO,TRANSACCION,PROCESO,EJECUTOR,FECHA)VALUES(" & REGISTRO_NOVO & ",'" & transaccion & "','" & proceso_realizado & "','" & ejecutor & "',getdate())"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
    End Sub
    Public Sub Eliminar()
        Dim c1 As String
        Dim dt1 As New DataTable
        Dim REGISTRO_NOVO As Double = TextBox1.Text
        Dim transaccion As String = "Eliminacion"
        Dim ejecutor As String = nombreusuario
        Dim proceso_realizado As String = "Base Personal"

        c1 = "INSERT INTO SEGURIDAD_BD(REGISTRO_NOVO,TRANSACCION,PROCESO,EJECUTOR,FECHA)VALUES(" & REGISTRO_NOVO & ",'" & transaccion & "','" & proceso_realizado & "','" & ejecutor & "',getdate())"
        Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dt1)
    End Sub

End Class