Imports System.Data.SqlClient

Public Class Reevaluaciones

    Private Sub Reevaluaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim consulta1 As String
        Dim dttable1 As New DataTable
        Dim dom As Integer = 2

        consulta1 = "select Nombre_Usuario from USUARIOS where Id_Dominio =" & dom & " order by Nombre_Usuario"
        Dim adaptadorsql1 As New SqlDataAdapter(consulta1, New SqlConnection(coneccion))
        adaptadorsql1.Fill(dttable1)

        Dim i As Integer

        For i = 0 To dttable1.Rows.Count - 1
            ComboBox1.Items.Add(dttable1.Rows(i).Item("Nombre_Usuario"))
        Next
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox2.Text = ""
        TextBox2.DataBindings.Clear()
        TextBox3.Text = ""
        ComboBox1.Text = ""
        ComboBox1.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        If TextBox1.Text <> "" Then

            Dim c1, c2 As String
            Dim dt1, dt2 As New DataTable
            Dim ID As Double = TextBox1.Text


            c1 = "SELECT * FROM ALARMAS_CORREOS WHERE REGISTRO_NOVO=" & ID & ""
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)

            If adaptadorsql1.Fill(dt1) <> 0 Then
                TextBox2.DataBindings.Add(New Binding("Text", dt1, "PROCESO"))
                TextBox3.DataBindings.Add(New Binding("Text", dt1, "USUARIO"))
            Else
            End If
            Dim USU As String = TextBox3.Text
            c2 = "SELECT Nombre_Usuario FROM USUARIOS WHERE USUARIO LIKE'" & USU & "'"
            Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
            adaptadorsql2.Fill(dt2)
            ComboBox1.DataBindings.Add(New Binding("Text", dt2, "Nombre_Usuario"))
        Else
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim c1, c2 As String
        Dim dt1, dt2 As New DataTable
        Dim USUARIO As String = ""
        Dim ID_REGISTRO As Integer = TextBox1.Text
        Dim PROCESO As String = TextBox2.Text
        Dim ESTADO As String = "Pendiente"
        Dim nameusu As String = ComboBox1.Text

        c2 = "Select Usuario from USUARIOS where Nombre_Usuario like '" & nameusu & "'"
        Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
        adaptadorsql2.Fill(dt2)

        Dim i As Integer

        For i = 0 To dt2.Rows.Count - 1
            USUARIO = dt2.Rows(i).Item("Usuario")
        Next
        If (MsgBox("Desea Cargar la Reevaluación?", MsgBoxStyle.YesNo, "Guardar")) = MsgBoxResult.Yes Then
            c1 = "INSERT INTO REEVALUACIONES (USUARIO,FECHA_REEVALUACION,ID_REGISTRO,PROCESO,ESTADO)VALUES('" & USUARIO & "',GETDATE()," & ID_REGISTRO & ",'" & PROCESO & "', '" & ESTADO & "')"
            Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
            adaptadorsql1.Fill(dt1)
            MsgBox("Reevaluacion cargada al usuario")
            ComboBox1.Text = ""
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
        Else
            Exit Sub
        End If

    End Sub
End Class