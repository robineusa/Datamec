Imports System.Data.SqlClient

Public Class Formato_Correo

    Private Sub Formato_Correo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label25.Text = nombreusuario
        TextBox1.Text = NOTA_ASESOR
        TextBox2.Text = NOMBRE_ASESOR
        TextBox3.Text = ALIADO_ASESOR
        TextBox4.Text = OPERACION_ASESOR
        TextBox5.Text = CUENTA_CORREO
        RichTextBox5.Text = ETIQUETA_LLAMADA
        RichTextBox4.Text = OBSERVACIONES_ASESOR
        TextBox6.Text = REGISTRO_NOVO_GENERAL

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click



        If TextBox6.Text = 0 Then
            MsgBox("Registro novo no pueder quedar en valor 0")
            TextBox6.BackColor = Color.Red
            TextBox6.Focus()
        Else
            Dim c1, c2 As String
                Dim dt1, dt2 As New DataTable

                Dim USUARIOMEC As String = ""
                Dim FECHA As String = Today
                Dim AÑO As Integer = Now.Year
                Dim MES As Integer = Now.Month
                Dim DIA As Integer = Now.Day
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
                Dim REGISTRO_NOVO As Double = TextBox6.Text


                c1 = "Select Usuario from USUARIOS where Nombre_Usuario like '" & nombreusuario & "'"
                Dim adaptadorsql1 As New SqlDataAdapter(c1, New SqlConnection(coneccion))
                adaptadorsql1.Fill(dt1)
                Dim i As Integer

                For i = 0 To dt1.Rows.Count - 1
                    USUARIOMEC = dt1.Rows(i).Item("Usuario")
                Next

                If (MsgBox("Desea Guardar Registro?", MsgBoxStyle.YesNo, "Guardar")) = MsgBoxResult.Yes Then
                    c2 = "INSERT INTO ALARMAS_CORREOS ( USUARIO,FECHA_GESTION,AÑO_GESTION,MES_GESTION,DIA_GESTION,PROCESO,FRENTE,IMPACTO,RESULTADO_ACOMPAÑAMIENTO,NOMBRE_ASESOR,ALIADO,OPERACION,CUENTA,MOTIVO_LLAMADA,OPORTUNIDAD_MEJORA,RECOMENDACIONES,OBSERCACIONES,UBICACION_LLAMADA,REGISTRO_NOVO)VALUES ('" & USUARIOMEC & "','" & FECHA & "'," & AÑO & "," & MES & "," & DIA & ",'" & PROCESO_GESTION & "','" & FRENTE_GESTION & "','" & IMPACTO & "','" & ACOMPAÑAMIENTO & "','" & NOMBRE_ASESORR & "','" & ALIADO & "','" & OPERACION & "','" & CUENTA & "','" & MOTIVO_LLAMADA & "','" & OPORTUNIDAD_MEJORA & "','" & RECOMENDACIONES & "','" & OBSERCACIONES & "','" & UBICACION_LLAMADA & "'," & REGISTRO_NOVO & ")"
                    Dim adaptadorsql2 As New SqlDataAdapter(c2, New SqlConnection(coneccion))
                    adaptadorsql2.Fill(dt2)
                    MsgBox("Registro Almacenado")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""

                    ComboBox1.Text = ""
                    RichTextBox1.Text = ""
                    RichTextBox2.Text = ""
                    RichTextBox3.Text = ""
                    RichTextBox4.Text = ""
                    RichTextBox5.Text = ""
                    Me.Close()
                Else
                    Exit Sub

                End If
            End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        TextBox6.BackColor = Color.WhiteSmoke
    End Sub
End Class