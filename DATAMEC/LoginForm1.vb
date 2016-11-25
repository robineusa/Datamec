Imports System.Data.SqlClient

Public Class LoginForm1

    ' TODO: inserte el código para realizar autenticación personalizada usando el nombre de usuario y la contraseña proporcionada 
    ' (Consulte http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' El objeto principal personalizado se puede adjuntar al objeto principal del subproceso actual como se indica a continuación: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' donde CustomPrincipal es la implementación de IPrincipal utilizada para realizar la autenticación. 
    ' Posteriormente, My.User devolverá la información de identidad encapsulada en el objeto CustomPrincipal
    ' como el nombre de usuario, nombre para mostrar, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        Try
            Dim consulta As String
            Dim Usuario As String = UsernameTextBox.Text
            Dim Clave As String = PasswordTextBox.Text
            Dim dt1 As New DataTable
            Dim nombre_gestor As String = ""
            Dim dominio As Integer

            consulta = "SELECT Nombre_Usuario,Usuario,Clave, Id_Dominio From USUARIOS Where Usuario='" & Usuario & "' and Clave='" & Clave & "'"
            Dim adaptadorSQL As New SqlDataAdapter(consulta, New SqlConnection(coneccion))
            adaptadorSQL.Fill(dt1)

            If adaptadorSQL.Fill(dt1) <> 0 Then
                Dim i As Integer

                For i = 0 To dt1.Rows.Count - 1
                    nombre_gestor = dt1.Rows(i).Item("Nombre_Usuario")
                Next

                Dim j As Integer

                For j = 0 To dt1.Rows.Count - 1
                    dominio = dt1.Rows(j).Item("Id_Dominio")
                Next
                If dominio = 1 Then
                    nombreusuario = nombre_gestor
                    MsgBox(nombreusuario, MsgBoxStyle.Information, "!Bienvenido(a)!")
                    Panel_Principal.Show()
                    Me.Visible = False
                Else
                    If dominio = 2 Then
                        nombreusuario = nombre_gestor
                        MsgBox(nombreusuario, MsgBoxStyle.Information, "!Bienvenido(a)!")
                        Panel_Principal.Show()
                        Me.Visible = False
                    Else
                        If dominio = 4 Then
                            nombreusuario = nombre_gestor
                            MsgBox(nombreusuario, MsgBoxStyle.Information, "!Bienvenido(a)!")
                            Panel_Principal.Show()
                            Me.Visible = False
                        Else
                            If dominio = 3 Then
                                nombreusuario = nombre_gestor
                                MsgBox(nombreusuario, MsgBoxStyle.Information, "!Bienvenido(a)!")
                                Panel_Principal.Show()
                                Me.Visible = False
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox("El usuario o contraseña son incorrectos")
            End If



        Catch ex As Exception When UsernameTextBox.Text = ""
            MsgBox("Debe Escribir el usuario" & ex.Message & vbCrLf & "Asegurese de ingresar un Usuario", MsgBoxStyle.Critical, "CUIDADO")
        Catch ex As Exception When PasswordTextBox.Text = ""
            MsgBox("Debe Escribir la contraseña" & ex.Message & vbCrLf & "Asegurese de ingresar se Contraseña", MsgBoxStyle.Critical, "CUIDADO")

        End Try
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
End Class
