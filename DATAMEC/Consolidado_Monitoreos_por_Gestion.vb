﻿Imports System.Data.SqlClient

Public Class Consolidado_Monitoreos_por_Gestion
    Dim bs As New BindingSource
    Dim adaptadorsql As SqlDataAdapter



    Private Sub Form7_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.MinDate = New DateTime(1985, 6, 12)
        DateTimePicker1.MaxDate = DateTime.Today
        DateTimePicker1.Format = DateTimePickerFormat.Short
        DateTimePicker1.CustomFormat = "MMMM dd, yyyy - dddd"

        DateTimePicker2.MinDate = New DateTime(1985, 6, 12)
        DateTimePicker2.MaxDate = DateTime.Today
        DateTimePicker2.Format = DateTimePickerFormat.Short
        DateTimePicker2.CustomFormat = "MMMM dd, yyyy - dddd"

        Dim gestion As String
        Dim dttable111 As New DataTable

        gestion = "SELECT PROCESO FROM CONSOLIDADO_MONITOREOS GROUP BY PROCESO"
        Dim adaptadorsql As New SqlDataAdapter(gestion, New SqlConnection(coneccion))
        adaptadorsql.Fill(dttable111)

        Dim i As Integer

        For i = 0 To dttable111.Rows.Count - 1
            CheckedListBox1.Items.Add(dttable111.Rows(i).Item("PROCESO"))
        Next


    End Sub
    Private Sub cargar_registros(ByVal sql As String, ByVal dv As DataGridView)
        Dim fechadesde As String = DateTimePicker1.Value
        Dim fechahasts As String = DateTimePicker2.Value
        Dim ges As String = CheckedListBox1.SelectedItem


        Dim consulta As String = "SELECT * FROM CONSOLIDADO_MONITOREOS WHERE PROCESO ='" & ges & "' AND FECHA_GESTION between '" & fechadesde & "' and  '" & fechahasts & "'"
        Dim dttable As New DataTable
        Dim bedit As Boolean
        adaptadorsql = New SqlDataAdapter(consulta, New SqlConnection(coneccion))
        Dim sqlbuilder As New SqlCommandBuilder(adaptadorsql)
        adaptadorsql.Fill(dttable)
        bs.DataSource = dttable

        If Not bs.DataSource Is Nothing Then
            adaptadorsql.Update(bs.DataSource) 'CType(bs.DataSource, DataTable))

        End If

        With DataGridView1
            .Refresh()
        End With
        bedit = False

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        pBar1.Value = 0.0
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        Try
            'Añadimos el Libro al programa, y la hoja al libro
            exLibro = exApp.Workbooks.Add
            exHoja = exLibro.Worksheets.Add()
            ' ¿Cuantas columnas y cuantas filas? 
            Dim NCol As Integer = ElGrid.ColumnCount
            Dim NRow As Integer = ElGrid.RowCount
            'Aqui recorremos todas las filas, y por cada fila todas las columnas y vamos 
            For i As Integer = 1 To NCol
                exHoja.Cells.Item(1, i) = ElGrid.Columns(i - 1).Name.ToString
                exHoja.Cells.Item(1, i).HorizontalAlignment = 3

            Next
            For Fila As Integer = 0 To NRow - 1
                For Col As Integer = 0 To NCol - 1
                    exHoja.Cells.Item(Fila + 2, Col + 1) = ElGrid.Rows(Fila).Cells(Col).Value
                    DataGridView1.Visible = False
                    pBar1.Visible = True
                    Label8.Visible = True
                    pBar1.Value = Fila
                    pBar1.Maximum = pBar1.Value + 25
                Next
            Next
            'Titulo en negrita, Alineado al centro y que el tamaño de la columna se ajuste 
            exHoja.Rows.Item(1).Font.Bold = 1
            exHoja.Rows.Item(1).HorizontalAlignment = 3
            exHoja.Columns.AutoFit()
            'Aplicación visible
            exApp.Application.Visible = True
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")
            Return False

        End Try
        pBar1.Hide()
        Label8.Visible = False
        DataGridView1.Visible = True
        MsgBox("La exportación a Excel fue exitosa!")
        Return True
    End Function

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        DataGridView1.Show()
        DataGridView1.DataSource = bs
        cargar_registros(coneccion, DataGridView1)

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Call GridAExcel(DataGridView1)
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        limpiar(Me)
        DataGridView1.Hide()

    End Sub
End Class