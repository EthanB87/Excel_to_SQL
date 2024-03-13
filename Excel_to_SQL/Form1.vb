Imports System.Data.SqlClient
Imports System.IO
Imports ExcelDataReader

Public Class Form1

    Dim tables As DataTableCollection

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using opnFileDlg As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"}
            If opnFileDlg.ShowDialog() = DialogResult.OK Then
                TextBox1.Text = opnFileDlg.FileName
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)
                Using stream = File.Open(opnFileDlg.FileName, FileMode.Open, FileAccess.Read)
                    Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                        Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                             .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                             .UseHeaderRow = True}})
                        tables = result.Tables
                        ComboBox1.Items.Clear()

                        For Each table As DataTable In tables
                            ComboBox1.Items.Add(table.TableName)
                        Next
                    End Using
                End Using
            End If
        End Using
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim dataTable As DataTable = tables(ComboBox1.SelectedItem.ToString())
        DataGridView1.DataSource = dataTable
        If dataTable IsNot Nothing Then
            Dim list As List(Of Dictionary(Of String, Object)) = New List(Of Dictionary(Of String, Object))()
            For i As Integer = 0 To dataTable.Rows.Count - 1
                Dim dict As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()
                For Each column As DataColumn In dataTable.Columns
                    Dim value As Object = dataTable.Rows(i)(column)
                    If TypeOf value Is DBNull Then
                        dict.Add(column.ColumnName, Nothing)
                    ElseIf TypeOf value Is Double Then
                        dict.Add(column.ColumnName, Convert.ToDouble(value))
                    ElseIf TypeOf value Is DateTime Then
                        dict.Add(column.ColumnName, Convert.ToDateTime(value))
                    Else
                        dict.Add(column.ColumnName, value)
                    End If
                Next
                list.Add(dict)
            Next
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = dataTable
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim list As List(Of Dictionary(Of String, Object)) = New List(Of Dictionary(Of String, Object))()

            Dim dataTable As DataTable = tables(ComboBox1.SelectedItem.ToString())
            If dataTable IsNot Nothing Then
                For i As Integer = 0 To dataTable.Rows.Count - 1
                    Dim rowDict As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()
                    For Each column As DataColumn In dataTable.Columns
                        rowDict.Add(column.ColumnName, dataTable.Rows(i)(column))
                    Next
                    list.Add(rowDict)
                Next
                Dim sql As String = "INSERT INTO (" & String.Join(",", list(0).Keys) & ") VALUES "
                Dim valueParams As New List(Of String)()
                For Each item As Dictionary(Of String, Object) In list
                    Dim values As New List(Of String)()
                    For Each value In item.Values
                        values.Add($"'{value}'")
                    Next
                    valueParams.Add("(" & String.Join(",", values) & ")")
                Next
                sql &= String.Join(",", valueParams)

                Using dataBase As IDbConnection = New SqlConnection("Your SQL connection string here")
                    dataBase.Open()
                    Using command As IDbCommand = dataBase.CreateCommand()
                        command.CommandText = sql
                        command.ExecuteNonQuery()
                    End Using
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
