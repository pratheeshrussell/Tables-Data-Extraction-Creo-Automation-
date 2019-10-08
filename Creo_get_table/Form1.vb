Imports pfcls
Imports System.IO
Public Class Form1
    'https://community.ptc.com/t5/Data-Exchange/Creo-VBA-API-for-drawing-table-update/td-p/397136
    'read last comment
    'https://community.ptc.com/t5/Detailing-MBD-MBE/Extracting-Dimension-from-Creo-model/td-p/217712
    Dim Casync As New pfcls.CCpfcAsyncConnection
    Dim asyncConnection As IpfcAsyncConnection
    Dim cAC As New CCpfcAsyncConnection
    Dim session As IpfcBaseSession
    Dim model1 As IpfcModel2D
    Dim name1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo a
        ConnectToCreo()
        Label1.Text = "Drawing Name: " & name1
        Dim mainmodel As IpfcModel
        For Each modelname As IpfcModel In model1.ListModels()
            mainmodel = modelname
            Label2.Text = "Model Name: " & modelname.FileName
        Next
        GetTablesList()

        DisconnectCreo()


    End Sub
    Private Sub GetTablesList()
        Dim tables As IpfcTables
        tables = model1.ListTables()
        Dim p = 0
        For Each singletable As IpfcTable In tables
            ComboBox1.Items.Add(p)
        Next
    End Sub
    Private Sub GetTablesVal()

        Dim tables As IpfcTables
        tables = model1.ListTables()
        Dim noRow
        Dim noCol
        Dim tableno = CInt(ComboBox1.Text)
        Dim singletable As IpfcTable = tables(tableno)

        noRow = singletable.GetRowCount
        noCol = singletable.GetColumnCount
        Label4.Text = "Rows: " & noRow
        Label5.Text = "Columns: " & noCol

        Dim tableCellCreate As New pfcls.CCpfcTableCell
        Dim tableCell As pfcls.IpfcTableCell
        ListView1.Items.Clear()
        For i = 1 To noRow
            For j = 1 To noCol
                Dim cellgettext
                Dim celltext = ""
                tableCell = tableCellCreate.Create(i, j)
                Try
                    cellgettext = singletable.GetText(tableCell, 0)
                    celltext = cellgettext.Item(0).ToString
                Catch ex As Exception
                    ' MsgBox(ex.Message)
                End Try

                Dim lvw As ListView = ListView1
                Dim newItem As ListViewItem
                newItem = lvw.Items.Add(i)
                newItem.SubItems.Add(j)
                newItem.SubItems.Add(celltext)


            Next
        Next
        Exit Sub

    End Sub
    Private Sub ConnectToCreo()
        asyncConnection = cAC.Connect("", "", DBNull.Value, DBNull.Value)
        session = asyncConnection.Session
        name1 = session.CurrentModel.FileName

        model1 = session.GetModel(name1, EpfcModelType.EpfcMDL_DRAWING)
    End Sub
    Private Sub DisconnectCreo()
        asyncConnection.Disconnect(1)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ConnectToCreo()
        GetTablesVal()
        DisconnectCreo()
    End Sub
End Class
