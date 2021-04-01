Public Class Form1
    Public csvImport As New ImportCsvFile
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        csvImport.ImportTextFiley()
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        csvImport.ImportCSV()
        MsgBox(csvImport.CSVdata(0, 1))
    End Sub
End Class