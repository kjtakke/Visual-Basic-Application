Public Class Form1
    Public csvImport As New ImportCsvFile
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        csvImport.CsfToArray()

    End Sub
End Class