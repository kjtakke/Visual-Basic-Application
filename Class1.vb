

Public Class ImportCsvFile
    Public myCSVdata(,) As String
    Private i As Single, j As Single, k As Single
    Public Sub CsfToArray()
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Dim objFSO As Object
        Dim objTF As Object
        Dim strIn As String
        Dim ary1() As String, ary2() As String
        Dim fileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
        Else
            GoTo en
        End If

        fileName = fd.FileName & fd.DefaultExt
        objFSO = CreateObject("Scripting.FileSystemObject")
        objTF = objFSO.OpenTextFile(fileName, 1)
        strIn = objTF.readall
        objTF.Close

        ary1 = Split(strIn, vbNewLine)

        ary2 = Split(ary1(0), ",")
        ReDim myCSVdata(0 To UBound(ary1), 0 To UBound(ary2))

        For i = 0 To UBound(ary1)
            ary2 = Split(ary1(i), ",")
            For j = 0 To UBound(ary2)
                myCSVdata(i, j) = ary2(j)
            Next j
        Next i
En:
    End Sub
End Class
