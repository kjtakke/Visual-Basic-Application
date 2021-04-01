

Public Class ImportCsvFile
    Public CSVdata(,) As String
    Private i As Single, j As Single, k As Single
    Public Sub ImportTextFiley()
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Dim fileName As String
        Dim objFSO As Object
        Dim objTF As Object
        Dim strIn As String
        Dim ary1() As String, ary2() As String


        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
        Else
            GoTo En
        End If

        fileName = fd.FileName & fd.DefaultExt
        objFSO = CreateObject("Scripting.FileSystemObject")
        objTF = objFSO.OpenTextFile(fileName, 1)
        strIn = objTF.readall
        objTF.Close

        ary1 = Split(strIn, vbNewLine)

        ary2 = Split(ary1(0), ",")
        ReDim CSVdata(0 To UBound(ary1), 0 To UBound(ary2))

        For i = 0 To UBound(ary1)
            ary2 = Split(ary1(i), ",")
            For j = 0 To UBound(ary2)
                CSVdata(i, j) = ary2(j)
            Next j
        Next i
En:
    End Sub



    Sub ImportCSV()

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Dim fileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
        Else
            GoTo En
        End If

        fileName = fd.FileName & fd.DefaultExt
        Dim rowCounter As Single = 0
        Dim columnCounter As Single = 0
        Dim totalRowCount As Single = 0

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strFileName)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    Dim currentField As String = 0
                    If rowCounter = 0 Then
                        For Each currentField In currentRow
                            totalRowCount += 1
                        Next
                        ReDim CSVdata(0 To rowCounter, 0 To totalRowCount)
                        MsgBox(totalRowCount)
                    End If
                    For Each currentField In currentRow
                        ReDim CSVdata(0 To rowCounter, 0 To totalRowCount)
                        CSVdata(rowCounter, columnCounter) = currentField
                        columnCounter += 1
                    Next
                Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                    'MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
                columnCounter = 0
                rowCounter += 1
            End While
        End Using
En:
    End Sub
End Class
