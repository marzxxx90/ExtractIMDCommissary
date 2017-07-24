Imports Microsoft.Office.Interop

Public Class frmExtractIMD
    Private path As String
    Private fillData As String, mySql As String
    Private SavePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Private HeaderCollection As String()
    Private tmpTableName As New TextBox
    Private tmp As String

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim ds As DataSet = LoadSQL(txtQuery.Text)

        For Each tbl As DataTable In ds.Tables

            fillData = tbl.TableName
            mysql = "SELECT * FROM " & fillData
            Dim MaxDS As Integer = 0, MaxRate As Integer = 0
            Dim i As Integer = 0

            For clm As Integer = 0 To ds.Tables(fillData).Columns.Count - 1

                AddaColumn(ds.Tables(fillData).Columns.Item(clm).ToString)
            Next

            For Each dr As DataRow In ds.Tables(fillData).Rows
                Dim dsNewRow As DataRow
                dsNewRow = ds.Tables(fillData).NewRow

                Dim lv As ListViewItem = lvIMD.Items.Add(IIf(IsDBNull(ds.Tables(fillData).Rows(i).Item(0)), " ", ds.Tables(fillData).Rows(i).Item(0)))
                For setColumn As Integer = 1 To ds.Tables(fillData).Columns.Count - 1
                    Console.WriteLine("Column " & ds.Tables(fillData).Rows(i).Item(setColumn))
                    lv.SubItems.Add(IIf(IsDBNull(ds.Tables(fillData).Rows(i).Item(setColumn)), " ", ds.Tables(fillData).Rows(i).Item(setColumn)))
                Next
                Application.DoEvents()
                i += 1
            Next
        Next
    End Sub

    Private Sub AddaColumn(ByRef ColumnString As String)
        Dim NewCH As New ColumnHeader

        NewCH.Text = ColumnString
        lvIMD.Columns.Add(NewCH)
        tmpTableName.AppendText(ColumnString & " ")
    End Sub

    Private Sub InsertArrayElement(Of T)( _
          ByRef sourceArray() As T, _
          ByVal insertIndex As Integer)

        Dim newPosition As Integer
        Dim counter As Integer

        newPosition = insertIndex
        If (newPosition < 0) Then newPosition = 0
        If (newPosition > sourceArray.Length) Then _
           newPosition = sourceArray.Length

        Array.Resize(sourceArray, sourceArray.Length + 1)

        For counter = sourceArray.Length - 2 To newPosition Step 0
            sourceArray(counter + 1) = sourceArray(counter)
        Next counter

        'sourceArray(newPosition) = newValue
    End Sub

    Private Sub btnExtract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtract.Click
        tmp = tmpTableName.Text.TrimEnd

        Dim tmpCount() As String = tmp.Split(CChar(" "))
        path = SavePath & "\" & "CommissaryIMD" & Date.Now.ToString("yyyyMMdd")
        ExtractToExcell(tmpCount, txtQuery.Text, path)

    End Sub


    Friend Sub ExtractToExcell(ByVal headers As String(), ByVal mySql As String, ByVal dest As String)
        Try
            If dest = "" Then Exit Sub
            Dim ds As DataSet = LoadSQL(mySql)
            pbProgress.Maximum = ds.Tables(0).Rows.Count

            'Load Excel
            Dim oXL As New Excel.Application
            If oXL Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")
                Return
            End If

            Dim oWB As Excel.Workbook
            Dim oSheet As Excel.Worksheet

            oXL = CreateObject("Excel.Application")
            oXL.Visible = False

            oWB = oXL.Workbooks.Add
            oSheet = oWB.ActiveSheet
            'oSheet.Name = GetOption("BranchCode") & "_DAILY"

            ' ADD BRANCHCODE
            InsertArrayElement(headers, 0)

            ' HEADERS
            Dim cnt As Integer = 0
            For Each hr In headers
                cnt += 1 : oSheet.Cells(1, cnt).value = hr
            Next

            ' EXTRACTING
            Console.Write("Extracting")
            Dim rowCnt As Integer = 2
            For Each dr As DataRow In ds.Tables(0).Rows
                For colCnt As Integer = 0 To headers.Count - 2
                    oSheet.Cells(rowCnt, colCnt + 1).value = dr(colCnt)
                Next
                pbProgress.Value = pbProgress.Value + 1
                rowCnt += 1

                Console.Write(".")
                Application.DoEvents()
            Next

            oWB.SaveAs(dest)
            oSheet = Nothing
            oWB.Close(False)
            oWB = Nothing
            oXL.Quit()
            oXL = Nothing

            pbProgress.Value = 0
            MsgBox("Extracting Done!", MsgBoxStyle.Information, "Data")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
