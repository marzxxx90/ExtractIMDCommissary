Imports System.Data.SqlClient

''' <summary>
''' This module declare the connection of the database.
''' 
''' </summary>
''' <remarks></remarks>
Friend Module database
    Public con As SqlConnection
    Public ReaderCon As SqlConnection
    Friend fbDataSet As New DataSet
    Friend conStr As String = String.Empty

    Private language() As String = _
        {"Connection error failed."} 'verification if the database is connected.

    Public Sub dbOpen()
        conStr = "Data Source=MISLAMION-PC;Initial Catalog=Commissary;Integrated Security=True"
        con = New SqlConnection(conStr)
        Try
            con.Open()
        Catch ex As Exception
            con.Dispose()
            MsgBox(language(0) + vbCrLf + ex.Message.ToString, vbCritical, "Connecting Error")
            Log_Report(ex.Message.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub dbClose()
        con.Close()
    End Sub

    Friend Function isReady() As Boolean
        Dim ready As Boolean = False
        Try
            dbOpen()
            ready = True
        Catch ex As Exception
            Console.WriteLine("[ERROR] " & ex.Message.ToString)
            Return False
        End Try

        Return ready
    End Function

    Friend Function SaveEntry(ByVal dsEntry As DataSet, Optional ByVal isNew As Boolean = True) As Boolean
        If dsEntry Is Nothing Then
            Return False
        End If

        dbOpen()

        Dim da As SqlDataAdapter
        Dim ds As New DataSet, mySql As String, fillData As String
        ds = dsEntry

        'Save all tables in the dataset
        For Each dsTable As DataTable In dsEntry.Tables
            fillData = dsTable.TableName
            mySql = "SELECT * FROM " & fillData
            If Not isNew Then
                Dim colName As String = dsTable.Columns(0).ColumnName
                Dim idx As Integer = dsTable.Rows(0).Item(0)
                mySql &= String.Format(" WHERE {0} = {1}", colName, idx)

                Console.WriteLine("ModifySQL: " & mySql)
            End If

            da = New SqlDataAdapter(mySql, con)
            Dim cb As New SqlCommandBuilder(da) 'Required in Saving/Update to Database
            da.Update(ds, fillData)
        Next

        dbClose()
        Return True
    End Function

    Friend Sub SQLCommand(ByVal sql As String)
        conStr = "Data Source=MISLAMION-PC;Initial Catalog=Commissary;Integrated Security=True"
        con = New SqlConnection(conStr)

        Dim cmd As SqlCommand
        cmd = New SqlCommand(sql, con)

        Try
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
            Log_Report(String.Format("[{0}] - ", sql) & ex.ToString)
            con.Dispose()
            Exit Sub
        End Try

        System.Threading.Thread.Sleep(1000)
    End Sub

    Friend Function LoadSQL(ByVal mySql As String, Optional ByVal tblName As String = "QuickSQL") As DataSet
        dbOpen() 'open the database.

        Dim da As SqlDataAdapter
        Dim ds As New DataSet, fillData As String = tblName
        Try
            da = New SqlDataAdapter(mySql, con)
            da.Fill(ds, fillData)
        Catch ex As Exception
            Console.WriteLine(">>>>>" & mySql)
            MsgBox(ex.ToString)
            Log_Report("LoadSQL - " & ex.ToString)
            ds = Nothing
        End Try

        dbClose()

        Return ds
    End Function

    Friend Function LoadSQL_byDataReader(ByVal mySql As String) As SqlDataReader
        Dim myCom As SqlCommand = New SqlCommand(mySql, ReaderCon)
        Dim reader As SqlDataReader = myCom.ExecuteReader()

        Return reader
    End Function

    Public Sub dbReaderOpen()
        conStr = "Data Source=MISLAMION-PC;Initial Catalog=Commissary;Integrated Security=True"

        ReaderCon = New SqlConnection(conStr)
        Try
            ReaderCon.Open() 'open the database.
        Catch ex As Exception
            ReaderCon.Dispose()
            MsgBox(language(0) + vbCrLf + ex.Message.ToString, vbCritical, "Connecting Error")
            Log_Report(ex.Message.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub dbReaderClose()
        ReaderCon.Close()
    End Sub

End Module
