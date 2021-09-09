
Option Explicit On

Imports System.Data.OleDb
Imports System.Data

Public Class OleFillingData

    Public Function GetDS(ByVal oleStr As String, ByVal DSname As String, ByVal oleConn As OleDbConnection) As DataSet
        Dim mycombods As New DataSet
        Dim mycomboadapter As New OleDbDataAdapter
        Try
            mycomboadapter = New OleDbDataAdapter(oleStr, oleConn)

            mycomboadapter.Fill(mycombods, DSname)
        Catch ex As Exception
            MsgBox("Failure in [GetDS]: " & ex.Message() & " (Error Code: " & Err.Number & ").")
        Finally
            GetDS = mycombods

            mycomboadapter.Dispose()
            mycombods.Dispose()

        End Try

        Return GetDS

    End Function



    Public Function InsertData(ByVal OleStr As String, ByVal oleConn As OleDbConnection, ByVal oleTrans As OleDbTransaction) As Integer
        Dim cmdData As New OleDbCommand
        Dim iRetVal As Integer = 0

        Try
            cmdData.Connection = oleConn
            cmdData.CommandType = CommandType.Text
            cmdData.CommandText = OleStr
            cmdData.Transaction = oleTrans

            iRetVal = cmdData.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Failure in [InsertData]: " & ex.Message() & " (Error Code: " & Err.Number & ").")
        Finally
            cmdData.Dispose()
        End Try

        Return iRetVal
    End Function

    Public Function DeleteData(ByVal OleStr As String, ByVal oleConn As OleDbConnection, ByVal oleTrans As OleDbTransaction) As Integer
        Dim cmdData As New OleDbCommand
        Dim iRetVal As Integer = 0

        Try
            cmdData.Connection = oleConn
            cmdData.CommandType = CommandType.Text
            cmdData.CommandText = OleStr
            cmdData.Transaction = oleTrans

            iRetVal = cmdData.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Failure in [DeleteData]: " & ex.Message() & " (Error Code: " & Err.Number & ").")
        Finally
            cmdData.Dispose()
        End Try

        Return iRetVal
    End Function

    Public Function UpdateData(ByVal OleStr As String, ByVal oleConn As OleDbConnection, ByVal oleTrans As OleDbTransaction) As Integer
        Dim cmdData As New OleDbCommand
        Dim iRetVal As Integer = 0

        Try
            cmdData.Connection = oleConn
            cmdData.CommandType = CommandType.Text
            cmdData.CommandText = OleStr
            cmdData.Transaction = oleTrans

            iRetVal = cmdData.ExecuteNonQuery()

        Catch ex As Exception

            MsgBox("Failure in [UpdateData]: " & ex.Message() & " (Error Code: " & Err.Number & ").")

        Finally
            cmdData.Dispose()
        End Try

        Return iRetVal
    End Function

    Public Function UpdateDataSyuk(ByVal OleStr As String, ByVal oleConn As OleDbConnection) As Integer
        Dim cmdData As New OleDbCommand
        Dim iRetVal As Integer = 0

        Try
            cmdData.Connection = oleConn
            cmdData.CommandType = CommandType.Text
            cmdData.CommandText = OleStr

            iRetVal = cmdData.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Failure in [UpdateData]: " & ex.Message() & " (Error Code: " & Err.Number & ").")
        Finally
            cmdData.Dispose()
        End Try

        Return iRetVal
    End Function

    Public Function IsDataExist(ByVal OleStr As String, ByVal oleConn As OleDbConnection, ByVal oleTrans As OleDbTransaction) As Integer
        Dim oleCMD As New OleDbCommand
        Dim dr As OleDbDataReader

        Try
            oleCMD.CommandText = OleStr
            oleCMD.CommandType = CommandType.Text
            oleCMD.Connection = oleConn
            oleCMD.Transaction = oleTrans

            dr = oleCMD.ExecuteReader

            If dr.HasRows Then
                While dr.Read
                    If Not IsDBNull(dr.Item(0)) Then
                        If dr.Item(0) > 0 Then
                            IsDataExist = 1
                        Else
                            IsDataExist = 0
                        End If
                    Else
                        IsDataExist = 0
                    End If
                End While
            Else
                IsDataExist = 0
            End If

        Catch ex As Exception
            MsgBox("Failure in [IsDataExist]: " & ex.Message() & " (Error Code: " & Err.Number & ").")
        Finally
            oleCMD.Dispose()
        End Try


        Return IsDataExist
    End Function

    '' Return ONE value from a specific row and column in STRING
    Public Function Get_Return_Value(ByVal OleStr As String, ByVal oleConn As OleDbConnection, ByVal oleTrans As OleDbTransaction) As String
        Dim OleCmd As New OleDbCommand
        Dim OleReader As OleDbDataReader
        Dim RetValue As String = Nothing

        Try
            OleCmd.CommandText = OleStr
            OleCmd.CommandType = CommandType.Text
            OleCmd.Connection = oleConn
            OleCmd.Transaction = oleTrans

            OleReader = OleCmd.ExecuteReader
            If OleReader.HasRows Then
                While OleReader.Read
                    If Not IsDBNull(OleReader.Item(0)) Then
                        RetValue = OleReader.Item(0)
                    Else
                        RetValue = Nothing
                    End If
                End While
            End If
            OleReader.Close()

        Catch ex As Exception
            MsgBox("Failure in [Get_Return_Value]: " & ex.Message() & " (Error Code: " & Err.Number & ").")
        Finally
            OleCmd.Dispose()
        End Try

        Return RetValue
    End Function

End Class



