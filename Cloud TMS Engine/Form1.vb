Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO

Public Class Form1
    Public connStrAcc As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" + Application.StartupPath + "\TMSEngine.mdb"
    Dim bRetBol As Boolean = False
    Dim txtFile As String
    Dim dir As String
    Dim entry As String()
    Dim fileName As String
    Dim sr As IO.StreamReader
    Dim fileContents As New System.Text.StringBuilder()
    Dim arr As String() = {""}
    Dim arr2(999999) As DateTime
    Dim read As String
    Dim info As String
    Dim newFileName As String
    Dim timeCount = 0
    Dim idCount = 0
    Dim splitCount = 0
    Dim sendCount = 0
    Dim insertCount = 0
    Dim failedCount = 0
    Dim time As String
    Dim id As String
    Dim datetime As DateTime

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Timer1.Enabled = True

    End Sub

    Private Function SaveToDB(conn, time, id, loc) As Boolean

        Dim fd As New OleFillingData
        Dim olePOSConn As New OleDbConnection(conn)
        Dim olePOSTrans As OleDbTransaction
        Dim sQuery As String
        Dim bretval

        Try
            If olePOSConn.State = ConnectionState.Closed Then
                olePOSConn.Open()
            End If

            id = id.ToString.PadLeft(5, "0")
            olePOSTrans = olePOSConn.BeginTransaction()
            If id.ToString.Contains("=") Or id.ToString.Contains(":") Then
                Return bRetBol
            End If

            Dim strSQL As String = "SELECT StaffID FROM AttendanceLog WHERE StaffID='" & id & "'AND datetime='" & time & "'"

            bretval = fd.IsDataExist(strSQL, olePOSConn, olePOSTrans)

            If bretval = 1 Then
                sQuery = "insert into AttendanceLog (StaffID,datetime,status,location,remark,syncd,synmd,synca) values 
                        ('" & id & "','" & time & "','1','" & loc & "','import','" & Format(DateTime.Now(), "yyyy-MM-dd HH:mm:ss") & "',
                        '" & Format(DateTime.Now(), "yyyy-MM-dd HH:mm:ss") & "','TMSEngine')"
                WriteAppLog(sQuery & "bRetBol:" & False & vbCrLf & "Duplicate entry id = [" & id & "],datetime = [" & Format(DateTime.Now(), "yyyy-MM-dd HH:mm:ss") & "]")
                failedCount += 1
            Else
                sQuery = "insert into AttendanceLog (StaffID,datetime,status,location,remark,syncd,synmd,synca) values 
                        ('" & id & "','" & time & "','1','" & loc & "','import','" & Format(DateTime.Now(), "yyyy-MM-dd HH:mm:ss") & "',
                        '" & Format(DateTime.Now(), "yyyy-MM-dd HH:mm:ss") & "','TMSEngine')"
                bRetBol = fd.InsertData(sQuery, olePOSConn, olePOSTrans)
                WriteAppLog(sQuery & "bRetBol:" & bRetBol)
                olePOSTrans.Commit()
                insertCount += 1
            End If


        Catch ex As Exception
            WriteAppLog("[INSERT ATTLOG]" & ex.Message)
        End Try

        If olePOSConn.State = ConnectionState.Open Then
            olePOSConn.Close()
        End If

        Return bRetBol
    End Function

    Private Function WriteToFile(arr, files)
        Dim sw As StreamWriter
        If File.Exists(files) = False Then
            sw = File.CreateText(files)
            sw.Close()

        End If
        File.WriteAllLines(files, arr)
        Return True
    End Function

    Private Function readFile() As String()
        Dim oleConn As OleDbConnection = New OleDbConnection(connStrAcc)
        Dim sQuery As String

        Dim oleTrans As OleDbTransaction
        Dim fd As New OleFillingData
        Dim dt As String
        Dim no As Integer = 0

        If oleConn.State = ConnectionState.Closed Then
            oleConn.Open()
        End If

        oleTrans = oleConn.BeginTransaction()
        sQuery = "select value1 from mastercode where type='path'"
        dt = fd.Get_Return_Value(sQuery, oleConn, oleTrans)

        If Not Directory.Exists(dt) Then
            WriteAppLog("[readFile]Invalid path")
        End If

        dir = dt
        entry = Directory.GetFiles(dir, "*.txt")


        If oleConn.State = ConnectionState.Open Then
            oleConn.Close()
        End If
        Return entry
    End Function

    Private Function moveFile(file, destination)
        Dim oleConn As OleDbConnection = New OleDbConnection(connStrAcc)
        Dim sQuery As String

        Dim oleTrans As OleDbTransaction
        Dim fd As New OleFillingData
        Dim dt As String
        Dim no As Integer = 0
        Dim newName
        Try
            If oleConn.State = ConnectionState.Closed Then
                oleConn.Open()
            End If

            oleTrans = oleConn.BeginTransaction()
            sQuery = "select value1 from mastercode where type='archive'"
            dt = fd.Get_Return_Value(sQuery, oleConn, oleTrans)
            If Not Directory.Exists(dt) Then
                Directory.CreateDirectory(dt)
                WriteAppLog("[moveFile]Creating archive file...")
            End If
            newName = dt + "\" + destination

            My.Computer.FileSystem.MoveFile(file, newName)
            If oleConn.State = ConnectionState.Open Then
                oleConn.Close()
            End If
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Private Function CreateConnStr() As String
        Dim connStr As String

        connStr = "Provider=SQLOLEDB;Persist Security Info=True;Initial Catalog=mais_ezhr;
                   User Id=workplace; Password=smartwp11;Data Source=127.0.0.1"

        Return connStr
    End Function
    Public Sub WriteAppLog(ByVal strLogMsg As String)
        Dim sw As StreamWriter
        Try
            Dim infoReader As System.IO.FileInfo
            Dim tempfile As Integer

            If Not Directory.Exists(Application.StartupPath & "\log") Then
                Directory.CreateDirectory(Application.StartupPath & "\log")
            End If

            If File.Exists(Application.StartupPath & "\log\SystemLog.txt") = True Then
                infoReader = My.Computer.FileSystem.GetFileInfo(Application.StartupPath & "\log\SystemLog.txt")
                tempfile = infoReader.Length

                If tempfile >= 5242880 Then     '' 5MB
                    If Not Directory.Exists(Application.StartupPath & "\log\archive\" & MonthName(Month(Date.Today())) & "-" & Year(Date.Today()) & "") Then
                        Directory.CreateDirectory(Application.StartupPath & "\log\archive\" & MonthName(Month(Date.Today())) & "-" & Year(Date.Today()) & "")
                    End If

                    infoReader.MoveTo(Application.StartupPath & "\log\archive\" & MonthName(Month(Date.Today())) & "-" & Year(Date.Today()) & "\logRetailPOS_" & CStr(Format(Date.Now(), "yyyyMMddHHmmss")) & ".txt")
                End If
            End If

            sw = File.AppendText(Application.StartupPath & "\log\SystemLog.txt")
            sw.WriteLine(Format(Date.Now(), "MM/dd/yyyy HH:mm:ss") & " : " & strLogMsg)

        Catch ex As Exception
            sw.WriteLine(Format(Date.Now(), "MM/dd/yyyy HH:mm:ss") & " : " & "Error in ~WriteAppLog(). " & ex.Message() & " (Error Code: " & Err.Number & ")")
        Finally
            sw.Close()
            sw.Dispose()
        End Try
    End Sub

    Public Function fileInUse(sFile As String) As Boolean
        Try
            Using f As New IO.FileStream(sFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            End Using
        Catch ex As Exception
            Return True
        End Try
        Return False
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Dim total As Integer

        'reset all value
        insertCount = 0
        failedCount = 0

        Timer1.Enabled = False
        WriteAppLog("Timer start... Updating Attendance Log")
        entry = readFile()

        For Each fileName In entry

            If fileInUse(fileName) = True Then
                WriteAppLog("File is in use... Please wait for next cycle")
            End If

            newFileName = fileName.Substring(fileName.LastIndexOf("\") + 1)
            newFileName = newFileName.Remove(newFileName.LastIndexOf("."))
            Dim newFileDir = Application.StartupPath & "\" & newFileName & ".txt"
            My.Computer.FileSystem.CopyFile(fileName, newFileDir)

            timeCount = 0
            idCount = 0
            splitCount = 0

            Array.Clear(arr, 0, arr.Length)
            Array.Clear(arr2, 0, arr2.Length)

            sr = New IO.StreamReader(newFileDir)
            read = sr.ReadToEnd
            arr = Split(read, " ")
            Try
                While splitCount < arr.Length - 1
                    If arr(splitCount).Contains("time=") Then

                        time = arr(splitCount) + " " + arr(splitCount + 1)
                        time = time.Substring(time.IndexOf(vbCrLf) + 1)
                        time = time.Substring(time.IndexOf(""""))
                        time = time.Trim("""")
                        datetime = DateTime.ParseExact(time, "yyyy-MM-dd HH:mm:ss", Nothing)

                        arr2(timeCount) = datetime

                        timeCount += 1
                    End If

                    'id in arr(0) is dev_id, normal id start at 1
                    If arr(splitCount).Contains("id=") Then
                        id = arr(splitCount)
                        id = id.Substring(id.IndexOf(""""))
                        id = id.Trim("""")
                        arr(idCount) = id
                        idCount += 1

                    End If

                    splitCount += 1
                End While
            Catch ex As Exception
                WriteAppLog("read error : " + ex.Message)
            End Try

            sr.Close()

            sendCount = 0
            Dim sendcount2 = 1

            My.Computer.FileSystem.DeleteFile(newFileDir)
            Try

                Dim loc = arr(0)

                While idCount > sendCount
                    id = arr(sendcount2)
                    datetime = arr2(sendCount)

                    SaveToDB(CreateConnStr, datetime, id, loc)

                    sendCount += 1
                    sendcount2 += 1
                End While

                Me.Text = "Moving " & fileName
                newFileName = newFileName + "_" + DateTime.Now.ToString("ddMMyyHHmm") + ".txt"

                moveFile(fileName, newFileName)
            Catch ex As Exception
                WriteAppLog("write error : " + ex.Message)
            End Try

            total = insertCount + failedCount

            '22 space before Failed so that it aligns with others
            WriteAppLog("Total no of data inserted : " & insertCount & " out of " & total &
                    vbCrLf & "                      Failed:" & failedCount)

        Next


        Threading.Thread.Sleep(500)
        Me.Close()


    End Sub
End Class
