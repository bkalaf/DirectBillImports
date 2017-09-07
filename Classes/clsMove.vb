Imports System
Imports System.IO
Imports System.Configuration

Public Class clsMove
    Public Sub MoveFiles()
        Dim myfiles As String()
        Dim FileName As String
        Dim sToFolder As String = InputBox("To Period (YYYYMM):")
        Dim sPeriod As String = MakePeriod(sToFolder)
        Dim outfile As StreamWriter
        Dim sTempFile As String = "c:\my stuff\usfg.txt"
        If File.Exists(sTempFile) Then File.Delete(sTempFile)
        outfile = File.CreateText(sTempFile)
        Dim sMainDir As String = ConfigurationSettings.AppSettings("GeoveraUSFGPath")

        myfiles = Directory.GetFiles(ConfigurationSettings.AppSettings("GeoveraFilePath"))
        Try
            For Each FileName In myfiles
                outfile.WriteLine(FileName)
                Dim JustFileName As String = Mid(FileName, 39, 22)

                If Not Directory.Exists(sMainDir & "FFB\" & sToFolder) Then
                    Directory.CreateDirectory(sMainDir & "FFB\" & sToFolder)
                End If
                If Not Directory.Exists(sMainDir & "Normal\" & sToFolder) Then
                    Directory.CreateDirectory(sMainDir & "Normal\" & sToFolder)
                End If
                Dim midname As String = Mid(JustFileName, 13, 3)
                If midname = "SIU" Then
                    Dim ToFolderAbbr = IIf(Mid(JustFileName, 16, 3) = "FFB", "FFB", "Normal")
                    Dim FileMonth As String = Mid(JustFileName, 3, 2) & Mid(JustFileName, 1, 2)
                    If FileMonth = sPeriod Then
                        If Not File.Exists(sMainDir & ToFolderAbbr & "\" & sToFolder & "\" & JustFileName & ".CSV") Then File.Copy(FileName, ConfigurationSettings.AppSettings("GeoveraUSFGPath") & ToFolderAbbr & "\" & sToFolder & "\" & JustFileName & ".CSV")
                    End If
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally

        End Try
    End Sub

    Public Sub ArchiveFiles()
        Dim myfiles As String()
        Dim FileName As String
        myfiles = Directory.GetFiles(ConfigurationSettings.AppSettings("GeoveraFilePath"))
        Try
            For Each FileName In myfiles
                Dim dPath As String = Replace(FileName, ConfigurationSettings.AppSettings("GeoveraFilePath"), ConfigurationSettings.AppSettings("GeoveraArchivePath"))
                File.Move(FileName, dPath)
            Next
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally

        End Try
    End Sub

    Private Function MakePeriod(ByVal sToFolder As String) As String
        Dim sAnswer As String = ""
        Dim sFirstDay As Date = CDate(Mid(sToFolder, 5, 2) & "/01/" & Mid(sToFolder, 1, 4))
        Dim sLastMonthFirstDay As Date = DateAdd(DateInterval.Month, 0, sFirstDay)
        sAnswer = IIf(Month(sLastMonthFirstDay) < 10, "0", "") & CStr(Month(sLastMonthFirstDay)) & Mid(CStr(Year(sLastMonthFirstDay)), 3, 2)
        Return sAnswer
    End Function

End Class