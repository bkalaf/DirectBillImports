Imports System
Imports System.IO
Imports System.Text
Imports System.Configuration

Public Class clsTextWriterTraceListener

    Dim myTracer As New TextWriterTraceListener(Console.Out)
    Dim myError As New TextWriterTraceListener()

    Public Sub CreateErrorWriter(ByVal sCarrier As String)

        'Year
        If Not Directory.Exists(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\") Then
            Directory.CreateDirectory(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\")
        End If
        'Month
        If Not Directory.Exists(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\" & Date.Now.ToString("MM") & "\") Then
            Directory.CreateDirectory(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\" & Date.Now.ToString("MM") & "\")
        End If
        'File
        If File.Exists(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\" & Date.Now.ToString("MM") & "\" & Date.Now.ToString("MMddyyyy") & ".log") Then
            File.Delete(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\" & Date.Now.ToString("MM") & "\" & Date.Now.ToString("MMddyyyy") & ".log")
            myError = New TextWriterTraceListener(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\" & Date.Now.ToString("MM") & "\" & Date.Now.ToString("MMddyyyy") & ".log")
        Else
            myError = New TextWriterTraceListener(ConfigurationSettings.AppSettings("LogPath") & sCarrier & "\" & Date.Now.ToString("yyyy") & "\" & Date.Now.ToString("MM") & "\" & Date.Now.ToString("MMddyyyy") & ".log")
        End If
    End Sub

    Public Sub AddListerners()
        Debug.Listeners.Add(myTracer)
        Debug.Listeners.Add(myError)
    End Sub

    Public Sub CloseWriter()
        Debug.Flush()
        Debug.Close()
    End Sub
End Class