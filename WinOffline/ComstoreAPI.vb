Partial Public Class WinOffline

    Public Class ComstoreAPI

        Public Shared Function GetParameterValue(ByVal ParameterSection As String,
                                                 ByVal ParameterName As String) As String

            ' Local variables
            Dim ExecutionString As String
            Dim ArgumentString As String
            Dim RunningProcess As Process
            Dim ProcessStartInfo As ProcessStartInfo
            Dim StandardOutput As String

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\" + "ccnfcmda.exe"
            ArgumentString = "-cmd getparametervalue -ps " + ParameterSection + " -pn " + ParameterName

            ' Create detached process for reading comstore
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Read standard output
            StandardOutput = RunningProcess.StandardOutput.ReadToEnd

            ' Close detached process
            RunningProcess.Close()

            ' Return -- remove carriage return (with full safety)
            Return StandardOutput.Replace(vbCr, "").Replace(vbLf, "")

        End Function

        Public Shared Function SetParameterValue(ByVal ParameterSection As String,
                                                 ByVal ParameterName As String,
                                                 ByVal NewValue As String) As String

            ' Local variables
            Dim ExecutionString As String
            Dim ArgumentString As String
            Dim RunningProcess As Process
            Dim ProcessStartInfo As ProcessStartInfo
            Dim StandardOutput As String

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\" + "ccnfcmda.exe"
            ArgumentString = "-cmd setparametervalue -ps " + ParameterSection + " -pn " + ParameterName + " -v """ + NewValue + """"

            ' Create detached process for reading comstore
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Read standard output
            StandardOutput = RunningProcess.StandardOutput.ReadToEnd

            ' Close detached process
            RunningProcess.Close()

            ' Return -- remove carriage return (with full safety)
            Return StandardOutput.Replace(vbCr, "").Replace(vbLf, "")

        End Function

        Public Shared Function DeleteParameter(ByVal ParameterSection As String,
                                               ByVal ParameterName As String) As String

            ' Local variables
            Dim ExecutionString As String
            Dim ArgumentString As String
            Dim RunningProcess As Process
            Dim ProcessStartInfo As ProcessStartInfo
            Dim StandardOutput As String

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\" + "ccnfcmda.exe"
            ArgumentString = "-cmd deleteparameter -ps " + ParameterSection + " -pn " + ParameterName

            ' Create detached process for reading comstore
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Read standard output
            StandardOutput = RunningProcess.StandardOutput.ReadToEnd

            ' Close detached process
            RunningProcess.Close()

            ' Return -- remove carriage return (with full safety)
            Return StandardOutput.Replace(vbCr, "").Replace(vbLf, "")

        End Function

        Public Shared Function DeleteParameterSection(ByVal ParameterSection As String) As String

            ' Local variables
            Dim ExecutionString As String
            Dim ArgumentString As String
            Dim RunningProcess As Process
            Dim ProcessStartInfo As ProcessStartInfo
            Dim StandardOutput As String

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\" + "ccnfcmda.exe"
            ArgumentString = "-cmd deleteparamsection -ps " + ParameterSection

            ' Create detached process for reading comstore
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Read standard output
            StandardOutput = RunningProcess.StandardOutput.ReadToEnd

            ' Close detached process
            RunningProcess.Close()

            ' Return -- remove carriage return (with full safety)
            Return StandardOutput.Replace(vbCr, "").Replace(vbLf, "")

        End Function

    End Class

End Class