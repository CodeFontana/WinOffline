'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   CertManager
' File Name:    CertManager.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Class CertManager

        ' Class variables
        Private Shared CertStoreFile As String
        Private Shared KeyStoreFile As String
        Private Shared CertDBFolder As String
        Private Shared CertificateStore As New ArrayList
        Private Shared PrivateKeyStore As New ArrayList

        Public Shared Sub RepairStore(ByVal CallStack As String)

            ' Local variables
            Dim CBBFolder As String
            Dim CBBFolderPrimaryBackup As String
            Dim CBBFolderSecondaryBackup As String
            Dim CurrentLine As String
            Dim PreviousLine As String
            Dim EntryKey As String
            Dim EntryType As String
            Dim EntryData As String
            Dim PrivateKeyList As New ArrayList
            Dim CertFileList As New ArrayList
            Dim AnonymousCertCounter As Integer = 0
            Dim StopCondition As Boolean = False
            Dim FileList As String()

            ' Update call stack
            CallStack += "CertMgr|"

            ' Build paths to cert files
            CBBFolder = Globals.SharedCompFolder + "CBB"
            CBBFolderPrimaryBackup = Globals.SharedCompFolder + "CBB.backup1"
            CBBFolderSecondaryBackup = Globals.SharedCompFolder + "CBB.backup2"
            CertStoreFile = Globals.SharedCompFolder + "CBB\certstor.dat"
            KeyStoreFile = Globals.SharedCompFolder + "CBB\cbbkstor.dat"
            CertDBFolder = Globals.SharedCompFolder + "CBB\certdb"

            ' *****************************
            ' - Perform verifications.
            ' *****************************

            ' Verify certdb folder exists
            If Not System.IO.Directory.Exists(CBBFolder) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Warning: Certificate folder does not exist: " + CBBFolder)
                Logger.WriteDebug(CallStack, "Warning: Unable to repair certificate store.")

                ' Return
                Return

            End If

            ' Verify cert store file exists
            If Not System.IO.File.Exists(CertStoreFile) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Warning: Certificate store file does not exist: " + CertStoreFile)
                Logger.WriteDebug(CallStack, "Warning: Unable to repair certificate store.")

                ' Return
                Return

            End If

            ' Verify private key store file exists
            If Not System.IO.File.Exists(KeyStoreFile) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Warning: Private key store file does not exist: " + KeyStoreFile)
                Logger.WriteDebug(CallStack, "Warning: Unable to repair certificate store.")

                ' Return
                Return

            End If

            ' *****************************
            ' - Backup CBB folder.
            ' *****************************

            ' Check for second backup folder
            If System.IO.Directory.Exists(CBBFolderSecondaryBackup) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Delete folder: " + CBBFolderSecondaryBackup)

                ' Delete the folder
                System.IO.Directory.Delete(CBBFolderSecondaryBackup, True)

            End If

            ' Check for primary backup folder
            If System.IO.Directory.Exists(CBBFolderPrimaryBackup) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Move folder: " + CBBFolderPrimaryBackup)
                Logger.WriteDebug(CallStack, "To: " + CBBFolderSecondaryBackup)

                ' Move the folder
                System.IO.Directory.Move(CBBFolderPrimaryBackup, CBBFolderSecondaryBackup)

            End If

            ' Backup CBB folder
            If System.IO.Directory.Exists(CBBFolder) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Backup folder: " + CBBFolder)
                Logger.WriteDebug(CallStack, "To: " + CBBFolderPrimaryBackup)

                ' Copy the folder
                If Not Utility.CopyDirectory(CBBFolder, CBBFolderPrimaryBackup) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Warning: Unable to backup CBB folder.")
                    Logger.WriteDebug(CallStack, "Warning: Unable to repair certificate store.")

                    ' Return
                    Return

                End If

            End If

            ' *****************************
            ' - Read certiciate and key stores.
            ' *****************************

            ' Read cert store
            ReadCertStore(CallStack)

            ' *****************************
            ' - Process certificate store.
            ' *****************************

            ' Iterate the certificate store
            For i As Integer = 0 To CertificateStore.Count - 1

                ' Reset variables
                EntryKey = ""
                EntryType = ""
                EntryData = ""

                ' Check for last entry
                If i > CertificateStore.Count - 1 Then

                    ' Check last entry
                    If CertificateStore.Item(i - 1).ToString.ToLower.Equals("end") Then

                        ' Add a blank space to bottom of array
                        CertificateStore.Add(" ")

                    ElseIf CertificateStore.Item(i - 1).ToString.ToLower.Equals("") Then

                        ' Convert blank line --> to blank space
                        CertificateStore.Item(i - 1) = " "

                    End If

                    ' Stop processing
                    Exit For

                End If

                ' Read new line
                CurrentLine = CertificateStore.Item(i)

                ' Read previous line
                If i > 0 Then

                    ' Read previous line
                    PreviousLine = CertificateStore.Item(i - 1)

                Else

                    ' Stub previous line
                    PreviousLine = "-1"

                End If

                ' Check current line
                If CurrentLine.ToLower.StartsWith("id=") Then

                    ' Parse tag info
                    EntryKey = CurrentLine.Substring(3)

                    ' Check tag length
                    If Not EntryKey.Length > 0 OrElse Not EntryKey.Contains(".") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Removing entry: " + CurrentLine)

                        ' Remove it
                        CertificateStore.RemoveAt(i)

                        ' Decrement i
                        If i > 0 Then i -= 1

                        ' Skip to end
                        Continue For

                    End If

                    ' Parse entry type (cert or tag)
                    EntryType = EntryKey.Substring(0, EntryKey.IndexOf("."))

                    ' Parse uuid
                    EntryKey = EntryKey.Substring(EntryKey.IndexOf(".") + 1)

                    ' Check type/uuid information, formatting, and tag is followed by sufficient lines of data
                    If Not (EntryType.ToLower.Equals("cert") Or EntryType.ToLower.Equals("tag")) OrElse
                        (EntryType.ToLower.Equals("cert") AndAlso (EntryKey.Length <> 40 Or Not Utility.IsHexString(EntryKey))) OrElse
                        (EntryType.ToLower.Equals("tag") AndAlso EntryKey.ToLower.Contains("itcm-anonymous")) OrElse
                        i + 3 > CertificateStore.Count - 1 OrElse
                        Not CertificateStore.Item(i + 1).ToString.ToLower.Equals("data=") OrElse
                        Not CertificateStore.Item(i + 3).ToString.ToLower.Equals("end") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Removing entry: " + CurrentLine)

                        ' Remove it
                        CertificateStore.RemoveAt(i)

                        ' Decrement i
                        If i > 0 Then i -= 1

                        ' Skip to end
                        Continue For

                    End If

                    ' Read entry data
                    EntryData = CertificateStore.Item(i + 2)

                    ' Check entry type
                    If EntryType.ToLower.Equals("tag") Then

                        ' Check data
                        If Not EntryData.ToLower.StartsWith("cn=") Or
                            EntryData.ToLower.Contains("itcm-self-signed") Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i))
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 1))
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 2))
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 3))

                            ' Remove the entire entry
                            CertificateStore.RemoveAt(i)
                            CertificateStore.RemoveAt(i)
                            CertificateStore.RemoveAt(i)
                            CertificateStore.RemoveAt(i)

                            ' Decrement i
                            If i > 0 Then i -= 1

                            ' Skip to end
                            Continue For

                        Else

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Valid tag:         " + CertificateStore.Item(i + 2))

                            ' Check spacing from previous entry
                            If Not PreviousLine.ToLower.Equals(" ") Then

                                ' Add a blank space
                                CertificateStore.Insert(i, " ")

                                ' Increment i
                                i += 1

                            End If

                            ' Entry is valid, increment i (to 'end' of section)
                            i += 3

                            ' Skip to end
                            Continue For

                        End If

                    ElseIf EntryType.ToLower.Equals("cert") Then

                        ' Check data
                        If Not EntryData.ToLower.StartsWith("subj") Or
                            EntryData.ToLower.Contains("itcm-self-signed") Or
                            Not EntryData.ToLower.Contains("file") Then

                            ' Check for anonymous certificate (for accountability)
                            If EntryData.ToLower.Contains("itcm-self-signed") Then

                                ' Increment the counter
                                AnonymousCertCounter += 1

                            End If

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i))
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 1))
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 2))
                            Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 3))

                            ' Remove the entire entry
                            CertificateStore.RemoveAt(i)
                            CertificateStore.RemoveAt(i)
                            CertificateStore.RemoveAt(i)
                            CertificateStore.RemoveAt(i)

                            ' Decrement i
                            If i > 0 Then i -= 1

                            ' Skip to end
                            Continue For

                        Else

                            ' Isolate the private key file
                            Dim FileString As String = EntryData.Substring(EntryData.IndexOf("file """) + 6)
                            FileString = FileString.Substring(0, FileString.Length - 1)

                            ' Verify file exists
                            If Not System.IO.File.Exists(FileString) Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i))
                                Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 1))
                                Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 2))
                                Logger.WriteDebug(CallStack, "Removing entry: " + CertificateStore.Item(i + 3))

                                ' Remove the entire entry
                                CertificateStore.RemoveAt(i)
                                CertificateStore.RemoveAt(i)
                                CertificateStore.RemoveAt(i)
                                CertificateStore.RemoveAt(i)

                                ' Decrement i
                                If i > 0 Then i -= 1

                                ' Skip to end
                                Continue For

                            End If

                            ' Add to certificate file list
                            CertFileList.Add(FileString.ToLower)

                            ' Add private key to list
                            PrivateKeyList.Add(EntryKey.ToLower)

                            ' Isolate the certificate subject
                            Dim CertSubject As String = CertificateStore.Item(i + 2).ToString.Substring(CertificateStore.Item(i + 2).ToString.IndexOf("CN="))
                            CertSubject = CertSubject.Substring(0, CertSubject.IndexOf(""""))

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Valid certificate: " + CertSubject)

                            ' Check spacing from previous entry
                            If Not PreviousLine.ToLower.Equals(" ") Then

                                ' Add a blank space
                                CertificateStore.Insert(i, " ")

                                ' Increment i
                                i += 1

                            End If

                            ' Entry is valid, increment i (to 'end' of section)
                            i += 3

                            ' Skip to end
                            Continue For

                        End If

                    End If

                ElseIf (CurrentLine.Equals("") And PreviousLine.Equals("")) Or
                    (CurrentLine.ToLower.Equals(" ") And PreviousLine.ToLower.Equals("")) Or
                    (CurrentLine.ToLower.Equals(" ") And PreviousLine.ToLower.Equals(" ")) Then

                    ' Remove current line
                    CertificateStore.RemoveAt(i)

                    ' Decrement i
                    If i > 0 Then i -= 1

                ElseIf (CurrentLine.ToLower.StartsWith(" ") And Not (PreviousLine.ToLower.Equals("") Or PreviousLine.ToLower.Equals(" "))) Or
                    CurrentLine.ToLower.StartsWith("version") Or
                    CurrentLine.ToLower.StartsWith("#") Then

                    ' Do nothing -- these are valid symbols.

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Removing entry: " + CurrentLine)

                    ' Remove current line
                    CertificateStore.RemoveAt(i)

                    ' Decrement i
                    If i > 0 Then i -= 1

                End If

            Next

            ' *****************************
            ' - Process private key store.
            ' *****************************

            ' Reset variables
            PreviousLine = "-1"
            CurrentLine = ""

            ' Iterate the certificate store
            For i As Integer = 0 To PrivateKeyStore.Count - 1

                ' Reset variables
                EntryKey = ""
                EntryType = ""
                StopCondition = False

                ' Check for last entry
                If i > PrivateKeyStore.Count - 1 Then

                    ' Check last entry
                    If PrivateKeyStore.Item(i - 1).ToString.ToLower.Equals("end") Then

                        ' Add a blank space to bottom of array
                        PrivateKeyStore.Add(" ")

                    ElseIf PrivateKeyStore.Item(i - 1).ToString.ToLower.Equals("") Then

                        ' Convert blank line --> to blank space
                        PrivateKeyStore.Item(i - 1) = " "

                    End If

                    ' Stop processing
                    Exit For

                End If

                ' Read new line
                CurrentLine = PrivateKeyStore.Item(i)

                ' Read previous line
                If i > 0 Then

                    ' Read previous line
                    PreviousLine = PrivateKeyStore.Item(i - 1)

                Else

                    PreviousLine = "-1"

                End If

                ' Check current line
                If CurrentLine.ToLower.StartsWith("id=") Then

                    ' Parse tag info
                    EntryKey = CurrentLine.Substring(3)

                    ' Check tag length
                    If Not EntryKey.Length > 0 OrElse Not EntryKey.Contains(".") And Not EntryKey.Equals("CBB_LOCAL_SYSTEM_KEY_NAME") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Removing entry: " + CurrentLine)

                        ' Remove it
                        PrivateKeyStore.RemoveAt(i)

                        ' Decrement i
                        If i > 0 Then i -= 1

                        ' Skip to end
                        Continue For

                    End If

                    ' There's a VALID EntryKey called "CBB_LOCAL_SYSTEM_KEY_NAME" which does not have a "."
                    If EntryKey.Contains(".") Then

                        ' Parse entry type (ex. cbbcstor or cfencrypt_fips)
                        EntryType = EntryKey.Substring(0, EntryKey.IndexOf("."))

                        ' Parse uuid
                        EntryKey = EntryKey.Substring(EntryKey.IndexOf(".") + 1)

                    Else

                        ' The entire EntryKey is the type
                        EntryType = EntryKey

                    End If

                    ' Check type/uuid information, formatting, and tag is followed by sufficient lines of data
                    If (EntryType.ToLower.Equals("cbbcstor") AndAlso (EntryKey.Length <> 40 Or Not Utility.IsHexString(EntryKey))) OrElse
                        (EntryType.ToLower.Equals("cbbcstor") AndAlso Not (PrivateKeyList.Contains(EntryKey.ToLower))) OrElse
                        i + 3 > PrivateKeyStore.Count - 1 OrElse
                        Not PrivateKeyStore.Item(i + 1).ToString.ToLower.Equals("data=") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Removing entry: " + CurrentLine)

                        ' Remove it
                        PrivateKeyStore.RemoveAt(i)

                        ' Decrement i
                        If i > 0 Then i -= 1

                        ' Skip to end
                        Continue For

                    End If

                    ' Iterate ahead to consume data
                    For j As Integer = i + 2 To PrivateKeyStore.Count - 1

                        ' Check for 'end'
                        If PrivateKeyStore.Item(j).ToString.ToLower.Equals("end") Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Valid entry: " + PrivateKeyStore.Item(i))

                            ' Check spacing from previous entry
                            If Not PreviousLine.ToLower.Equals(" ") Then

                                ' Add a blank space
                                PrivateKeyStore.Insert(i, " ")

                                ' Increment i
                                i += 1

                            End If

                            ' Set stop condition flag
                            StopCondition = True

                            ' Entry is valid, move i up to j's position
                            i = j

                            ' Skip to end
                            Exit For

                        End If

                    Next

                    ' Check stop condition flag
                    If Not StopCondition Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Removing entry: " + CurrentLine)

                        ' Remove it
                        PrivateKeyStore.RemoveAt(i)

                        ' Decrement i
                        If i > 0 Then i -= 1

                    End If

                    ' Skip to end
                    Continue For

                ElseIf (CurrentLine.Equals("") And PreviousLine.Equals("")) Or
                    (CurrentLine.ToLower.Equals(" ") And PreviousLine.ToLower.Equals("")) Or
                    (CurrentLine.ToLower.Equals(" ") And PreviousLine.ToLower.Equals(" ")) Then

                    ' Remove current line
                    PrivateKeyStore.RemoveAt(i)

                    ' Decrement i
                    If i > 0 Then i -= 1

                ElseIf (CurrentLine.ToLower.StartsWith(" ") And Not (PreviousLine.ToLower.Equals("") Or PreviousLine.ToLower.Equals(" "))) Or
                    CurrentLine.ToLower.StartsWith("version") Or
                    CurrentLine.ToLower.StartsWith("#") Then

                    ' Do nothing -- these are valid symbols.

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Removing entry: " + CurrentLine)

                    ' Remove current line
                    PrivateKeyStore.RemoveAt(i)

                    ' Decrement i
                    If i > 0 Then i -= 1

                End If

            Next

            ' *****************************
            ' - Process certdb folder.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Found (" + AnonymousCertCounter.ToString + ") anonymous certificate(s).")

            ' Get certdb file listing
            FileList = System.IO.Directory.GetFiles(CertDBFolder)

            ' Iterate the certdb file list
            For Each CertFile As String In FileList

                ' Check for orphan file
                If Not CertFileList.Contains(CertFile.ToLower) Then

                    ' Delete the file
                    Utility.DeleteFile(CallStack, CertFile)

                End If

            Next

            ' Write updated certificate store
            WriteCertStore(CallStack)

            ' Trigger new anonymous cert
            TriggerNewCert(CallStack)

        End Sub

        Private Shared Sub ReadCertStore(ByVal CallStack As String)

            ' Local variables
            Dim CertFileReader As System.IO.StreamReader
            Dim strLine As String

            ' *****************************
            ' - Read certificate store.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Read certificate store: " + CertStoreFile)

            ' Open the certificate store file
            CertFileReader = New System.IO.StreamReader(CertStoreFile)

            ' Loop cert store file contents
            Do While CertFileReader.Peek() >= 0

                ' Read a line
                strLine = CertFileReader.ReadLine()

                ' Add to array
                CertificateStore.Add(strLine)

            Loop

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CertStoreFile)

            ' Close file
            CertFileReader.Close()

            ' *****************************
            ' - Read private key store.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Read private key store: " + KeyStoreFile)

            ' Open the private key store file
            CertFileReader = New System.IO.StreamReader(KeyStoreFile)

            ' Loop private key store contents
            Do While CertFileReader.Peek() >= 0

                ' Read a line
                strLine = CertFileReader.ReadLine()

                ' Add to array
                PrivateKeyStore.Add(strLine)

            Loop

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + KeyStoreFile)

            ' Close file
            CertFileReader.Close()

        End Sub

        Private Shared Sub WriteCertStore(ByVal CallStack As String)

            ' Local variables
            Dim CertFileWriter As System.IO.StreamWriter

            ' *****************************
            ' - Write new certificate store.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Write certificate store: " + CertStoreFile)

            ' Open the cert store for overwrite
            CertFileWriter = New IO.StreamWriter(CertStoreFile, False)

            ' Iterate the cert store
            For Each CertLine As String In CertificateStore

                ' Write the line to file
                CertFileWriter.WriteLine(CertLine)

            Next

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CertStoreFile)

            ' Close file
            CertFileWriter.Close()

            ' *****************************
            ' - Write new private key store.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Write private key store: " + KeyStoreFile)

            ' Open the private key store for overwrite
            CertFileWriter = New IO.StreamWriter(KeyStoreFile, False)

            ' Iterate the private key store
            For Each KeyLine As String In PrivateKeyStore

                ' Write the line to file
                CertFileWriter.WriteLine(KeyLine)

            Next

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + KeyStoreFile)

            ' Close file
            CertFileWriter.Close()

        End Sub

        Private Shared Sub TriggerNewCert(ByVal CallStack As String)

            ' Local variables
            Dim ExecutionString As String
            Dim ArgumentString As String
            Dim RunningProcess As Process
            Dim ProcessStartInfo As ProcessStartInfo
            Dim StandardOutput As String
            Dim RemainingOutput As String
            Dim ConsoleOutput As String

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\cacertutil.exe"
            ArgumentString = "list -v"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString)
            ProcessStartInfo.Arguments = ArgumentString
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

        End Sub

    End Class

End Class