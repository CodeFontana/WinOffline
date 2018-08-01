'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline/HistoryVector
' File Name:    HistoryVector.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Class HistoryVector

        Private ProductComponent As String
        Private HistoryEntries As New ArrayList

        Public Sub New(ByVal Component As String, ByVal Entries As ArrayList)

            ' Initialize history vector
            ProductComponent = Component
            HistoryEntries = Entries

        End Sub

        Public Function IsEmpty() As Boolean

            ' Check for "NoHistory" placeholder
            If HistoryEntries.Item(0).Equals("NoHistory") Then

                ' Return
                Return True

            Else

                ' Return
                Return False

            End If

        End Function

        Public Function IsApplyAction() As Boolean

            ' Local variables
            Dim EntryString As String

            ' Check for positive number of entries
            If HistoryEntries.Count > 0 Then

                ' Get the first element
                ' Example: "[Thu Mar 26 15:35:01 2015] - PTF Wizard installed T5J1378 (DTSVMG)"
                ' Example: "[Thu Apr 23 15:36:33 2015] - PTF Wizard backed out T55V065 (DTSVMG)"
                ' Example: "[Sun Feb 15 14:34:44 2015] - PTF Wizard installed INSTALL-ENC-CERTS (DTSVMG)"
                EntryString = HistoryEntries.Item(0)

                ' Check for installed keyword
                If EntryString.ToLower.Contains("installed") Then

                    ' Return
                    Return True

                Else

                    ' Return
                    Return False

                End If

            End If

            ' Return
            Return False

        End Function

        Public Function GetPatchName() As String

            ' Local variables
            Dim EntryString As String

            ' Check for positive number of entries
            If HistoryEntries.Count > 0 Then

                ' Check for empty history
                If IsEmpty() Then Return ProductComponent

                ' Get the first element
                ' Example: "[Thu Mar 26 15:35:01 2015] - PTF Wizard installed T5J1378 (DTSVMG)"
                ' Example: "[Thu Apr 23 15:36:33 2015] - PTF Wizard backed out T55V065 (DTSVMG)"
                ' Example: "[Sun Feb 15 14:34:44 2015] - PTF Wizard installed INSTALL-ENC-CERTS (DTSVMG)"
                ' Exmaple: "Shared Components -- SC;NoHistory"
                EntryString = HistoryEntries.Item(0)

                ' Check entry keyword
                If EntryString.ToLower.Contains("installed") Then

                    ' Trim patch name
                    EntryString = EntryString.Substring(EntryString.IndexOf("installed") + 10)
                    EntryString = EntryString.Substring(0, EntryString.IndexOf("(") - 1)

                    ' Return
                    Return EntryString

                ElseIf EntryString.ToLower.Contains("backed out") Then

                    ' Trim patch name
                    EntryString = EntryString.Substring(EntryString.IndexOf("backed out") + 11)
                    EntryString = EntryString.Substring(0, EntryString.IndexOf("(") - 1)

                    ' Return
                    Return EntryString

                Else

                    ' Return
                    Return Nothing

                End If

            End If

            ' Return
            Return Nothing

        End Function

        Public Function GetInstalledFiles() As ArrayList

            ' Local variables
            Dim InstalledFiles As New ArrayList

            ' Iterate history file entries
            For Each p As String In HistoryEntries

                ' Check for installed file
                If p.ToLower.StartsWith("installedfile=") Then

                    ' Trim filename
                    p = p.Substring(p.IndexOf("=") + 1)
                    p = p.Substring(0, p.IndexOf(","))

                    ' Check for duplicate file name (corrupt history file case)
                    If Not InstalledFiles.Contains(p) Then

                        ' Add filename to list
                        InstalledFiles.Add(p)

                    End If

                End If

            Next

            ' Return
            Return InstalledFiles

        End Function

        Public Function GetInstallFilesRaw() As String

            ' Local variables
            Dim ReturnString As String = ""

            ' Iterate history file entries
            For Each p As String In HistoryEntries

                ' Check for installed file
                If p.ToLower.StartsWith("installedfile=") Then

                    ' Add installed file info to list
                    ReturnString += p + Environment.NewLine

                End If

            Next

            ' Return list
            Return ReturnString

        End Function

        Public Function GetProductComponent() As String
            Return ProductComponent
        End Function

        Public Function GetHeader() As String
            Return HistoryEntries.Item(0)
        End Function

        Public Overrides Function toString() As String

            ' Local variables
            Dim ReturnString As String = ""

            ' Add product component
            ReturnString += ProductComponent + Environment.NewLine()

            ' Iterate history file entries
            For Each strLine As String In HistoryEntries

                ' Add history entry
                ReturnString += strLine + ";"

            Next

            ' Return
            Return ReturnString

        End Function

        Public Shared Function ReadCache(ByVal CallStack As String) As ArrayList

            ' Local variables
            Dim CacheFile As String = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-history.cache"
            Dim CacheReader As System.IO.StreamReader
            Dim strLine As String = ""
            Dim strLine2 As String = ""
            Dim strLine3 As String = ""
            Dim Counter As Integer = 0
            Dim CodeList As New ArrayList
            Dim HistoryEntries As New ArrayList
            Dim RestoredHistory As HistoryVector
            Dim RestoredManifest As New ArrayList

            ' Write debug
            Logger.WriteDebug(CallStack, "Open file: " + CacheFile)

            ' Check if cache file exists
            If Not System.IO.File.Exists(CacheFile) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Cache file does not exist.")

                ' Return
                Return Nothing

            End If

            ' Open the cache file
            CacheReader = New System.IO.StreamReader(CacheFile)

            ' Iterate cache file contents
            Do While CacheReader.Peek() >= 0

                ' Read the product component
                strLine = CacheReader.ReadLine

                ' Read the history entries
                strLine2 = CacheReader.ReadLine

                ' Clear history entries
                HistoryEntries = New ArrayList

                ' Process the history entries
                While strLine2.Contains(";") And strLine2.Length > 1

                    ' Parse a history entry
                    strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                    ' Add the entry to the array
                    HistoryEntries.Add(strLine3)

                    ' Trim the list of entries
                    strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

                End While

                ' Create a HistoryVector object
                RestoredHistory = New HistoryVector(strLine, HistoryEntries)

                ' Add to restored manifest
                RestoredManifest.Add(RestoredHistory)

                ' Check event action for debug
                If RestoredHistory.IsEmpty Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Read: " + RestoredHistory.GetPatchName + ", NoHistory")

                ElseIf RestoredHistory.IsApplyAction Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Read: " + RestoredHistory.GetPatchName + ", Applied")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Read: " + RestoredHistory.GetPatchName + ", Removed")

                End If

            Loop

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

            ' Close the cache file
            CacheReader.Close()

            ' Return
            Return RestoredManifest

        End Function

        Public Shared Sub WriteCache(ByVal CallStack As String, ByVal HistoryManifest As ArrayList)

            ' Local variables
            Dim CacheFile As String = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-history.cache"
            Dim CacheWriter As System.IO.StreamWriter
            Dim CurEvent As HistoryVector

            ' *****************************
            ' - Flush prior cache file.
            ' *****************************

            ' Verify any prior cache file is deleted before creating a new file
            If System.IO.File.Exists(CacheFile) Then

                ' Delete prior cache file
                Utility.DeleteFile(CallStack, CacheFile)

            ElseIf Not System.IO.Directory.Exists(Globals.WinOfflineTemp) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Warning: Temporary folder already cleaned up!")
                Logger.WriteDebug(CallStack, "Warning: Cache file write aborted!")

                ' Return
                Return

            End If

            ' *****************************
            ' - Create a new cache file.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Write cache file: " + CacheFile)

            ' Open the cache file for write
            CacheWriter = New System.IO.StreamWriter(CacheFile, False)

            ' Iterate the history manifest
            For n As Integer = 0 To HistoryManifest.Count() - 1

                ' Get/cast a history event
                CurEvent = DirectCast(HistoryManifest.Item(n), HistoryVector)

                ' Check event action for debug message
                If CurEvent.IsEmpty Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Write: " + CurEvent.GetPatchName + ", NoHistory")

                ElseIf CurEvent.IsApplyAction Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Write: " + CurEvent.GetPatchName + ", Applied")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Write: " + CurEvent.GetPatchName + ", Removed")

                End If

                ' Write history to cache
                CacheWriter.WriteLine(CurEvent.toString)

            Next

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

            ' Close cache file
            CacheWriter.Close()

        End Sub

    End Class

End Class