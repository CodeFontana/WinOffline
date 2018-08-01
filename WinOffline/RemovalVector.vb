'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline.RemovalVector
' File Name:    PatchVector.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Class RemovalVector

        Private _RemovalItem As String                      ' Name of patch to be removed.
        Private _RemovalAction As Integer                   ' Action code/result from removing patch.
        Private _HistoryFile As String                      ' History file related to removal operation.
        Private _RemovalFileName As ArrayList               ' Stores list of files removed.
        Private _FileRemovalResult As ArrayList             ' Stores result of removal operation.
        Private _CommentString As String                    ' Stores a comment.
        Public Const NO_ACTION As Integer = -2              ' Removal action code: No action.
        Public Const SKIPPED As Integer = -1                ' Removal action code: Skipped.
        Public Const REMOVAL_OK As Integer = 0              ' Removal action code: Removal ok.
        Public Const REMOVAL_FAIL As Integer = 1            ' Removal action code: Removal fail.
        Public Const FILE_SKIPPED As Integer = -1           ' File replacement code: Skipped.
        Public Const FILE_OK As Integer = 0                 ' File replacement code: File ok.
        Public Const FILE_REBOOT_REQUIRED As Integer = 1    ' File replacement code: Reboot required.
        Public Const FILE_FAILED As Integer = 2             ' File replacement code: File failed.

        Public Sub New(ByVal RemovalItem As String)

            ' Initialize removal vector
            _RemovalItem = RemovalItem
            _RemovalAction = NO_ACTION
            _RemovalFileName = New ArrayList
            _FileRemovalResult = New ArrayList
            _CommentString = ""

        End Sub

        Public Property RemovalItem As String
            Get
                Return _RemovalItem
            End Get
            Set(value As String)
                _RemovalItem = value
            End Set
        End Property

        Public Property RemovalAction As Integer
            Get
                Return _RemovalAction
            End Get
            Set(value As Integer)
                _RemovalAction = value
            End Set
        End Property

        Public Property HistoryFile As String
            Get
                Return _HistoryFile
            End Get
            Set(value As String)
                _HistoryFile = value
            End Set
        End Property

        Public Property RemovalFileName As ArrayList
            Get
                Return _RemovalFileName
            End Get
            Set(value As ArrayList)
                _RemovalFileName = value.Clone
            End Set
        End Property

        Public Property FileRemovalResult As ArrayList
            Get
                Return _FileRemovalResult
            End Get
            Set(value As ArrayList)
                _FileRemovalResult = value.Clone
            End Set
        End Property

        Public Property CommentString As String
            Get
                Return _CommentString
            End Get
            Set(value As String)
                _CommentString = value
            End Set
        End Property

        Public Overrides Function toString() As String

            ' Local variables
            Dim ReturnString As String = ""

            ' Add removal item (patch name)
            ReturnString += _RemovalItem + Environment.NewLine

            ' Add removal action code
            ReturnString += _RemovalAction.ToString + Environment.NewLine

            ' Add patch history file
            ReturnString += _HistoryFile

            ' Add file replacement list
            ReturnString += Environment.NewLine() + "FILE_REPLACE:"

            ' Iterate replacement file list
            For Each strLine As String In RemovalFileName

                ' Add replacement file
                ReturnString += strLine + ";"

            Next

            ' Add file replacement result list
            ReturnString += Environment.NewLine() + "FILE_RESULT:"

            ' Iterate file replacement result list
            For Each strLine As String In FileRemovalResult

                ' Add replacement result code
                ReturnString += strLine + ";"

            Next

            ' Add removal comment
            ReturnString += Environment.NewLine + _CommentString

            ' Return output string
            Return ReturnString

        End Function

        Public Shared Function ReadCache(ByVal CallStack As String) As ArrayList

            ' Local variables
            Dim CacheFile As String = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-removal.cache"
            Dim CacheReader As System.IO.StreamReader
            Dim strLine As String = ""
            Dim strLine2 As String = ""
            Dim strLine3 As String = ""
            Dim CodeList As New ArrayList
            Dim rVector As RemovalVector
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

                ' Read a removal item
                strLine = CacheReader.ReadLine

                ' Create a removal vector
                rVector = New RemovalVector(strLine)

                ' Read removal action
                strLine = CacheReader.ReadLine
                rVector.RemovalAction = Integer.Parse(strLine)

                ' Read history file
                strLine = CacheReader.ReadLine
                rVector.HistoryFile = strLine

                ' Read removal filenames
                strLine2 = CacheReader.ReadLine
                strLine2 = strLine2.Replace("FILE_REPLACE:", "")

                ' Clear return code list
                CodeList = New ArrayList

                ' Process filenames
                While strLine2.Contains(";") And strLine2.Length > 1

                    ' Parse a filename
                    strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                    ' Add to the array
                    CodeList.Add(strLine3)

                    ' Trim the list
                    strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

                End While

                ' Set removal filenames
                rVector.RemovalFileName = CodeList

                ' Read removal replacement codes
                strLine2 = CacheReader.ReadLine
                strLine2 = strLine2.Replace("FILE_RESULT:", "")

                ' Clear return code list
                CodeList = New ArrayList

                ' Process the return codes
                While strLine2.Contains(";") And strLine2.Length > 1

                    ' Parse a return code
                    strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                    ' Add to the array
                    CodeList.Add(strLine3)

                    ' Trim the list
                    strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

                End While

                ' Set file replacement codes
                rVector.FileRemovalResult = CodeList

                ' Read comment string
                strLine = CacheReader.ReadLine
                rVector.CommentString = strLine

                ' Add to removal manifest
                RestoredManifest.Add(rVector)

                ' Write debug
                Logger.WriteDebug(CallStack, "Read: " + rVector.RemovalItem)

            Loop

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

            ' Close the cache file
            CacheReader.Close()

            ' Return
            Return RestoredManifest

        End Function

        Public Shared Sub WriteCache(ByVal CallStack As String, ByVal RemovalManifest As ArrayList)

            ' Local variables
            Dim CacheFile As String = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-removal.cache"
            Dim CacheWriter As System.IO.StreamWriter

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

            ' Iterate the removal manifest
            For Each RemoveItem As RemovalVector In RemovalManifest

                ' Write debug
                Logger.WriteDebug(CallStack, "Write: " + RemoveItem.RemovalItem)

                ' Write the patch to cache
                CacheWriter.WriteLine(RemoveItem)

            Next

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

            ' Close cache file
            CacheWriter.Close()

        End Sub

        Public Shared Function ReadString(ByVal CacheReader As System.IO.StreamReader) As RemovalVector

            ' Local variables
            Dim strLine As String = ""
            Dim strLine2 As String = ""
            Dim strLine3 As String = ""
            Dim CodeList As New ArrayList
            Dim rVector As RemovalVector

            ' Read a removal item
            strLine = CacheReader.ReadLine

            ' Create a removal vector
            rVector = New RemovalVector(strLine)

            ' Read removal action
            strLine = CacheReader.ReadLine
            rVector.RemovalAction = Integer.Parse(strLine)

            ' Read history file
            strLine = CacheReader.ReadLine
            rVector.HistoryFile = strLine

            ' Read removal filenames
            strLine2 = CacheReader.ReadLine
            strLine2 = strLine2.Replace("FILE_REPLACE:", "")

            ' Clear return code list
            CodeList = New ArrayList

            ' Process filenames
            While strLine2.Contains(";") And strLine2.Length > 1

                ' Parse a filename
                strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                ' Add to the array
                CodeList.Add(strLine3)

                ' Trim the list
                strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

            End While

            ' Set removal filenames
            rVector.RemovalFileName = CodeList

            ' Read removal replacement codes
            strLine2 = CacheReader.ReadLine
            strLine2 = strLine2.Replace("FILE_RESULT:", "")

            ' Clear return code list
            CodeList = New ArrayList

            ' Process the return codes
            While strLine2.Contains(";") And strLine2.Length > 1

                ' Parse a return code
                strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                ' Add to the array
                CodeList.Add(strLine3)

                ' Trim the list
                strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

            End While

            ' Set file replacement codes
            rVector.FileRemovalResult = CodeList

            ' Read comment string
            strLine = CacheReader.ReadLine
            rVector.CommentString = strLine

            ' Return
            Return rVector

        End Function

    End Class

End Class