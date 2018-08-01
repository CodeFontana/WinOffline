'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline.PatchVector
' File Name:    PatchVector.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Class PatchVector

        Private _FileName As FileVector                     ' The patch file (Ex: C:\SomeDirectory\SomePatch.jcl)
        Private _InstructionList As ArrayList               ' List of patch instructions.
        Private _PatchAction As Integer                     ' Action code/result from applying patch.
        Private _PreCmdReturnCodes As ArrayList             ' Stores return codes from all PRESYSCMD scripts.
        Private _SysCmdReturnCodes As ArrayList             ' Stores return codes from all SYSCMD scripts.
        Private _FileReplaceResult As ArrayList             ' Stores result of file replacement operations.
        Private _CommentString As String                    ' Stores a comment.
        Private _SourceReplaceList As New ArrayList         ' Source location filename for each file replacement.
        Private _DestReplaceList As New ArrayList           ' Destination location filename for each file replacement.
        Private _ReplaceFolder As String                    ' Replacement location folder for storing original files.
        Private _ReplaceSubFolder As New ArrayList          ' Sub folder within replacement folder, for storing each original file.
        Private _SkipIfNotFoundList As New ArrayList        ' File replacements to skip, if destination file is not originally present (i.e. source file is new).
        Public Const NOT_APPLICABLE As Integer = -4         ' Apply action code: Not applicable.
        Public Const ALREADY_APPLIED As Integer = -3        ' Apply action code: Already applied.
        Public Const NO_ACTION As Integer = -2              ' Apply action code: No action.
        Public Const SKIPPED As Integer = -1                ' Apply action code: Skipped.
        Public Const APPLY_OK As Integer = 0                ' Apply action code: Apply ok.
        Public Const APPLY_FAIL As Integer = 1              ' Apply action code: Apply fail.
        Public Const EXECUTE_OK As Integer = 2              ' Apply action code: Execute ok.
        Public Const EXECUTE_FAIL As Integer = 3            ' Apply action code: Execute fail.
        Public Const FILE_SKIPPED As Integer = -1           ' File replacement code: Skipped.
        Public Const FILE_OK As Integer = 0                 ' File replacement code: File ok.
        Public Const FILE_REBOOT_REQUIRED As Integer = 1    ' File replacement code: Reboot required.
        Public Const FILE_FAILED As Integer = 2             ' File replacement code: File fail.
        Private ClientAutoProductCodes = New ArrayList({"BITCM", "DTSVMG", "DTMGSU", "TNGAMO", "TNGSDO", "TNGRCO"})
        Private SharedComponentProductCodes As New ArrayList({"DTMINF"})
        Private CAMProductCodes As New ArrayList({"CCSCAM", "TNGCAM"})
        Private SSAProductCodes As New ArrayList({"SSA"})
        Private DataTransportProductCodes As New ArrayList({"TNGDTS", "TNGDTS-30", "DTS"})
        Private ExplorerGUIProductCodes As New ArrayList({"EGC30N"})

        Public Sub New(ByVal FileName As String,
                       ByVal InstructionList As ArrayList)

            ' Initialize patch vector
            _FileName = New FileVector(FileName)
            _InstructionList = New ArrayList(InstructionList)
            _PatchAction = NO_ACTION
            _PreCmdReturnCodes = New ArrayList
            _SysCmdReturnCodes = New ArrayList
            _FileReplaceResult = New ArrayList
            _CommentString = ""

            ' Iterate file replacements -- Build source file replacement list
            For Each ReplacedFile As String In GetShortNameReplaceList()

                ' Add filename to source replacment list
                _SourceReplaceList.Add(_FileName.GetFilePath + "\" + ReplacedFile)

                ' Stub a file replacement result
                _FileReplaceResult.Add(-1)

            Next

            ' Iterate file replacements -- Build destination file replacement list
            For Each ReplacedFile As String In GetRawReplaceList()

                ' Base path to destination depends on product code
                If IsClientAuto() Then

                    ' Add to destination file replacement list
                    _DestReplaceList.Add(Globals.DSMFolder + ReplacedFile)

                ElseIf IsSharedComponent() Then

                    ' Special case -- CAPKI may not be with other shared components
                    If ReplacedFile.ToLower.StartsWith("capki\capki") Then

                        ' Use CAPKI folder
                        _DestReplaceList.Add(Globals.CAPKIFolder + ReplacedFile.ToLower.Replace("capki\capki", ""))

                    Else

                        ' Add to destination file replacement list
                        _DestReplaceList.Add(Globals.SharedCompFolder + ReplacedFile)

                    End If

                ElseIf IsCAM() Then

                    ' Add to destination file replacement list
                    _DestReplaceList.Add(Globals.CAMFolder + "\" + ReplacedFile)

                ElseIf IsSSA() Then

                    ' Add to destination file replacement list
                    _DestReplaceList.Add(Globals.SSAFolder + ReplacedFile)

                ElseIf IsDataTransport() Then

                    ' Add to destination file replacement list
                    _DestReplaceList.Add(Globals.DTSFolder + "\" + ReplacedFile)

                ElseIf IsExplorerGUI() Then

                    ' Add to destination file replacement list
                    _DestReplaceList.Add(Globals.EGCFolder + ReplacedFile)

                End If

                ' Check for subfolder in replacement file path
                If ReplacedFile.Contains("\") Then

                    ' Store reference to sub folder
                    _ReplaceSubFolder.Add(ReplacedFile.Substring(0, ReplacedFile.LastIndexOf("\")))

                Else

                    ' Stub empty subfolder string
                    _ReplaceSubFolder.Add("")

                End If

            Next

            ' Set replaced subfolder path based on product code
            If IsClientAuto() Then _ReplaceFolder = Globals.DSMFolder + "REPLACED\" + _FileName.GetFriendlyName.ToUpper
            If IsSharedComponent() Then _ReplaceFolder = Globals.SharedCompFolder + "REPLACED\" + _FileName.GetFriendlyName.ToUpper
            If IsCAM() Then _ReplaceFolder = Globals.CAMFolder + "\REPLACED\" + _FileName.GetFriendlyName.ToUpper
            If IsSSA() Then _ReplaceFolder = Globals.SSAFolder + "REPLACED\" + _FileName.GetFriendlyName.ToUpper
            If IsDataTransport() Then _ReplaceFolder = Globals.DTSFolder + "\REPLACED\" + _FileName.GetFriendlyName.ToUpper
            If IsExplorerGUI() Then _ReplaceFolder = Globals.EGCFolder + "REPLACED\" + _FileName.GetFriendlyName.ToUpper

            ' Build "skip if not found" list (files to be replaced only if they exist on the target)
            _SkipIfNotFoundList = GetSkipFileList()

        End Sub

        Public Function GetRawInstruction(ByVal Name As String) As String

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Check for instruction match by name
                If strLine.ToLower.StartsWith(Name.ToLower) Then

                    ' Return the raw instruction
                    Return strLine

                End If

            Next

            ' Return emptry string
            Return ""

        End Function

        Public Function GetInstruction(ByVal Name As String) As String

            ' Local variables
            Dim ReturnString As String = ""

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Match instruction by name
                If strLine.ToLower.StartsWith(Name.ToLower) And strLine.Contains(":") Then

                    ' Parse instruction detail
                    strLine = strLine.Substring(strLine.IndexOf(":") + 1)

                    ' Check for additional instruction specifications
                    If strLine.Contains(":") Then

                        ' Trim additional instruction specifications
                        strLine = strLine.Substring(0, strLine.IndexOf(":"))

                    End If

                    ' Check for empty return string
                    If ReturnString.Equals("") Then

                        ' Add instruction
                        ReturnString += strLine

                    Else

                        ' Add delimiter and instruction
                        ReturnString += ";" + strLine

                    End If

                    ' Break the loop
                    Exit For

                End If

            Next

            ' Return
            Return ReturnString
        End Function

        Public Function GetRawReplaceList() As ArrayList

            ' Local variables
            Dim ReturnList As New ArrayList

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Check for file replacement instruction
                If strLine.ToLower.StartsWith("file:") Then

                    ' Parse instruction
                    strLine = strLine.Substring(strLine.IndexOf(":") + 1)

                    ' Check for additional instruction specifications
                    If strLine.Contains(":") Then

                        ' Trim additional instruction specifications
                        strLine = strLine.Substring(0, strLine.IndexOf(":"))

                    End If

                    ' Add file replacement to list
                    ReturnList.Add(strLine)

                End If

            Next

            ' Return
            Return ReturnList

        End Function

        Public Function GetShortNameReplaceList() As ArrayList

            ' Local variables
            Dim ReturnList As New ArrayList

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Check for file replacement instruction
                If strLine.ToLower.StartsWith("file:") Then

                    ' Parse instruction
                    strLine = strLine.Substring(strLine.IndexOf(":") + 1)

                    ' Check for additional instruction specifications
                    If strLine.Contains(":") Then

                        ' Trim additional instruction specifications
                        strLine = strLine.Substring(0, strLine.IndexOf(":"))

                    End If

                    ' Check for folder path separator
                    If strLine.Contains("\") Then

                        ' Trim folder path
                        strLine = strLine.Substring(strLine.LastIndexOf("\") + 1)

                    End If

                    ' Add file shortname to list
                    ReturnList.Add(strLine)

                End If

            Next

            ' Return
            Return ReturnList

        End Function

        Public Function GetSkipFileList() As ArrayList

            ' Local variables
            Dim ReturnList As New ArrayList

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Check for file replacement instruction, with :skipifnotfound: specification
                If strLine.ToLower.StartsWith("file:") And strLine.ToLower.Contains(":skipifnotfound:") Then

                    ' Trim file replacement
                    strLine = strLine.Substring(strLine.IndexOf(":") + 1)
                    strLine = strLine.Substring(0, strLine.IndexOf(":"))

                    ' Check for folder path separator
                    If strLine.Contains("\") Then

                        ' Trim folder path
                        strLine = strLine.Substring(strLine.LastIndexOf("\") + 1)

                    End If

                    ' Add file shortname to list
                    ReturnList.Add(strLine.ToLower)

                End If

            Next

            ' Return
            Return ReturnList

        End Function

        Public Function GetPreCommandList() As ArrayList

            ' Local variables
            Dim ReturnList As New ArrayList

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Check for presyscmd
                If strLine.ToLower.StartsWith("presyscmd:") Then

                    ' Trim instruction
                    strLine = strLine.Substring(strLine.IndexOf(":") + 1)
                    strLine = strLine.Replace(" ", "")

                    ' Check for multiple listing delimeter
                    If strLine.Contains(",") Then

                        ' Split delimited text to list
                        ReturnList.AddRange(strLine.Split(","))

                    Else

                        ' Add to list
                        ReturnList.Add(strLine)

                    End If

                End If

            Next

            ' Return list
            Return ReturnList
        End Function

        Public Function GetSysCommandList() As ArrayList

            ' Local variables
            Dim ReturnList As New ArrayList

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Check for syscmd
                If strLine.ToLower.StartsWith("syscmd:") Then

                    ' Trim instruction
                    strLine = strLine.Substring(strLine.IndexOf(":") + 1)
                    strLine = strLine.Replace(" ", "")

                    ' Check for multiple listing delimeter
                    If strLine.Contains(",") Then

                        ' Split delimited text to list
                        ReturnList.AddRange(strLine.Split(","))

                    Else

                        ' Add to list
                        ReturnList.Add(strLine)

                    End If

                End If

            Next

            ' Return
            Return ReturnList

        End Function

        Public Function GetDependsList() As ArrayList

            ' Local variables
            Dim ReturnList As New ArrayList

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Check for dependency tag
                If strLine.ToLower.StartsWith("depends:") Then

                    ' Trim instruction
                    strLine = strLine.Substring(strLine.IndexOf(":") + 1)
                    strLine = strLine.Replace(" ", "")

                    ' Check for multiple listing delimeter
                    If strLine.Contains(",") Then

                        ' Split delimited text to list
                        ReturnList.AddRange(strLine.Split(","))

                    Else

                        ' Add to list
                        ReturnList.Add(strLine)

                    End If

                End If

            Next

            ' Return
            Return ReturnList

        End Function

        Public Function GetAllDependentFiles() As ArrayList

            ' Local variables
            Dim ReturnArray As New ArrayList

            ' Iterate file shortname replacement list
            For Each strLine As String In GetShortNameReplaceList()

                ' Add to list
                ReturnArray.Add(strLine)

            Next

            ' Iterate presyscmd file list
            For Each strLine As String In GetPreCommandList()
                ' Add to list
                ReturnArray.Add(strLine)

            Next

            ' Iterate syscmd file list
            For Each strLine As String In GetSysCommandList()

                ' Add to list
                ReturnArray.Add(strLine)

            Next

            ' Iterate dependent file list
            For Each strLine As String In GetDependsList()

                ' Add to list
                ReturnArray.Add(strLine)

            Next

            ' Return
            Return ReturnArray

        End Function

        Public Function GetApplyPTFDateFormat() As String

            ' Local variables
            Dim ReturnString As String = ""
            Dim CurrentTime As DateTime = DateTime.Now

            ' Build applyptf formatted date
            ReturnString += CurrentTime.ToString("D").Substring(0, 3) + " "
            ReturnString += CurrentTime.ToString("M").Substring(0, 3) + " "
            ReturnString += CurrentTime.Day.ToString + " "
            ReturnString += CurrentTime.TimeOfDay.ToString.Substring(0, CurrentTime.TimeOfDay.ToString.IndexOf(".")) + " "
            ReturnString += CurrentTime.Year.ToString

            ' Return
            Return ReturnString

        End Function

        Public Function GetApplyPTFSummary() As String

            ' Local variables
            Dim ReturnString As String = "["
            Dim CurrentTime As DateTime = DateTime.Now

            ' Build applyptf entry header
            ReturnString += GetApplyPTFDateFormat() + "] - PTF Wizard "

            ' Check patch action code
            If _PatchAction = 0 Then

                ' Build applyptf entry header
                ReturnString += "installed " + _FileName.GetFriendlyName + " (" + GetInstruction("PRODUCT") + ")" + Environment.NewLine

                ' Build applyptf entry properties
                ReturnString += GetRawInstruction("RELEASE").Replace(":", "=") + Environment.NewLine
                ReturnString += GetRawInstruction("GENLEVEL").Replace(":", "=") + Environment.NewLine
                ReturnString += GetRawInstruction("COMPONENT").Replace(":", "=") + Environment.NewLine
                ReturnString += GetRawInstruction("PREREQS").Replace(":", "=") + Environment.NewLine
                ReturnString += GetRawInstruction("MPREREQS").Replace(":", "=") + Environment.NewLine
                ReturnString += GetRawInstruction("COREQS").Replace(":", "=") + Environment.NewLine
                ReturnString += GetRawInstruction("MCOREQS").Replace(":", "=") + Environment.NewLine
                ReturnString += GetRawInstruction("SUPERSEDE").Replace(":", "=")

                ' Iterate destination file replacement list
                For x As Integer = 0 To _DestReplaceList.Count - 1

                    ' Read a desintation file
                    Dim strLine As String = _DestReplaceList.Item(x)

                    ' Check for positive file replacement code
                    If _FileReplaceResult.Item(x) = FILE_OK Or _FileReplaceResult.Item(x) = FILE_REBOOT_REQUIRED Then

                        ' Build installed file entry
                        ReturnString += Environment.NewLine + "INSTALLEDFILE=" + strLine + ",TIME="
                        ReturnString += System.IO.File.GetLastWriteTime(strLine).ToString("D").Substring(0, 3) + " "
                        ReturnString += System.IO.File.GetLastWriteTime(strLine).ToString("M").Substring(0, 3) + " "
                        ReturnString += System.IO.File.GetLastWriteTime(strLine).Day.ToString + " "
                        ReturnString += System.IO.File.GetLastWriteTime(strLine).TimeOfDay.ToString.Substring(0, CurrentTime.TimeOfDay.ToString.IndexOf(".")) + " "
                        ReturnString += System.IO.File.GetLastWriteTime(strLine).Year.ToString

                    End If

                Next

                ' Append new line (after list of installedfile= tags, or after supersede=
                ReturnString += Environment.NewLine

            ElseIf _PatchAction = 1 Then

                ' Build applyptf entry header
                ReturnString += "backed out " + _FileName.GetFriendlyName + " (" + GetInstruction("PRODUCT") + ")" + Environment.NewLine

                ' Return
                Return ReturnString

            End If

            ' Check for double new line
            If ReturnString.EndsWith(Environment.NewLine + Environment.NewLine) Then

                ' Trim additional new line
                ReturnString = ReturnString.Remove(ReturnString.LastIndexOf(Environment.NewLine))

            End If

            ' Return
            Return ReturnString

        End Function

        Public Function WereAllFileReplacementsSkipped() As Boolean
            If _FileReplaceResult.Count > 0 Then
                For Each result As Integer In _FileReplaceResult
                    If Not result = FILE_SKIPPED Then Return False
                Next
                Return True
            End If
            Return False
        End Function

        Public Function IsClientAuto() As Boolean
            For Each ProductCode As String In ClientAutoProductCodes
                If ProductCode.ToLower.Equals(GetInstruction("PRODUCT").ToLower) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function IsSharedComponent() As Boolean
            For Each ProductCode As String In SharedComponentProductCodes
                If ProductCode.ToLower.Equals(GetInstruction("PRODUCT").ToLower) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function IsCAM() As Boolean
            For Each ProductCode As String In CAMProductCodes
                If ProductCode.ToLower.Equals(GetInstruction("PRODUCT").ToLower) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function IsSSA() As Boolean
            For Each ProductCode As String In SSAProductCodes
                If ProductCode.ToLower.Equals(GetInstruction("PRODUCT").ToLower) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function IsDataTransport() As Boolean
            For Each ProductCode As String In DataTransportProductCodes
                If ProductCode.ToLower.Equals(GetInstruction("PRODUCT").ToLower) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function IsExplorerGUI() As Boolean
            For Each ProductCode As String In ExplorerGUIProductCodes
                If ProductCode.ToLower.Equals(GetInstruction("PRODUCT").ToLower) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Property PatchFile As FileVector
            Get
                Return _FileName
            End Get
            Set(value As FileVector)
                _FileName = value
            End Set
        End Property

        Public Property InstructionList As ArrayList
            Get
                Return _InstructionList
            End Get
            Set(value As ArrayList)
                _InstructionList = value
            End Set
        End Property

        Public Property PatchAction As Integer
            Get
                Return _PatchAction
            End Get
            Set(value As Integer)
                _PatchAction = value
            End Set
        End Property

        Public Property PreCmdReturnCodes As ArrayList
            Get
                Return _PreCmdReturnCodes
            End Get
            Set(value As ArrayList)
                _PreCmdReturnCodes = value.Clone
            End Set
        End Property

        Public Property SysCmdReturnCodes As ArrayList
            Get
                Return _SysCmdReturnCodes
            End Get
            Set(value As ArrayList)
                _SysCmdReturnCodes = value.Clone
            End Set
        End Property
        Public Property FileReplaceResult As ArrayList
            Get
                Return _FileReplaceResult
            End Get
            Set(value As ArrayList)
                _FileReplaceResult = value.Clone
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

        Public ReadOnly Property SourceReplaceList As ArrayList
            Get
                Return _SourceReplaceList
            End Get
        End Property

        Public ReadOnly Property DestReplaceList As ArrayList
            Get
                Return _DestReplaceList
            End Get
        End Property

        Public ReadOnly Property ReplaceFolder As String
            Get
                Return _ReplaceFolder
            End Get
        End Property

        Public ReadOnly Property ReplaceSubFolder As ArrayList
            Get
                Return _ReplaceSubFolder
            End Get
        End Property

        Public ReadOnly Property SkipIfNotFoundList As ArrayList
            Get
                Return _SkipIfNotFoundList
            End Get
        End Property

        Public Overrides Function toString() As String

            ' Local variables
            Dim ReturnString As String = ""

            ' Add patch file name
            ReturnString += _FileName.GetFileName() + Environment.NewLine()

            ' Iterate the patch instuction list (JCL)
            For Each strLine As String In InstructionList

                ' Add instruction with delimeter
                ReturnString += strLine + ";"

            Next

            ' Add patch apply action
            ReturnString += Environment.NewLine() + _PatchAction.ToString

            ' Add presyscmd header
            ReturnString += Environment.NewLine() + "PRESYSCMD:"

            ' Iterate presyscmd return codes
            For Each strLine As String In _PreCmdReturnCodes

                ' Add return code
                ReturnString += strLine + ";"

            Next

            ' Add syscmd header
            ReturnString += Environment.NewLine() + "SYSCMD:"

            ' Iterate syscmd return codes
            For Each strLine As String In _SysCmdReturnCodes

                ' Add return code
                ReturnString += strLine + ";"

            Next

            ' Add file replacement return code header
            ReturnString += Environment.NewLine() + "FILE_REPLACE:"

            ' Iterate file replcement results
            For Each strLine As String In _FileReplaceResult

                ' Add replacement code
                ReturnString += strLine + ";"

            Next

            ' Add comment string
            ReturnString += Environment.NewLine + _CommentString

            ' Return
            Return ReturnString

        End Function

        Public Shared Function ReadCache(ByVal CallStack As String) As ArrayList

            ' Local variables
            Dim CacheFile As String = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-patch.cache"
            Dim CacheReader As System.IO.StreamReader
            Dim strLine As String = ""
            Dim strLine2 As String = ""
            Dim strLine3 As String = ""
            Dim InstructionList As New ArrayList
            Dim CodeList As New ArrayList
            Dim pVector As PatchVector
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

                ' Read patch filename
                strLine = CacheReader.ReadLine

                ' Read the jcl instruction list
                strLine2 = CacheReader.ReadLine

                ' Clear instruction list
                InstructionList = New ArrayList

                ' Process the jcl instruction list
                While strLine2.Contains(";") And strLine2.Length > 1

                    ' Parse an instruction from the list
                    strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                    ' Add the instuction to the array
                    InstructionList.Add(strLine3)

                    ' Remove the instruction from the list
                    strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

                End While

                ' Create a PatchVector object
                pVector = New PatchVector(strLine, InstructionList)

                ' Read patch action
                strLine = CacheReader.ReadLine()
                pVector.PatchAction = Integer.Parse(strLine)

                ' Read PRESYSCMD return codes
                strLine2 = CacheReader.ReadLine

                ' Trim presyscmd header
                strLine2 = strLine2.Replace("PRESYSCMD:", "")

                ' Clear return code list
                CodeList = New ArrayList

                ' Process presyscmd return codes
                While strLine2.Contains(";") And strLine2.Length > 1

                    ' Parse a return code
                    strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                    ' Add to the array
                    CodeList.Add(strLine3)

                    ' Trim the list
                    strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

                End While

                ' Set PRESYSCMD return code list
                pVector.PreCmdReturnCodes = CodeList

                ' Read SYSCMD return codes
                strLine2 = CacheReader.ReadLine

                ' Trim header
                strLine2 = strLine2.Replace("SYSCMD:", "")

                ' Clear return code list
                CodeList = New ArrayList

                ' Process syscmd return codes
                While strLine2.Contains(";") And strLine2.Length > 1

                    ' Parse a return code
                    strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                    ' Add to the array
                    CodeList.Add(strLine3)

                    ' Trim the list
                    strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

                End While

                ' Set SYSCMD return code property
                pVector.SysCmdReturnCodes = CodeList

                ' Read file replacement codes
                strLine2 = CacheReader.ReadLine

                ' Trim header
                strLine2 = strLine2.Replace("FILE_REPLACE:", "")

                ' Clear return code list
                CodeList = New ArrayList

                ' Process file replacement return codes
                While strLine2.Contains(";") And strLine2.Length > 1

                    ' Parse a return code
                    strLine3 = strLine2.Substring(0, strLine2.IndexOf(";"))

                    ' Add to the array
                    CodeList.Add(strLine3)

                    ' Trim the list
                    strLine2 = strLine2.Substring(strLine2.IndexOf(";") + 1)

                End While

                ' Set file replacement codes
                pVector.FileReplaceResult = CodeList

                ' Read comment string
                strLine = CacheReader.ReadLine
                pVector.CommentString = strLine

                ' Add to patch manifest
                RestoredManifest.Add(pVector)

                ' Write debug
                Logger.WriteDebug(CallStack, "Read: " + pVector.PatchFile.GetFriendlyName)

            Loop

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

            ' Close the cache file
            CacheReader.Close()

            ' Return
            Return RestoredManifest

        End Function

        Public Shared Sub WriteCache(ByVal CallStack As String, ByVal PatchManifest As ArrayList)

            ' Local variables
            Dim CacheFile As String = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-patch.cache"
            Dim CacheWriter As System.IO.StreamWriter
            Dim CurPatch As PatchVector

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

            ' Iterate the patch manifest
            For n As Integer = 0 To PatchManifest.Count - 1

                ' Get/cast a patch
                CurPatch = DirectCast(PatchManifest.Item(n), PatchVector)

                ' Write debug
                Logger.WriteDebug(CallStack, "Write: " + CurPatch.PatchFile.GetFriendlyName)

                ' Write the patch to cache
                CacheWriter.WriteLine(CurPatch.toString)

            Next

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

            ' Close cache file
            CacheWriter.Close()

        End Sub

    End Class

End Class