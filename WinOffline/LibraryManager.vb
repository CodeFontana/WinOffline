Imports System.Data.SqlClient
Imports System.Threading

Partial Public Class WinOffline

    Public Class LibraryManager

        Public Shared LibraryEvents As New List(Of String)

        Public Shared Sub RepairLibrary(ByVal CallStack As String)

            ' Local variables
            Dim ITCMManager As Boolean = False
            Dim ITCMServer As Boolean = False
            Dim ConnectionString As String = Nothing
            Dim DbConnection As SqlConnection = Nothing
            Dim DomainUUID As String = Nothing
            Dim DomainId As String = Nothing
            Dim DomainType As String = Nothing
            Dim ManagerName As String = Nothing
            Dim AgentUUID As String = Nothing
            Dim ServerUUID As String = Nothing
            Dim LibraryFile As String
            Dim LibraryReader As System.IO.StreamReader
            Dim LibraryData As New List(Of String)
            Dim rswFileReader As System.IO.StreamReader
            Dim rswFileLine As String = Nothing
            Dim DbRegSwUUID As New List(Of String)
            Dim DbRegSwItemName As New List(Of String)
            Dim DbSwProcs As New List(Of String)
            Dim DctRegSw As New List(Of String)
            Dim DctRegSwFolder As New List(Of String)
            Dim CurrentLine As String
            Dim PreviousLine As String
            Dim DctMissingArcCount As Integer = 0
            Dim DctMissingVolCount As Integer = 0
            Dim DctMissingRegInfoCount As Integer = 0
            Dim DctDuplicateCount As Integer = 0
            Dim DctNotInDbCount As Integer = 0
            Dim DbNotInDct As New List(Of String)
            Dim ItemUUID As String = Nothing
            Dim ItemFolder As String = Nothing
            Dim ItemName As String = Nothing
            Dim ItemVersion As String = Nothing
            Dim ItemType As String = Nothing
            Dim ItemSubFolders As String() = Nothing
            Dim FolderExists As Boolean = False
            Dim VolumeExists As Boolean = False
            Dim RegInfoExists As Boolean = False
            Dim ArchivedPackage As Boolean = False
            Dim FolderList As String()
            Dim MatchFound As Boolean = False
            Dim ArcsForDeletion As New List(Of String)
            Dim RcpsForDeletion As New List(Of String)
            Dim LibraryFileBackup1 As String
            Dim LibraryFileBackup2 As String
            Dim LibraryFileWriter As System.IO.StreamWriter
            Dim SqlData As SqlDataReader
            Dim SqlCmd As SqlCommand
            Dim ReturnCode As Integer = 0

            ' Update call stack
            CallStack += "LibraryMgr|"

            ' *****************************
            ' - Perform initial verifications.
            ' *****************************

            ' Verify ITCM is installed
            If Not Utility.IsITCMInstalled Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: ITCM is not installed.")

                ' Return
                Return

            End If

            ' Verify ITCM functionality
            If Globals.ITCMFunction.ToLower.Contains("manager") Then

                ' Set flag
                ITCMManager = True

            ElseIf Globals.ITCMFunction.ToLower.Contains("server") Then

                ' Set flag
                ITCMServer = True

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Library cleanup only applicable to ITCM managers or servers.")

                ' Return
                Return

            End If

            ' Check if library path exists
            If Globals.SDLibraryFolder IsNot Nothing AndAlso System.IO.Directory.Exists(Globals.SDLibraryFolder) Then

                ' Append backslash (if needed)
                If Not Globals.SDLibraryFolder.EndsWith("\") Then Globals.SDLibraryFolder += "\"

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Library folder does not exist: " + Globals.SDLibraryFolder)

                ' Return
                Return

            End If

            ' Set library file
            LibraryFile = Globals.SDLibraryFolder + "library.dct"
            LibraryFileBackup1 = Globals.SDLibraryFolder + "library.dct.backup1"
            LibraryFileBackup2 = Globals.SDLibraryFolder + "library.dct.backup2"

            ' Check if library file exists
            If Not System.IO.File.Exists(LibraryFile) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Library file does not exist: " + LibraryFile)

                ' Return
                Return

            End If

            ' *****************************
            ' - Perform database verifications.
            ' *****************************

            ' Verify database connection info
            If Globals.DatabaseServer Is Nothing Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Database server name was not detected or provided.")
                Logger.WriteDebug(CallStack, "Info: Provide ""-dbserver serverName"" switch.")

                ' Return
                Return

            End If

            ' Provide default values, if necessary
            If Globals.DatabaseInstance Is Nothing Then Globals.DatabaseInstance = "MSSQLSERVER"
            If Globals.DatabasePort Is Nothing Then Globals.DatabasePort = "1433"
            If Globals.DatabaseName Is Nothing Then Globals.DatabaseName = "mdb"

            ' Verify if database user was provided
            If Globals.DbUser Is Nothing Then

                ' Assign current user
                Globals.DbUser = Globals.ProcessIdentity.Name

                ' Write debug
                Logger.WriteDebug(CallStack, "Logon as: " + Globals.DbUser)

            ElseIf Globals.DbPassword Is Nothing Then

                ' Write directly to attached console (this is not logged)
                Console.WriteLine()
                Console.WriteLine(CallStack.Substring(CallStack.IndexOf("LibraryMgr")) + "Logon as: " + Globals.DbUser)
                Console.Write(CallStack.Substring(CallStack.IndexOf("LibraryMgr")) + "Enter password: ")

                ' Hook the keyboard for low-level input (we are not a console app)
                KeyboardHook.SetHook()

                ' Wait for keyboard release
                While (KeyboardHook.KeyboardHooked)

                    ' Process message queue
                    Windows.Forms.Application.DoEvents()

                    ' Rest thread
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                End While

                ' Retrieve password
                Globals.DbPassword = KeyboardHook.CapturedString.ToString

            End If

            ' *****************************
            ' - Build connection string.
            ' *****************************

            ' Build connection string based on user account (SQL or trusted)
            If Globals.DbUser.Contains("\") Then

                ' Build connection string (windows authentication)
                ConnectionString = "Server=" + Globals.DatabaseServer
                ConnectionString += "\" + Globals.DatabaseInstance
                ConnectionString += "," + Globals.DatabasePort
                ConnectionString += ";Database=" + Globals.DatabaseName
                ConnectionString += ";Trusted_Connection=True"
                ConnectionString += ";Application Name=" + Globals.ProcessFriendlyName + " " + Globals.AppVersion

            Else

                ' Build connection string (SQL authentication)
                ConnectionString = "Server=" + Globals.DatabaseServer
                ConnectionString += "\" + Globals.DatabaseInstance
                ConnectionString += "," + Globals.DatabasePort
                ConnectionString += ";Database=" + Globals.DatabaseName
                ConnectionString += ";User Id=" + Globals.DbUser
                ConnectionString += ";Password=" + Globals.DbPassword
                ConnectionString += ";Application Name=" + Globals.ProcessFriendlyName + " " + Globals.AppVersion

            End If

            ' *****************************
            ' - Verify connection to SQL.
            ' *****************************

            ' Encapsulate sql connection
            Try

                ' Initialize sql connection instance
                DbConnection = New SqlConnection(ConnectionString)

                ' Open sql connection
                DbConnection.Open()

                ' Query domain_uuid
                DomainUUID = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select set_val_uuid from ca_settings with (nolock) where set_id=1")
                Logger.WriteDebug(CallStack, "Domain UUID: " + DomainUUID)

                ' Query domain_id
                DomainId = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select domain_id from ca_n_tier with (nolock) where domain_uuid=" + DomainUUID)
                Logger.WriteDebug(CallStack, "Domain ID: " + DomainId)

                ' Query domain_type
                DomainType = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select domain_type from ca_n_tier with (nolock) where domain_uuid=" + DomainUUID)
                Logger.WriteDebug(CallStack, "Domain type: " + DomainType)

                ' Query manager name (label)
                ManagerName = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select label from ca_manager with (nolock) where domain_uuid=" + DomainUUID)
                Logger.WriteDebug(CallStack, "Manager name: " + ManagerName)

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Database exception: " + ex.Message)

                ' Close the database connection
                DbConnection.Close()

                ' Write debug
                Logger.WriteDebug(CallStack, "Database connection closed.")

                ' Return
                Return

            End Try

            ' *****************************
            ' - Verify ITCM server is registered with ITCM manager (ITCM server only).
            ' *****************************

            ' ITCM server only
            If ITCMServer Then

                ' Encapsulate sql
                Try

                    ' Query ca_agent record with matching host_uuid
                    AgentUUID = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select dis_hw_uuid from ca_discovered_hardware with (nolock) where host_uuid='" + Globals.HostUUID.ToUpper + "'")

                    ' Check UUID (Ex: 0xD5BEF3369CAC5540994D1EF0ECC38FEF)
                    If AgentUUID Is Nothing OrElse AgentUUID.Length <> 34 Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Error: Scalability server not found in manager database.")

                        ' Close the database connection
                        DbConnection.Close()

                        ' Return
                        Return

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Agent UUID: " + AgentUUID)

                    End If

                    ' Query ca_server record with matching dis_hw_uuid
                    ServerUUID = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select server_uuid from ca_server with (nolock) where dis_hw_uuid=" + AgentUUID + "")

                    ' Check UUID
                    If ServerUUID Is Nothing OrElse ServerUUID.Length <> 34 Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Error: Scalability server is unregistered with manager database.")

                        ' Close the database connection
                        DbConnection.Close()

                        ' Return
                        Return

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Server UUID: " + ServerUUID)

                    End If

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Database exception: " + ex.Message)

                    ' Close the database connection
                    DbConnection.Close()

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Database connection closed.")

                    ' Return
                    Return

                End Try

            End If

            ' *****************************
            ' - Read database packages.
            ' *****************************

            ' Encapsulate sql
            Try

                ' Query UUID of REGISTERED software packages from database
                DbRegSwUUID = DatabaseAPI.SqlSelectScalarList(CallStack, DbConnection, "select cast(objectid as uniqueidentifier) from usd_rsw where itemtype<>5")
                Logger.WriteDebug(CallStack, "Database Software: " + DbRegSwUUID.Count.ToString)

                ' Query itemname of REGISTERED software packages from database
                DbRegSwItemName = DatabaseAPI.SqlSelectScalarList(CallStack, DbConnection, "select itemname from usd_rsw where itemtype<>5")

                ' Query list of software procedures from database
                DbSwProcs = DatabaseAPI.SqlSelectScalarList(CallStack, DbConnection, "select cast(objectid as uniqueidentifier) from usd_actproc")
                Logger.WriteDebug(CallStack, "Database Procedures: " + DbSwProcs.Count.ToString)

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Database exception: " + ex.Message)

                ' Close the database connection
                DbConnection.Close()

                ' Write debug
                Logger.WriteDebug(CallStack, "Database connection closed.")

                ' Return
                Return

            End Try

            ' *****************************
            ' - Read library file.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Read library file: " + LibraryFile)

            ' Open library file
            LibraryReader = New System.IO.StreamReader(LibraryFile)

            ' Loop library file contents
            Do While LibraryReader.Peek() >= 0

                ' Add to array
                LibraryData.Add(LibraryReader.ReadLine())

            Loop

            ' Write debug (suppress)
            'Logger.WriteDebug(CallStack, "Close file: " + LibraryFile)

            ' Close file
            LibraryReader.Close()

            ' *****************************
            ' - Process library file.
            ' *****************************

            ' Iterate the library (Yes, not count -1)
            For i = 0 To LibraryData.Count

                ' Check i's position
                If i >= LibraryData.Count - 1 Then

                    ' Loop while trailing whitespace
                    While String.IsNullOrWhiteSpace(LibraryData.Item(LibraryData.Count - 1))

                        ' Remove trailing whitespace
                        LibraryData.RemoveAt(LibraryData.Count - 1)

                    End While

                    ' Add whitespace
                    LibraryData.Add("")

                    ' Stop processing
                    Exit For

                ElseIf i = 0 And String.IsNullOrWhiteSpace(LibraryData.Item(0)) Then

                    ' Remove white space
                    LibraryData.RemoveAt(0)

                    ' Decrement i
                    i -= 1

                    ' Start from the top
                    Continue For

                End If

                ' Read new line
                CurrentLine = LibraryData.Item(i)

                ' Check i position
                If i > 0 Then

                    ' Read previous line
                    PreviousLine = LibraryData.Item(i - 1)

                Else

                    ' Stub previous line
                    PreviousLine = "-1"

                End If

                ' Check current line
                If CurrentLine.ToLower.StartsWith("[") Then

                    ' Check for recognizable tags and required entries
                    If CurrentLine.ToLower.Equals("[locale]") AndAlso
                        i + 1 < LibraryData.Count - 1 AndAlso
                        LibraryData.Item(i + 1).ToLower.Contains("codepage") Then

                        ' Increment i ahead
                        i += 1

                        ' Skip to end
                        Continue For

                    ElseIf i + 4 <= LibraryData.Count - 1 AndAlso
                            (LibraryData.Item(i + 1).ToLower.Contains("path=") OrElse LibraryData.Item(i + 1).ToLower.Contains("path =")) AndAlso
                            (LibraryData.Item(i + 2).ToLower.Contains("itemname=") OrElse LibraryData.Item(i + 2).ToLower.Contains("itemname =")) AndAlso
                            (LibraryData.Item(i + 3).ToLower.Contains("itemversion=") OrElse LibraryData.Item(i + 3).ToLower.Contains("itemversion =")) AndAlso
                            (LibraryData.Item(i + 4).ToLower.Contains("itemtype=") OrElse LibraryData.Item(i + 4).ToLower.Contains("itemtype =")) Then

                        ' Set package properties for quick reference
                        ItemFolder = Globals.SDLibraryFolder + LibraryData.Item(i + 1).Substring(LibraryData.Item(i + 1).IndexOf("=") + 1).Trim
                        ItemName = LibraryData.Item(i + 2).Substring(LibraryData.Item(i + 2).IndexOf("=") + 1).Trim
                        ItemVersion = LibraryData.Item(i + 3).Substring(LibraryData.Item(i + 3).IndexOf("=") + 1).Trim
                        ItemType = LibraryData.Item(i + 4).Substring(LibraryData.Item(i + 4).IndexOf("=") + 1).Trim

                        ' Reset flags
                        FolderExists = False
                        VolumeExists = False
                        RegInfoExists = False
                        ArchivedPackage = False

                        ' Checked for archived package type
                        If ItemType.Equals("6") Then ArchivedPackage = True

                        ' Check if folder exists
                        If System.IO.Directory.Exists(ItemFolder) Then

                            ' Set flag
                            FolderExists = True

                            ' Get the directory listing
                            ItemSubFolders = System.IO.Directory.GetDirectories(ItemFolder)

                            ' Iterate folder list (looking for .vol folder)
                            For x As Integer = 0 To ItemSubFolders.Length - 1

                                ' Check for .vol folder
                                If ItemSubFolders(x).ToLower.EndsWith(".vol") Then

                                    ' Set flag
                                    VolumeExists = True

                                ElseIf ItemSubFolders(x).ToLower.EndsWith("reginfo") Then

                                    ' Check if rsw.dat file exists
                                    If System.IO.File.Exists(ItemSubFolders(x) + "\rsw.dat") Then

                                        ' Open rsw.dat file
                                        rswFileReader = New System.IO.StreamReader(ItemSubFolders(x) + "\rsw.dat")

                                        ' Iterate rsw.dat file contents
                                        Do While rswFileReader.Peek() >= 0

                                            ' Read line from rsw.dat
                                            rswFileLine = rswFileReader.ReadLine()

                                            ' Look for package uuid
                                            If rswFileLine.ToLower.Contains("uuid =") Then

                                                ' Set ItemUUID (rsw uuid from rsw.dat file)
                                                ItemUUID = rswFileLine.Substring(rswFileLine.ToLower.IndexOf("uuid = ") + 7).ToLower

                                                'Set flag
                                                RegInfoExists = True

                                            End If

                                            ' Check exist condition
                                            If RegInfoExists Then Exit Do

                                        Loop

                                        ' Close file
                                        rswFileReader.Close()

                                    End If

                                End If

                                ' Check for exit condition
                                If (VolumeExists AndAlso RegInfoExists) OrElse (ArchivedPackage AndAlso RegInfoExists) Then Exit For

                            Next

                        Else

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Missing arc folder: " + ItemName + " " + ItemVersion)

                            ' Increment counter
                            DctMissingArcCount += 1

                        End If

                        ' Check if volume and reginfo folders were found
                        If FolderExists AndAlso Not VolumeExists AndAlso Not ArchivedPackage Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Missing volume folder: " + ItemName + " " + ItemVersion)

                            ' Increment counter
                            DctMissingVolCount += 1

                        ElseIf FolderExists AndAlso Not RegInfoExists Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Missing reginfo\rsw.dat file: " + ItemName + " " + ItemVersion)

                            ' Increment counter
                            DctMissingRegInfoCount += 1

                        End If

                        ' Perform validity checks before committing registered software
                        If Not (FolderExists AndAlso (VolumeExists OrElse ArchivedPackage) AndAlso RegInfoExists) Then

                            ' Remove the entire entry
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)

                            ' Decrement i
                            If i > 0 Then i -= 1

                            ' Skip to end
                            Continue For

                        ElseIf Not DbRegSwUUID.Contains(ItemUUID.ToLower) Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Missing from database: " + ItemName + " " + ItemVersion)

                            ' Remove the entire entry
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)

                            ' Decrement i
                            If i > 0 Then i -= 1

                            ' Increment counter
                            DctNotInDbCount += 1

                            ' Skip to end
                            Continue For

                        ElseIf DctRegSw.Contains(ItemUUID.ToLower) Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Duplicate in library: " + ItemName + " " + ItemVersion + " (UUID=" + ItemUUID + ")")

                            ' Remove the entire entry
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)
                            LibraryData.RemoveAt(i)

                            ' Decrement i
                            If i > 0 Then i -= 1

                            ' Increment counter
                            DctDuplicateCount += 1

                            ' Skip to end
                            Continue For

                        Else

                            ' Add to library registered software
                            DctRegSw.Add(ItemUUID.ToLower)

                            ' Save arc folder name (we'll need it later to ensure arc folder is not removed)
                            DctRegSwFolder.Add(LibraryData.Item(i + 1).Substring(LibraryData.Item(i + 1).IndexOf("=") + 1).Trim.ToLower)

                            ' Write Debug (suppress)
                            'Logger.WriteDebug(CallStack, "ItemUUID: " + ItemUUID)
                            'Logger.WriteDebug(CallStack, "ItemFolder: " + ItemFolder)
                            'Logger.WriteDebug(CallStack, "ItemName: " + ItemName)
                            'Logger.WriteDebug(CallStack, "ItemVersion: " + ItemVersion)
                            'Logger.WriteDebug(CallStack, "ItemType: " + ItemType)

                            ' Check previous line for whitespace
                            If Not PreviousLine.Equals("-1") AndAlso Not String.IsNullOrWhiteSpace(PreviousLine) Then


                                ' Add whitespace
                                LibraryData.Insert(i, "")

                                ' Increment i
                                i += 5

                            Else

                                ' Increment i
                                i += 4

                            End If

                            ' Skip to end
                            Continue For

                        End If

                        ' Reset variables
                        ItemUUID = Nothing
                        ItemFolder = Nothing
                        ItemName = Nothing
                        ItemVersion = Nothing
                        ItemType = Nothing
                        ItemSubFolders = Nothing
                        FolderExists = False
                        VolumeExists = False
                        ArchivedPackage = False
                        RegInfoExists = False

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Invalid entry: " + CurrentLine)

                        ' Remove it
                        LibraryData.RemoveAt(i)

                        ' Decrement i
                        If i > 0 Then i -= 1

                        ' Skip to end
                        Continue For

                    End If

                ElseIf String.IsNullOrWhiteSpace(CurrentLine) AndAlso String.IsNullOrWhiteSpace(PreviousLine) Then

                    ' Remove current line
                    LibraryData.RemoveAt(i)

                    ' Decrement i
                    i -= 1

                ElseIf String.IsNullOrWhiteSpace(CurrentLine) AndAlso Not String.IsNullOrWhiteSpace(PreviousLine) Then

                    ' Do nothing -- single CR/LF is fine.
                    Continue For

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Invalid entry: " + CurrentLine)

                    ' Remove current line
                    LibraryData.RemoveAt(i)

                    ' Decrement i
                    If i > 0 Then i -= 1

                End If

            Next

            ' *****************************
            ' - Process database for orphans (ITCM manager only).
            ' *****************************

            ' ITCM manager only (manager has full library)
            If ITCMManager Then

                ' Iterate database registered software
                For i As Integer = 0 To DbRegSwUUID.Count - 1

                    ' Check if database software was captured in the library file
                    If Not DctRegSw.Contains(DbRegSwUUID.Item(i).ToLower) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Database orphan: " + DbRegSwItemName.Item(i) + " (UUID=" + DbRegSwUUID.Item(i) + ")")

                        ' Add to removal list
                        DbNotInDct.Add(DbRegSwUUID.Item(i))

                    End If

                Next

            End If

            ' *****************************
            ' - Sweep for orphaned folders.
            ' *****************************

            ' Get the directory listing
            FolderList = System.IO.Directory.GetDirectories(Globals.SDLibraryFolder)

            ' Iterate the folder list
            For i As Integer = 0 To FolderList.Length - 1

                ' Check for .arc folder
                If FolderList(i).ToLower.EndsWith(".arc") Then

                    ' Iterate valid library arcs
                    For j As Integer = 0 To DctRegSwFolder.Count - 1

                        ' Check for a match (folder matches software in library and is not on the removal list)
                        If FolderList(i).Substring(FolderList(i).LastIndexOf("\") + 1).ToLower.Equals(DctRegSwFolder.Item(j).ToLower) AndAlso
                            Not DbNotInDct.Contains(FolderList(i).Substring(FolderList(i).LastIndexOf("\") + 1).ToLower.Replace(".arc", "")) Then

                            ' Set flag
                            MatchFound = True

                            ' Exit inner loop
                            Exit For

                        End If

                    Next

                    ' Check if a match was not found
                    If Not MatchFound Then

                        ' Add orphan to cleanup list
                        ArcsForDeletion.Add(FolderList(i))

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Library orphan: " + FolderList(i).Replace(Globals.SDLibraryFolder, ""))

                    End If

                    ' Reset flag
                    MatchFound = False

                ElseIf ITCMManager AndAlso DbSwProcs.Count > 0 AndAlso FolderList(i).ToLower.EndsWith(".rcp") Then

                    ' Iterate library arcs
                    For j As Integer = 0 To DbSwProcs.Count - 1

                        ' Check for a match
                        If FolderList(i).Substring(FolderList(i).LastIndexOf("\") + 1).ToLower.Replace(".rcp", "").Equals(DbSwProcs.Item(j).ToLower) Then

                            ' Set flag
                            MatchFound = True

                            ' Exit inner loop
                            Exit For

                        End If

                    Next

                    ' Check if a match was not found
                    If Not MatchFound Then

                        ' Add orphan to cleanup list
                        RcpsForDeletion.Add(FolderList(i))

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Library orphan: " + FolderList(i).Replace(Globals.SDLibraryFolder, ""))

                    End If

                    ' Reset flag
                    MatchFound = False

                End If

            Next

            ' *****************************
            ' - Report results.
            ' *****************************

            ' Write debug
            Logger.WriteDebug(CallStack, "Software read from database: " + DbRegSwUUID.Count.ToString)
            Logger.WriteDebug(CallStack, "Software read from library: " + (DctRegSw.Count + DctMissingArcCount + DctMissingVolCount + DctMissingRegInfoCount + DctDuplicateCount + DctNotInDbCount).ToString)
            Logger.WriteDebug(CallStack, "Library items to cleanup: " + (DctMissingArcCount + DctMissingVolCount + DctMissingRegInfoCount + DctDuplicateCount + DctNotInDbCount).ToString)
            Logger.WriteDebug(CallStack, "Library folders to cleanup: " + (ArcsForDeletion.Count + RcpsForDeletion.Count).ToString)
            Logger.WriteDebug(CallStack, "Database items to cleanup: " + DbNotInDct.Count.ToString)

            ' Add to library events
            LibraryEvents.Add("Software read from database: " + DbRegSwUUID.Count.ToString)
            LibraryEvents.Add("Software read from library: " + (DctRegSw.Count + DctMissingArcCount + DctMissingVolCount + DctMissingRegInfoCount + DctDuplicateCount + DctNotInDbCount).ToString)
            LibraryEvents.Add("Library items to remove: " + (DctMissingArcCount + DctMissingVolCount + DctMissingRegInfoCount + DctDuplicateCount + DctNotInDbCount).ToString)
            LibraryEvents.Add("Library folders to cleanup: " + (ArcsForDeletion.Count + RcpsForDeletion.Count).ToString)
            LibraryEvents.Add("Database items to cleanup: " + DbNotInDct.Count.ToString)

            ' Check for library cleanup switch
            If Globals.CleanupSDLibrarySwitch AndAlso Not Globals.CheckSDLibrarySwitch Then

                ' *****************************
                ' - Cleanup database.
                ' *****************************

                ' Check for database cleanup items
                If DbNotInDct.Count > 0 Then

                    ' Local counters
                    Dim SoftwareAffectedCount As Integer = 0
                    Dim ProcedureAffectedCount As Integer = 0
                    Dim ApplicationAffectedCount As Integer = 0
                    Dim GroupLinksAffectedCount As Integer = 0
                    Dim RetryCount As Integer = 0
                    Dim RetryLimit As Integer = 3
                    Dim WasSuccessful As Boolean = False

                    ' Encapsulate sql retry loop
                    Do

                        ' Encapsulate sql
                        Try

                            ' Iterate database cleanup items
                            For i As Integer = DbNotInDct.Count - 1 To 0 Step -1

                                ' Convert the uniqueidentifier into string (really binary(16) for the package's usd_rsw.objectid)
                                Dim rsw_objectid As String = "0x" + BitConverter.ToString(Guid.Parse(DbNotInDct.Item(i)).ToByteArray).Replace("-", "")

                                ' Retrieve software properties
                                Dim rsw_itemname As String = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select itemname from usd_rsw where objectid=" + rsw_objectid)
                                Dim rsw_itemversion As String = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select itemversion from usd_rsw where objectid=" + rsw_objectid)

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Delete from database: " + rsw_itemname + " " + rsw_itemversion + "(objectid=" + rsw_objectid + ")")

                                ' Retrive list of procedures associated with the rsw
                                Dim actproc_list As List(Of String) = DatabaseAPI.SqlSelectScalarList(CallStack, DbConnection, "select objectid from usd_actproc where rsw=" + rsw_objectid)

                                ' Iterate list of procedures associated with rsw
                                For Each actproc_uuid As String In actproc_list

                                    ' Retrieve procedure name
                                    Dim actproc_name As String = DatabaseAPI.SqlSelectScalar(CallStack, DbConnection, "select itemname from usd_actproc where objectid=" + actproc_uuid)

                                    ' Delete any associated application
                                    SqlCmd = New SqlCommand("delete from usd_applic where actproc=" + actproc_uuid, DbConnection)
                                    SqlData = SqlCmd.ExecuteReader()

                                    ' Increment record counter
                                    ApplicationAffectedCount += SqlData.RecordsAffected

                                    ' Close sql data reader
                                    SqlData.Close()

                                    ' Delete the procedure
                                    SqlCmd = New SqlCommand("delete from usd_actproc where objectid=" + actproc_uuid, DbConnection)
                                    SqlData = SqlCmd.ExecuteReader()

                                    ' Increment record counter
                                    ProcedureAffectedCount += SqlData.RecordsAffected

                                    ' Close sql data reader
                                    SqlData.Close()

                                Next

                                ' Delete the software
                                SqlCmd = New SqlCommand("delete from usd_rsw where objectid=" + rsw_objectid, DbConnection)
                                SqlData = SqlCmd.ExecuteReader()

                                ' Increment record counter
                                SoftwareAffectedCount += SqlData.RecordsAffected

                                ' Close sql data reader
                                SqlData.Close()

                                ' Delete the software group links
                                SqlCmd = New SqlCommand("delete from usd_link_swg_sw where sw=" + rsw_objectid, DbConnection)
                                SqlData = SqlCmd.ExecuteReader()

                                ' Increment record counter
                                GroupLinksAffectedCount += SqlData.RecordsAffected

                                ' Close sql data reader
                                SqlData.Close()

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Software deleted: " + SoftwareAffectedCount.ToString)
                                Logger.WriteDebug(CallStack, "Procedures deleted: " + ProcedureAffectedCount.ToString)
                                Logger.WriteDebug(CallStack, "Applications deleted: " + ApplicationAffectedCount.ToString)
                                Logger.WriteDebug(CallStack, "Group links deleted: " + GroupLinksAffectedCount.ToString)

                                ' Add to library events
                                LibraryEvents.Add("Software deleted: " + SoftwareAffectedCount.ToString)
                                LibraryEvents.Add("Procedures deleted: " + ProcedureAffectedCount.ToString)
                                LibraryEvents.Add("Applications deleted: " + ApplicationAffectedCount.ToString)
                                LibraryEvents.Add("Group links deleted: " + GroupLinksAffectedCount.ToString)

                                ' Dequeue element (to avoid unnecessary retries)
                                DbNotInDct.RemoveAt(i)

                                ' Reset counters
                                SoftwareAffectedCount = 0
                                ProcedureAffectedCount = 0
                                ApplicationAffectedCount = 0
                                GroupLinksAffectedCount = 0
                                RetryCount = 0

                            Next

                            ' All items processed -- set success flag
                            WasSuccessful = True

                        Catch ex As Exception

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Database exception: " + ex.Message)

                            ' Check for timeout
                            If ex.Message.ToLower.Contains("execution timeout expired") OrElse ex.Message.ToLower.Contains("timeout period") Then

                                ' Update retry count
                                RetryCount += 1

                                ' Rest thread
                                Thread.Sleep(2000)

                            Else

                                ' Don't retry
                                RetryCount = RetryLimit

                            End If

                        End Try

                    Loop Until WasSuccessful = True OrElse RetryCount >= RetryLimit

                End If

                ' *****************************
                ' - Cleanup library file.
                ' *****************************

                ' Check for library changes
                If (DctMissingArcCount + DctMissingVolCount + DctMissingRegInfoCount + DctDuplicateCount + DctNotInDbCount) > 0 Then

                    ' *****************************
                    ' - Increment library backup.
                    ' *****************************

                    ' Check for existing secondary backup file
                    If System.IO.File.Exists(LibraryFileBackup2) Then

                        ' Delete file
                        Utility.DeleteFile(CallStack, LibraryFileBackup2)

                    End If

                    ' Check for existing primary backup file
                    If System.IO.File.Exists(LibraryFileBackup1) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Rename file: " + LibraryFileBackup1)
                        Logger.WriteDebug(CallStack, "To: " + LibraryFileBackup2)

                        ' Rename the file
                        System.IO.File.Move(LibraryFileBackup1, LibraryFileBackup2)

                    End If

                    ' Backup current library file
                    If System.IO.File.Exists(LibraryFile) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Copy file: " + LibraryFile)
                        Logger.WriteDebug(CallStack, "To: " + LibraryFileBackup1)

                        ' Copy file
                        System.IO.File.Copy(LibraryFile, LibraryFileBackup1)

                    End If

                    ' *****************************
                    ' - Write new library file.
                    ' *****************************

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Write library file: " + LibraryFile)

                    ' Open library file for overwrite
                    LibraryFileWriter = New IO.StreamWriter(LibraryFile, False)

                    ' Iterate the library data
                    For Each LibraryLine As String In LibraryData

                        ' Write the line to file
                        LibraryFileWriter.WriteLine(LibraryLine)

                    Next

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Close file: " + LibraryFile)

                    ' Close file
                    LibraryFileWriter.Close()

                End If

                ' *****************************
                ' - Cleanup folders.
                ' *****************************

                ' Cleanup arc folders
                For Each folder As String In ArcsForDeletion

                    ' Delete folder
                    Utility.DeleteFolder(CallStack, folder)

                Next

                ' Cleanup rcp folders
                For Each folder As String In RcpsForDeletion

                    ' Delete folder
                    Utility.DeleteFolder(CallStack, folder)

                Next

                ' *****************************
                ' - Sync staging library (ITCM servers only).
                ' *****************************

                ' Check itcm function
                If ITCMServer Then

                    ' Sync staging library
                    If Utility.IsProcessRunning("caf") Then

                        ' Write Debug
                        Logger.WriteDebug(CallStack, "Library sync: sd_sscmd.exe deldctbu")

                        ' Run library sync command
                        ReturnCode = Utility.RunCommand("sd_sscmd.exe", "deldctbu", False)

                        ' Write Debug
                        Logger.WriteDebug(CallStack, "Return code: " + ReturnCode.ToString)

                        ' Add to library events
                        LibraryEvents.Add("Library sync: " + "[Return Code: " + ReturnCode.ToString + "]")

                    End If

                End If

            End If

            ' *****************************
            ' - Close database connection.
            ' *****************************

            ' Close the database connection
            DbConnection.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "Database connection closed.")

        End Sub

    End Class

End Class