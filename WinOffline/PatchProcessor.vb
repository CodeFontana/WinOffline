Partial Public Class WinOffline

    Public Shared Function PatchProcessor(ByVal CallStack As String) As Integer

        ' Local variables
        Dim FileList As String()            ' Directory listing of files.
        Dim PatchFileName As String         ' Name of patch file (i.e. C:\SomeDirectory\SomePatch.caz).
        Dim PatchPath As String             ' Path to patch file (i.e. C:\SomeDirectory).
        Dim PatchShortName As String        ' Name of patch file (i.e. SomePatch.caz).
        Dim PatchFriendlyName As String     ' Name of patch file (i.e. SomePatch).
        Dim RunLevel As Integer = 0         ' Tracks the health of the function and calls to external functions.

        ' Update call stack
        CallStack += "PatchProcessor|"

        ' *****************************
        ' - Get a directory listing.
        ' *****************************

        ' Get the directory listing
        FileList = System.IO.Directory.GetFiles(Globals.WorkingDirectory)

        ' *****************************
        ' - Process patch files.
        ' *****************************

        ' Iterate the file list
        For n As Integer = 0 To FileList.Length - 1

            ' Check for patch files
            If FileList(n).ToLower.EndsWith(".caz") Or
                FileList(n).ToLower.EndsWith(".jcl") Then

                ' Get all versions of the filename
                PatchFileName = FileList(n)
                PatchPath = FileVector.GetFilePath(PatchFileName)
                PatchShortName = FileVector.GetShortName(PatchFileName)
                PatchFriendlyName = FileVector.GetFriendlyName(PatchFileName)

                ' Write debug
                Logger.WriteDebug(CallStack, "Detected patch file: " + PatchShortName)

                ' Check for duplicate patch
                If Manifest.PatchManifestContains(PatchFriendlyName) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Warning: Patch name is duplicate, skipping: " + PatchFriendlyName)

                    ' Continue loop
                    Continue For

                End If

                ' Route CAZ files through decompression
                If PatchShortName.ToLower.EndsWith(".caz") Then

                    ' *****************************
                    ' - Decompress the CAZ file.
                    ' *****************************

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Decompress patch: " + PatchShortName)

                    ' Decompress the CAZ file
                    RunLevel = DecompressPatch(CallStack, PatchFileName)

                    ' Check result
                    If RunLevel <> 0 Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Error: Decompression failure.")

                        ' Return
                        Return 1

                    End If

                ElseIf PatchShortName.ToLower.EndsWith(".jcl") Then

                    ' *****************************
                    ' - Migrate the JCL file.
                    ' *****************************

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Migrate patch: " + PatchShortName)

                    ' Migrate the JCL file
                    RunLevel = MigratePatch(CallStack, PatchFileName)

                    ' Check result
                    If RunLevel <> 0 Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Error: Migration failure.")

                        ' Return
                        Return 2

                    End If

                End If

                ' *****************************
                ' - Process the patch JCL file.
                ' *****************************

                ' Process the patch
                RunLevel = ProcessPatch(CallStack, PatchFriendlyName)

                ' Check result
                If RunLevel <> 0 Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Failed to process the patch.")

                    ' Return
                    Return 3

                End If

            End If

        Next

        ' Write debug
        Logger.WriteDebug(CallStack, " (" + Manifest.PatchManifestCount.ToString + ") patches processed.")

        ' Return
        Return RunLevel

    End Function

    Public Shared Function MigratePatch(ByVal CallStack As String,
                                        ByVal JCLFileName As String) As Integer

        ' Local variables
        Dim TargetFileName As String        ' Destination filename (e.g. C:\Windows\Temp\WinOffline\SomePatch\SomePatch.jcl).
        Dim RunLevel As Integer = 0         ' Tracks the health of the function and calls to external functions.

        ' Update call stack
        CallStack += "Migrate|"

        ' *****************************
        ' - Create unique patch folder and copy JCL to the new folder.
        ' *****************************

        ' Calculate destination filename
        TargetFileName = Globals.WinOfflineTemp + "\" + FileVector.GetFriendlyName(JCLFileName) + "\" + FileVector.GetShortName(JCLFileName)

        ' Prepare migration destination
        Try

            ' Write debug
            Logger.WriteDebug(CallStack, "Create folder: " + FileVector.GetFilePath(TargetFileName))

            ' Create patch directory
            System.IO.Directory.CreateDirectory(FileVector.GetFilePath(TargetFileName))

            ' Write debug
            Logger.WriteDebug(CallStack, "Copy File: " + JCLFileName)
            Logger.WriteDebug(CallStack, "To: " + TargetFileName)

            ' Copy patch files to directory
            System.IO.File.Copy(JCLFileName, TargetFileName, True)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught migrating patch to destination.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 1

        End Try

        ' Return
        Return RunLevel

    End Function

    Public Shared Function DecompressPatch(ByVal CallStack As String,
                                           ByVal CAZFileName As String) As Integer

        ' Local variables
        Dim TargetFileName As String                        ' Destination filename (e.g. C:\Windows\Temp\WinOffline\SomePatch\SomePatch.caz).
        Dim ExecutionString As String                       ' Command line to be executed externally to the application.
        Dim ArgumentString As String                        ' Arguments passed on the command line for the external execution.
        Dim CAZipXPProcessStartInfo As ProcessStartInfo     ' Process startup settings for configuring the bahviour of the process.
        Dim RunningProcess As Process                       ' A process shell for executing the command line.
        Dim ExitCode As Integer                             ' External process return code.
        Dim FileList As String()                            ' Directory listing of patch files.
        Dim JCLShortName As String                          ' Name of jcl file found in caz (i.e. SomePatch.jcl)
        Dim JCLCounter As Integer = 0                       ' Counter of JCL files.
        Dim RunLevel As Integer = 0                         ' Tracks the health of the function and calls to external functions.

        ' Update call stack
        CallStack += "Decompress|"

        ' *****************************
        ' - Create unique patch folder and copy CAZ to the new folder.
        ' *****************************

        ' Calculate names
        TargetFileName = Globals.WinOfflineTemp + "\" + FileVector.GetFriendlyName(CAZFileName) + "\" + FileVector.GetShortName(CAZFileName)

        ' Prepare decompression destination
        Try

            ' Write debug
            Logger.WriteDebug(CallStack, "Create folder: " + FileVector.GetFilePath(TargetFileName))

            ' Create patch directory
            System.IO.Directory.CreateDirectory(FileVector.GetFilePath(TargetFileName))

            ' Write debug
            Logger.WriteDebug(CallStack, "Copy File: " + CAZFileName)
            Logger.WriteDebug(CallStack, "To: " + TargetFileName)

            ' Copy patch files to directory
            System.IO.File.Copy(CAZFileName, TargetFileName, True)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught preparing patch destination.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 1

        End Try

        ' *****************************
        ' - Decompress the CAZ file
        ' *****************************

        ' Build cazipxp execution string
        ExecutionString = Globals.WinOfflineTemp + "\cazipxp.exe"
        ArgumentString = "-u " + """" + TargetFileName + """"

        ' Write debug
        Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

        ' Create detached process for cazipxp
        CAZipXPProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
        CAZipXPProcessStartInfo.WorkingDirectory = FileVector.GetFilePath(TargetFileName)
        CAZipXPProcessStartInfo.UseShellExecute = False
        CAZipXPProcessStartInfo.RedirectStandardOutput = True
        CAZipXPProcessStartInfo.CreateNoWindow = True

        ' Start detached process
        RunningProcess = Process.Start(CAZipXPProcessStartInfo)

        ' Wait for detached process to exit
        RunningProcess.WaitForExit()

        ' Write debug
        Logger.WriteDebug("------------------------------------------------------------")
        Logger.WriteDebug(RunningProcess.StandardOutput.ReadToEnd.ToString)
        Logger.WriteDebug("------------------------------------------------------------")

        ' Get the exitcode
        ExitCode = RunningProcess.ExitCode

        ' Write debug
        Logger.WriteDebug(CallStack, "Exit code: " + ExitCode.ToString)

        ' Close detached process
        RunningProcess.Close()

        ' *****************************
        ' - Verify decompression was successful.
        ' *****************************

        ' Check cazipxp exit code
        If ExitCode <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Cazipxp utility decompression failed.")

            ' Create exception
            Manifest.UpdateManifest(CallStack,
                                    Manifest.EXCEPTION_MANIFEST,
                                    {"Error: Cazipxp utility decompression failed.",
                                    "Reason: A non-zero exit code was returned from the utility."})

            ' Return -- Failed to decompress .caz file
            Return 2

        End If

        ' *****************************
        ' - Verify CAZ extracted a single JCL file.
        ' *****************************

        ' Get the patch directory file listing
        FileList = System.IO.Directory.GetFiles(FileVector.GetFilePath(TargetFileName))

        ' Loop the file list -- Check and decompress any CAZ files
        For n As Integer = 0 To FileList.Length - 1

            ' Check for a JCL file
            If FileList(n).ToLower.Contains(".jcl") Then

                ' Read the JCL name
                JCLShortName = FileVector.GetShortName(FileList(n))

                ' Write debug
                Logger.WriteDebug(CallStack, "Detected instruction file: " + JCLShortName)

                ' Increment the counter
                JCLCounter = JCLCounter + 1

            End If

        Next

        ' Check the JCL counter
        If JCLCounter = 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Patch does not contain an instruction file.")

            ' Create exception
            Manifest.UpdateManifest(CallStack,
                                    Manifest.EXCEPTION_MANIFEST,
                                    {"Error: Patch is missing an instruction (.jcl) file.",
                                    "Please verify the contents of the " + FileVector.GetFriendlyName(CAZFileName) + " patch."})

            ' Return -- No JCL contained in CAZ.
            Return 3

        ElseIf JCLCounter > 1 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Patch file contains multiple instruction files.")

            ' Create exception
            Manifest.UpdateManifest(CallStack,
                                    Manifest.EXCEPTION_MANIFEST,
                                    {"Error: Patch file contains multiple instruction (.jcl) files.",
                                    "Please verify the contents of the " + FileVector.GetFriendlyName(CAZFileName) + " patch."})

            ' Return -- Too many JCLs in CAZ.
            Return 4

        Else

            ' Write debug
            Logger.WriteDebug(CallStack, "Verified: Patch contains only a single instruction file.")

        End If

        ' Return
        Return RunLevel

    End Function

    Public Shared Function ProcessPatch(ByVal CallStack As String,
                                        ByVal PatchFriendlyName As String) As Integer

        ' Local variables
        Dim PatchFileName As String = ""                    ' Patch file being processed.
        Dim PatchStreamReader As System.IO.StreamReader     ' Input stream for reading the patch file.
        Dim InstructionList As New ArrayList                ' Array for storing instruction file.
        Dim NewPatch As PatchVector                         ' Resulting patch object.
        Dim strLine As String                               ' String for parsing instructions.
        Dim FileList As String()                            ' Directory listing of files.
        Dim DependentFlag As Boolean = False                ' Flag for processing patch dependencies.
        Dim RunLevel As Integer = 0                         ' Tracks the health of the function and calls to external functions.

        ' Update call stack
        CallStack += "Process|"

        ' *****************************
        ' - Locate the patch instruction file.
        ' *****************************

        ' Get the patch directory file listing
        FileList = System.IO.Directory.GetFiles(Globals.WinOfflineTemp + "\" + PatchFriendlyName)

        ' Loop the file list -- find the .jcl file
        For n As Integer = 0 To FileList.Length - 1

            ' Match the JCL file
            If FileList(n).ToLower.Contains(".jcl") Then

                ' Assign .jcl file
                PatchFileName = FileList(n)

            End If

        Next

        ' Verify a .jcl was located
        If PatchFileName.Equals("") Then Return 1

        ' *****************************
        ' - Open the patch instruction file.
        ' *****************************

        Try

            ' Write debug
            Logger.WriteDebug(CallStack, "Open file: " + PatchFileName)

            ' Open the .jcl reader stream
            PatchStreamReader = New System.IO.StreamReader(PatchFileName)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught opening patch instruction file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 2

        End Try

        ' *****************************
        ' - Read the instructions file.
        ' *****************************

        ' Write debug
        Logger.WriteDebug("############################################################")

        ' Loop patch file contents
        Do While PatchStreamReader.Peek() >= 0

            ' Read a line
            strLine = PatchStreamReader.ReadLine()

            ' Write debug -- Write file contents only
            Logger.WriteDebug(strLine)

            ' Update the array
            InstructionList.Add(strLine)

        Loop

        ' Write debug
        Logger.WriteDebug("############################################################")
        Logger.WriteDebug(CallStack, "Close file: " + PatchFileName)

        ' Close the patch stream reader
        PatchStreamReader.Close()

        ' *****************************
        ' - Verify the patch instructions
        ' *****************************

        ' Create a new patch object
        NewPatch = New PatchVector(PatchFileName, InstructionList)

        ' Verify a product code is specified
        If NewPatch.GetInstruction("product").Equals("") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Patch file is missing a product code (i.e. 'PRODUCT:' tag).")

            ' Create exception
            Manifest.UpdateManifest(CallStack,
                                    Manifest.EXCEPTION_MANIFEST,
                                    {"Error: Patch file is missing a product code (i.e. 'PRODUCT:' tag).",
                                    "Please verify the " + NewPatch.PatchFile.GetShortName + " patch."})

            ' Return
            Return 3

        End If

        ' Verify patch contains a supported product code
        If Not (NewPatch.IsClientAuto Or
                NewPatch.IsSharedComponent Or
                NewPatch.IsCAM Or
                NewPatch.IsSSA Or
                NewPatch.IsDataTransport Or
                NewPatch.IsExplorerGUI) Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Unsupported product code.")

            ' Create exception
            Manifest.UpdateManifest(CallStack,
                                    Manifest.EXCEPTION_MANIFEST,
                                    {"Error: Unsupported product code.",
                                    "Please verify the PRODUCT: tag in the patch file."})

            ' Return
            Return 4

        End If

        ' Verify the patch contains at least one valid instruction
        If NewPatch.GetShortNameReplaceList.Count +
            NewPatch.GetPreCommandList.Count +
            NewPatch.GetSysCommandList.Count = 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Patch file is missing instructions (i.e. 'FILE:', 'PRESYSCMD:' or 'SYSCMD:' tags).")

            ' Create exception
            Manifest.UpdateManifest(CallStack,
                                    Manifest.EXCEPTION_MANIFEST,
                                    {"Error: Patch file is missing instructions (i.e. 'FILE:', 'PRESYSCMD:' or 'SYSCMD:' tags).",
                                    "Please verify the " + NewPatch.PatchFile.GetShortName + " patch."})

            ' Return
            Return 5

        End If

        ' *****************************
        ' - Resolve dependent files.
        ' *****************************

        ' Iterate the file replacement list
        For Each Dependent As String In NewPatch.GetAllDependentFiles

            ' Reset dependency flag
            DependentFlag = False

            ' List files from the patch folder
            FileList = System.IO.Directory.GetFiles(NewPatch.PatchFile.GetFilePath)

            ' Loop the file list -- Resolve file dependency
            For n As Integer = 0 To FileList.Length - 1

                ' Check for a match
                If FileList(n).ToLower.Contains(Dependent.ToLower) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Dependency verified: " + Dependent)

                    ' Set flag
                    DependentFlag = True

                    ' Break loop
                    Exit For

                End If

            Next

            ' Check flag, move to next dependency
            If DependentFlag Then Continue For

            ' List files from the working directory
            FileList = System.IO.Directory.GetFiles(Globals.WorkingDirectory)

            ' Loop the file list -- Resolve file dependency
            For n As Integer = 0 To FileList.Length - 1

                ' Check for a match
                If FileList(n).ToLower.Contains(Dependent.ToLower) Then

                    ' Copy dependent file
                    Try

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Copy File: " + Globals.WorkingDirectory + Dependent)
                        Logger.WriteDebug(CallStack, "To: " + NewPatch.PatchFile.GetFilePath + "\" + Dependent)

                        ' Copy required file
                        System.IO.File.Copy(Globals.WorkingDirectory + Dependent, NewPatch.PatchFile.GetFilePath + "\" + Dependent, True)

                    Catch ex As Exception

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Error: Exception caught copying dependent file.")
                        Logger.WriteDebug(ex.Message)
                        Logger.WriteDebug(ex.StackTrace)

                        ' Create exception
                        Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                        ' Return
                        Return 6

                    End Try

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Dependency verified: " + Dependent)

                    ' Set flag
                    DependentFlag = True

                    ' Break loop
                    Exit For

                End If

            Next

            ' Dependent not found
            If DependentFlag = False Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Missing dependency: " + Dependent)

                ' Create exception
                Manifest.UpdateManifest(CallStack,
                                        Manifest.EXCEPTION_MANIFEST,
                                        {"Error: Missing dependency: " + Dependent,
                                        "Please verify the " + NewPatch.PatchFile.GetShortName + " patch."})

                ' Return
                Return 7

            End If

        Next

        ' *****************************
        ' - Update patch manifest.
        ' *****************************

        Logger.WriteDebug(CallStack, "Patch file: " + NewPatch.PatchFile.GetShortName)
        Logger.WriteDebug(CallStack, "Patch folder: " + NewPatch.PatchFile.GetFilePath)
        Logger.WriteDebug(CallStack, "Product Tag: " + NewPatch.GetInstruction("product"))

        ' Commit patch to manifest
        Manifest.UpdateManifest(CallStack, Manifest.PATCH_MANIFEST, NewPatch)

        ' Return
        Return RunLevel

    End Function

End Class