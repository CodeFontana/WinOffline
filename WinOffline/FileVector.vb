'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline/FileVector
' File Name:    FileVector.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Class FileVector

        Private FileName As String          ' Ex: C:\SomeDirectory\SomePatch.jcl
        Private FilePath As String          ' Ex: C:\SomeDirectory
        Private ShortName As String         ' Ex: SomePatch.jcl
        Private FriendlyName As String      ' Ex: SomePatch

        Public Sub New(ByVal TargetFile As String)
            FileName = TargetFile
            FilePath = TargetFile.Substring(0, TargetFile.LastIndexOf("\"))
            ShortName = TargetFile.Substring(TargetFile.LastIndexOf("\") + 1)
            FriendlyName = ShortName.Substring(0, ShortName.IndexOf("."))
        End Sub

        Public Function GetFileName() As String
            Return FileName
        End Function

        Public Function GetFilePath() As String
            Return FilePath
        End Function

        Public Function GetShortName() As String
            Return ShortName
        End Function

        Public Function GetFriendlyName() As String
            Return FriendlyName
        End Function

        Public Sub SetFullPath(ByVal NewFilePath As String)
            FileName = NewFilePath
        End Sub

        Public Sub SetFilePath(ByVal NewSourceDir As String)
            FilePath = NewSourceDir
        End Sub

        Public Sub SetShortName(ByVal NewShortName As String)
            ShortName = NewShortName
        End Sub

        Public Sub SetFriendlyName(ByVal NewFriendlyName As String)
            FriendlyName = NewFriendlyName
        End Sub

        Public Shared Function GetFilePath(ByVal FileName As String) As String
            If FileName.Contains("\") Then
                Return FileName.Substring(0, FileName.LastIndexOf("\"))
            Else
                Return Nothing
            End If
        End Function

        Public Shared Function GetShortName(ByVal FileName As String) As String
            Return FileName.Substring(FileName.LastIndexOf("\") + 1)
        End Function

        Public Shared Function GetFriendlyName(ByVal FileName As String) As String
            Dim ShortName As String = GetShortName(FileName)
            Return ShortName.Substring(0, ShortName.IndexOf("."))
        End Function

        Public Shared Function GetFriendlyNameFromShortName(ByVal ShortName As String) As String
            Return ShortName.Substring(0, ShortName.IndexOf("."))
        End Function

    End Class

End Class