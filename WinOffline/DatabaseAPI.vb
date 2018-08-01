'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline/DatabaseAPI
' File Name:    DatabaseAPI.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Threading

Partial Public Class WinOffline

    Public Class DatabaseAPI

        Public Shared ConnectionString As String = Nothing
        Public Shared DatabaseThread As Thread = Nothing
        Public Shared CancelSignal As Boolean = False
        Public Shared DomainUuid As String = Nothing
        Public Shared DomainId As String = Nothing
        Public Shared DomainType As String = Nothing
        Public Shared ManagerName As String = Nothing

        Public Shared Function SQLFunctionDispatch(ByVal Callstack As String) As Integer

            ' Local variables
            Dim ConnectionString As String = Nothing
            Dim DbConnection As SqlConnection = Nothing

            ' Update call stack
            Callstack += "DatabaseAPI|"

            ' *****************************
            ' - Perform verifications.
            ' *****************************

            ' Verify ITCM is installed
            If Not Utility.IsITCMInstalled Then

                ' Write debug
                Logger.WriteDebug(Callstack, "Error: ITCM is not installed.")

                ' Return
                Return 0

            End If

            ' Verify ITCM function
            If Not Globals.ITCMFunction.ToLower.Contains("manager") Then

                ' Write debug
                Logger.WriteDebug(Callstack, "Error: Database functions only available on ITCM managers.")

                ' Return
                Return 0

            End If

            ' Verify database server
            If Globals.DatabaseServer Is Nothing Then

                ' Write debug
                Logger.WriteDebug(Callstack, "Error: Database server name from comstore is empty.")

                ' Return
                Return 0

            End If

            ' Verify if database user was provided
            If Globals.DbUser Is Nothing Then

                ' Assign current user
                Globals.DbUser = Globals.ProcessIdentity.Name

                ' Write debug
                Logger.WriteDebug(Callstack, "Logon as: " + Globals.DbUser)

            ElseIf Globals.DbPassword Is Nothing Then

                ' Write directly to attached console (this is not logged)
                Console.WriteLine()
                Console.WriteLine(Callstack.Substring(Callstack.IndexOf("DatabaseAPI")) + "Logon as: " + Globals.DbUser)
                Console.Write(Callstack.Substring(Callstack.IndexOf("DatabaseAPI")) + "Enter password: ")

                ' Hook the keyboard for low-level input
                KeyboardHook.SetHook()

                ' Wait for keyboard release
                While (KeyboardHook.KeyboardHooked)

                    ' Do events
                    System.Windows.Forms.Application.DoEvents()

                    ' Rest thread
                    System.Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

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

            ' Encapsulate sql connection
            Try

                ' Initialize sql connection instance
                DbConnection = New SqlConnection(ConnectionString)

                ' Open sql connection
                DbConnection.Open()

                ' Query domain_uuid
                DomainUuid = SqlSelectScalar(Callstack, DbConnection, "select set_val_uuid from ca_settings with (nolock) where set_id=1")
                Logger.WriteDebug(Callstack, "Domain UUID: " + DomainUuid)

                ' Query domain_id
                DomainId = SqlSelectScalar(Callstack, DbConnection, "select domain_id from ca_n_tier with (nolock) where domain_uuid=" + DomainUuid)
                Logger.WriteDebug(Callstack, "Domain ID: " + DomainId)

                ' Query domain_type
                DomainType = SqlSelectScalar(Callstack, DbConnection, "select domain_type from ca_n_tier with (nolock) where domain_uuid=" + DomainUuid)
                Logger.WriteDebug(Callstack, "Domain type: " + DomainType)

                ' Query manager name (label)
                ManagerName = SqlSelectScalar(Callstack, DbConnection, "select label from ca_manager with (nolock) where domain_uuid=" + DomainUuid)
                Logger.WriteDebug(Callstack, "Manager name: " + ManagerName)

                ' Check for utility switches
                If Globals.MdbOverviewSwitch Then

                    ' Write mdb overview
                    Logger.WriteDebug(Callstack, "Domain managers: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_n_tier with (nolock) where domain_type=0"))
                    Logger.WriteDebug(Callstack, "Scalability Servers: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_server with (nolock)"))
                    Logger.WriteDebug(Callstack, "Registered Agents: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=1"))
                    Logger.WriteDebug(Callstack, "Registered Users: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=2"))
                    Logger.WriteDebug(Callstack, "Registered Profiles: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=4"))
                    Logger.WriteDebug(Callstack, "Obsolete Agents [>=90 days]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=1 and last_run_date<DATEDIFF(s,'19700101',GETUTCDATE())-7776000"))
                    Logger.WriteDebug(Callstack, "Obsolete Agents [>=1 year]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=1 and last_run_date<DATEDIFF(s,'19700101',GETUTCDATE())-31536000"))
                    Logger.WriteDebug(Callstack, "Obsolete Users [>=90 days]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=2 and last_run_date<DATEDIFF(s,'19700101',GETUTCDATE())-7776000"))
                    Logger.WriteDebug(Callstack, "Obsolete Users [>=1 year]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=2 and last_run_date<DATEDIFF(s,'19700101',GETUTCDATE())-31536000"))
                    Logger.WriteDebug(Callstack, "Obsolete Profiles [>=90 days]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=4 and last_run_date<DATEDIFF(s,'19700101',GETUTCDATE())-7776000"))
                    Logger.WriteDebug(Callstack, "Obsolete Profiles [>=1 year]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_agent with (nolock) where agent_type=4 and last_run_date<DATEDIFF(s,'19700101',GETUTCDATE())-31536000"))
                    Logger.WriteDebug(Callstack, "Group Definitions [Static]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_group_def with (nolock) where member_type=1 and query_uuid is null"))
                    Logger.WriteDebug(Callstack, "Group Definitions [Dynamic]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_group_def with (nolock) where member_type=1 and query_uuid is not null"))
                    Logger.WriteDebug(Callstack, "Query Definitions [Simple]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_query_def with (nolock)"))
                    Logger.WriteDebug(Callstack, "Query Definitions [Complex]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_query_def_contents with (nolock)"))
                    Logger.WriteDebug(Callstack, "Software Packages: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_rsw with (nolock) where itemtype<>5"))
                    Logger.WriteDebug(Callstack, "Job Containers [In-Progress]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_job_cont with (nolock) where type=3 and comment<>'CA-Software Delivery Reserved Group' and objectid in (select jcont from usd_link_jc_act with (nolock) where activity in (select objectid from usd_activity with (nolock) where state not in (2,3)))"))
                    Logger.WriteDebug(Callstack, "Job Containers [Completed]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_job_cont with (nolock) where type=3 and comment<>'CA-Software Delivery Reserved Group' and objectid in (select jcont from usd_link_jc_act with (nolock) where activity in (select objectid from usd_activity with (nolock) where state in (2,3)))"))
                    Logger.WriteDebug(Callstack, "Software Jobs [In-Progress]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_applic with (nolock) where status in (7,8,27)"))
                    Logger.WriteDebug(Callstack, "Software Jobs [Completed]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_applic with (nolock) where status=9"))
                    Logger.WriteDebug(Callstack, "Staging Jobs [In-Progress]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_applic with (nolock) where status in (2,3,17,18,19,20,21)"))
                    Logger.WriteDebug(Callstack, "Staging Jobs [Completed]: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_applic with (nolock) where status=4"))
                    Logger.WriteDebug(Callstack, "Software Based Policies: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from usd_job_cont with (nolock) where type=4 and comment<>'CA-Software Delivery Reserved Group'"))
                    Logger.WriteDebug(Callstack, "Query Based Policies: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from polidef with (nolock)"))
                    Logger.WriteDebug(Callstack, "Engine Instances: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ca_engine with (nolock)"))
                    Logger.WriteDebug(Callstack, "Engine Tasks: " + SqlSelectScalar(Callstack, DbConnection, "select count(*) from ncjobcfg with (nolock) where jotype<>2 and domainid=" + DomainId))

                ElseIf Globals.MdbCleanAppsSwitch Then

                    ' Local variables
                    Dim SqlScript As String = My.Resources.SQL_CleanApps_Unified
                    Dim SqlBatches As String()
                    Dim SqlCmd As SqlCommand
                    Dim SqlData As SqlDataReader

                    ' Attach informational message handler to connection
                    AddHandler DbConnection.InfoMessage, New SqlInfoMessageEventHandler(AddressOf OnMdbCleanAppsInfoMessage)

                    ' Enumerate sql batches based on "GO"
                    SqlBatches = Regex.Split(SqlScript, "^\s*GO\s*\d*\s*($|\-\-.*$)", RegexOptions.Multiline Or RegexOptions.IgnorePatternWhitespace Or RegexOptions.IgnoreCase)

                    ' Iterate each batch of queries
                    For Each QueryLine As String In SqlBatches

                        ' Check for empty line
                        If QueryLine.Equals("") Then Continue For

                        ' Build sql command from user query
                        SqlCmd = New SqlCommand(QueryLine, DbConnection)

                        ' Execute query -- ExecuteReader method
                        SqlData = SqlCmd.ExecuteReader()

                        ' Close sql data reader
                        SqlData.Close()

                    Next

                End If

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(Callstack, "Exception: " + ex.Message)

            Finally

                ' Check if database connection is open
                If Not DbConnection.State = ConnectionState.Closed Then

                    ' Close the database connection
                    DbConnection.Close()

                    ' Write debug
                    Logger.WriteDebug(Callstack, "Database connection closed.")

                End If

            End Try

            ' Return
            Return 0

        End Function

        Private Shared Sub OnMdbCleanAppsInfoMessage(ByVal sender As Object, ByVal e As SqlInfoMessageEventArgs)
            Logger.WriteDebug(Logger.LastCallStack, e.Message)
        End Sub

        Public Shared Function SqlSelectScalar(ByVal CallStack As String,
                                               ByVal DatabaseConnection As SqlConnection,
                                               ByVal QueryText As String) As String

            ' Local variables
            Dim SqlCmd As SqlCommand
            Dim SqlData As SqlDataReader
            Dim ScalarValue As String = ""

            ' Encapsulate sql operation
            Try

                ' Check connection state
                If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then

                    ' Return
                    Return Nothing

                End If

                ' Build sql command from user query
                SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

                ' Write debug (suppress for now)
                'Logger.WriteDebug(CallStack, QueryText)

                ' Execute query -- ExecuteReader method
                SqlData = SqlCmd.ExecuteReader()

                ' Check if data returns any rows
                If SqlData.Read() AndAlso SqlData.HasRows Then

                    ' Check column position for null
                    If SqlData.IsDBNull(0) Then

                        ' Add string null instead of actual null
                        ScalarValue = "Null"

                    ElseIf SqlData.GetFieldType(0) Is GetType(Byte()) Then

                        ' Add string representation of binary data
                        ScalarValue = "0x" + BitConverter.ToString(SqlData.GetSqlBinary(0).Value).Replace("-", "")

                    Else

                        ' Add raw data
                        ScalarValue = SqlData(0).ToString

                    End If

                End If

                ' Close sql data reader
                SqlData.Close()

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Database Exception: " + ex.Message)

                ' Return
                Return ex.Message

            End Try

            ' Return
            Return ScalarValue

        End Function

        Public Shared Function SqlSelectScalarList(ByVal CallStack As String,
                                                   ByVal DatabaseConnection As SqlConnection,
                                                   ByVal QueryText As String) As List(Of String)

            ' Local variables
            Dim SqlCmd As SqlCommand
            Dim SqlData As SqlDataReader
            Dim ResultArray As New List(Of String)

            ' Encapsulate sql operation
            Try

                ' Check connection state
                If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then

                    ' Return
                    Return Nothing

                End If

                ' Build sql command from user query
                SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

                ' Write debug (suppress for now)
                'Logger.WriteDebug(CallStack, QueryText)

                ' Execute query -- ExecuteReader method
                SqlData = SqlCmd.ExecuteReader()

                ' Loop through sql data results
                Do

                    ' Check if data returns any rows
                    If SqlData.HasRows Then

                        ' Iterate each record
                        While SqlData.Read

                            ' Check column position for null
                            If SqlData.IsDBNull(0) Then

                                ' Add string null instead of actual null
                                ResultArray.Add("Null")

                            ElseIf SqlData.GetFieldType(0) Is GetType(Byte()) Then

                                ' Add string representation of binary data
                                ResultArray.Add("0x" + BitConverter.ToString(SqlData.GetSqlBinary(0).Value).Replace("-", ""))

                            Else

                                ' Add raw data
                                ResultArray.Add(SqlData(0).ToString)

                            End If

                        End While

                    End If

                Loop While SqlData.NextResult

                ' Close sql data reader
                SqlData.Close()

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Database Exception: " + ex.Message)

            End Try

            ' Return
            Return ResultArray

        End Function

        Public Shared Function SqlSelectGrid(ByVal CallStack As String,
                                             ByVal DatabaseConnection As SqlConnection,
                                             ByVal QueryText As String,
                                             Optional ByVal IncludeColumnHeaders As Boolean = False) As List(Of List(Of String))

            ' Local variables
            Dim SqlCmd As SqlCommand
            Dim SqlData As SqlDataReader
            Dim RowArray As New List(Of String)
            Dim ResultArray As New List(Of List(Of String))

            ' Encapsulate sql operation
            Try

                ' Check connection state
                If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then

                    ' Return
                    Return Nothing

                End If

                ' Build sql command from user query
                SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

                ' Write debug (suppress for now)
                'Logger.WriteDebug(CallStack, QueryText)

                ' Execute query -- ExecuteReader method
                SqlData = SqlCmd.ExecuteReader()

                ' Loop through sql data results
                Do

                    ' Check if data returns any rows
                    If SqlData.HasRows Then

                        ' Check if column headers should be included in resultset
                        If IncludeColumnHeaders Then

                            ' Iterate through table columns names
                            For i As Integer = 0 To SqlData.FieldCount - 1

                                ' Check for a column name
                                If SqlData.GetName(i).Equals("") Then

                                    ' Stub an empty column name
                                    RowArray.Add("(No column name)")

                                Else

                                    ' Set column name
                                    RowArray.Add(SqlData.GetName(i))

                                End If

                            Next

                            ' Add columns names as first row of the result
                            ResultArray.Add(RowArray)

                        End If

                        ' Iterate each record
                        While SqlData.Read

                            ' Initialize array for row data
                            RowArray = New List(Of String)

                            ' Iterate each column of the record
                            For i As Integer = 0 To SqlData.FieldCount - 1

                                ' Check column position for null
                                If SqlData.IsDBNull(i) Then

                                    ' Add string null instead of actual null
                                    RowArray.Add("Null")

                                ElseIf SqlData.GetFieldType(i) Is GetType(Byte()) Then

                                    ' Add string representation of binary data
                                    RowArray.Add("0x" + BitConverter.ToString(SqlData.GetSqlBinary(i).Value).Replace("-", ""))

                                Else

                                    ' Add raw data
                                    RowArray.Add(SqlData(i))

                                End If

                            Next

                            ' Add row (arraylist) to result (arraylist)
                            ResultArray.Add(RowArray)

                        End While

                    End If

                Loop While SqlData.NextResult

                ' Close sql data reader
                SqlData.Close()

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Database Exception: " + ex.Message)

            End Try

            ' Return
            Return ResultArray

        End Function

    End Class

End Class