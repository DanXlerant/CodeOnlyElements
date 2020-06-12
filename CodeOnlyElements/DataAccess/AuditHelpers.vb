Imports XLerant.DataAccessLayer
Imports System.Data.SqlClient
Imports XLerant.CommonDataSets

Module AuditHelpers

    ''' <summary>
    ''' Enter an audit record. Convenience wrapper around sproc execution.
    ''' </summary>
    ''' <param name="UserID">The user ID of the responsible user.</param>
    ''' <param name="VersionID">The affected version</param>
    ''' <param name="SourceUnitID">The unit causing the change (if any)</param>
    ''' <param name="TargetUnitID">The unit affected by the change (if any)</param>
    ''' <param name="Category">The logical OR bitmask of the relevant categories</param>
    ''' <param name="Message">The human-readable text message associated with this action.</param>
    ''' <remarks></remarks>
    Public Sub SaveLogEntry(ByVal TenantID As Integer, ByVal UserID As Integer, ByVal VersionID As Integer, ByVal SourceUnitID As Integer, ByVal TargetUnitID As Integer, ByVal Category As Integer,
                            ByVal Message As String, Optional ByVal projectionID As Integer = -1, Optional ByVal ProjectID As Integer = -1)
        Dim da As New DataAccess(TenantID)
        Dim cmd As New SqlCommand("spAuditRecordCreate")
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("UserID", UserID)
        cmd.Parameters.AddWithValue("VersionID", VersionID)
        cmd.Parameters.AddWithValue("SourceUnitID", SourceUnitID)
        cmd.Parameters.AddWithValue("TargetUnitID", TargetUnitID)
        cmd.Parameters.AddWithValue("ProjectionID", projectionID)
        cmd.Parameters.AddWithValue("ProjectID", ProjectID)
        cmd.Parameters.AddWithValue("Category", Category)
        cmd.Parameters.AddWithValue("StatusMessage", Message)
        cmd.Parameters.AddWithValue("@TenantID", TenantID)
        cmd.Connection = New SqlConnection(TenantLookup.GetTenantConnectionString(TenantID))
        'da.ExecuteSProc(cmd)
        Dim dtEmails As New Emails.EmailsDataTable
        Dim sqlDA As New SqlDataAdapter(cmd)
        sqlDA.Fill(dtEmails)
        SendEmails(dtEmails)
    End Sub

    ''' <summary>
    ''' Show the audit trail for a given company, version, and unit
    ''' </summary>
    ''' <param name="VersionID"></param>
    ''' <param name="UnitID"></param>
    ''' <remarks></remarks>
    Public Sub ShowAuditTrail(ByVal ThisPage As Page, ByVal TenantID As Integer, ByVal VersionID As Integer, ByVal UnitID As Integer)
        DisplayBudgetHistoryDialog(ThisPage, TenantID, VersionID, UnitID)



        'Dim da As New DataAccess(TenantID)
        'Dim dsAudit As System.Data.DataSet
        'Dim cmd As New SqlCommand("spAuditRecordsGet")
        'cmd.Parameters.AddWithValue("VersionID", VersionID)
        'cmd.Parameters.AddWithValue("TargetUnitID", UnitID)
        'dsAudit = da.ExecReadSProc(cmd)
        'Dim msg As StringBuilder = New StringBuilder()
        'For Each row As DataRow In dsAudit.Tables(0).Rows
        '    msg.AppendFormat("{0}: Version '{1}', unit '{2}', user '{3} {4}': {5}", row("AsOf"), row("VersionDescription"), row("TargetUnitDescription"), row("UserFirstName"), row("UserLastName"), row("StatusMessage"))
        '    msg.AppendLine()
        'Next
        'DisplayLongMessage(ThisPage, msg.ToString, "History", 850, 435)
    End Sub

    Public Sub WriteUserMessageToAdministrator(ByVal TenantID As Integer, ByVal FromUserID As Integer, ByVal CompanyID As Integer, ByVal MessageTypeID As Integer, ByVal MessageText As String, _
      Optional ByVal VersionID As Integer = -1, Optional ByVal UnitID As Integer = -1)

        Dim da As New DataAccess(TenantID)

        ' Find the user(s) that have the SuperAdmin capability
        Dim SQL As New System.Text.StringBuilder
        SQL.Append("SELECT Users.UserID, RoleCapabilityMap.RoleID, RoleCapabilityMap.CapabilityID ")
        SQL.Append("FROM RoleCapabilityMap ")
        SQL.Append("INNER JOIN LookupCapabilities ON RoleCapabilityMap.CapabilityID=LookupCapabilities.CapabilityID ")
        SQL.Append("INNER JOIN Users ON Users.RoleID=RoleCapabilityMap.RoleID ")
        SQL.Append("WHERE CapabilityCode='SUPERADMIN' ")
        SQL.Append("AND Users.CompanyID=" & CompanyID)
        ' Send the message to each user
        Dim ds As DataSet = Nothing
        Try
            ds = da.GetData(SQL.ToString)
        Catch ex As Exception
            Dim strWhatHappened As String = "An error occurred while getting super administrator information."
            Dim ueh As New ASPUnhandledException.Handler
            ueh.HandleDataException(SQL.ToString, ex, strWhatHappened)
        End Try
        For Each dr As DataRow In ds.Tables(0).Rows
            WriteUserMessage(TenantID, FromUserID, dr("UserID"), MessageTypeID, MessageText, VersionID, UnitID)
        Next
    End Sub
    Public Sub WriteUserMessage(ByVal TenantID As Integer, ByVal FromUserID As Integer, ByVal ToUserID As Integer, ByVal MessageTypeID As Integer, ByVal MessageText As String, _
Optional ByVal VersionID As Integer = -1, Optional ByVal UnitID As Integer = -1)

        ' Since we can't store CR/LF or tab characters in our varchar field, let's display the message
        ' to the user using RTF.  (The FarPoint grid support RTF cell types but not, sadly, HTML cell types.)
        ' The only thing we convert to RTF for now is CR/LF and Tab.
        'Dim rtfText As String = rtfPrefix() & MessageText.Replace(ControlChars.CrLf, "\line ").Replace(ControlChars.Tab, "\tab ")
        Dim AdjustedText As String = MessageText.Replace(ControlChars.CrLf, Chr(13) & Chr(10)).Replace(ControlChars.Tab, Chr(9))


        Dim da As New DataAccess(TenantID)
        Try
            Dim cmd As New SqlCommand("spUserMessagesInsert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@UserID", ToUserID)
            cmd.Parameters.AddWithValue("@FromUserID", FromUserID)
            cmd.Parameters.AddWithValue("@MessageTypeID", MessageTypeID)
            cmd.Parameters.AddWithValue("@VersionID", VersionID)
            cmd.Parameters.AddWithValue("@UnitID", UnitID)
            cmd.Parameters.AddWithValue("@MessageText", MessageText)
            da.ExecuteSProc(cmd)

        Catch ex As Exception
            Dim strWhatHappened As String = "An error occurred trying to write user message: " & MessageText
            Throw New XLerant.BudgetPakException(strWhatHappened, ex)
        End Try
    End Sub

End Module
