Attribute VB_Name = "mod_ChangeLog"
Option Compare Database
Option Explicit

' =============================================================================
' mod_ChangeLog
' Standard module — logs data changes to tblChangeLog for cross-session
' pub/sub synchronization via frmObserver + clsPubSubBroker.
'
' Mirrors the Kanbantt pattern: every repo CRUD calls LogChange() so the
' hidden observer form can detect and broadcast changes to other sessions.
' =============================================================================

' Logs a change to tblChangeLog
Public Sub LogChange(ByVal strChangeType As String, ByVal lngRecordId As Long, ByVal strAction As String)
    On Error Resume Next
    Dim strSQL As String
    strSQL = "INSERT INTO tblChangeLog ([ChangeType], [RecordID], [Action], [ChangedBy], [ChangedOn]) " & _
             "VALUES ('" & Replace(strChangeType, "'", "''") & "', " & lngRecordId & ", '" & Replace(strAction, "'", "''") & "', '" & Replace(Environ$("USERNAME"), "'", "''") & "', '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "')"
    CurrentDb.Execute strSQL, dbFailOnError
End Sub

' Purges old change log entries
Public Sub PurgeChangeLog(Optional ByVal lngRetainDays As Long = 7)
    On Error Resume Next
    Dim strSQL As String
    strSQL = "DELETE FROM tblChangeLog WHERE [ChangedOn] < " & _
             "#" & Format$(DateAdd("d", -lngRetainDays, Now()), "yyyy-mm-dd hh:nn:ss") & "#"
    CurrentDb.Execute strSQL, dbFailOnError
End Sub
