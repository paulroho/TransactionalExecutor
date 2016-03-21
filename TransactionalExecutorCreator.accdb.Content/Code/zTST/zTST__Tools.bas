Option Compare Database
Option Explicit

Private m_UpdateCnt As Integer

Public Function UpdateTextInTable() As String
   Dim Rst As DAO.Recordset
   Dim TextUpdatedTo As String
   
   Set Rst = CurrentDb().OpenRecordset("ATable")
   
   m_UpdateCnt = m_UpdateCnt + 1
   TextUpdatedTo = "Update #" & m_UpdateCnt & " (" & DateTime.Now & ")" ' Ensure that the texts differ from one call to the next
   
   With Rst
      .Edit
      .Fields("TextField").Value = TextUpdatedTo
      .Update
   End With
   
   UpdateTextInTable = TextUpdatedTo
   
   If Not (Rst Is Nothing) Then Rst.Close
   Set Rst = Nothing
End Function

Public Function GetActualTextInTableViaDefaultWorkspace() As String
   Dim Wsp As DAO.Workspace
   
   ' Use the default workspace to read from inside the current transaction (if any)
   Set Wsp = DBEngine.Workspaces(0)
   
   GetActualTextInTableViaDefaultWorkspace = GetTextInTable(Wsp)
   
   Set Wsp = Nothing
End Function

Public Function GetActualTextInTable() As String
   Dim Wsp As DAO.Workspace
   
   ' Do not use the default workspace to read outside the current transaction
   Set Wsp = DBEngine.CreateWorkspace("ReadWorkspace", "Admin", "")
   
   GetActualTextInTable = GetTextInTable(Wsp)
   
   If Not (Wsp Is Nothing) Then Wsp.Close
   Set Wsp = Nothing
End Function

Private Function GetTextInTable(ByVal Workspace As DAO.Workspace) As String
   Dim Dbs As DAO.Database
   Dim Rst As DAO.Recordset
   
   Set Dbs = Workspace.OpenDatabase(CurrentDb().Name)
   Set Rst = Dbs.OpenRecordset("ATable")
   
   GetTextInTable = Rst!TextField
   
   If Not (Rst Is Nothing) Then Rst.Close
   Set Rst = Nothing
   If Not (Dbs Is Nothing) Then Dbs.Close
   Set Dbs = Nothing
End Function

Public Function IsInTransaction() As Boolean
   On Error Resume Next
   
   DBEngine.Rollback
   
   Select Case Err.Number
      Case 0
         IsInTransaction = True
      Case 3034   ' "You tried to commit or rollback a transaction without first beginning a transaction."
         IsInTransaction = False
      Case Else
         On Error GoTo 0
         Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   End Select
End Function