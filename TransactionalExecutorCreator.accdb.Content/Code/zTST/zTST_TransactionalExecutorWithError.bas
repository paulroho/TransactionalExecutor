Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private Const ThrownErrorNumber As Long = vbObjectError + 1234
Private Const ThrownErrorDescription As String = "Bad things happened."
Private Const ThrownErrorSource As String = "The source of the evil."

Private WithEvents Executor As AT_TransactionalExecutor
Attribute Executor.VB_VarHelpID = -1
Private m_FiredBeforeCommit As Boolean
Private m_FiredAfterCommitEvent As Boolean
Private m_FiredAfterRollbackEvent As Boolean
Private m_ErrorStateFromAfterRollbackEvent As AT_ErrorState
Private m_TextReadViaDefaultWorkspaceInEventAfterRollback As String
Private m_AfterExecuteWasFired As Boolean

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager

Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set Executor = New AT_TransactionalExecutor
   m_FiredBeforeCommit = False
   m_FiredAfterCommitEvent = False
   m_FiredAfterRollbackEvent = False
   Set m_ErrorStateFromAfterRollbackEvent = Nothing
   m_TextReadViaDefaultWorkspaceInEventAfterRollback = vbNullString
   m_AfterExecuteWasFired = False
End Sub
Public Sub TearDown()
   Set Executor = Nothing
End Sub

Public Sub DoesNotFireEventAfterCommit()
   Executor.Execute
   Assert.IsFalse m_FiredAfterCommitEvent, "The event 'AfterCommit' should not be fired."
End Sub

Public Sub FiresEventAfterRollback()
   Executor.Execute
   Assert.IsTrue m_FiredAfterRollbackEvent, "The event 'AfterRollback' should be fired."
End Sub

Public Sub RollsBackOperationsDoneInExecute()
   Dim OriginalTextInTable As String
   
   OriginalTextInTable = GetActualTextInTable()
   
   ' Act
   Executor.Execute
   
   Assert.AreEqual OriginalTextInTable, GetActualTextInTable(), "The database should not be updated."
End Sub

Public Sub HasAlreadyRolledBackWhenEventAfterRollbackIsFired()
   Dim OriginalTextInTable As String
   
   OriginalTextInTable = GetActualTextInTable()
   
   ' Act
   Executor.Execute
   
   Assert.AreEqual OriginalTextInTable, m_TextReadViaDefaultWorkspaceInEventAfterRollback
End Sub

Public Sub LeavesNoOpenTransaction()
   Executor.Execute
   
   Assert.IsFalse IsInTransaction()
End Sub

Public Sub ProvidesTheErrorInTheAfterRollbackEvent()
   Executor.Execute
   Assert.AreEqual ThrownErrorNumber, m_ErrorStateFromAfterRollbackEvent.Number, "The number of the thrown error should be provided."
   Assert.AreEqual ThrownErrorDescription, m_ErrorStateFromAfterRollbackEvent.Description, "The description of the thrown error should be provided."
   Assert.AreEqual ThrownErrorSource, m_ErrorStateFromAfterRollbackEvent.Source, "The source of the thrown error should be provided."
End Sub

Public Sub BeforeCommitIsNotFired()
   Executor.Execute
   Assert.IsFalse m_FiredBeforeCommit
End Sub

Public Sub FiresEventAfterExecute()
   Executor.Execute
   Assert.IsTrue m_AfterExecuteWasFired
End Sub



' ___ Executor Event Handlers ___



Private Sub Executor_BeforeCommit(Cancel As Boolean)
   m_FiredBeforeCommit = True
End Sub

Private Sub Executor_Execute(ByVal ErrorState As AT_ErrorState)
On Error GoTo Err_
   UpdateTextInTable
   
   Err.Raise ThrownErrorNumber, _
             ThrownErrorSource, _
             ThrownErrorDescription
   Exit Sub
Err_:
   ErrorState.SetError Err
End Sub

Private Sub Executor_AfterExecute()
   m_AfterExecuteWasFired = True
End Sub

Private Sub Executor_AfterCommit()
   m_FiredAfterCommitEvent = True
End Sub

Private Sub Executor_AfterRollback(ByVal ErrorState As AT_ErrorState)
   m_FiredAfterRollbackEvent = True
   Set m_ErrorStateFromAfterRollbackEvent = ErrorState
   m_TextReadViaDefaultWorkspaceInEventAfterRollback = GetActualTextInTableViaDefaultWorkspace()
End Sub