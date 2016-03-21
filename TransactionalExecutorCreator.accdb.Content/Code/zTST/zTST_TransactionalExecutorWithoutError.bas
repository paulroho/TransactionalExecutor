Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private WithEvents Executor As AT_TransactionalExecutor
Attribute Executor.VB_VarHelpID = -1
Private m_FiredAfterCommitEvent As Boolean
Private m_FiredAfterRollbackEvent As Boolean
Private m_TextWrittenToTable As String
Private m_TextReadInAfterCommitHandler As String
Private m_SetCancelInBeforeCommitEvent As Boolean
Private m_TextReadInBeforeCommitEvent As String
Private m_AfterExecuteWasFired As Boolean
Private m_AfterExecuteWasAlreadyFiredInBeforeCommitEvent As Boolean

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager

Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set Executor = New AT_TransactionalExecutor
   m_FiredAfterCommitEvent = False
   m_FiredAfterRollbackEvent = False
   m_TextWrittenToTable = vbNullString
   m_TextReadInAfterCommitHandler = vbNullString
   m_SetCancelInBeforeCommitEvent = False
   m_TextReadInBeforeCommitEvent = vbNullString
   m_AfterExecuteWasFired = False
   m_AfterExecuteWasAlreadyFiredInBeforeCommitEvent = False
End Sub
Public Sub TearDown()
   Set Executor = Nothing
End Sub

Public Sub FiresEventAfterCommit()
   Executor.Execute
   Assert.IsTrue m_FiredAfterCommitEvent, "The event 'AfterCommit' should be fired."
End Sub

Public Sub CommitsOperationsDoneInExecute()
   Executor.Execute
   Assert.AreEqual m_TextWrittenToTable, GetActualTextInTable(), "The database should be updated."
End Sub

Public Sub HasAlreadyCommittedWhenEventAfterCommitIsFired()
   Executor.Execute
   Assert.AreEqual m_TextWrittenToTable, m_TextReadInAfterCommitHandler
End Sub

Public Sub LeavesNoOpenTransaction()
   Executor.Execute
   Assert.IsFalse IsInTransaction()
End Sub

Public Sub DoesNotFireEventAfterRollback()
   Executor.Execute
   Assert.IsFalse m_FiredAfterRollbackEvent, "The event 'AfterRollback' should not be fired."
End Sub

Public Sub DoesNotFireEventAfterCommitIfCancelIsSetInBeforeCommitEvent()
   m_SetCancelInBeforeCommitEvent = True
   
   Executor.Execute
   
   Assert.IsFalse m_FiredAfterCommitEvent
End Sub

Public Sub HasNotCommittedYetInBeforeCommitEvent()
   Dim OriginalTextInTable As String
   
   OriginalTextInTable = GetActualTextInTable()
   
   ' Act
   Executor.Execute
   
   Assert.AreEqual OriginalTextInTable, m_TextReadInBeforeCommitEvent
End Sub

Public Sub RollsBackIfCancelIsSetInBeforeCommitEvent()
   Dim OriginalTextInTable As String
   
   OriginalTextInTable = GetActualTextInTable()
   m_SetCancelInBeforeCommitEvent = True
   
   Executor.Execute
   
   Assert.AreEqual OriginalTextInTable, GetActualTextInTableViaDefaultWorkspace()
End Sub

Public Sub FiresEventAfterExecute()
   Executor.Execute
   Assert.IsTrue m_AfterExecuteWasFired
End Sub

Public Sub AfterExecuteWasAlreadyFiredWhenBeforeCommitIsFired()
   Executor.Execute
   Assert.IsTrue m_AfterExecuteWasAlreadyFiredInBeforeCommitEvent
End Sub



' ___ Executor Event Handlers ___



Private Sub Executor_Execute(ByVal ErrorState As AT_ErrorState)
   m_TextWrittenToTable = UpdateTextInTable()
   ' no error is raised
End Sub

Private Sub Executor_AfterExecute()
   m_AfterExecuteWasFired = True
End Sub

Private Sub Executor_BeforeCommit(ByRef Cancel As Boolean)
   Cancel = m_SetCancelInBeforeCommitEvent
   m_TextReadInBeforeCommitEvent = GetActualTextInTable()
   m_AfterExecuteWasAlreadyFiredInBeforeCommitEvent = m_AfterExecuteWasFired
End Sub

Private Sub Executor_AfterCommit()
   m_FiredAfterCommitEvent = True
   m_TextReadInAfterCommitHandler = GetActualTextInTable()
End Sub

Private Sub Executor_AfterRollback(ByVal ErrorState As AT_ErrorState)
   m_FiredAfterRollbackEvent = True
End Sub