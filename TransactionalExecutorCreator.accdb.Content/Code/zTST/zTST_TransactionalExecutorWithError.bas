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
Private m_FiredCommittedEvent As Boolean
Private m_FiredRolledBackEvent As Boolean
Private m_ErrorStateFromRolledBackEvent As AT_ErrorState
Private m_TextReadViaDefaultWorkspaceInEventRolledBack As String

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager

Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set Executor = New AT_TransactionalExecutor
   m_FiredCommittedEvent = False
   m_FiredRolledBackEvent = False
   Set m_ErrorStateFromRolledBackEvent = Nothing
   m_TextReadViaDefaultWorkspaceInEventRolledBack = vbNullString
End Sub
Public Sub TearDown()
   Set Executor = Nothing
End Sub

Public Sub DoesNotFireEventCommitted()
   Executor.Execute
   Assert.IsFalse m_FiredCommittedEvent, "The event 'Committed' should not be fired."
End Sub

Public Sub FiresEventRolledBack()
   Executor.Execute
   Assert.IsTrue m_FiredRolledBackEvent, "The event 'RolledBack' should be fired."
End Sub

Public Sub RollsBackOperationsDoneInExecute()
   Dim OriginalTextInTable As String
   
   OriginalTextInTable = GetActualTextInTable()
   
   ' Act
   Executor.Execute
   
   Assert.AreEqual OriginalTextInTable, GetActualTextInTable(), "The database should not be updated."
End Sub

Public Sub HasAlreadyRolledBackWhenEventRolledBackIsFired()
   Dim OriginalTextInTable As String
   
   OriginalTextInTable = GetActualTextInTable()
   
   ' Act
   Executor.Execute
   
   Assert.AreEqual OriginalTextInTable, m_TextReadViaDefaultWorkspaceInEventRolledBack
End Sub

Public Sub LeavesNoOpenTransaction()
   Executor.Execute
   
   Assert.IsFalse IsInTransaction()
End Sub

Public Sub ProvidesTheErrorInTheRolledBackEvent()
   Executor.Execute
   Assert.AreEqual ThrownErrorNumber, m_ErrorStateFromRolledBackEvent.Number, "The number of the thrown error should be provided."
   Assert.AreEqual ThrownErrorDescription, m_ErrorStateFromRolledBackEvent.Description, "The description of the thrown error should be provided."
   Assert.AreEqual ThrownErrorSource, m_ErrorStateFromRolledBackEvent.Source, "The source of the thrown error should be provided."
End Sub



' ___ Executor Event Handlers ___



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

Private Sub Executor_Committed()
   m_FiredCommittedEvent = True
End Sub

Private Sub Executor_RolledBack(ByVal ErrorState As AT_ErrorState)
   m_FiredRolledBackEvent = True
   Set m_ErrorStateFromRolledBackEvent = ErrorState
   m_TextReadViaDefaultWorkspaceInEventRolledBack = GetActualTextInTableViaDefaultWorkspace()
End Sub