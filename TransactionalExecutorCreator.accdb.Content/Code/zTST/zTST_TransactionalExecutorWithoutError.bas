Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private WithEvents Executor As AT_TransactionalExecutor
Attribute Executor.VB_VarHelpID = -1
Private m_FiredCommittedEvent As Boolean
Private m_FiredRolledBackEvent As Boolean

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager

Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set Executor = New AT_TransactionalExecutor
End Sub
Public Sub TearDown()
   Set Executor = Nothing
End Sub

Public Sub FiresEventCommitted()
   Executor.Execute
   Assert.IsTrue m_FiredCommittedEvent, "The event 'Committed' should be fired."
End Sub

Public Sub DoesNotFireEventRolledBack()
   Executor.Execute
   Assert.IsFalse m_FiredRolledBackEvent, "The event 'RolledBack' should not be fired."
End Sub



' ___ Executor Event Handlers ___



Private Sub Executor_Execute(ByVal ErrorState As AT_ErrorState)
   ' no error is raised
End Sub

Private Sub Executor_Committed()
   m_FiredCommittedEvent = True
End Sub

Private Sub Executor_RolledBack(ByVal ErrorState As AT_ErrorState)
   m_FiredRolledBackEvent = True
End Sub