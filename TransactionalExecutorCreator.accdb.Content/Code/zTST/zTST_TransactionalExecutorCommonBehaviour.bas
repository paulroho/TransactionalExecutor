Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private WithEvents Executor As AT_TransactionalExecutor
Attribute Executor.VB_VarHelpID = -1
Private m_FiredExecuteEvent As Boolean
Private m_FiredBeforeExecuteEvent As Boolean
Private m_BeforeExecuteWasAlreadyFiredInExecuteEvent As Boolean
Private m_WasNotInTransactionOnBeforeExecuteEvent As Boolean
Private m_SetCancelInBeforeExecuteEvent As Boolean

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager

Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set Executor = New AT_TransactionalExecutor
   m_FiredExecuteEvent = False
   m_FiredBeforeExecuteEvent = False
   m_BeforeExecuteWasAlreadyFiredInExecuteEvent = False
   m_WasNotInTransactionOnBeforeExecuteEvent = False
   m_SetCancelInBeforeExecuteEvent = False
End Sub
Public Sub TearDown()
   Set Executor = Nothing
End Sub

Public Sub Execute_FiresBeforeExecuteEvent()
   Executor.Execute
   Assert.IsTrue m_FiredBeforeExecuteEvent
End Sub

Public Sub Execute_FiresExecuteEvent()
   Executor.Execute
   Assert.IsTrue m_FiredExecuteEvent
End Sub

Public Sub ExecuteEventIsFiredAfterTheBeforeExecuteEvent()
   Executor.Execute
   Assert.IsTrue m_BeforeExecuteWasAlreadyFiredInExecuteEvent
End Sub

Public Sub OnTheBeforeExecuteEventNoTransactionIsStarted()
   Executor.Execute
   Assert.IsTrue m_WasNotInTransactionOnBeforeExecuteEvent
End Sub

Public Sub ExecuteEventDoesNotFireWhenCancelIsSetInBeforeExecuteEvent()
   m_SetCancelInBeforeExecuteEvent = True
   
   Executor.Execute
   
   Assert.IsFalse m_FiredExecuteEvent
End Sub


' ___ Executor Event Handlers ___



Private Sub Executor_BeforeExecute(Cancel As Boolean)
   m_FiredBeforeExecuteEvent = True
   m_WasNotInTransactionOnBeforeExecuteEvent = Not IsInTransaction
   Cancel = m_SetCancelInBeforeExecuteEvent
End Sub

Private Sub Executor_Execute(ByVal ErrorState As AT_ErrorState)
   m_FiredExecuteEvent = True
   m_BeforeExecuteWasAlreadyFiredInExecuteEvent = m_FiredBeforeExecuteEvent
End Sub