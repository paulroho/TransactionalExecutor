Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private Const ApplicationDefinedOrObjectDefinedError As Long = 8

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub


Public Sub SetError_OnceSet_CannotBeSetAgain()
   Dim ErrorState As AT_ErrorState
   Dim TestErrObject As ErrObject
   
   Set ErrorState = New AT_ErrorState
   
   ' Set it once
   Set TestErrObject = LetAnErrorHappen()
   Debug.Assert TestErrObject.Number <> 0
   ErrorState.SetError TestErrObject
   
   ' Try to set it a second time
   Set TestErrObject = LetAnErrorHappen()
   Debug.Assert TestErrObject.Number <> 0
   
On Error GoTo Err_

   ' Act
   ErrorState.SetError TestErrObject
   
On Error GoTo 0
   Assert.Fail "An error should have been raised."

Err_:
   Assert.AreEqual ApplicationDefinedOrObjectDefinedError, Err.Number, "The error should be raised with the correct number."
   Assert.AreEqual "The error cannot be set another time after it has been set once.", Err.Description, "The error should be raised with the correct description."
End Sub



' ___ Private Members ___



Private Function LetAnErrorHappen() As ErrObject
   Dim Dummy As Long
   
   On Error Resume Next
   Dummy = 1 / 0
   
   Set LetAnErrorHappen = Err
End Function