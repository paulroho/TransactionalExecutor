Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event BeforeExecute(ByRef Cancel As Boolean)
Public Event Execute(ByVal ErrorState As AT_ErrorState)
Public Event BeforeCommit(ByRef Cancel As Boolean)
Public Event AfterCommit()
Public Event AfterRollback(ByVal ErrorState As AT_ErrorState)

Public Property Get Self() As AT_TransactionalExecutor
   Set Self = Me
End Property

Public Sub Execute()
   Dim Cancel As Boolean
   
   RaiseEvent BeforeExecute(Cancel)
   If Cancel Then Exit Sub
   
   With New AT_ErrorState
      
      DBEngine.BeginTrans
      RaiseEvent Execute(.Self)
      
      If .ErrorOccurred Then
         DBEngine.Rollback
         RaiseEvent AfterRollback(.Self)
      Else
         If CanCommit() Then
            DBEngine.CommitTrans
            RaiseEvent AfterCommit
         Else
            DBEngine.Rollback
         End If
      End If
   End With
   
End Sub

Private Function CanCommit() As Boolean
   Dim Cancel As Boolean
   
   Cancel = False
   RaiseEvent BeforeCommit(Cancel)
   
   CanCommit = Not Cancel
End Function