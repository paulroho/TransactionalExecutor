Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event Execute(ByVal ErrorState As AT_ErrorState)
Public Event Committed()
Public Event RolledBack(ByVal ErrorState As AT_ErrorState)

Public Property Get Self() As AT_TransactionalExecutor
   Set Self = Me
End Property

Public Sub Execute()
   With New AT_ErrorState
      RaiseEvent Execute(.Self)
      
      If .ErrorOccurred Then
         RaiseEvent RolledBack(.Self)
      Else
         RaiseEvent Committed
      End If
   End With
End Sub