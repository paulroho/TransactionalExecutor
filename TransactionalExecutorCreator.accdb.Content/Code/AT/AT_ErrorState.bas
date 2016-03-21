Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_ErrNumber As Long
Private m_ErrDescription As String
Private m_ErrSource As String

Public Property Get Self() As AT_ErrorState
   Set Self = Me
End Property

Public Sub SetError(ByVal ErrorObject As ErrObject)
   If m_ErrNumber <> 0 Then Err.Raise 8, TypeName(Me) & ".SetError()", "The error cannot be set another time after it has been set once."
   
   With ErrorObject
      m_ErrNumber = .Number
      m_ErrDescription = .Description
      m_ErrSource = .Source
   End With
End Sub

Public Property Get ErrorOccurred() As Boolean
   ErrorOccurred = m_ErrNumber <> 0
End Property

Public Property Get Number() As Long
   Number = m_ErrNumber
End Property

Public Property Get Description() As String
   Description = m_ErrDescription
End Property

Public Property Get Source() As String
   Source = m_ErrSource
End Property