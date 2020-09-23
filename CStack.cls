VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Stack() As Variant ' used to hold the stack data
Private m_StackSize As Long ' size of the stack
Private First As Long ' first item in the stack

Public Property Get Size() As Long
    Size = m_StackSize ' get stack size
End Property

Public Property Let Size(ByVal vNewValue As Long)
    m_StackSize = vNewValue + 1 ' add one to the stack size
    ReDim Preserve Stack(0 To vNewValue) ' resize the stack
End Property

Public Sub Reset()
    ' reset stack data
    First = 0
    m_StackSize = 0
    Erase Stack
End Sub

Public Function Top() As Variant
    Top = Stack(First - 1) ' return the top item
End Function

Function Bottom() As Variant
    Bottom = Stack(0) ' Return the last item
End Function
Public Function StackEmpty() As Boolean
    StackEmpty = (First = 0) 'check if the stack is empty
End Function

Public Function StackFull() As Boolean
    StackFull = (First = m_StackSize - 1) 'check if the stack is full
End Function

Public Sub Push(vData As Variant)
    ' push some data onto the stack
    Stack(First) = vData
    First = First + 1
End Sub

Public Function Pop() As Variant
    'pop the data back of the stack
    First = First - 1
    Pop = Stack(First)
End Function
