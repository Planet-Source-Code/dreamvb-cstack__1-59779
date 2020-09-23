VERSION 5.00
Begin VB.Form frmStack 
   Caption         =   "Stack Demo"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfull 
      Caption         =   "IsFull"
      Height          =   435
      Left            =   6120
      TabIndex        =   7
      Top             =   3150
      Width           =   1395
   End
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "IsEmpty"
      Height          =   435
      Left            =   6120
      TabIndex        =   6
      Top             =   2610
      Width           =   1395
   End
   Begin VB.CommandButton cmdsize 
      Caption         =   "StackSize"
      Height          =   435
      Left            =   6120
      TabIndex        =   5
      Top             =   2070
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Bottom Item"
      Height          =   435
      Left            =   6120
      TabIndex        =   4
      Top             =   1560
      Width           =   1395
   End
   Begin VB.CommandButton cmdtop 
      Caption         =   "Get Top Item"
      Height          =   435
      Left            =   6120
      TabIndex        =   3
      Top             =   1065
      Width           =   1395
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   6120
      TabIndex        =   2
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton cmdpop 
      Caption         =   "Pop"
      Height          =   435
      Left            =   6120
      TabIndex        =   1
      Top             =   570
      Width           =   1395
   End
   Begin VB.CommandButton cmdPush 
      Caption         =   "Push"
      Height          =   435
      Left            =   6120
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "frmStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myStack As New CStack

Function CheckISEmpty() As Boolean
    CheckISEmpty = False
    If myStack.StackEmpty = True Then ' check if the stack is empty
        MsgBox "There is no data on the stack to pop off" ' now it was not
        CheckISEmpty = True
    End If
End Function

Private Sub cmdEmpty_Click()
    MsgBox myStack.StackEmpty ' is the stack empty
End Sub

Private Sub cmdexit_Click()
    myStack.Reset
    MsgBox "Stack class by Ben Jones.", vbInformation
    End
End Sub

Private Sub cmdfull_Click()
    MsgBox myStack.StackFull ' is the stack full
End Sub

Private Sub cmdpop_Click()

    If CheckISEmpty Then Exit Sub
    
    Me.Print vbCrLf + "Pop the data off the stack"
    
    Do While Not myStack.StackEmpty 'while stack is not empty
        Me.Print myStack.Pop ' pop the data from the stack to the screen
        DoEvents
    Loop
    
End Sub

Private Sub cmdPush_Click()
Dim StrNum As String, c As String * 1
Dim x As Integer

    StrNum = "0123456789" ' data to be placed on the stack
    ' Note the stack class uses a varient for all it's data of couse
    ' you can chnage it for your own datatypes.
    
    myStack.Reset ' reset the stack first
    myStack.Size = Len(StrNum) ' resize the stack
    
    Me.Cls
    Me.Print "Place some data on the stack"
    
    For x = 1 To Len(StrNum)
        If Not myStack.StackFull() Then ' check if the stack is full
            c = Mid(StrNum, x, 1) ' extract one char at a time
            myStack.Push c ' push the char onto the stack
            Me.Print c ' show output
        End If
    Next
    
End Sub

Private Sub cmdsize_Click()
    MsgBox myStack.Size 'Get stack size
End Sub

Private Sub cmdtop_Click()
    If CheckISEmpty Then Exit Sub 'check if stack is empty
    MsgBox myStack.Top 'get the top item
End Sub

Private Sub Command1_Click()
    If CheckISEmpty Then Exit Sub 'check if stack is empty
    MsgBox myStack.Bottom 'get the top item
End Sub

