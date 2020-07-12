VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "COMMAND BUTTON"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim tmpValue&
  Dim Align&
  Dim ret&
  'Contoh berikut akan membuat tulisan di Command1
  'menjadi rapat ke atas ketika Anda mengklik tombol
  'tersebut.
  fAlignment& = A_TOP
  tmpValue& = GetWindowLong&(Command1.hwnd, _
              GWL_STYLE) And Not BS_RIGHT
  ret& = SetWindowLong&(Command1.hwnd, GWL_STYLE, _
         tmpValue& Or fAlignment&)
  Command1.Refresh
End Sub

Private Sub Command2_Click()
  Dim tmpValue&
  Dim Align&
  Dim ret&
  'Contoh berikut akan membuat tulisan di Command2
  'menjadi rapat ke bawah ketika Anda mengklik tombol
  'tersebut.
  fAlignment& = A_BOTTOM
  tmpValue& = GetWindowLong&(Command2.hwnd, _
              GWL_STYLE) And Not BS_RIGHT
  ret& = SetWindowLong&(Command2.hwnd, GWL_STYLE, _
         tmpValue& Or fAlignment&)
  Command2.Refresh
End Sub

Private Sub Command3_Click()
  Dim tmpValue&
  Dim Align&
  Dim ret&
  'Contoh berikut akan membuat tulisan di Command3
  'menjadi rapat ke kiri ketika Anda mengklik tombol
  'tersebut.
  fAlignment& = A_LEFT
  tmpValue& = GetWindowLong&(Command3.hwnd, _
              GWL_STYLE) And Not BS_RIGHT
  ret& = SetWindowLong&(Command3.hwnd, GWL_STYLE, _
         tmpValue& Or fAlignment&)
  Command3.Refresh
End Sub

Private Sub Command4_Click()
Dim tmpValue&
Dim Align&
Dim ret&
  'Contoh berikut membuat tulisan di Command4 menjadi
  'rapat ke kanan ketika Anda mengklik tombol tersebut.
  fAlignment& = A_RIGHT
  tmpValue& = GetWindowLong&(Command4.hwnd, _
              GWL_STYLE) And Not BS_RIGHT
  ret& = SetWindowLong&(Command4.hwnd, GWL_STYLE, _
         tmpValue& Or fAlignment&)
  Command4.Refresh
End Sub


