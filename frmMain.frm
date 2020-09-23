VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "modSnap Demo"
   ClientHeight    =   1575
   ClientLeft      =   6180
   ClientTop       =   7350
   ClientWidth     =   2850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   2850
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag and size me around!"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' NOTE:
'   I've used VB's built-in subclassing [GHEY] in this
'    demo to save space. Get MsgBlaster or some other
'    decent subclassing control/library.

Option Explicit

Dim cx As Long
Dim cy As Long

Private Sub Form_Load()
    gHW = Me.hWnd
    Label1.Move 0, (Me.ScaleHeight / 2) - (Label1.Height / 2), Me.ScaleWidth
    Hook
End Sub

Private Sub Form_Resize()
    Label1.Move 0, (Me.ScaleHeight / 2) - (Label1.Height / 2), Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook
End Sub
