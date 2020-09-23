VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boxes"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic3d 
      Height          =   7680
      Left            =   2520
      ScaleHeight     =   7620
      ScaleWidth      =   7620
      TabIndex        =   5
      Top             =   0
      Width           =   7680
   End
   Begin VB.TextBox txtPrimCount 
      Height          =   285
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "12"
      Top             =   1785
      Width           =   345
   End
   Begin VB.CheckBox chkLights 
      Caption         =   "Lights"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   1155
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.CheckBox chkIndex 
      Caption         =   "Indexed"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   960
   End
   Begin VB.CheckBox chkAlpha 
      Caption         =   "Transparent"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   525
      Width           =   1230
   End
   Begin VB.Timer tmrRender 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin MSComCtl2.UpDown udcPrimCount 
      Height          =   285
      Left            =   450
      TabIndex        =   4
      Top             =   1785
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   12
      BuddyControl    =   "txtPrimCount"
      BuddyDispid     =   196609
      OrigLeft        =   600
      OrigTop         =   1785
      OrigRight       =   855
      OrigBottom      =   2070
      Max             =   12
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call sub_CleanUp
    End If
End Sub

Private Sub Form_Load()
    Call sub_Setup
    Call sub_InitD3D(True, pic3d.hWnd) 'Setup the "Screen"
    Call sub_CreateObjects 'Create the objects
    tmrRender.Interval = 10
    tmrRender.Enabled = True
End Sub

Private Sub tmrRender_Timer()
    Call sub_Render
End Sub
