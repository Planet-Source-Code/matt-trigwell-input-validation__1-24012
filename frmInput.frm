VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input Validation Example"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmInput 
      Caption         =   "Input Validation Example"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton optInput 
         Caption         =   "Date  dd-mm-yy"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.OptionButton optInput 
         Caption         =   "Numeric Only"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optInput 
         Caption         =   "Text Only A-z"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin VB.OptionButton optInput 
         Caption         =   "Currency $0,000.00"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CheckBox chkUppercase 
         Caption         =   "Uppercase"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton optInput 
         Caption         =   "Date   dd/mm/yy"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblHowToUse 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmInput.frx":0000
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub optInput_Click(Index As Integer)
txtInput.Text = ""
txtInput.SetFocus
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
'Description: - The ValidateInput function checks to see if the KeyCode is allowed
'Parameters:  - KeyAscii - The Keycode of the key pressed by the user
'Returnvalue: -
'Programmer : - Matt Trigwell 11/06/01
'Comments:    - If the Keycode is not allowed it returns 0 which cancels the keypress



If (optInput(0).Value = True) Then

    'Date   dd/mm/yy
    KeyAscii = ValidateInput(KeyAscii, Date_Slash_Input)
    
ElseIf (optInput(1).Value = True) Then

    'Date  dd-mm-yy
    KeyAscii = ValidateInput(KeyAscii, Date_Dash_Input)
    
ElseIf (optInput(2).Value = True) Then

    'Numeric Only
    KeyAscii = ValidateInput(KeyAscii, Numeric_Input)
    
    
ElseIf (optInput(3).Value = True) Then

    'Text Only A-z
    KeyAscii = ValidateInput(KeyAscii, Text_Input, chkUppercase.Value)
    
ElseIf (optInput(4).Value = True) Then

    'Currency $0,000.00
    KeyAscii = ValidateInput(KeyAscii, Currency_Input)
    
End If


End Sub
