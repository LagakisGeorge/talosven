VERSION 4.00
Begin VB.Form Preheat 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Motor Preheating ..."
   ClientHeight    =   555
   ClientLeft      =   735
   ClientTop       =   1635
   ClientWidth     =   6945
   ControlBox      =   0   'False
   Height          =   960
   HelpContextID   =   1023
   Left            =   675
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Top             =   1290
   Width           =   7065
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   6630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   250
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   15
   End
End
Attribute VB_Name = "Preheat"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Function SendBlock(div_adr, cmd$, index_1) As Boolean
On Error Resume Next
tmp% = SendAT6400Block(div_adr, cmd$, 0)
SendBlock = True
End Function
Private Sub Form_GotFocus()
Screen.MousePointer = 11
asw = Me.Caption
FindValves
Me.Caption = asw + " ( " + Str(Motor_Preheating / 1000) + " Sec )"
cmd$ = "!" + Motor_Output_on + ":"
 temp% = SendBlock(768, cmd$, 0)
tot = Label2.Width
Num = 20
For I = 1 To Num
    MilSec Motor_Preheating / Num
    Label1.Width = tot * I / 20
    Label2.Caption = Format(I / 20, "##%")
Next I
Screen.MousePointer = 1
Unload Me
End Sub

