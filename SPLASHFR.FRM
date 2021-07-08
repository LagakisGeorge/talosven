VERSION 4.00
Begin VB.Form SplashFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1125
   ClientLeft      =   1860
   ClientTop       =   2670
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Height          =   1530
   Left            =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SPLASHFR.frx":0000
   ScaleHeight     =   1125
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Top             =   2325
   Width           =   6105
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Robolab 240 Dispenser System"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   2175
      TabIndex        =   1
      Top             =   600
      Width           =   3465
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   225
      Picture         =   "SPLASHFR.frx":757B2
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Talos : System Monitor - Joystick"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   13.5
         underline       =   -1  'True
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   900
      TabIndex        =   0
      Top             =   75
      Width           =   4740
   End
End
Attribute VB_Name = "SplashFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
