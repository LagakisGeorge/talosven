VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   Height          =   6345
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1170
   Width           =   6810
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   840
      Left            =   510
      TabIndex        =   0
      Top             =   1335
      Width           =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set DB = OpenDatabase("C:\TALOS\RECIPIES.MDB")

DB.Execute "ADD FILED"


End Sub


