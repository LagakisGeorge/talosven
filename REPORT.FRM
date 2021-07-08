VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   2505
   ClientTop       =   1455
   ClientWidth     =   6690
   Height          =   4545
   Left            =   2445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6690
   Top             =   1110
   Width           =   6810
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1575
      TabIndex        =   1
      Top             =   615
      Width           =   2805
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      Exclusive       =   0   'False
      Height          =   360
      Left            =   1575
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "dokprosp"
      Top             =   3420
      Width           =   2085
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "REPORT.frx":0000
      Height          =   1425
      Left            =   240
      OleObjectBlob   =   "REPORT.frx":000E
      TabIndex        =   0
      Top             =   2070
      Width           =   5205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set R = Data1.Recordset
R.MoveFirst
Printer.Font.Name = "Arial Greek"
Do While Not R.EOF
   Printer.Print
   Printer.Print R("PERIGR"), R("REGR") / R("GR") * R("EK")
   M = R("EPIL")
   R.MoveNext
Loop
Printer.Print M
Printer.EndDoc


End Sub
