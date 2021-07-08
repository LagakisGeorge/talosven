VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   2340
   ClientTop       =   1380
   ClientWidth     =   6690
   Height          =   6345
   Left            =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1035
   Width           =   6810
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\watch.mdb"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "moves"
      Top             =   4440
      Width           =   3495
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "WATCH_SH.frx":0000
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "WATCH_SH.frx":000E
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
