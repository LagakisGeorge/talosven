VERSION 4.00
Begin VB.Form dokimes 
   Caption         =   "dokimes"
   ClientHeight    =   5940
   ClientLeft      =   1350
   ClientTop       =   5085
   ClientWidth     =   10830
   Height          =   6345
   Left            =   1290
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   10830
   Top             =   4740
   Width           =   10950
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Text            =   "586"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\TESTS.MDB"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "tdos_ygr"
      Top             =   5640
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Milisec"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Παρατηρησεις"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   960
      Top             =   3960
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
      cdtimeout       =   0
      commport        =   1
      ctstimeout      =   0
      dsrtimeout      =   0
      dtrenable       =   -1  'True
      handshaking     =   0
      inbuffersize    =   1024
      inputlen        =   0
      interval        =   1000
      nulldiscard     =   0   'False
      outbuffersize   =   512
      parityreplace   =   "?"
      rthreshold      =   0
      rtsenable       =   0   'False
      settings        =   "9600,n,8,1"
      sthreshold      =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Level"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Times"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DOKIMES.frx":0000
      Height          =   5775
      Left            =   6120
      OleObjectBlob   =   "DOKIMES.frx":000E
      TabIndex        =   3
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Μgrs"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "dokimes"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public Sub RobSend(cmd$)
   m_Cmd$ = cmd$
   TEMP% = SendATBlock(768, cmd$, 0)
End Sub

Private Sub Command1_Click()
Dim r, k, start, mdiaf
system_ready = 1
mdiaf = 0
Set r = Data1.Recordset

For k = 1 To Val(Text2.text)
' kod6400 = 100000 + a_Pane(Move)
 '             cmd$ = "!var12=" + Str(kod6400) + ":!var13=" + Str(a_Fere(Move)) + ":"
  '            DUM = wait_to_understand(cmd$)
Label7.Caption = k

Valve_on = valve_1_on
Valve_oFF = valve_1_off

tara = gzygisi4(0, mscomm1)

If Val(Text3.text) >= 600 Then
    mparox = 780
Else
    mparox = 780 - (600 - Val(Text3.text))
End If

DT = 1000 * (Val(Text1.text) - 100) / mparox





'=============
cmd$ = "!" + Valve_on + ":"
RobSend (cmd$)
'============
start = GetCurrentTime()
Do
Loop Until GetCurrentTime() - start > DT

'=============
cmd$ = "!" + Valve_oFF + ":"
RobSend (cmd$)
'=============




MilSec 2000
    arw = gzygisi4(0, mscomm1)
    asw = (Val(arw) - Val(tara)) * 1000 'καθαρό βάρος
    Label2.Caption = Int(asw)



parox = asw / DT
If parox < 0.5 Then parox = 0.5
If parox > 0.78 Then parox = 0.78



dt2 = (Val(Text1.text) - asw) / parox '* 1000
 
 
 
 
 start2 = GetCurrentTime()
'=============
 cmd$ = "!" + Valve_on + ":"
RobSend (cmd$)
'============
 
 Do While GetCurrentTime() - start2 < dt2
      
 Loop

'=============
 cmd$ = "!" + Valve_oFF + ":"
RobSend (cmd$)
'============

MilSec 2000
    arw = gzygisi4(0, mscomm1)
    asw = (Val(arw) - Val(tara)) * 1000 'καθαρό βάρος
    Label2.Caption = Int(asw)


r.AddNew

r("level") = Val(Text3.text)

r("apait") = Val(Text1.text)
r("pragm") = Val(asw)
r("parat") = Text4.text
r("time") = DT

mdiaf = Val(r("pragm")) - Val(r("apait"))
r("diafora") = mdiaf
r.update

Text3.text = Val(Text3.text) - asw / 1000
r.MoveLast
' 780 500gr   420 200gr   Text3.text = Val(Text3.text) + mdiaf
Next
End Sub

'
'
Private Sub oldCommand1_Click()
Dim r, k, start, mdiaf
system_ready = 1
mdiaf = 0
Set r = Data1.Recordset

For k = 1 To Val(Text2.text)
' kod6400 = 100000 + a_Pane(Move)
 '             cmd$ = "!var12=" + Str(kod6400) + ":!var13=" + Str(a_Fere(Move)) + ":"
  '            DUM = wait_to_understand(cmd$)
Label7.Caption = k

Valve_on = valve_1_on
Valve_oFF = valve_1_off

tara = gzygisi4(0, mscomm1)
start = GetCurrentTime()
'=============
 cmd$ = "!" + Valve_on + ":"
RobSend (cmd$)
'============

Do
    arw = gzygisi4(0, mscomm1)
    asw = (Val(arw) - Val(tara)) * 1000 'καθαρό βάρος
    Label2.Caption = Int(asw)
Loop Until Val(asw) + 1000 >= Val(Text1.text)
      
 
    '=============
    cmd$ = "!" + Valve_oFF + ":"
    RobSend (cmd$)
    '=============
    
    DT = (GetCurrentTime() - start) / 1000

MilSec 2000
    arw = gzygisi4(0, mscomm1)
    asw = (Val(arw) - Val(tara)) * 1000 'καθαρό βάρος
    Label2.Caption = Int(asw)



parox = asw / DT

dt2 = (Val(Text1.text) - asw) / parox * 1000
 
 
 
 
 start2 = GetCurrentTime()
'=============
 cmd$ = "!" + Valve_on + ":"
RobSend (cmd$)
'============
 
 Do While GetCurrentTime() - start2 < dt2
      
 Loop

'=============
 cmd$ = "!" + Valve_oFF + ":"
RobSend (cmd$)
'============

MilSec 2000
    arw = gzygisi4(0, mscomm1)
    asw = (Val(arw) - Val(tara)) * 1000 'καθαρό βάρος
    Label2.Caption = Int(asw)


r.AddNew

r("level") = Val(Text3.text)

r("apait") = Val(Text1.text)
r("pragm") = Val(asw)
r("parat") = Text4.text
r("time") = DT + dt2

mdiaf = Val(r("pragm")) - Val(r("apait"))
r("diafora") = mdiaf
r.update

Text3.text = Val(Text3.text) - asw / 1000
r.MoveLast
' 780 500gr   420 200gr   Text3.text = Val(Text3.text) + mdiaf
Next
End Sub

Private Sub Form_Load()
   FindValves
   
    mscomm1.CommPort = Balance_Port
    mscomm1.Settings = Balance_Settings
    If mscomm1.PortOpen = False Then mscomm1.PortOpen = True
    mscomm1.Settings = Balance_Settings
    mscomm1.InputLen = 0
End Sub


