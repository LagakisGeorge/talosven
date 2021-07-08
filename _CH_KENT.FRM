VERSION 4.00
Begin VB.Form ch_kentr 
   Caption         =   "Form2"
   ClientHeight    =   5940
   ClientLeft      =   1620
   ClientTop       =   2640
   ClientWidth     =   6690
   Height          =   6345
   Left            =   1560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   2295
   Width           =   6810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Ακύρωση"
      Height          =   885
      Left            =   4245
      TabIndex        =   7
      Top             =   2010
      Width           =   2115
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1200
      Top             =   4020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Εξοδος"
      Height          =   855
      Left            =   4290
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Data XIMITECH 
      Caption         =   "XIMITECH"
      Connect         =   "dBASE IV;"
      DatabaseName    =   "c:\talos"
      Exclusive       =   0   'False
      Height          =   405
      Left            =   3765
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM XIMITECH WHERE ADDR_PROT>0  ORDER BY ADDR_PROT;"
      Top             =   4755
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Data mpoyk 
      Caption         =   "Data1"
      Connect         =   "dBASE IV;"
      DatabaseName    =   "c:\talos"
      Exclusive       =   0   'False
      Height          =   405
      Left            =   1020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "MPOYK"
      Top             =   4755
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εκκίνηση"
      Height          =   885
      Left            =   750
      TabIndex        =   5
      Top             =   2025
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   975
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1350
      TabIndex        =   0
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Ελεγχος Κεντραρισμάτων"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1065
      TabIndex        =   4
      Top             =   255
      Width           =   4995
   End
   Begin VB.Label Label2 
      Caption         =   "Εως θέση "
      Height          =   285
      Left            =   3135
      TabIndex        =   3
      Top             =   1035
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Από θέση "
      Height          =   285
      Left            =   405
      TabIndex        =   1
      Top             =   1050
      Width           =   840
   End
End
Attribute VB_Name = "ch_kentr"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Function is_Enable(mpoykali)
 is_Enable = 0
 mpoyk.Recordset.MoveFirst
 Do While Not mpoyk.Recordset.EOF
    If Val(mpoyk.Recordset("address")) = mpoykali Then
       If Not IsNull(mpoyk.Recordset("Enabled")) Then
         is_Enable = mpoyk.Recordset("Enabled")
         Exit Do
       End If
    End If
    mpoyk.Recordset.MoveNext
 Loop
 a = 1
End Function


Function is_On_Table(mpoykali)
 is_On_Table = 0
 XIMITECH.Recordset.MoveFirst
 Do While Not XIMITECH.Recordset.EOF
    If XIMITECH.Recordset("addr_prot") = mpoykali Then
         is_On_Table = 1
         Exit Do
    End If
    XIMITECH.Recordset.MoveNext
 Loop
End Function



Private Sub Command1_Click()


MsgBox mL_Res(822) ' "Βάλτε ένα γιάλινο μπουκάλι στην θέση ποτηριών 1 και αφήστε την 2η θέση ποτηριών κενή"
Command1.Enabled = False
Command2.Enabled = False


'Timer1.Enabled = True
dum = begin_talos(0)





For k = Val(Text1.Text) To Val(Text2.Text)
   If Command3.Enabled = False Then Exit For
   If is_Enable(k) > 0 Then
        
       If is_On_Table(k) = 0 Then
          Command1.Caption = Str(k)
          ' pane
           cmd$ = "!var12=" + Str(12000 + k) + ":!var13=" + Str(1) + ":"
           dum = wait_to_understand(cmd$)
           dum = Wait_Robot(0)
       
           'fere
           cmd$ = "!var12=" + Str(6000 + k) + ":!var13=" + Str(1) + ":"
           dum = wait_to_understand(cmd$)
           dum = Wait_Robot(0)
        Else
           Command1.Caption = Str(k)
         'fere
           cmd$ = "!var12=" + Str(6000 + k) + ":!var13=" + Str(2) + ":"
           dum = wait_to_understand(cmd$)
           dum = Wait_Robot(0)
           
          ' pane
           cmd$ = "!var12=" + Str(12000 + k) + ":!var13=" + Str(2) + ":"
           dum = wait_to_understand(cmd$)
           dum = Wait_Robot(0)
           
        End If
   End If
 
Next


MsgBox mL_Res(823) ' "Τέλος διαδικασίας"
Command2.Enabled = True

End Sub


Private Sub Command2_Click()
MDIForm1.arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.Bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True

Unload Me
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
End Sub

Private Sub Form_Load()
MDIForm1.arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.Bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False


Label1.Caption = mL_Res(825) ' Από θέση
Label2.Caption = mL_Res(826) 'Εως θέση
Label3.Caption = mL_Res(827) ' Ελεγχος Κεντραρισμάτων
Command1.Caption = mL_Res(828) ' Εκκίνηση
Command3.Caption = mL_Res(829) 'Ακύρωση
Command2.Caption = mL_Res(830) 'Εξοδος

End Sub


Private Sub Timer1_Timer()
 
       If system_ready = 1 Then
            Elegxo_Katastash
        End If
 
End Sub
