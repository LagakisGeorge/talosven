VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PELATES 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   1335
   ClientWidth     =   9645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9645
   WindowState     =   2  'Maximized
   Begin VB.TextBox MaskEdBox1 
      Height          =   435
      Left            =   5160
      TabIndex        =   9
      Top             =   1845
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   7230
      TabIndex        =   5
      Top             =   1845
      Width           =   1830
   End
   Begin VB.CommandButton DELETE 
      Caption         =   "Διαγραφή"
      Height          =   495
      Left            =   7860
      TabIndex        =   4
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton EDIT 
      Caption         =   "Διόρθωση"
      Height          =   495
      Left            =   6525
      TabIndex        =   3
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton ADD 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Νέα Εγγραφή"
      Height          =   495
      Left            =   5220
      TabIndex        =   2
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton EXODOS 
      Caption         =   "Εξοδος"
      Height          =   495
      Left            =   7860
      TabIndex        =   0
      Top             =   4140
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "PELATES"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      Top             =   4950
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3570
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_PELATES.frx":0000
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "_PELATES.frx":0014
      TabIndex        =   6
      Top             =   1290
      Width           =   4590
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   7215
      TabIndex        =   8
      Top             =   1455
      Width           =   1860
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   5175
      TabIndex        =   7
      Top             =   1455
      Width           =   1920
   End
   Begin VB.Shape Shape1 
      Height          =   2145
      Left            =   4815
      Top             =   1305
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Πελάτες"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3990
      TabIndex        =   1
      Top             =   6270
      Visible         =   0   'False
      Width           =   2385
   End
End
Attribute VB_Name = "PELATES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const BLE = &HFF0000
Const ASPRO = &HFFFFFF
Dim katax
Dim Addcaption
Dim Editcaption
Dim Deletecaption
Dim field, field2
Dim validate
Private Sub Command1_Click()
'  Data1.Recordset.Close
  Unload Me
End Sub

Private Sub Add_Click()
' καταχώρησ

If ADD.Caption = katax Then
   maskedbox1.Enabled = False
   Data1.Recordset.AddNew
   Data1.Recordset.update
   Data1.Recordset.Move 0, Data1.Recordset.LastModified
   Data1.Recordset.EDIT
   Data1.Recordset(field) = Trim(maskedbox1)
   If field2 <> " " Then
      Data1.Recordset(field2) = Val(Text1.text)
   End If
   Data1.Recordset.update
   
'   Data1.Recordset.Edit
 '  Data1.Recordset.update
   
   DBGrid1.Refresh
   
   EDIT.Enabled = True
   DELETE.Enabled = True
   ADD.Caption = Addcaption
   maskedbox1.BackColor = ASPRO
   maskedbox1.ForeColor = BLE
   maskedbox1.text = String(Data1.Recordset(field).Size, " ")
   validate = False
   Text1.text = Space(Data1.Recordset(field).Size)
   validate = True


Else 'νεα εγγραφή
   EDIT.Enabled = False
   DELETE.Enabled = False
   ADD.Caption = katax '"Καταχώρηση"
   maskedbox1.BackColor = BLE
   maskedbox1.ForeColor = ASPRO
    maskedbox1.Enabled = True
    maskedbox1.SetFocus
End If

End Sub
Private Sub Delete_Click()

 '
   
   If Data1.Recordset.RecordCount = 0 Then
      MsgBox mL_Res(201) ' 201 "Δεν υπάρχουν εγγραφές στο αρχείο."
      Exit Sub
   End If
   If Data1.Recordset.eof Or Data1.Recordset.BOF Then
      MsgBox mL_Res(202) '202 "Δεν έχετε ορίσει ποιά εγγραφή θέλετε."
      Exit Sub
   End If

   If IsNull(Data1.Recordset(field)) Then
      maskedbox1 = String(Data1.Recordset(field).Size, " ")
   Else
      maskedbox1 = left(Data1.Recordset(field) + Space(50), Data1.Recordset(field).Size)
   End If
   response = MsgBox(maskedbox1 + Chr(13) + mL_Res(203), vbYesNo)
      '203   "Να Διαγραφή ; "
   If response = vbYes Then
      Data1.Recordset.DELETE
      Data1.Recordset.MoveNext
      If Data1.Recordset.eof Then
        If Data1.Recordset.RecordCount > 0 Then
           Data1.Recordset.MoveFirst
        End If
      End If
   End If
   maskedbox1.text = String(Data1.Recordset(field).Size, " ")


End Sub

Private Sub Edit_Click()
 
 On Error Resume Next
 
If IsNull(Data1.Recordset(field)) Then
      maskedbox1.Enabled = True
     Data1.Recordset.EDIT
     Data1.Recordset(field) = "."
     Data1.Recordset.update
     'MsgBox "Η εγγραφή δεν διορθώνεται γιατί είναι μηδενική.Δώσε νέα εγγραφή."
    'Exit Sub
End If

 
 ' καταχώρηση
If EDIT.Caption = katax Then
    maskedbox1.Enabled = False
   DBGrid1.Enabled = True
   Data1.Recordset.EDIT
  If IsNumeric(maskedbox1) Then maskedbox1 = str(maskedbox1)
     Data1.Recordset(field) = Trim(maskedbox1)
    If field2 <> " " Then
      Data1.Recordset("apor") = Val(Text1.text)
   End If
     
   Data1.Recordset.update
     
   
   
   DBGrid1.Refresh
   
   ADD.Enabled = True
   DELETE.Enabled = True
   EDIT.Caption = Editcaption
   maskedbox1.BackColor = ASPRO
   maskedbox1.ForeColor = BLE
   
   maskedbox1.text = String(Data1.Recordset(field).Size, " ")
 If field2 <> " " Then
   validate = False
   Text1.text = Space(Data1.Recordset(field2).Size)
   DBGrid1.Enabled = True
 End If

Else ' εγγραφή
    maskedbox1.Enabled = True
   If Data1.Recordset.RecordCount = 0 Then
      MsgBox mL_Res(204)
      ' 204"Δεν υπάρχουν εγγραφές στο αρχείο."
      Exit Sub
   End If
   If Data1.Recordset.eof Or Data1.Recordset.BOF Then
      MsgBox mL_Res(205)
      '205 "Δεν έχετε ορίσει ποιά εγγραφή θέλετε."
      Exit Sub
   End If
   
   DBGrid1.Enabled = False
   maskedbox1.text = left(Data1.Recordset(field) + Space(50), Data1.Recordset(field).Size)
   
If field2 <> " " Then
   validate = False
   If IsNull(Data1.Recordset(field2)) Then
      Text1.text = Space(Data1.Recordset(field).Size)
   Else
      Text1.text = Data1.Recordset(field2)
   End If
   validate = True

End If
   ADD.Enabled = False
   DELETE.Enabled = False
   EDIT.Caption = katax '"Καταχώρηση"
   maskedbox1.SetFocus
End If
 
 On Error GoTo 0
 
End Sub

Private Sub EXIT_Click()


End Sub


Private Sub exodos_Click()
If ADD.Caption = katax Or EDIT.Caption = katax Then
   response = MsgBox(mL_Res(206), vbYesNo)
   ' 206 "Να καταχωρηθούν οι αλλαγές; "
   If response = vbYes Then
      If ADD.Caption = katax Then
          Add_Click
      Else
          Edit_Click
      End If
   Else
   End If
End If
     ' MsgBox
   Unload Me
End Sub

Private Sub Form_DblClick()
 Dim a
      
      a = save_colors(Me, Me.BackColor)
End Sub

Private Sub Form_Load()


MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False


a = find_colors(Me)


katax = mL_CapRes(214) '214 "Καταχώρηση"
Addcaption = mL_CapRes(215) '215 "Νέα Εγγραφή"
Editcaption = mL_CapRes(216) '216 "Διόρθωση"
Deletecaption = mL_CapRes(213) '213 "Διαγραφή"

ADD.Caption = Addcaption
EDIT.Caption = Editcaption
DELETE.Caption = Deletecaption
exodos.Caption = mL_CapRes(212) '"Εξοδος"
validate = True
   
If Menu = 1 Then  'pelates
  Label1.Caption = mL_CapRes(217) '"Πελάτες"
  Data1.RecordSource = "select * from PELATES order by onoma;"
  Data1.Refresh
  
  field = "onoma"
  field2 = " "
 '
  DBGrid1.Columns(0).Caption = mL_CapRes(218)  '"ONOMA"
  
   DBGrid1.Columns(0).DataField = field
   DBGrid1.Refresh
   Label2.Caption = DBGrid1.Columns(0).Caption
  
  
  
  
  Text1.Visible = False
  pelates.Caption = mL_CapRes(217) ' "Πελάτες"
  
  
ElseIf Menu = 2 Then
   Label1.Caption = mL_CapRes(219) '"Διαδικασίες"
   Data1.RecordSource = "select *from eidh_baf;"
   Data1.Refresh
   field = "ono_baf"
   field2 = " "
   DBGrid1.Font.Bold = True
   DBGrid1.Columns(0).Caption = mL_CapRes(220) ' "Βαφή"
   DBGrid1.Columns(0).DataField = "Ono_baf"
   DBGrid1.Refresh
   Label2.Caption = DBGrid1.Columns(0).Caption
   Text1.Visible = False
 
  
ElseIf Menu = 3 Then  'Πανιά
   DBGrid1.AllowUpdate = True
   
   
   
   Label1.Caption = mL_CapRes(221) ' "Είδη Υφασμάτων"
   Data1.RecordSource = "select *from pania;"
   Data1.Refresh
   field = "pani"
   field2 = "apor"
   '   DBGrid1.Font.Bold = True
   DBGrid1.Columns(0).Caption = mL_CapRes(222) '222 "Υφασμα"
   DBGrid1.Columns(0).DataField = "pani"
   
   DBGrid1.Columns(1).Caption = mL_CapRes(223) '223 "Συντ.Απορφ"
   DBGrid1.Columns(1).DataField = "apor"
    DBGrid1.Columns(1).width = 1300
   
   DBGrid1.Columns(2).Caption = "Loss in Weight %" '223 "Συντ.Απορφ"
   DBGrid1.Columns(2).DataField = "FYRA"
   DBGrid1.Columns(2).width = 1300
   
   
   
   Label2.Caption = DBGrid1.Columns(0).Caption
   Label3.Caption = DBGrid1.Columns(1).Caption
  

End If
   DBGrid1.Refresh



End Sub

Private Sub Form_Paint()
     Screen.MousePointer = 1
     L = Data1.Recordset(field).Size
     maskedbox1.mask = String$(L, "C")

     
End Sub

Private Sub Form_Unload(Cancel As Integer)


MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True




End Sub

Private Sub MaskEdBox1_GotFocus()
   maskedbox1.BackColor = BLE
   maskedbox1.ForeColor = ASPRO

End Sub

Private Sub MaskEdBox1_LostFocus()
   maskedbox1.ForeColor = BLE
   maskedbox1.BackColor = ASPRO

End Sub


Private Sub MaskEdBox2_GotFocus()
   Maskedbox2.ForeColor = ASPRO
   Maskedbox2.BackColor = BLE

End Sub

Private Sub MaskEdBox2_LostFocus()
   Maskedbox2.ForeColor = BLE
   Maskedbox2.BackColor = ASPRO

End Sub

Private Sub Text1_Change()
  If Not IsNumeric(Text1) And validate = True Then
      Text1 = ""
      
      MsgBox mL_Res(207)
      ' 207 "Επιτρέπονται μόνο αριθμοί!"
  
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.BackColor = BLE
   Text1.ForeColor = ASPRO

End Sub


Private Sub Text1_LostFocus()
   Text1.BackColor = ASPRO
   Text1.ForeColor = BLE
   If Val(Text1.text) > 99 Then
      MsgBox mL_Res(208)
      '208 "μέγιστος επιτρεπόμενος αριθμός 10"
      Text1.text = "   "
    End If
   If Val(Text1.text) <= 0 Then
      MsgBox mL_Res(209)
      ' 209 "Δεν επιτρέπεται το 0"
      Text1.text = "   "
      Text1.SetFocus
    End If
    
    
    
    
End Sub


