VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Farben 
   ClientHeight    =   6375
   ClientLeft      =   2040
   ClientTop       =   2205
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   5880
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      Height          =   540
      Left            =   6990
      TabIndex        =   1
      Top             =   5070
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   9
      Left            =   3225
      Top             =   5175
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   8
      Left            =   3225
      Top             =   4650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   7
      Left            =   3225
      Top             =   4125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   6
      Left            =   3225
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   5
      Left            =   3225
      Top             =   3075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   4
      Left            =   3225
      Top             =   2550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   3
      Left            =   3225
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   2
      Left            =   3225
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   1
      Left            =   3225
      Top             =   975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   0
      Left            =   3195
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Γραμματοσειρά"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   22
      Top             =   60
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Χρώμα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   435
      TabIndex        =   21
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα ανεύρεσης  χημικών"
      Height          =   315
      Index           =   0
      Left            =   4050
      TabIndex        =   20
      Top             =   465
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα διόρθωσης χημικών"
      Height          =   315
      Index           =   1
      Left            =   4080
      TabIndex        =   19
      Top             =   1005
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα εργασιών"
      Height          =   315
      Index           =   2
      Left            =   4080
      TabIndex        =   18
      Top             =   1530
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα Κατασκευής συνταγών"
      Height          =   315
      Index           =   3
      Left            =   4080
      TabIndex        =   17
      Top             =   2055
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Λίστα αναμονής"
      Height          =   315
      Index           =   4
      Left            =   4080
      TabIndex        =   16
      Top             =   2580
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα εκτέλεσης συνταγών"
      Height          =   315
      Index           =   5
      Left            =   4080
      TabIndex        =   15
      Top             =   3105
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα ανανέωσης διαλυμάτων"
      Height          =   315
      Index           =   6
      Left            =   4080
      TabIndex        =   14
      Top             =   3630
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα αυτόματης κατασκευής διαλυμάτων"
      Height          =   390
      Index           =   7
      Left            =   4080
      TabIndex        =   13
      Top             =   4080
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Τοποθέτηση μπουκαλιών σε τραπέζι"
      Height          =   390
      Index           =   8
      Left            =   4080
      TabIndex        =   12
      Top             =   4680
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Φόρμα πελατών"
      Height          =   390
      Index           =   9
      Left            =   4080
      TabIndex        =   11
      Top             =   5205
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα πελατών"
      Height          =   390
      Index           =   9
      Left            =   420
      TabIndex        =   10
      Top             =   5250
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Τοποθέτηση μπουκαλιών σε τραπέζι"
      Height          =   390
      Index           =   8
      Left            =   420
      TabIndex        =   9
      Top             =   4725
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα αυτόματης κατασκευής διαλυμάτων"
      Height          =   390
      Index           =   7
      Left            =   420
      TabIndex        =   8
      Top             =   4125
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα ανανέωσης διαλυμάτων"
      Height          =   315
      Index           =   6
      Left            =   420
      TabIndex        =   7
      Top             =   3675
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα εκτέλεσης συνταγών"
      Height          =   315
      Index           =   5
      Left            =   420
      TabIndex        =   6
      Top             =   3150
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Λίστα αναμονής"
      Height          =   315
      Index           =   4
      Left            =   420
      TabIndex        =   5
      Top             =   2625
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα Κατασκευής συνταγών"
      Height          =   315
      Index           =   3
      Left            =   420
      TabIndex        =   4
      Top             =   2100
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα εργασιών"
      Height          =   315
      Index           =   2
      Left            =   420
      TabIndex        =   3
      Top             =   1575
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα διόρθωσης χημικών"
      Height          =   315
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   1050
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Φόρμα ανεύρεσης  χημικών"
      Height          =   315
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   525
      Width           =   2505
   End
End
Attribute VB_Name = "Farben"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FI(0 To 10)



Private Sub Command1_Click()
     Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\syntages.mdb")
     Set r = dbsjet.OpenRecordset("form_colors")
     If r.RecordCount = 0 Then
        r.AddNew
        r.update
     End If
           r.MoveFirst
           r.EDIT
           For k = 0 To 9
              If Not IsNull(Label1(k).BackColor) Then
                       r(FI(k)) = str(Label1(k).BackColor)
              End If
         Next
         r.update
         
         r.MoveNext
         r.EDIT
          For k = 0 To 9
               s = ""
               s = s + IIf(Label2(k).FontBold, "1", "0")
               s = s + IIf(Label2(k).FontItalic, "1", "0")
               s = s + IIf(Label2(k).FontStrikethru, "1", "0")
               s = s + IIf(Label2(k).FontUnderline, "1", "0")
               s = s + Right$("0000000" + LTrim(str(Label2(k).FontSize)), 7)
               s = s + Label2(k).FontName
               r(FI(k)) = s
          Next
           r.update
         
         
         
         
      r.Close
      'find_colors
      Unload Me
End Sub


Private Sub Form_Load()
Dim Ki
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False

'Ki = Farben.Controls.Count
'For k = 1 To Ki
' Farben.Controls(6).FontSize = 15
'Next

   FI(0) = "f0"  '   : fi(1) = "f1": fi(2) = "f2": fi(3) = "f3": fi(4) = "f4": fi(5) = "f5":fi(6)="f6"
   For k = 1 To 9
        FI(k) = "f" & LTrim(str(k))
   Next
   
   
     Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\syntages.mdb")
     Set r = dbsjet.OpenRecordset("form_colors")
          
      On Error Resume Next
          r.MoveFirst
      
           For k = 0 To 9
                 If Not IsNull(r(FI(k))) Then
                       Label1(k).BackColor = Val(r(FI(k)))
                End If
          Next
          
          r.MoveNext
      
           For k = 0 To 9
                 If Not IsNull(r(FI(k))) Then
                       Label2(k).FontBold = IIf(Mid$(r(FI(k)), 1, 1) = "1", True, False)
                       Label2(k).FontItalic = IIf(Mid$(r(FI(k)), 2, 1) = "1", True, False)
                       Label2(k).FontStrikethru = IIf(Mid$(r(FI(k)), 3, 1) = "1", True, False)
                       Label2(k).FontUnderline = IIf(Mid$(r(FI(k)), 4, 1) = "1", True, False)
                       Label2(k).FontSize = Val(Mid$(r(FI(k)), 5, 7))
                       Label2(k).FontName = Mid$(r(FI(k)), 12, 30)
                  End If
           Next
                
                
                
                
          
          
          
      On Error GoTo 0
      r.Close
       
'      Label2(Index).FontName = CommonDialog1(Index).FontName
 '     Label2(Index).FontBold = CommonDialog1(Index).FontBold
  '    Label2(Index).FontSize = CommonDialog1(Index).FontSize
   '   Label2(Index).FontItalic = CommonDialog1(Index).FontItalic
      
    '  Label2(Index).FontStrikethru = CommonDialog1(Index).FontStrikethru
     ' Label2(Index).FontUnderline = CommonDialog1(Index).FontUnderline
      'On Error GoTo 0
       
Label1(0).Caption = mL_CapRes(870) '"Φόρμα ανεύρεσης  χημικών"  //l1(0)
Label1(1).Caption = mL_CapRes(871) ' "Φόρμα διόρθωσης  χημικών" //    1
Label1(2).Caption = mL_CapRes(872) ' "Φόρμα εργασιών" // 2
Label1(3).Caption = mL_CapRes(873) ' "Φόρμα κατασκευής συνταγών" //3
Label1(4).Caption = mL_CapRes(874) ' "Λίστα αναμονής" //4
Label1(5).Caption = mL_CapRes(875) ' "Φόρμα εκτέλεσης συνταγών" //5
Label1(6).Caption = mL_CapRes(876) ' "Φόρμα ανανέωσης διαλυμάτων"  //6
Label1(7).Caption = mL_CapRes(877) ' "Φόρμα αυτόματης κατασκευής διαλυμάτων" //7
Label1(8).Caption = mL_CapRes(878) ' "Τοποθέτηση μπουκαλιών σε τραπέζι"  //8
Label1(9).Caption = mL_CapRes(879) '"Φόρμα πελατών/πανιών/διαδικασιών"  //9

Label2(0).Caption = mL_CapRes(870) '"Φόρμα ανεύρεσης  χημικών"  //l1(0)
Label2(1).Caption = mL_CapRes(871) ' "Φόρμα διόρθωσης  χημικών" //    1
Label2(2).Caption = mL_CapRes(872) ' "Φόρμα εργασιών" // 2
Label2(3).Caption = mL_CapRes(873) ' "Φόρμα κατασκευής συνταγών" //3
Label2(4).Caption = mL_CapRes(874) ' "Λίστα αναμονής" //4
Label2(5).Caption = mL_CapRes(875) ' "Φόρμα εκτέλεσης συνταγών" //5
Label2(6).Caption = mL_CapRes(876) ' "Φόρμα ανανέωσης διαλυμάτων"  //6
Label2(7).Caption = mL_CapRes(877) ' "Φόρμα αυτόματης κατασκευής διαλυμάτων" //7
Label2(8).Caption = mL_CapRes(878) ' "Τοποθέτηση μπουκαλιών σε τραπέζι"  //8
Label2(9).Caption = mL_CapRes(879) '"Φόρμα πελατών/πανιών/διαδικασιών"  //9






Label3.Caption = mL_CapRes(880) ' "Χρώμα" // l3
Label4.Caption = mL_CapRes(881) ' "Γραμματοσειρά"  //l4

       
Command1.Caption = mL_CapRes(882) ' "εξοδος" // l3
       
       
       
       
       
       
       
       
End Sub


 








Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True
End Sub


Private Sub Label1_Click(Index As Integer)

      CommonDialog1(Index).ShowColor
      Label1(Index).BackColor = CommonDialog1(Index).Color


End Sub


Private Sub Label1_DblClick(Index As Integer)
    On Error Resume Next
      CommonDialog1(Index).Flags = cdlCFScreenFonts
      CommonDialog1(Index).ShowFont
      Label2(Index).FontName = CommonDialog1(Index).FontName
      Label2(Index).FontBold = CommonDialog1(Index).FontBold
      Label2(Index).FontSize = CommonDialog1(Index).FontSize
      Label2(Index).FontItalic = CommonDialog1(Index).FontItalic
      
      Label2(Index).FontStrikethru = CommonDialog1(Index).FontStrikethru
      Label2(Index).FontUnderline = CommonDialog1(Index).FontUnderline
      On Error GoTo 0

End Sub


Private Sub Label2_Click(Index As Integer)
     On Error Resume Next
      CommonDialog1(Index).Flags = cdlCFScreenFonts
      CommonDialog1(Index).ShowFont
      Label2(Index).FontName = CommonDialog1(Index).FontName
      Label2(Index).FontBold = CommonDialog1(Index).FontBold
      Label2(Index).FontSize = CommonDialog1(Index).FontSize
      Label2(Index).FontItalic = CommonDialog1(Index).FontItalic
      
      Label2(Index).FontStrikethru = CommonDialog1(Index).FontStrikethru
      Label2(Index).FontUnderline = CommonDialog1(Index).FontUnderline
      On Error GoTo 0
End Sub


