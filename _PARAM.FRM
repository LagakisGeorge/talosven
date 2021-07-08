VERSION 4.00
Begin VB.Form PARAMETROIOLD 
   ClientHeight    =   6780
   ClientLeft      =   1605
   ClientTop       =   225
   ClientWidth     =   5445
   Height          =   7185
   Left            =   1545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   5445
   Top             =   -120
   Width           =   5565
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      Height          =   360
      Left            =   6720
      TabIndex        =   25
      Top             =   4920
      Width           =   1245
   End
   Begin VB.TextBox Text9 
      DataField       =   "TIME_PREHEAT"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   6600
      TabIndex        =   23
      Top             =   405
      Width           =   1245
   End
   Begin VB.TextBox Text8 
      DataField       =   "MAXTIMEWATER"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   4560
      TabIndex        =   21
      Top             =   4920
      Width           =   1245
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Numbers on Labels"
      Height          =   435
      Left            =   690
      TabIndex        =   20
      Top             =   2340
      Width           =   2610
   End
   Begin VB.TextBox Text7 
      DataField       =   "fornext_big"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   6840
      TabIndex        =   17
      Top             =   4215
      Width           =   1245
   End
   Begin VB.TextBox Text6 
      DataField       =   "fornext_small"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   4560
      TabIndex        =   16
      Top             =   4245
      Width           =   1245
   End
   Begin VB.TextBox Text5 
      DataField       =   "metatopish_W"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   6825
      TabIndex        =   14
      Top             =   1200
      Width           =   1245
   End
   Begin VB.TextBox Text4 
      DataField       =   "MPANIO"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   4545
      TabIndex        =   12
      Top             =   3480
      Width           =   1245
   End
   Begin VB.TextBox Text3 
      DataField       =   "uesh_alatioy"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   6825
      TabIndex        =   10
      Top             =   3450
      Width           =   1245
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Δοκιμαστική συνταγή"
      DataField       =   "dokimastikh_syntagh"
      DataSource      =   "Data1"
      Height          =   435
      Left            =   690
      TabIndex        =   9
      Top             =   1860
      Width           =   2610
   End
   Begin VB.TextBox Text2 
      DataField       =   "AXRHSIMOPOIHTES_ORES"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   6810
      TabIndex        =   7
      Top             =   2400
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      DataField       =   "default_baros_panioy"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   4560
      TabIndex        =   5
      Top             =   2385
      Width           =   1245
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      Exclusive       =   0   'False
      Height          =   330
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "users"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      Height          =   570
      Left            =   600
      TabIndex        =   3
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Αλλάζω fonts"
      DataField       =   "allazo_fonts"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   690
      TabIndex        =   2
      Top             =   1335
      Width           =   2610
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Αλλάζω χρώματα"
      DataField       =   "allazo_xromata"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   690
      TabIndex        =   1
      Top             =   855
      Width           =   2610
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      Exclusive       =   0   'False
      Height          =   360
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "parametroi"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Προζυγισμένα πανιά"
      DataField       =   "prozygismena"
      DataSource      =   "Data1"
      Height          =   330
      Left            =   690
      TabIndex        =   0
      Top             =   375
      Width           =   2610
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Water Dispensing Parameter"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   26
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Preheating time"
      Height          =   285
      Left            =   6615
      TabIndex        =   24
      Top             =   45
      Width           =   2610
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Max.Time in Recipe Water Dispense (s)"
      Height          =   285
      Left            =   3120
      TabIndex        =   22
      Top             =   4680
      Width           =   3315
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Big Step"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6135
      TabIndex        =   19
      Top             =   3855
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Small Step"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3855
      TabIndex        =   18
      Top             =   3885
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Displacement of W"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Default Liquor Ratio"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   13
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salt Position"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   11
      Top             =   3090
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Οριο ωρών για δοκιμαστικές συνταγές"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6105
      TabIndex        =   8
      Top             =   1845
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Default βάρος πανιού"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   2025
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_PARAM.frx":0000
      Height          =   1500
      Left            =   3690
      OleObjectBlob   =   "_PARAM.frx":000E
      TabIndex        =   4
      Top             =   285
      Width           =   2670
   End
   Begin VB.Shape Shape1 
      Height          =   3690
      Left            =   525
      Top             =   285
      Width           =   3045
   End
End
Attribute VB_Name = "PARAMETROIOLD"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub Check5_Click()
If Check5 Then numbers_on_labels = 1
End Sub


Private Sub Command1_Click()
'Option Explicit
  'Data1.Refresh
   Set r = Data1.Recordset
   
   'r.Edit
   'r.update
    prozygismena = IIf(Check1, -1, 0)
  allazo_xromata = IIf(Check2, -1, 0)
  allazo_fonts = IIf(Check3, -1, 0)
 AXRHSIMOPOIHTES_ORES = r("AXRHSIMOPOIHTES_ORES")
 UESH_ALATIOY = Val(Text3.text)
 dokimastikh_syntagh = IIf(Check4, -1, 0)
 Default_baros_panioy = Val(Text1.text)
 
 r.MoveNext
 
 r.Edit
 r("maxtimewater") = Text10.text
 Avance_Online_Water = Val(Text10.text)
 r.update
 
 
 
  Unload Me
End Sub


Private Sub Command2_Click()


  


End Sub

Private Sub Form_Load()
Check1.Caption = mL_CapRes(750)  ' "Προζυγισμένα πανιά"
Check2.Caption = mL_CapRes(751)  'Αλλάζω χρώματα
Check3.Caption = mL_CapRes(752)  'Αλλάζω fonts
Check4.Caption = mL_CapRes(753)  'Δοκιμαστική συνταγή
Check5.Caption = mLResnew(924, "Numbers on Labels", 0)
DBGrid1.Caption = mL_CapRes(754) 'Χειριστές
DBGrid1.Columns(0).Caption = mL_CapRes(755) ' χειριστής
Label2.Caption = mL_CapRes(757) ' Οριο ωρών για δοκιμαστικές συνταγές
Label1.Caption = mL_CapRes(756) 'Default βάρος πανιού
Label9.Caption = mLResnew(925, "Preheating Time ms(1000=1sec*", 0)
Label10.Caption = mLResnew(926, "Water Dispensing Parameter*", 0)

Command1.Caption = mL_CapRes(759) 'Εξοδος

If DOYLEYO_ALATI = 0 Then
     Text3.Visible = False
     Label3.Visible = False
End If







'default μπάνιο
Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
ok = 0
For k = 0 To r.Fields.Count - 1
     If UCase(r.Fields(k).Name) = "MPANIO" Then M_AP = 1
Next

r.MoveNext

On Error GoTo addmore
Text10.text = r("maxtimewater")
'Avance_Online_Water = r("maxtimewater")

r.Close
On Error Resume Next
   If M_AP = 0 Then mydb.Execute "ALTER TABLE PARAMETROI ADD COLUMN MPANIO TEXT(10)"
'------------
Exit Sub

addmore:
       r.AddNew
       r.update
       Resume Next
End Sub

Private Sub Text3_LostFocus()
     If Not IsNumeric(Text3.text) Then
        Text3.text = 0
     End If
     If Val(Text3.text) >= Ar_1_mpoykal And Val(Text3.text) <= Ar_tel_mpoykal Then
    Else
      Text3.text = 0
    End If
End Sub
