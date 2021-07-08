VERSION 4.00
Begin VB.Form Form2 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   6810
   ClientLeft      =   2340
   ClientTop       =   2085
   ClientWidth     =   8685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   7215
   Left            =   2280
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   8685
   Top             =   1740
   Width           =   8805
   Begin VB.CommandButton Command15 
      Caption         =   " ΑΠΟ ζυγό  σε ανάδευση"
      Height          =   525
      Left            =   7185
      TabIndex        =   26
      Top             =   5280
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\WATCH.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "commun"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.CommandButton Command14 
      Caption         =   "φερε ποτηρι σε ζυγό ΔΙΑΛ"
      Height          =   525
      Left            =   4455
      TabIndex        =   24
      Top             =   4725
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Κεφαλή σκόνης σε ζυγό"
      Height          =   555
      Left            =   45
      TabIndex        =   23
      Top             =   4635
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command12 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   2820
      TabIndex        =   22
      Top             =   3645
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   4560
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Please Minimize Window"
      Height          =   2220
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   7455
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   300
      Left            =   8760
      TabIndex        =   15
      Text            =   "201"
      Top             =   6240
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.CommandButton Command10 
      Caption         =   "πανε-ελα"
      Height          =   390
      Left            =   7680
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ΕΠΕΣΤΡΕΨΕ ΜΠΟΥΚΑΛΑ ΣΤΗΝ ΘΕΣΗ ΤΗΣ"
      Height          =   615
      Left            =   6105
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ΔΕΞΙΑ ΣΕ ΔΟΣΟΜΕΤΡΗΣΗ(65)"
      Height          =   645
      Left            =   1305
      TabIndex        =   12
      Top             =   5265
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ARISTERH SE DOSOM"
      Height          =   600
      Left            =   2955
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NERO SE DOSOM"
      Height          =   675
      Left            =   45
      TabIndex        =   10
      Top             =   5265
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ΦΕΡΕ ΜΠΟΥΚΑΛΑ ΣΕ ΚΕΦΑΛΕΣ"
      Height          =   615
      Left            =   4845
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ΠΑΡΕ ποτηρι ΑΠΟ ζυγό"
      Height          =   525
      Left            =   5760
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "φερε ποτηρι σε ζυγό ΣΥΝΤ"
      Height          =   525
      Left            =   4455
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eπέστρεψε μπουκάλα απο χειριστή"
      Height          =   630
      Left            =   1320
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   3060
      TabIndex        =   3
      Text            =   "61"
      Top             =   6405
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3060
      TabIndex        =   2
      Text            =   "1"
      Top             =   5985
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "φέρε μπουκάλα σε χειριστή"
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Vers.2.05.27"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   27
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Version με νέες συνταταγμένες"
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Executing Command"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "!var12=110:"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   7455
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "ΔΕΞΙΑ"
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "ΑΡΙΣΤΕΡΗ"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "μπουκάλι"
      Height          =   285
      Left            =   3765
      TabIndex        =   5
      Top             =   6420
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "ποτήρι"
      Height          =   285
      Left            =   3825
      TabIndex        =   4
      Top             =   5985
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Const gripper_anoiktos = 0
Const gripper_kleistos = 1
Const Keno = "..."
Dim ArMp_Mixer19 ' ;Μπουκαλα se Mixer
Dim ArMp_ArKef21 ' ;ArMp_ArKef21=Μπουκαλα στην αριστερη κεφαλη
Dim ArMp_DexKef22 '= μπουκαλα στην δεξια κεφαλη
Dim F_0epis_1seZyg35 'Flag =0 επιστροφη στην αρχικη θεση ,= 1 σε ζυγαρια γ κεφαλες
Dim Proth_fora
Dim X_Mpoykal_36 '= var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
Dim Y_Mpoykal_37 '= var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
Dim X_2Mpoykal_38 '= var(38) '   Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
Dim Y_2Mpoykal_39 '= var(39) '    Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
Dim TaxyThtes
'Dim var(44) ' Ποσες μπουκαλες υπαρχουν στις κε­αλες
Dim W_dex_kef_zyg65 ' = var(65) '
Dim W_ar_kef_zyg66 ' = var(66) '
Dim W_Ner_kef_zyg67 ' = var(67) '
Dim W_Skon_kef_zyg68 ' = var(68)

Dim X_Zyg_85 '= var(85) '    Συντεταγμενη ζυγαριας κατα Χ
Dim Y_Zyg_86 '= var(86) '     Συντεταγμενη ζυγαριας κατα Υ
Dim Z_Zyg_87 '= var(887) '     Συντεταγμενη ζυγαριας κατω Ζ
Dim Z_Mpoyk_92 ' κατω Z  μπουκάλας
Dim X_Mpoyk_80  ' ;var80     x μπουκάλας
Dim Y_Mpoyk_81  ' ;VAR81     Συντεταγμενη κατα Υ της προς ευρεση μπουκαλας

Dim Ano_Z_93 '= var(93)
Dim X_Pot95 ' VAR95 Num 1εως ;VAR51, Beaker: Συντεταγμενη 1ου ποτηριου κατα Χ
Dim Y_Pot96 ';VAR96     Num 1 εως ;VAR51, Beaker: υντεταγμενη 1ου ποτηριου κατα Υ
Dim Z_Pot97 ' ;VAR97     Συντεταγμενη 1ου ποτηριου κατω Ζ
Dim Epitaxynseis
Dim Y_diadromos62 '= var(75) + var(58) + var(58) ';Διαδρομος μπουκαλιων
Dim dejia, aristerh




Function Exo_Parea()
    Dim timing1, tim2
On Error Resume Next
           Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
           Set r = dbsJet.OpenRecordset("commun")
           timing1 = r("ora")
           r.Close


If IsNull(timing1) Then
   Exo_Parea = 0
Else
  If Abs(Val(timing1) - GetCurrentTime()) < 3000 Then
     Exo_Parea = 1
  Else
     Exo_Parea = 0
  End If
End If
On Error GoTo 0

End Function

Sub PAO_ARXH()
    
' ΑΝΕΒΑΣΜΑ
    anixe
    
    
    
    
    ' ανέβα πάνω Ζ
    z0 = Ano_Z_93
    
    Find_Axes
    z0 = Ano_Z_93
    RobSend ("!d , ," + Str(Ano_Z_93) + ":")
    RobSend_w ("!go,,1:")



' πήγαινε πάνω από το αντικείμενο

'If dejia + aristerh = 0 Then
     RobSend ("!comexc1:")
    Find_Axes
                                        'kai den eisai pros kefales meria
If z0 - Zt > 11000 And (Not (X_arxh < 10000 + var(70) And Y_arxh < var(62))) Then ' psilotera kata  5cm toylaxiston
  '    RobSend ("!gowhen ,,2pm>" + LTrim(var(62)) + ":")
        
        
    ' einai pano stis kefales   na kanei ligo piso kai na ksekinisei
       If x0 < 10000 + var(70) And y0 < var(62) Then
            Find_Axes
            y0 = var(62)
           RobSend ("!d ," + Str(y0) + ":")
           RobSend_w ("!go ,1:")
           Find_Axes
       End If
       
    If var(12) > 1500 And var(12) < 1800 Then
       x0 = X_arxh: y0 = Y_arxh: w0 = var(68): z0 = Z_arxh + 25000   ' normal 11000 karaso 25000
       RobSend ("!d " + Str(x0) + "," + Str(y0) + "," + Str(z0) + "," + Str(w0) + ":")
       RobSend_w ("!go 1,1,1,1:")
    Else
      x0 = X_arxh: y0 = Y_arxh: z0 = Z_arxh + 25000   ' normal 11000 karaso 25000
      RobSend ("!d " + Str(x0) + "," + Str(y0) + "," + Str(z0) + ":")
      RobSend_w ("!go 1,1,1:")
    End If
    
    
    
Else


    If var(12) > 1500 And var(12) < 1800 Then
       x0 = X_arxh: y0 = Y_arxh: w0 = var(68)
       RobSend ("!d " + Str(X_arxh) + "," + Str(Y_arxh) + ",," + Str(w0) + ":")
       RobSend_w ("!go 1,1,,1:")
    Else
      x0 = X_arxh: y0 = Y_arxh
      RobSend ("!d " + Str(X_arxh) + "," + Str(Y_arxh) + ":")
      RobSend_w ("!go 1,1:")
    End If
   
End If
     
     
     
     
     RobSend ("!comexc0:")


'Else
    



' πηγαινε τον W σε θέση παρκαρίσματος  όταν είναι
' ;Επαναφορά ποτηριού απο ζυγό στην θέση του 4000-5000
' ή  ποτήρι  από χειριστη σε ζυγό 3000-4000
' ή ποτήρι σε ζυγαριά για ΔΙΑΛΥΜΑ  19000-19020
' ή μπουκάλι από ανάδευση σε ζυγό  101
'ή μπουκάλι από ζυγό σε ανάδευση  100
If (var(12) > 3000 And var(12) < 5000) Or (var(12) >= 19000 And var(12) <= 19030) Or var(12) = 100 Or (var(12) >= 101 And var(12) < 130) Then
    ' παρκάρισε τον W
    Find_Axes
    w0 = var(61)
    RobSend ("!d ,,," + Str(w0) + ":")
    RobSend_w ("!go,,,1:")
End If

' κατέβα κάτω Ζ
    Find_Axes
    z0 = Z_arxh
    RobSend ("!d , ," + Str(Z_arxh) + ":")
    RobSend_w ("!go,,1:")



End Sub

Sub metak_fortomenos()

' ΑΝΕΒΑΣΜΑ
    
    ' ανέβα πάνω Ζ
    z0 = Ano_Z_93
    
    Find_Axes
    z0 = Ano_Z_93
    RobSend ("!d , ," + Str(Ano_Z_93) + ":")
    RobSend_w ("!go,,1:")



 
 If var(12) > 2000 And var(12) < 2500 Then   '      ;Επαναφορά μπουκάλας στη βαση της
      ' πρώτα πήγαινε μέχρι το var62    MONO TON Y
      Find_Axes
      y0 = Y_diadromos62 ' var(62)
      RobSend ("!d ," + Str(y0) + ":")
      RobSend_w ("!go ,1:")
  End If
     
If var(12) > 1000 And var(12) < 1900 Then   ' ;Τοποθέτηση μπουκάλας σε κεφαλή δοσομέτρησης
   ' πρώτα πήγαινε μέχρι το ΜΕΧΡΙ ΤΟ var62 ΚΑΙ ΣΥΓΧΡΟΝΩΣ ΤΟΝ Χ
    Find_Axes
    x0 = Xt: y0 = Y_diadromos62 ' var(62)
    RobSend ("!d " + Str(x0) + "," + Str(y0) + ":")
    RobSend_w ("!go 1,1:")
End If

    
' πηγαινε στο τελικο σημείο
    Find_Axes
    x0 = Xt: y0 = Yt
    RobSend ("!d " + Str(x0) + "," + Str(y0) + ":")
    RobSend_w ("!go 1,1:")



    ' κατέβα κάτω Ζ
    Find_Axes
    z0 = Zt
    RobSend ("!d , ," + Str(Zt) + ":")
    RobSend_w ("!go,,1:")

End Sub






Private Sub Command1_Click()
'μεταφορά μπουκαλιού var(12) σε θέση ποτηριού var(13)
' var(12):μπουκάλι   var(13):ποτήρι
var(12) = 6000 + Val(Text2.Text)
var(13) = Val(Text1.Text)
var(94) = 18
find_uesh ("!var12=" + Str(var(12)) + ":!var13=" + Str(var(13)) + ":")

Movement

End Sub

Private Sub Command10_Click()
  ' επεστρεψε  ΜΠΟΥΚΑΛΙ  απο  ΚΕΦΑΛΗ
   cmd$ = "!var12=" + Str(100000 + Val(Text2.Text)) + ":!var13=" + Str(Val(Text3.Text)) + ":"
   
   var(12) = 2000 + Val(Text2.Text)
   find_uesh (cmd$)
   Movement
             
             
             
   var(36) = var(38)
   var(37) = var(39)
  ' φερε ΜΠΟΥΚΑΛΙ  σε ΚΕΦΑΛΗ
   var(12) = 1000 + Val(Text2.Text)
   Movement





End Sub
Sub park()
     '  ΠΑΡΚΑΡΙΣΜΑ
     RobSend ("d ,," + Str(var(93)) + ":")
     Find_Axes
     z0 = var(93)
     RobSend_w ("go,,1:")
     
     Find_Axes
     x0 = -1000: y0 = -4500
     RobSend ("!d " + Str(x0) + "," + Str(y0) + ":")
     RobSend_w ("!go11:")
     
     Find_Axes
     w0 = var(61)
     RobSend ("!d,,," + Str(w0) + ":")
     RobSend_w ("!go,,,1:")
  
End Sub
'
'
'
Private Sub Begining()

Dim dbsJet As Database
Dim r As Recordset, MHome
Dim cmd$, Runner, tim

If Proth_fora = 1 Then Exit Sub
Proth_fora = 1
Command11.Caption = "Robot program is executing "
Runner = 0
Form2.WindowState = 1

' σβηνω καμιά εντολή ξεχασμένη
        Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
        Set r = dbsJet.OpenRecordset("commun")
         r.Edit
          r("diabasthke") = 1
         r("ektelesthke") = 0
         r("cmd") = Keno
         r.Update
         r.Close
       
'DDE
       Label1.Caption = Keno


cmd$ = "!out." + LTrim(Str(bit_ok)) + "-1:"
     temp% = SendATBlock(768, cmd$, 0)



'LOOP MAIN
Do While True

RobSend ("!INFNC2-A:")
RobSend ("!out.7-0:")
RobSend ("!out.7-1:")
RobSend (TaxyThtes)
RobSend (Epitaxynseis)
RobSend (Epibradynseis)
tim = Now

'διαβάζω την  εντολη cmd$ από το αρχείο Commun
Do While True
        Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
        Set r = dbsJet.OpenRecordset("commun")
        
        If Command12.Enabled = False Then
               RobSend ("!k")
             If Lim1 = 17 Then
                RobSend ("!lh 3,3,3,3:")
                RobSend ("!ls 0,0,0,0:")
             End If
                End ' πάτησε εξοδο ο Χρήστης
        End If
        If r("bges_exo") = 1 Then  ' απο πρόγραμμα συνταγών (όταν κλείνει)
               RobSend ("!k")
            If Lim1 = 17 Then
               MilSec 100
                RobSend ("!lh 3,3,3,3:")
                MilSec 100
                RobSend ("!ls 0,0,0,0:")
            End If
                End
        End If
        
        If Label1.Caption <> Keno Then Exit Do
        
        'If r("diabasthke") = 0 And r("cmd") <> Keno Then Exit Do
        
        DoEvents
        r.Close
        
        '-------------------------------------------
        ' γράφει την ώρα στο πεδίο ORA για να βλέπω αν τρέχει ήδη το πρόγραμμα
        Runner = Runner + 1
        If Runner = 100 Then
           Runner = 0
           Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
           Set r = dbsJet.OpenRecordset("commun")
           On Error Resume Next
           r.Edit
           r("ora") = Str(GetCurrentTime())
           r.Update
           r.Close
           On Error GoTo 0
        End If
        '------------------------------------
Loop
         
         r.Edit
         r("ora") = Str(GetCurrentTime())
         cmd$ = r("cmd")
         Label1.Caption = cmd$
         r("diabasthke") = 1
         r("ektelesthke") = 0
         r("cmd") = Keno
         r.Update
         r.Close
         
         
' DDE
         cmd$ = Label1.Caption
         Label1.Caption = Keno
         var(13) = 0
         
         
             
             
             Set r = dbsJet.OpenRecordset("moves")
             r.AddNew
               r("movement") = "--------- " + left$(cmd$, 130)
               r("datetime") = Now
             r.Update
             r.Close
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         '   ξεχωρίζω το var12
    m = InStr(1, cmd$, "var12")
    var(12) = Val(Mid$(cmd$, InStr(1, cmd$, "=", 1) + 1, 8))
    
    X_arxh = 0
    Y_arxh = 0
    Z_arxh = 0
        
    Xt = 0
    Yt = 0
    Zt = 0
    
    find_uesh (cmd$)
   
   
   RobSend ("!out.14-0:")
   
   
 If var(12) = 90 Then ' δεξιά κεφαλή σε ζυγαριά
        Find_Axes
        w0 = var(65)
        RobSend ("!d ,,," + Str(w0) + ":")
        RobSend_w ("!go,,,1:")
ElseIf var(12) >= 20000 And var(12) < 100000 Then
        druv
ElseIf var(12) = 91 Then ' αριστερή κεφαλή σε ζυγαριά
        Find_Axes
        w0 = var(66)
        RobSend ("!d ,,," + Str(w0) + ":")
        RobSend_w ("!go,,,1:")
ElseIf var(12) = 93 Then ' αριστερή κεφαλή σε ζυγαριά
        Find_Axes
        w0 = var(68)
        RobSend ("!d ,,," + Str(w0) + ":")
        RobSend_w ("!go,,,1:")




ElseIf var(12) = 82 Then ' parking
       park
        
     
ElseIf var(12) > 100000 Then   ' πανε ελα
   
   ' επεστρεψε  ΜΠΟΥΚΑΛΙ  απο  ΚΕΦΑΛΗ
    var(13) = Val(Mid$(cmd$, InStr(1, cmd$, "var13=", 1) + 6, 7))
    var(12) = var(12) - 100000
    var(12) = 2000 + var(12)
    find_uesh (cmd$)
    Movement
             
             
  ' φερε ΜΠΟΥΚΑΛΙ  σε ΚΕΦΑΛΗ
   var(36) = var(38)  ' συντεταγμένες της μπουκάλας που θα έρθει στις κεφαλές
   var(37) = var(39)
   var(12) = 1000 + var(13)
   Movement

 ElseIf var(12) >= 1000 And var(12) <= 1900 Then
     var(13) = IIf(InStr(1, cmd$, "var13") > 0, 1, 0)
     'var13 = 1 ' συνταγές (αλάτι)
     'var13 = 0 'dialymata
      Movement
 Else
        Movement
 End If
   
   RobSend ("!out.14-1:")
       'ενημερώνω το αρχείο Commun οτι εκτελέστηκε η εντολή
        Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
        Set r = dbsJet.OpenRecordset("commun")
         r.Edit
         r("ektelesthke") = 1
         r.Update
         r.Close

MilSec 100
Loop

End Sub

Private Sub Command11_Click()
 ' a = begin_talos(0)
 ' Begining
End Sub

Private Sub Command11_GotFocus()
Dim lsneg, lspos
  If Exo_Parea = 1 Then End
    
  a = begin_talos(0)
  MilSec 200

 Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
     Set r = dbsJet.OpenRecordset("parametroi")
     lsneg = r("lsneg")
     lspos = r("lspos")
     r.Close



Lim1 = 15
Lim2 = 16

'    RobSend ("!lh 0,0,0,0:")
 '   RobSend ("!ls 3,3,3,3:")
  '  RobSend ("!lsneg " + lsneg + ":")
   ' RobSend ("!lspos " + lspos + ":")
    'Lim1 = 17
   ' Lim2 = 18

  
  
  Begining
End Sub


Private Sub Command12_Click()
   Command12.Enabled = False
End Sub

Private Sub Command13_Click()
    Find_Axes
    w0 = var(68)
    RobSend ("!d ,,," + Str(w0) + ":")
    RobSend_w ("!go,,,1:")
End Sub

Private Sub Command14_Click()
   ' φερε ποτήρι σε ζυγό Γ ΙΑ  ΔΙΑΛΥΜΑ
   var(12) = 19000 + Val(Text1.Text)
   cmd$ = "!var12=" + Str(var(12)) + ":"
   find_uesh (cmd$)
   Movement
End Sub

Private Sub Command15_Click()
'  από ζυγό σε αναδευση
   var(12) = 100
   cmd$ = "!var12=100:"
   find_uesh (cmd$)
   Movement
End Sub

Private Sub Command2_Click()
' επιστροφή μπουκαλιού από χειριστή στην θέση του

var(12) = 12000 + Val(Text2.Text)
var(13) = Val(Text1.Text)

cmd$ = "!var12=" + Str(var(12)) + ":!var13=" + Str(var(13)) + ":"
find_uesh (cmd$)
Movement

End Sub


Sub find_arxh()
'--------------------------------------------

If var(12) = 110 Then ' zygos=>mixer 1
   ' συντ.ζυγού
    X_arxh = var(85): Y_arxh = var(86): Z_arxh = var(87)
   ' συντεταγμένες ανάδευσης
    Xt = var(70): Yt = var(71): Zt = var(72)
End If
If var(12) = 111 Then ' mixer 1  =>  zygos
   ' συντ.ζυγού
    X_arxh = var(70): Y_arxh = var(71): Z_arxh = var(72)
   ' συντεταγμένες ανάδευσης
    Xt = var(85): Yt = var(86): Zt = var(87)
End If
If var(12) = 112 Then ' zygos=>mixer 1 χωρίς να το αφήνει
   ' συντ.ζυγού
    X_arxh = var(85): Y_arxh = var(86): Z_arxh = var(87)
   ' συντεταγμένες ανάδευσης
    Xt = var(70): Yt = var(71): Zt = var(72)
End If







If var(12) = 120 Then ' zygos=>mixer 2
   ' συντ.ζυγού
    X_arxh = var(85): Y_arxh = var(86): Z_arxh = var(87)
      ' συντεταγμένες ανάδευσης
    Xt = var(75): Yt = var(76): Zt = var(73)
End If

If var(12) = 121 Then ' mixer 2 =>  zygos
      ' συντεταγμένες ανάδευσης
    X_arxh = var(75): Y_arxh = var(76): Z_arxh = var(73)
   ' συντ.ζυγού
    Xt = var(85): Yt = var(86): Zt = var(87)
End If
If var(12) = 122 Then ' zygos => mixer 2   χωρίς να το αφήνει
   ' συντ.ζυγού
    X_arxh = var(85): Y_arxh = var(86): Z_arxh = var(87)
      ' συντεταγμένες ανάδευσης
    Xt = var(75): Yt = var(76): Zt = var(73)
End If









' ;Μεταφορά μπουκαλιων στο χειριστή
If var(12) > 6000 And var(12) < 12000 Then
   ' συντεταγμένες μπουκαλιού
    X_arxh = var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
    Y_arxh = var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
    Z_arxh = var(35)
    
    ' αποθηκεύω τις συντεταγμένες του ποτηριού σε άλλες μεταβλητές
    Xt = var(24)
    Yt = var(77)
    Zt = var(89)
End If
    
    
' ;Επαναφορά μπουκαλιων στη θέση τους  από χειριστη
If var(12) > 12000 And var(12) < 18000 Then
       
    ' αποθηκεύω τις συντεταγμένες του ποτηριού σε άλλες μεταβλητές
    X_arxh = var(24)
    Y_arxh = var(77)
    Z_arxh = var(89)
       
       
    ' συντετ. μπουκαλιού
    Xt = var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
    Yt = var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
    Zt = var(35)
       
    '   apo_pot_se_buk
End If



' ποτήρι  από χειριστη σε ζυγό (ΣΥΝΤΑΓΗ 3001-3999) ή  (ΔΙΑΛΥΜΑ 19000-19030)
If var(12) > 3000 And var(12) < 4000 Or (var(12) >= 19000 And var(12) <= 19030) Then
    ' συντεταγμένες του ποτηριού
    X_arxh = var(24)
    Y_arxh = var(77)
    Z_arxh = var(89)
   
   ' συντ.ζυγού
    Xt = var(85)
    Yt = var(86)
    Zt = var(87)
End If

If var(12) = 101 Then ' ΑΠΟ ΑΝΑΔΕΥΣΗ ΣΕ ΖΥΓΟ
    ' συντεταγμένες ανάδευσης
    X_arxh = var(70)
    Y_arxh = var(71)
    Z_arxh = var(72)
   
   ' συντ.ζυγού
    Xt = var(85)
    Yt = var(86)
    Zt = var(87)
End If

If var(12) >= 200 And var(12) <= 230 Then ' ΑΠΟ ΑΝΑΔΕΥΣΗ ΣΕ θέση ποτηριού
    ' συντεταγμένες ανάδευσης
    X_arxh = var(70)
    Y_arxh = var(71)
    Z_arxh = var(72)
   
   ' συντ.ποτηριού
    Xt = var(24)
    Yt = var(77)
    Zt = var(89)
End If




' ;Επαναφορά ποτηριού απο ζυγό στην θέση του
If var(12) > 4000 And var(12) < 5000 Then
    
   ' συντ.ζυγού
    X_arxh = var(85)
    Y_arxh = var(86)
    Z_arxh = var(87)
   
    ' συντ.θέσης ποτηριού
    Xt = var(24)
    Yt = var(77)
    Zt = var(89)
End If

' ;μεταφορά  μπουκαλιού ΔΙΑΛΥΜΑΤΟΣ απο ζυγό στην ανάδευση
If var(12) = 100 Then
   ' συντ.ζυγού
    X_arxh = var(85)
    Y_arxh = var(86)
    Z_arxh = var(87)
   
    'συντ.ανάδευσης
    Xt = var(70)
    Yt = var(71)
    Zt = var(72)
End If

' ;μεταφορά  GRIPPER SE BOLAN
If var(12) = 94 Then
   ' συντ.ζυγού
       
       get_uesh (400)
       
       X_arxh = x0 ' var(78)
       Y_arxh = y0 ' var(73)
       Z_arxh = var(88)
   
    'συντ.ανάδευσης
    Xt = X_arxh
    Yt = Y_arxh
    Zt = Z_arxh
End If


If var(12) > 2000 And var(12) < 2500 Then
'   ;Επαναφορά μπουκάλας στη βαση της

    ' συντετ. μπουκαλιού
    Xt = var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
    Yt = var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
    Zt = var(35)

'Y_arxh = var(75)
'Z_arxh = var(76)

Select Case axe(4)
   Case var(65)  ' ΔΕΞΙΑ ΣΕ ΖΥΓΟ
        'If aristerh + dejia = 1 Then 'τελευταία μπουκάλα στην δεξιά κεφαλή
        If dejia = 1 And aristerh = 0 Then 'τελευταία μπουκάλα στην δεξιά κεφαλή
           get_uesh 401
           X_arxh = x0: Y_arxh = y0: Z_arxh = z0
           'X_arxh = var(78)
           dejia = 0
        Else
           get_uesh 402
           X_arxh = x0: Y_arxh = y0: Z_arxh = z0
           'X_arxh = var(78) - Abs(var(66) - var(65)) * var(31)
           aristerh = 0
        End If
   Case var(66)  ' ΑΡΙΣΤΕΡΗ ΣΕ ΖΥΓΟ
        If dejia = 0 And aristerh = 1 Then 'τελευταία μπουκάλα στην αριστ.κεφαλή
            get_uesh 412
            X_arxh = x0: Y_arxh = y0: Z_arxh = z0
            'X_arxh = var(78)
            aristerh = 0
        Else
            get_uesh 411
            X_arxh = x0: Y_arxh = y0: Z_arxh = z0
            'X_arxh = var(74)
            dejia = 0
        End If
   Case var(68)    ' ΣΚΟΝΗ
       
       get_uesh 400
            X_arxh = x0: Y_arxh = y0: Z_arxh = z0
       
       'X_arxh = var(78)
       'Y_arxh = var(73)
       'Z_arxh = var(69)
End Select



End If





' ;Τοποθέτηση μπουκάλας σε κεφαλή δοσομέτρησης
If var(12) > 1000 And var(12) <= 1300 Then

    ' συντετ. μπουκαλιού
    X_arxh = var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
    Y_arxh = var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
    Z_arxh = var(35)



'Yt = var(75)
'Zt = var(76)

Select Case axe(4)
   Case var(65) 'ΔΕΞΙΑ ΣΕ ΖΥΓΟ
       'Xt = var(78) - Abs(var(66) - var(65)) * var(31)
       aristerh = 1
       get_uesh 402
       Xt = x0: Yt = y0: Zt = z0

       
   Case var(66) 'ΑΡΙΣΤΕΡΗ ΣΕ ΖΥΓΟ
       'Xt = var(74)
        get_uesh 411
        Xt = x0: Yt = y0: Zt = z0
        dejia = 1
   Case var(67) 'νερό
       'If var(67) < var(65) Then ' η βαλβίδα νερού είναι ΔΕΞΙΑ της δεξιας κεφαλής
        '       Xt = var(78) - Abs(var(67) - var(65)) * var(31)
       'Else ' η βαλβίδα νερού είναι ΑΡΙΣΤΕΡΑ της δεξιας κεφαλής
        '       Xt = var(78) + Abs(var(67) - var(65)) * var(31)
      'End If
       get_uesh 421
       Xt = x0: Yt = y0: Zt = z0
      
       dejia = 1
   Case var(68) 'ΣΚΟΝΗ
       get_uesh 400
       Xt = x0: Yt = y0: Zt = z0
   
       'Xt = var(78)
       'Yt = var(73)
       'Zt = var(69)

  Case Else
   Stamathma_Talos "ΔΕΝ ΒΡΗΚΑ ΣΥΝΤΕΤ.ΚΕΦΑΛΗΣ"
   End
End Select


End If


' ;Τοποθέτηση μπουκάλας skonhs σε κεφαλή δοσομέτρησης
If var(12) > 1500 And var(12) <= 1800 Then

    ' συντετ. μπουκαλιού
    X_arxh = var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
    Y_arxh = var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
    Z_arxh = var(35)

   'Case var(68) 'ΣΚΟΝΗ
       get_uesh 400
       Xt = x0: Yt = y0: Zt = z0
End If







End Sub


Sub druv()

' RobSend ("!out .16 - 0:")
            
            var(8) = "."
           'Αν δεν εισαι ετοιμος druvanisma ready
           If var(6) = 0 Then
               var(6) = 1
           End If
               var(1) = var(12) - 20000
               druvan
 '          RobSend ("!out .16 - 1:")

End Sub


Sub klise()

If FindInput(Gripper_Input) = gripper_anoiktos Then  ' anoiktos
dz = -var(99)
     z0 = z0 + dz
    RobSend_w ("!ma0000:!d,," + Str(dz) + ":!go,,1:")
    RobSend ("!WAIT (mov = bxx0x):")
    RobSend ("!ma1111:")  '  ;Κατεβαινει 1 mm για να αποφυγει διαφορες Ζ
    RobSend ("!" + Gripper_Off + ":")
   Do While FindInput(Gripper_Input) = gripper_anoiktos '  kleistos
      DoEvents
   Loop
   var(10) = 1
  ' RobSend ("!t0 0.5:")
  MilSec 500
End If

End Sub


Sub anixe()

If FindInput(Gripper_Input) = gripper_kleistos Then  ' kleistos
     RobSend ("!" + Gripper_On + ":")
     Do While FindInput(Gripper_Input) = gripper_kleistos  ' sigk gripper_anoiktos
       DoEvents
    Loop
    var(10) = 0
'    MilSec 500
End If

End Sub

Sub Find_Axes()
   x0 = axe(1)
   y0 = axe(2)
   z0 = axe(3)
   w0 = axe(4)
End Sub

Sub Aneba_Pano_Z()

  ' ανέβα πάνω Ζ
    z0 = Ano_Z_93
   
    Find_Axes
    z0 = Ano_Z_93
    RobSend ("!d , ," + Str(Ano_Z_93) + ":")
    RobSend_w ("!go,,1:")
    
End Sub

Sub Movement()

Ano_Z_93 = var(93)

'1. προσδιορίζω τις συντεταγμένες του αντικειμένου
'      που θα πιάσω και τις συντεταγμένες τέρματος
find_arxh

Write_No_home 1


 ' απο ανάδ->ζυγό  ( το έχει ήδη πιασμένο το μπουκάλι)  101
 'απο ανάδ=>θέση ποτηριού  200-220  ( το έχει ήδη πιασμένο το μπουκάλι)  101
If var(12) = 101 Or (var(12) >= 200 And var(12) <= 230) Then
Else
    '2. Κινούμαι πρός το αντικείμενο
    PAO_ARXH

    '3. Κλείνω τον gripper
    klise
End If



If var(12) = 94 Then '  ΜΟΝΟ ΝΑ ΠΑΩ ΜΕΧΡΙ ΤΟ ΒΟΛΑΝ
     Exit Sub
End If


'5. Mετακινούμαι με το αντικείμενο προς το τέρμα
 metak_fortomenos


If var(12) = 100 Or var(12) = 112 Or var(12) = 122 Then ' απο ζυγό=>ανάδευση  ( δεν αφήνει το μπουκάλι μόνο του στην ανάδευση)

ElseIf var(12) > 1000 And var(12) <= 1999 And axe(4) = var(68) Then
    'ποτηρι για διάλυμα σε κεφαλή ΣΚΟΝΗΣ
    anixe


  If var(13) = 0 Then ' διαλυματα
    Find_Axes
    z0 = var(88) ' υψος βολάν
    RobSend_w ("!d,," + Str(var(88)) + ":!go,,1:")
    klise
  End If
Else
      '6. Ανοίγω τον gripper
       anixe
       ' 7 ano Z
        Aneba_Pano_Z
End If

 
'8 Κινησεις τακτοποίησης
kinhseis_taktopoihs

Write_No_home 0

End Sub
Sub kinhseis_taktopoihs()

' ετοιμάζω την κεφαλή νερού για δοσομέτρηση
If var(12) > 3000 And var(12) < 4000 Then
    Find_Axes
    w0 = var(67)
    RobSend ("!d ,,," + Str(w0) + ":")
    RobSend_w ("!go,,,1:")
     
     ' ετοιμάζω την κεφαλή σκόνης για δοσομέτρηση
     
ElseIf (var(12) >= 2000 And var(12) <= 2499 And var(13) = 0) Then 'pane mpoykali kai oxi pane-fere(var(13)>0)
          
    Find_Axes
    x0 = var(85): y0 = var(86) ' pano apo zygaria
    RobSend ("!d " + Str(x0) + "," + Str(y0) + ":")
    RobSend_w ("!go 1,1:")
   
     
     
ElseIf (var(12) >= 19000 And var(12) <= 19030) Then
 'sigk   Find_Axes
 'sigk   w0 = var(68)
 'sigk   RobSend ("!d ,,," + Str(w0) + ":")
 'sigk   RobSend_w ("!go,,,1:")
End If


End Sub

Private Sub Command3_Click()
  ' φερε ποτήρι σε ζυγό
   var(12) = 3000 + Val(Text1.Text)
   cmd$ = "!var12=" + Str(var(12)) + ":"
   find_uesh (cmd$)
   Movement
End Sub

Private Sub Command4_Click()
 ' παρε  ποτήρι από ζυγό
   var(12) = 4000 + Val(Text1.Text)
   cmd$ = "!var12=" + Str(var(12)) + ":"
   find_uesh (cmd$)
   Movement
End Sub

Private Sub Command5_Click()
  ' φερε ΜΠΟΥΚΑΛΙ  σε ΚΕΦΑΛΗ
   var(12) = 1000 + Val(Text2.Text)
   cmd$ = "!var12=" + Str(var(12)) + ":"
   find_uesh (cmd$)
   Movement
 
End Sub

Private Sub Command6_Click()
    Find_Axes
    w0 = var(67)
    RobSend ("!d ,,," + Str(w0) + ":")
    RobSend_w ("!go,,,1:")
End Sub

Private Sub Command7_Click()
   Find_Axes
    w0 = var(66)
    RobSend ("!d ,,," + Str(w0) + ":")
    RobSend_w ("!go,,,1:")
End Sub

Private Sub Command8_Click()

    Find_Axes
    w0 = var(65)
    RobSend ("!d ,,," + Str(w0) + ":")
    RobSend_w ("!go,,,1:")
End Sub


Private Sub Command9_Click()
  ' επεστρεψε  ΜΠΟΥΚΑΛΙ  από  ΚΕΦΑΛΗ
   var(12) = 2000 + Val(Text2.Text)
   cmd$ = "!var12=" + Str(var(12)) + ":"
   find_uesh (cmd$)
   Movement
End Sub

Private Sub Form_Load()

#If Win32 Then
   result% = SetDevice(0, 768)
#End If


'Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
'Set r = dbsJet.OpenRecordset("moves")
'If r.RecordCount > 10000 Then
'   r.Close
'   dbsJet.Execute "delete all from moves where  DateDiff('d',r('datetime'),Now)>1"
'End If









If App.PrevInstance Then End
RobSend ("!INFNC2-A:")
RobSend ("!out.7-0:")
RobSend ("!out.7-1:")

'RobSend ("!out.7-0:")
'RobSend ("!out.7-1:")


 Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = dbsJet.OpenRecordset("Ταχυτητες")
     
    Do Until r.EOF
         id_tax = r("id_tax")
         If id_tax = 5 Then
            vx = r("Ταχυτητα")
            a_x = r("Επιταχυνση")
            e_x = r("Επιβραδυνση")
            
            
         End If
         If id_tax = 6 Then
            vy = r("Ταχυτητα")
            a_y = r("Επιταχυνση")
            e_y = r("Επιβραδυνση")
            
         End If
         If id_tax = 7 Then
            vz = r("Ταχυτητα")
            a_z = r("Επιταχυνση")
            e_z = r("Επιβραδυνση")
            
         End If
         If id_tax = 8 Then
            vw = r("Ταχυτητα")
            a_w = r("Επιταχυνση")
            e_w = r("Επιβραδυνση")
            
         End If
         r.MoveNext
    Loop
    r.Close




TaxyThtes = "!v " + Str(vx) + "," + Str(vy) + "," + Str(vz) + "," + Str(vw) + ":"
Epitaxynseis = "!a " + Str(a_x) + "," + Str(a_y) + "," + Str(a_z) + "," + Str(a_w) + ":"
Epibradynseis = "!ad " + Str(e_x) + "," + Str(e_y) + "," + Str(e_z) + "," + Str(e_w) + ":"

Proth_fora = 0
dejia = 0: aristerh = 0
x0 = 0: y0 = 0: z0 = 0: w0 = 0
system_ready = IsOSLoaded(768)
FindValves
  Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r = dbsJet.OpenRecordset("var")
            For k = 1 To 99: var(k) = 0: Next
            Do Until r.EOF
              v = r("field_name")
              t = r("field_val")
              If IsNull(t) Then
                t = 0
              End If
                k = Val(Mid$(v, 4, 2))
                If k > 0 And k < 100 Then var(k) = t
              
              'cmd$ = "!" + V + "=" + Str$(t) + ":"
              
              'temp% = SendAT6400Block(768, cmd$, 0)
             r.MoveNext
        Loop
      r.Close

    'RobSend ("!v 32,32,32,1.2:")
   

Y_diadromos62 = var(62)    ' var(75) + var(58) + var(58) ';Διαδρομος μπουκαλιων
var(31) = (var(74) - var(78)) / (var(66) - var(65))




End Sub
Private Sub Timer1_Timer()
   Label12.Caption = aristerh
   Label14.Caption = dejia
   Elegxo_Katastash
End Sub




Sub druvan()

'If var(8) = "." Then
 '  var(8) = TaxyThtes '3v  'ΑΠΟΘΗΚΕΥΩ την ταχύτητα του 3ου άξονα
   RobSend ("!v , , 0.4:")
'End If

If var(1) = 0 Then
   druvfn
   RobSend (TaxyThtes)
   Exit Sub
End If

RobSend ("!ma0000:")

' ανέβασμα ντρουβανίσματος
Find_Axes
dz = var(1)
'                                  z0 = z0 + dz
If z0 + dz > Ano_Z_93 Then
    dz = Ano_Z_93 - z0
    var(1) = dz
End If
 
 z0 = z0 + dz
    
RobSend_w ("!d,," + Str(dz) + ":!go,,1:")


' κατέβασμα ντρουβανίσματος
var(1) = 0 - var(1)
Find_Axes
dz = var(1)
z0 = z0 + dz
RobSend_w ("!d,," + Str(dz) + ":!go,,1:")

RobSend ("!ma1111:")
End Sub


Sub druvfn() '   ;druvanisma finish
Dim aa, x1, y1, z1
Dim dddb, r, Vim

    Set dddb = Workspaces(0).OpenDatabase("c:\talos\recipies.MDB")
    Set r = dddb.OpenRecordset("parametroi")
      Vim = r("metatopish_W")
    dddb.Close



' επαναφέρω την αρχική ταχύτητα
RobSend (TaxyThtes)
anixe ' αφήνει το ποτήρι της σκόνης
RobSend ("!ma1111:")

get_uesh 400
z1 = z0
Find_Axes
z0 = z1 ' var(69) ' κατέβα να πιάσεις το μπουκάλι της σκόνης
RobSend_w ("!d,," + Str(z0) + ":!go,,1:")

klise

Find_Axes
z0 = Ano_Z_93
RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")

x0 = var(36)
y0 = var(37)
z0 = Ano_Z_93
' w0 = var(67) + Vim

robsend_3w ("!d " + Str(x0) + "," + Str(y0) + "," + Str(z0) + ":!go111:")


' κατεβάζει το μπουκάλι
Find_Axes
z0 = var(35)    '  var(92)
'w0 = var(67)
RobSend ("!d,," + Str(z0) + ":")
robsend_3w ("!go,,1:")

anixe

'Aneba_Pano_Z

    z0 = Ano_Z_93
    Find_Axes
    z0 = Ano_Z_93
    RobSend ("!d , ," + Str(Ano_Z_93) + ":")
    robsend_3w ("!go,,1:")







' ερχεται πάνω απο την ζυγαριά
Find_Axes
' συντ.ζυγού
    x0 = var(85)
    y0 = var(86)

RobSend ("!d " + Str(x0) + "," + Str(y0) + ":")
robsend_3w ("!go11:")


End Sub

Sub Write_No_home(ByVal Yes_No As Integer)
' 1 γραφει no home    0 : Home=OK

     Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r = dbsJet.OpenRecordset("mem3")

     If Yes_No = 1 Then
        r.Edit
        r("HOME") = "NO HOME"
        r("XRHSTHS") = "BUSY"
        r.Update
     Else
        r.Edit
        r("XRHSTHS") = "NO BUSY"
        r("HOME") = "HOME=OK"
        r.Update
     End If
     
     
     
     
       r.Close

End Sub
