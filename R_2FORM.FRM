VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   -1575
   ClientTop       =   870
   ClientWidth     =   9660
   Height          =   7380
   Left            =   -1635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   Top             =   525
   Width           =   9780
   Begin VB.CommandButton Command8 
      Caption         =   "ΦΕΡΕ ΠΟΤΗΡΙ ΣΤΟΝ ΖΥΓΟ ΓΙΑ ΣΚΟΝΗ"
      Height          =   330
      Left            =   225
      TabIndex        =   30
      Top             =   4185
      Width           =   4560
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   3135
      TabIndex        =   13
      Top             =   3720
      Width           =   1125
   End
   Begin VB.CommandButton Command7 
      Caption         =   "πάρε ένα μπουκάλι από κεφαλ΄ξ-φέρε άλλο μπουκάλι σε κεφαλή"
      Height          =   285
      Left            =   2565
      TabIndex        =   12
      Top             =   5895
      Width           =   5100
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ΕΠΕΣΤΡΕΨΕ ΜΠΟΥΚΑΛΙ ΑΠΟ ΚΕΦΑΛΗ ΣΤΗΝ ΘΕΣΗ ΤΟΥ"
      Height          =   330
      Left            =   45
      TabIndex        =   11
      Top             =   5580
      Width           =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "φερε μπουκάλι σε ζυγαριά "
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Top             =   5895
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1485
      TabIndex        =   8
      Text            =   "18"
      Top             =   3735
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1485
      TabIndex        =   6
      Top             =   3195
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eπιστροφή μπουκαλιού από χειριστή"
      Height          =   285
      Left            =   2565
      TabIndex        =   5
      Top             =   6120
      Width           =   3360
   End
   Begin VB.CommandButton Command4 
      Caption         =   "φερε μουκάλι σε χειριστή"
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   6165
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "φερε ποτηρι σε ζυγαριά "
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   6390
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3150
      Top             =   5805
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ΕΠΙΣΤΡΟΦΗ ΠΟΤΗΡΙΟΥ ΑΠΟ ΖΥΓΑΡΙΑ "
      Height          =   330
      Left            =   2565
      TabIndex        =   0
      Top             =   6390
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\COORD_TB.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "VAR"
      Top             =   5670
      Width           =   1140
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   390
      X2              =   5070
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Label Label20 
      Caption         =   "var75"
      Height          =   210
      Left            =   5145
      TabIndex        =   29
      Top             =   2325
      Width           =   945
   End
   Begin VB.Line Line4 
      X1              =   9885
      X2              =   9615
      Y1              =   735
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   9285
      X2              =   9615
      Y1              =   750
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   9600
      X2              =   9600
      Y1              =   2475
      Y2              =   345
   End
   Begin VB.Label Label19 
      Caption         =   "24,77,89"
      Height          =   240
      Left            =   7485
      TabIndex        =   28
      Top             =   1785
      Width           =   660
   End
   Begin VB.Label Label18 
      Caption         =   "(36,37)  και  2ο  (38,39)"
      Height          =   270
      Left            =   6135
      TabIndex        =   27
      Top             =   135
      Width           =   2145
   End
   Begin VB.Shape Shape27 
      BackStyle       =   1  'Opaque
      Height          =   780
      Left            =   6375
      Shape           =   3  'Circle
      Top             =   75
      Width           =   1035
   End
   Begin VB.Label Label17 
      Caption         =   "var73 y κεφ.σκόνης"
      Height          =   240
      Left            =   60
      TabIndex        =   26
      Top             =   2820
      Width           =   1560
   End
   Begin VB.Label Label16 
      Caption         =   "var7,var71,var72"
      Height          =   255
      Left            =   6510
      TabIndex        =   25
      Top             =   1155
      Width           =   1620
   End
   Begin VB.Label Label15 
      Caption         =   "var62"
      Height          =   210
      Left            =   5520
      TabIndex        =   24
      Top             =   555
      Width           =   945
   End
   Begin VB.Line Line1 
      X1              =   765
      X2              =   5445
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Label Label14 
      Caption         =   "var85,var86,var87"
      Height          =   255
      Left            =   2190
      TabIndex        =   23
      Top             =   2820
      Width           =   1620
   End
   Begin VB.Label Label13 
      Caption         =   "var70,var71,var72"
      Height          =   255
      Left            =   4470
      TabIndex        =   22
      Top             =   2580
      Width           =   1620
   End
   Begin VB.Shape Shape26 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   360
      Left            =   3975
      Shape           =   3  'Circle
      Top             =   2475
      Width           =   570
   End
   Begin VB.Shape Shape25 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   345
   End
   Begin VB.Label Label11 
      Caption         =   "var68"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   1005
      Width           =   600
   End
   Begin VB.Label Label12 
      Caption         =   "var67"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label10 
      Caption         =   "var66"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   915
      TabIndex        =   19
      Top             =   1665
      Width           =   600
   End
   Begin VB.Label Label9 
      Caption         =   "var65"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   495
      TabIndex        =   18
      Top             =   1995
      Width           =   480
   End
   Begin VB.Label Label8 
      Caption         =   "var61"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   15
      TabIndex        =   17
      Top             =   2415
      Width           =   600
   End
   Begin VB.Label Label7 
      Caption         =   "var74 (x) της δεξιάς κεφαλής"
      Height          =   240
      Left            =   3765
      TabIndex        =   16
      Top             =   1680
      Width           =   3045
   End
   Begin VB.Label Label6 
      Caption         =   "var75 (y) κεφαλών,VAR76   Ζ κεφαλων   "
      Height          =   240
      Left            =   1560
      TabIndex        =   15
      Top             =   105
      Width           =   5070
   End
   Begin VB.Label Label5 
      Caption         =   "var78 (x δεξιάς κεφαλής)"
      Height          =   300
      Left            =   2970
      TabIndex        =   14
      Top             =   2010
      Width           =   3105
   End
   Begin VB.Shape Shape24 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2850
      Top             =   1080
      Width           =   150
   End
   Begin VB.Shape Shape23 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   345
   End
   Begin VB.Shape Shape22 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1035
      Width           =   345
   End
   Begin VB.Shape Shape21 
      Height          =   240
      Left            =   3645
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   345
   End
   Begin VB.Shape Shape20 
      Height          =   255
      Left            =   2280
      Top             =   1020
      Width           =   1800
   End
   Begin VB.Shape Shape19 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2505
      Top             =   1380
      Width           =   150
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   2085
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   345
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   2655
      Shape           =   3  'Circle
      Top             =   1335
      Width           =   345
   End
   Begin VB.Shape Shape16 
      Height          =   240
      Left            =   3300
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   345
   End
   Begin VB.Shape Shape15 
      Height          =   255
      Left            =   1935
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Shape Shape14 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   1830
      Shape           =   3  'Circle
      Top             =   1665
      Width           =   345
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   345
   End
   Begin VB.Shape Shape12 
      Height          =   240
      Left            =   3045
      Shape           =   3  'Circle
      Top             =   1665
      Width           =   345
   End
   Begin VB.Shape Shape11 
      Height          =   255
      Left            =   1680
      Top             =   1665
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   390
      Left            =   2340
      Top             =   2400
      Width           =   705
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   675
      Shape           =   3  'Circle
      Top             =   2445
      Width           =   345
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   1170
      Shape           =   3  'Circle
      Top             =   2445
      Width           =   345
   End
   Begin VB.Shape Shape8 
      Height          =   240
      Left            =   1695
      Shape           =   3  'Circle
      Top             =   2445
      Width           =   345
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   1185
      Shape           =   3  'Circle
      Top             =   1995
      Width           =   345
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   1755
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   345
   End
   Begin VB.Shape Shape5 
      Height          =   240
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   1995
      Width           =   345
   End
   Begin VB.Shape Shape4 
      Height          =   255
      Left            =   1035
      Top             =   1995
      Width           =   1800
   End
   Begin VB.Shape Shape3 
      Height          =   225
      Left            =   660
      Top             =   2460
      Width           =   1425
   End
   Begin VB.Shape Shape2 
      Height          =   540
      Left            =   6930
      Top             =   2385
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "μπουκάλι"
      Height          =   330
      Left            =   225
      TabIndex        =   9
      Top             =   3735
      Width           =   1230
   End
   Begin VB.Label Label3 
      Caption         =   "ποτήρι"
      Height          =   330
      Left            =   225
      TabIndex        =   7
      Top             =   3195
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   5265
      Width           =   6765
   End
   Begin VB.Label Label1 
      Caption         =   "gripper"
      Height          =   285
      Left            =   3150
      TabIndex        =   1
      Top             =   3375
      Width           =   780
   End
   Begin VB.Shape Shape28 
      BackStyle       =   1  'Opaque
      Height          =   780
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   1545
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Dim ArMp_Mixer19 ' ;Μπουκαλα se Mixer
Dim ArMp_ArKef21 ' ;ArMp_ArKef21=Μπουκαλα στην αριστερη κεφαλη
Dim ArMp_DexKef22 '= μπουκαλα στην δεξια κεφαλη
Dim F_0epis_1seZyg35 'Flag =0 επιστροφη στην αρχικη θεση ,= 1 σε ζυγαρια γ κεφαλες

Dim X_Mpoykal_36 '= var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
Dim Y_Mpoykal_37 '= var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
Dim X_2Mpoykal_38 '= var(38) '   Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
Dim Y_2Mpoykal_39 '= var(39) '    Αποστολγ συντεταγεμβνης Υ μπουκαλιιν

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




Sub kefbuk() '  ;Επιστροφη μπουκαλας απο κεφαλες
     F_0epis_1seZyg35 = 0 ' ;Για επιστροφη
     anixe
     chakef
     'trace0
     vrskef
     rota
     If var(43) = 1 Then
         Find_Axes
         y0 = Y_diadromos62
         RobSend_w ("!d" + Str(var(56)) + "," + Str(Y_diadromos62) + ":!go,1:")
         x0 = var(56)
         RobSend ("!go1:")
       End If
       talos4
  If var(14) = 1 Then strofi
     topkef
     Find_Axes
     x0 = X_Mpoyk_80: y0 = Y_Mpoyk_81: z0 = Ano_Z_93
     RobSend ("!d " + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + "," + Str(Ano_Z_93) + ":")
     panbuk

End Sub





Sub metato()
' βλεπει αν χρειάζεται μετακίνηση ο αγκώνας της ΙΡΙΣ
       var(7) = 0
        var(44) = axe(3) '    ;Αν εισαι αλλού
        If var(44) <> Ano_Z_93 Then
            Find_Axes
            z0 = Ano_Z_93
            RobSend ("!d,," + Str(Ano_Z_93) + ":") '   ; Πάνε στο πανω Υ μπουκαλιων
            RobSend_w ("!go,,1:")
            'wait(3pm=var93)
        End If

        var(44) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19
        If var(44) > 0 Then   '      ' ;Αν εχεις εστω και μια μπουκαλα στις κεφαλες
            var(44) = axe(2) 'pm
           If var(44) < Y_diadromos62 Then
               var(44) = axe(1) ' pm
              If var(44) > -70000 Then
                 RobSend ("!d," + Str(Y_diadromos62) + ":") '     ;Πανε στριψε στο διαδρομο
                 Find_Axes
                 y0 = Y_diadromos62
                 RobSend_w ("!go,1:")
                 'WAIT (mov = bx0)
                 var(44) = 0
              End If
           End If
        End If
       
        If axe(1) < -125000 Then '   ; αν είσαι πολύ άκρη έλα στο -125000
              var(7) = axe(1) 'pm
              Find_Axes
              x0 = -125000
              RobSend_w ("!d " + Str(x0) + ":!go1")
             ' WAIT (mov = b0)
        End If
     
  End Sub




Sub strofi()  ' ;Προγραμμα στροφης

metato ' ;κοιτα αν πρεπει να μετακινηθεις και μετακινησου
If var(14) = 0 Then '     ; αν εισαι προς computer
         RobSend ("!out.8-0:!out.9-1:")  '    ; στριψε βραχίονα πρός τοίχο
         var(14) = 1 ';Στραμμένος πρός τοίχο
         Do While FindInput(15) = 1
               DoEvents
         Loop
         'wait(in=bxxxxxxxxxxxxxx0)   ; 15bit=0  περιμένω μέχρι να ρθει στόν τοίχο
         MilSec 1000
         Exit Sub
Else   '               ; αλλιως αν εισαι προς τοιχο
     RobSend ("!out.8-1:!out.9-0:")  '  ;  βραχίονας πρός computer
     var(14) = 0 ';Στραμμένος πρός computer
         Do While FindInput(14) = 0
               DoEvents
         Loop
     ' wait(in=bxxxxxxxxxxxxx0)   ; 14bit=0   για να ρθει προς computer
      MilSec 1000
   If var(7) > 0 Then   '           ;Αν εχεις μετακινηθει κατα Χ πριν στριψεις
      
      Find_Axes
      x0 = var(7)
      RobSend_w ("!d" + Str(var(7)) + ":!go1:")
      var(7) = 0 ';Ξανατραβα εκει που ησουν
      Exit Sub
   End If
   
If var(7) < 0 Then '           ;Αν εχεις μετακινηθει κατα Χ πριν στριψεις
   Find_Axes
   x0 = var(7)
   RobSend_w ("!d" + Str(var(7)) + ":!go1:")
   var(7) = 0 ' Ξανατραβα εκει που ησουν
   Exit Sub
End If
End Sub




Sub ferbuk()
'def ferbuk



anixe
 
 If ERGOSTASIO = "iris" Then
    Find_Axes
    z0 = Ano_Z_93
    RobSend ("!d,," + Str(Ano_Z_93) + ":")
    RobSend_w ("!go,,1:")
     If var(14) <> var(15) Then strofi
End If


rota

If var(43) = 0 Then ';Κατ'ευθειαν σκυψιμο ισιας πορειας
   
   var(44) = axe(1) 'pm
    If var(44) > X_Mpoyk_80 Then
      var(44) = var(44) - X_Mpoyk_80
    Else
      var(44) = X_Mpoyk_80 - var(44)
    End If

   var(45) = axe(2) 'pm
   If var(45) > Y_Mpoyk_81 Then
      var(45) = var(45) - Y_Mpoyk_81
   Else
      var(45) = Y_Mpoyk_81 - var(45)
   End If
  If var(44) > var(45) Then
   If X_Mpoyk_80 > axe(1) Then 'pm)
      var(44) = X_Mpoyk_80 - var(55)
      RobSend ("!gowhen(,,1pm>" + Str(var(44)) + "):")
   End If
   If X_Mpoyk_80 < axe(1) Then ' pm)
      var(44) = X_Mpoyk_80 + var(55)
      RobSend ("!gowhen(,,1pm<" + Str(var(44)) + "):")
   End If
  Else
   
     If Y_Mpoyk_81 > axe(2) Then 'pm)
        var(45) = Y_Mpoyk_81 - var(55)
        RobSend ("!gowhen(,,2pm>" + Str(var(45)) + "):")
     End If
     If Y_Mpoyk_81 < axe(2) Then 'pm)
        var(45) = Y_Mpoyk_81 + var(55)
        RobSend ("!gowhen(,,2pm<" + Str(var(45)) + "):")
     End If
  End If
  


Find_Axes
x0 = X_Mpoyk_80: y0 = Y_Mpoyk_81: z0 = Z_Mpoyk_92
RobSend ("!d" + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + "," + Str(Z_Mpoyk_92) + ":") ' ;για ισια

 
 If ERGOSTASIO = "iris" Then
    If (ArMp_ArKef21 = 0) Then
             RobSend_w ("!go1111:")
        Else
             RobSend_w ("!go111:")
     End If
 Else
        RobSend_w ("!go111:")
 End If
  



End If   ' var43=0

If var(43) = 1 Then '   ;για κλικ μονο-διπλο
   var(44) = axe(1) 'pm
   If var(44) > X_Mpoyk_80 Then
      var(44) = var(44) - X_Mpoyk_80
   Else
      var(44) = X_Mpoyk_80 - var(44)
   End If
   
   var(45) = Y_diadromos62
   If var(45) > Y_Mpoyk_81 Then
      var(45) = var(45) - Y_Mpoyk_81
   Else
      var(45) = Y_Mpoyk_81 - var(45)
   End If
  
  
  If var(44) > var(45) Then
     If X_Mpoyk_80 > axe(1) Then
        var(44) = X_Mpoyk_80 - var(55)
        RobSend ("!gowhen(2pm>" + Str(Y_diadromos62) + ",,1pm>" + Str(var(44)) + "):")
     End If
     If X_Mpoyk_80 < axe(1) Then
        var(44) = X_Mpoyk_80 + var(55)
        RobSend ("!gowhen(2pm>" + Str(Y_diadromos62) + ",,1pm<" + Str(var(44)) + "):")
     End If
  Else
     var(45) = Y_Mpoyk_81 - var(55)
     RobSend ("!gowhen(2pm>" + Str(Y_diadromos62 + ",,2pm>" + Str(var(45))) + "):")
  End If
   RobSend_w ("!d" + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + "," + Str(Z_Mpoyk_92) + ":!go111:")
End If



If var(43) = 2 Then '   ;για κλικ διπλο-μονο
  var(44) = axe(1) 'pm
  If var(44) > X_Mpoyk_80 Then
     var(44) = X_Mpoyk_80 + var(55)
     var(45) = Y_Mpoyk_81 + var(55)
     RobSend ("!gowhen(,1pm<" + Str(var(44)) + ",2pm<" + Str(var(45)) + "):")
  Else
     var(44) = X_Mpoyk_80 - var(55)
     var(45) = Y_Mpoyk_81 + var(55)
     RobSend ("!gowhen(,1pm>" + Str(var(44)) + ",2pm<" + Str(var(45)) + "):")
 End If
   Find_Axes
   x0 = X_Mpoyk_80: y0 = Y_Mpoyk_81: z0 = Z_Mpoyk_92
   RobSend_w ("!d " + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + "," + Str(Z_Mpoyk_92) + ":!go111")
End If


If var(43) = 3 Then   ' ;για δυο κλικ
  Find_Axes
  y0 = Y_diadromos62
  RobSend_w ("!d," + Str(Y_diadromos62) + ":!go,1:")
    var(44) = axe(1) 'pm
  If var(44) > X_Mpoyk_80 Then
     var(44) = var(79) - X_Mpoyk_80
     var(45) = Y_diadromos62 - Y_Mpoyk_81
    If var(44) > var(45) Then
       var(44) = X_Mpoyk_80 + var(55)
       RobSend ("!gowhen(,1pm<" + Str(var(79)) + ",1pm<" + Str(var(44)) + "):")
    Else
      var(45) = Y_Mpoyk_81 + var(55)
      RobSend ("!gowhen(,1pm<" + Str(var(79)) + ",2pm<" + Str(var(45)) + "):")
    End If
  Else
     var(44) = X_Mpoyk_80 - var(74)
     var(45) = Y_diadromos62 - Y_Mpoyk_81
    If var(44) > var(45) Then
       var(44) = X_Mpoyk_80 - var(55)
       RobSend ("!gowhen(,1pm>" + Str(var(74)) + ",1pm>" + Str(var(44)) + "):")
    Else
       var(45) = Y_Mpoyk_81 + var(55)
       RobSend ("!gowhen(,1pm>" + Str(var(74)) + ",2pm<" + Str(var(45)) + "):")
    End If
 End If

      Find_Axes
     x0 = X_Mpoyk_80: y0 = Y_Mpoyk_81: z0 = Z_Mpoyk_92
     RobSend_w ("!d " + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + "," + Str(Z_Mpoyk_92) + ":!go111:")
End If

 RobSend ("!comexc0:")
 Find_Axes
  z0 = Z_Mpoyk_92

  RobSend_w ("!d,," + Str(Z_Mpoyk_92) + ":!go,,1:")
  klise
  
  Find_Axes
  z0 = Ano_Z_93
  RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")

End Sub




Sub bukpot()
'     Mεταφορα μπουκαλιων στο χρηστη


var(30) = 1
var(3) = var(1)
talos4
RobSend ("!COMEXC0:")
Find_Axes
z0 = Ano_Z_93: w0 = W_dex_kef_zyg65
RobSend_w ("!d , ," + Str(Ano_Z_93) + "," + Str(W_dex_kef_zyg65) + ":!go0011:")
ferbuk
var(1) = var(2) ';Αλλαγη μεταβλητης
var(49) = X_Pot95

Find_Axes
x0 = var(49): y0 = Y_Pot96
RobSend_w ("!d" + Str(var(49)) + "," + Str(Y_Pot96) + ":!go11:")

Find_Axes
z0 = Z_Pot97
RobSend_w ("!d,," + Str(Z_Pot97) + ":!go,,1:")

anixe

Find_Axes
z0 = Ano_Z_93
RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")
var(44) = var(51) - var(1)
End Sub
Sub talos4()
'  Υπολογισμός στήλης και γραμμης που βρίσκεται η μπουκαλα
var(34) = 0
    X_Mpoyk_80 = X_Mpoykal_36 '                  ;Συντεταγμενη Χ
    Y_Mpoyk_81 = Y_Mpoykal_37 '                 ;Συντεταγμενη Y
 End Sub



Sub potbuk() ' Mεταφορα μπουκαλιων απο χρηστη στη θεση τους

RobSend ("!COMEXC0:")
var(30) = 1
Find_Axes
z0 = Ano_Z_93: w0 = W_dex_kef_zyg65
RobSend_w ("!d , ," + Str(Ano_Z_93) + "," + Str(W_dex_kef_zyg65) + "!:go0011:")
Ferpot '------
var(3) = var(1)
var(1) = var(2) ';Κρατα μεταβλ.ποτηριου
talos4 '------
panbuk '------

End Sub


Sub panbuk()
 '    πηγαινει απο κεφαλη στη θεση μπουκαλας φορτωμενος
rota


If var(43) = 2 Then
    Find_Axes
    y0 = Y_diadromos62
    RobSend_w ("!d," + Str(Y_diadromos62) + ":!go,1:")
End If

Find_Axes
x0 = X_Mpoyk_80: y0 = Y_Mpoyk_81
RobSend ("!d " + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + ":")


If ERGOSTASIO = "iris" Then
   If var(14) <> var(15) Then strofi
End If

If var(43) = 0 Or var(30) = 1 Then
    RobSend_w ("!go11:") ' ;Αν ειναι χειριστη για χειριστη ή πας μακρυα
End If

If var(43) = 1 And var(30) = 0 Then
   RobSend_w ("!go11:")
   RobSend ("!comexc0:")
End If

If var(43) = 2 And var(30) = 0 Then
  If X_Mpoyk_80 < var(79) Then RobSend_w ("!gowhen(,1pm <=" + Str(var(79)) + "):!go11:")
  If X_Mpoyk_80 > var(74) Then RobSend ("!gowhen(,1pm >=" + Str(var(74)) + "):!go11:")
End If
RobSend ("!comexc0:")

If ERGOSTASIO = "iris" Then
      Find_Axes
      x0 = X_Mpoyk_80: y0 = Y_Mpoyk_81
      RobSend_w ("!d " + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + ":!go11:")
End If

Find_Axes

z0 = Z_Mpoyk_92

RobSend ("!d,," + Str(Z_Mpoyk_92) + ":") '            ; συντ. μπουκάλας κάτω Z
RobSend_w ("!go,,1:") ' κατεβαινει

anixe

Find_Axes
z0 = Ano_Z_93

RobSend ("!d , , " + Str(Ano_Z_93) + ":")
RobSend_w ("!go,,1:") ' σηκωνεται  στο πάνω z  μπουκαλιων
End Sub






Sub Find_Axes()
   x0 = axe(1)
   y0 = axe(2)
   z0 = axe(3)
   w0 = axe(4)
End Sub

Sub Ferpot()

' ;Μεταφορα ποτηριων απο τη θεση τους
RobSend ("!comexc0:")
anixe
Find_Axes
z0 = Ano_Z_93

RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")

var(25) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19
 
If ERGOSTASIO = "iris" Then
 If var(25) > 0 Then ' ;Αν υπαρχει εστω και 1 μπουκ. στις κεφαλες
   Form1.Caption = "Υπαρχει μπουκαλα στις κεφαλες"
   Exit Sub
 End If
End If

   var(49) = X_Pot95 ' ΣΥΝΤΕΤΑΓΜΕΝΗ ΠΡΩΤΟΥ ΠΟΤΗΡΙΟΥ ΚΑΤΑ Χ
 If var(30) = 1 Then ' ΜΕΤΑΦΟΡΑ ΣΕ ΧΕΙΡΙΣΤΗ=1
    var(44) = W_dex_kef_zyg65 ' θέση 4ου : 1η κεφ σε ζυγ
  Else
    var(44) = var(61) 'θέση παρκαρίσματος 4 αξ
 End If
  
  'ΠΑΝΕ ΠΑΝΩ ΑΠΟ ΤΟ ΠΟΤΗΡΙ
   '95= Χ  ποτηριού     96=Υ ποτηριού   93=καθ.ανω Ζ
x0 = var(49): y0 = Y_Pot96: z0 = Ano_Z_93: w0 = var(44)
RobSend ("!d " + Str(var(49)) + "," + Str(Y_Pot96) + "," + Str(Ano_Z_93) + "," + Str(var(44)) + ":")
 
 If ERGOSTASIO = "iris" Then
        RobSend_w ("!go1111:")
        If var(14) = 1 Then strofi
 Else
        Find_Axes
        z0 = Ano_Z_93
        RobSend_w ("!go,,1:")
        
 End If
  
   RobSend ("!d,," + Str(Z_Pot97) + ":")
 
 If ERGOSTASIO = "iris" Then
   Find_Axes
   z0 = Z_Pot97: RobSend_w ("!go,,1:")
 Else
  
     var(44) = axe(1)  ' pm
     If var(44) > var(49) Then
        var(44) = var(44) - var(49)
     Else
        var(44) = var(49) - var(44)
     End If
     
     var(45) = axe(2) 'pm
     If var(45) > Y_Pot96 Then
        var(45) = var(45) - Y_Pot96
     Else
        var(45) = Y_Pot96 - var(45)
     End If
     
     If var(44) > var(45) Then
       ' αν η απόσταση που θα διανύσει ο Χ είναι μεγαλύτερη από το Υ
       ' τότε δεν θα κατεβάσεις το Ζ παρά μόνον όταν
       ' το Χ πλησιάσει σε απόσταση var(55)=4000 το ζητούμενο Χ
       
       
         If var(49) > axe(1) Then 'pm)
            var(44) = var(49) - var(55)
            RobSend ("!gowhen(,,1pm>" + Str(var(44)) + "):")
         End If
         If var(49) < axe(1) Then 'pm)
            var(44) = var(49) + var(55)
            RobSend ("!gowhen(,,1pm<" + Str(var(44)) + ":")
         End If
     Else
         If Y_Pot96 > axe(2) Then 'pm)
            var(45) = Y_Pot96 - var(55)
            RobSend ("!gowhen(,,2pm>" + Str(var(45)) + ":")
         End If
         If Y_Pot96 < 2 Then 'pm)
            var(45) = Y_Pot96 + var(55)
            RobSend ("!gowhen(,,2pm<" + Str(var(45)) + ":")
         End If
     End If
     
     x0 = var(49): y0 = Y_Pot96:  w0 = var(44)
     z0 = Z_Pot97
     RobSend_w ("!go1111 :")
     Find_Axes
     z0 = Z_Pot97
     RobSend_w ("!d,," + Str(Z_Pot97) + ":!go,,1:")
 End If
 
 klise
 
 Find_Axes
 z0 = Ano_Z_93
 RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")
 
 If var(30) = 1 Then '              ;Αν ειναι μπουκαλια από χειριστη για επιστροφη στη θεση τους
         Exit Sub                     ';Σταματα εδω
 End If
  Find_Axes
  x0 = X_Zyg_85: y0 = Y_Zyg_86
  RobSend_w ("!d " + Str(X_Zyg_85) + "," + Str(Y_Zyg_86) + ":!go11:")
  potzyg
 End Sub




Sub anixe()

If FindInput(Gripper_Input) = 0 Then  ' kleistos
     RobSend ("!" + Gripper_On + ":")
     Do While FindInput(Gripper_Input) = 0
       DoEvents
    Loop
    var(10) = 0
End If

End Sub
Sub potzyg() 'Τοποθετηση σε ζυγαρια

  RobSend ("comexc0")
                    ' Χ ζυγού            Υ  ζυγού             Ζ ανω                  W parking
    
        
    x0 = X_Zyg_85: y0 = Y_Zyg_86: z0 = Ano_Z_93: w0 = var(61)
    RobSend ("!d" + Str(X_Zyg_85) + "," + Str(Y_Zyg_86) + "," + Str(Ano_Z_93) + "," + Str(var(61)) + ":") ' ΧΥ  πανω Ζ ζυγαρ
    RobSend_w ("!go1111:") '              ;Ελα από πανω
    
    z0 = Z_Zyg_87
    RobSend ("!d,," + Str(Z_Zyg_87) + ":") ' κατω Ζ ζυγ
    RobSend_w ("!go,,1:") '               ;Κατεβα
  
  
  If var(10) = 0 Then     '  Αν εισαι ανοικτος
     klise
  Else
    anixe
  End If
    
' ανέβα πάνω Ζ
    z0 = Ano_Z_93
    RobSend ("!d , ," + Str(Ano_Z_93) + ":")
    RobSend_w ("go,,1:")
End Sub


Sub panpot() ' ;Επιστροφη ποτηριου

' αριθμός μπουκαλιών σε κεφαλές & μιξερ
var(25) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19

var(44) = var(61) ' θέση παρκαρίσματος
Find_Axes
w0 = var(44)
' παρκάρισε τον 4ο άξονα
RobSend_w ("!d,,," + Str(var(44)) + ":go,,,1:")

F_0epis_1seZyg35 = 0

anixe

potzyg 'πιάσε το ποτήρι και ανέβα στο πάνω Ζ

var(49) = X_Pot95  ' συντεταγμένη ποτηριού κατά Χ

Find_Axes
x0 = var(49): y0 = Y_Pot96
                    'Χ ποτηριού         Υ ποτηριού
RobSend ("!d " + Str(var(49)) + "," + Str(Y_Pot96) + ":")
RobSend_w ("!go11:")

Find_Axes
z0 = Z_Pot97
                    ' κάτω Ζ ποτηριού
  RobSend_w ("!d,," + Str(Z_Pot97) + ":go,,1:")
  anixe
  


Find_Axes
z0 = Ano_Z_93
        
        ' άνω Ζ
  RobSend_w ("!d,," + Str(Ano_Z_93) + ":go,,1:")

End Sub








Sub klise()

If FindInput(Gripper_Input) = 1 Then  ' anoiktos
dz = -200
     z0 = z0 + dz
    RobSend_w ("!ma0000:!d,," + Str(dz) + ":!go,,1:")
    RobSend ("!WAIT (mov = bxx0x):")
    RobSend ("!ma1111:")  '  ;Κατεβαινει 1 mm για να αποφυγει διαφορες Ζ
    RobSend ("!" + Gripper_Off + ":")
   Do While FindInput(Gripper_Input) = 1 '  kleistos
      DoEvents
   Loop
   var(10) = 1
  ' RobSend ("!t0 0.5:")
  MilSec 500
End If

End Sub



Sub akyro_fer_pot() ' ;Μεταφορα ποτηριων απο τη θεση τους

RobSend ("!comexc0:")
anixe

var(25) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19    ' ΑΡΙΘΜΟΣ ΜΠΟΥΚΑΛΙΩΝ ΣΕ ΚΑΦΑΛΕΣ & ΜΙΞΕΡ

var(49) = X_Pot95  ' ΣΥΝΤΕΤΑΓΜΕΝΗ ΠΡΩΤΟΥ ΠΟΤΗΡΙΟΥ ΚΑΤΑ Χ

 If var(30) = 1 Then  ' ΜΕΤΑΦΟΡΑ ΣΕ ΧΕΙΡΙΣΤΗ=1
      var(44) = W_dex_kef_zyg65 ' θέση 4ου : 1η κεφ σε ζυγ
 Else
      var(44) = var(61) 'θέση παρκαρίσματος 4 αξ
 End If
 
 
 'ΚΙΝΗΣΗ ΣΤΟ ΑΝΩ Ζ   ( Χ ποτ           Υ ποτ                 άνω Ζ
 RobSend ("!d " + Str(var(49)) + "," + Str(Y_Pot96) + "," + Str(Ano_Z_93) + "," + Str(var(44)) + ":go,,1:")
 
 ' κάτω Ζ
 RobSend ("!d , , " + Str(Z_Pot97) + ":")

    
    ' Κατ'ευθειαν σκυψιμο ισιας πορειας
   
    
    
    ' υπολογίζω την απόσταση var(44)  της αρπάγης από τον στόχο κατα Χ
    
    var(44) = axe(1) '1pm
    If var(44) > var(49) Then
           var(44) = var(44) - var(49)
    Else
           var(44) = var(49) - var(44)
    End If
   
   
    ' υπολογίζω την απόσταση var(45) της αρπάγης από τον στόχο κατα Υ
   var(45) = axe(2) ' 2pm
   If var(45) > Y_Pot96 Then
      var(45) = var(45) - Y_Pot96
  Else
      var(45) = Y_Pot96 - var(45)
  End If
  
 
 
 If var(44) > var(45) Then
       ' αν η απόσταση που θα διανύσει ο Χ είναι μεγαλύτερη από το Υ
       ' τότε δεν θα κατεβάσεις το Ζ παρά μόνον όταν
       ' το Χ πλησιάσει σε απόσταση var(55)=4000 το ζητούμενο Χ
       
       
       If var(49) > axe(1) Then    ' 1pm)
              var(44) = var(49) - var(55)
              RobSend ("!gowhen(,,1pm>" + Str(var(44)) + "):")
       End If
       If var(49) < axe(1) Then    'pm)
            var(44) = var(49) + var(55)
            RobSend ("!gowhen(,,1pm<" + Str(var(44)) + "):")
       End If
 Else
       
       ' αν η απόσταση που θα διανύσει ο Υ είναι μεγαλύτερη από το Χ
       ' τότε δεν θα κατεβάσεις το Ζ παρά μόνον όταν
       ' το Υ πλησιάσει σε απόσταση var(55)=4000 το ζητούμενο Υ
       
       If Y_Pot96 > axe(2) Then  'pm)
           var(45) = Y_Pot96 - var(55)
            RobSend ("gowhen(,,2pm>" + Str(var(45)) + "):")
        End If
        If Y_Pot96 < axe(2) Then  ' pm)
              var(45) = Y_Pot96 + var(55)
              RobSend ("!gowhen(,,2pm<" + Str(var(45)) + "):")
         End If
  End If

                                ' κατω Z ποτηριού
  RobSend ("!go1111 :  d,," + Str((Z_Pot97)) + " :  !go,,1:")
  
 klise
 RobSend ("!d,,(Ano_Z_93 :  !go,,1:")
 
 If var(30) = 1 Then  'Αν ειναι μπουκαλια από χειριστη για επιστροφη στη θεση τους
         Exit Sub                     'Σταματα εδω
 End If
 
                 ' Χ ζυγαριάς          Υ ζυγαριάς
 RobSend ("!d " + Str(X_Zyg_85) + "," + Str((Y_Zyg_86)) + ":!go11:")
 
 potzyg
 
 End Sub



Private Sub akyro_POTHRI_SE_ZYGARIA()

Dim Pothri

Pothri = 1

var(12) = 3001
var(94) = 18  'αριθμός ποτηριών

var(99) = 3000 + var(94) + 1
If (var(12) > 3000 And var(12) < var(99)) Then ';Μεταφορά ποτηριων στη ζυγαρια
        var(1) = var(12)
        If (var(11) = 1) Then
           Form1 = "Η ζυγαρια ειναι κατηλειμενη με αλλο ποτηρι"
           Exit Sub
        End If
        var(1) = var(1) - 3000
        F_0epis_1seZyg35 = 1 ' ΣΕ ΖΥΓΑΡΙΑ - ΚΕΦΑΛΕΣ          0=ΕΠΙΣΤΡΟΦΗ ΣΤΗΝ ΑΡΧΙΚΗ ΘΕΣΗ
        var(30) = 0  ' ΜΕΤΑΦΟΡΑ ΟΧΙ ΣΕ ΧΕΙΡΙΣΤΗ
        'WRVAR1
        Ferpot
        RobSend ("d,,," + Str(W_Ner_kef_zyg67) + ":go,,,1:")
        var(1) = 0:   var(2) = 0: var(11) = 1 ';Εναρξη δοσομετρησης
        RobSend ("OUT.14 - 1:")
        Form1 = "Αρχη Συνταγης"
        Exit Sub
End If




End Sub



Private Sub apo_zygo_se_pothri()

var(11) = 1
' επιστροφή ποτηριού ( ζυγαρια ->θέση ποτηριού)
'var(12) = 4001
var(94) = 18
                ' αριθμός ποτηριών
var(99) = 4000 + var(94) + 1

If var(12) > 4000 And var(12) < var(99) Then  ';Επαναφορά ποτηριων στη θεση τους
   
   var(1) = var(12)
    var(1) = var(1) - 4000
    
    F_0epis_1seZyg35 = 0 ' ΣΕ ΖΥΓΑΡΙΑ - ΚΕΦΑΛΕΣ          0=ΕΠΙΣΤΡΟΦΗ ΣΤΗΝ ΑΡΧΙΚΗ ΘΕΣΗ
    var(30) = 0  ' ΜΕΤΑΦΟΡΑ ΟΧΙ ΣΕ ΧΕΙΡΙΣΤΗ
    
    panpot
    
    var(1) = 0: var(2) = 0
    RobSend ("!out.14 - 1:")
        If (var(11) = 1) Then Form1.Caption = "Τελος Συνταγης"
        If (var(11) = 2) Then Form1 = "Τελος Παρασκευης Διαλυματος"
        var(11) = 0
End If



End Sub



Private Sub pane_pothri_se_zygo()
' ΠΑΝΕ ΠΟΤΗΡΙ ΣΕ ΖΥΓΑΡΙΑ   ( ΙΡΙΣ  )
     '  var(11) = 0 '     ΝΑ ΣΒΗΣΤΕΙ
        RobSend ("!v 8,12,10,1:")
        var(1) = var(12)
        If var(11) = 1 Then
          Form1.Caption = "Η ζυγαρια ειναι κατηλειμενη με αλλο ποτηρι"
          Exit Sub
        End If
        var(1) = var(1) - 3000
        F_0epis_1seZyg35 = 1 ' ΣΕ ΖΥΓΑΡΙΑ - ΚΕΦΑΛΕΣ          0=ΕΠΙΣΤΡΟΦΗ ΣΤΗΝ ΑΡΧΙΚΗ ΘΕΣΗ
        var(30) = 0  ' ΜΕΤΑΦΟΡΑ ΟΧΙ ΣΕ ΧΕΙΡΙΣΤΗ
        
        Ferpot
        
        RobSend_w ("!d,,,(var67):!go,,,1:")
        
        var(1) = 0:   var(2) = 0: var(11) = 1 ';Εναρξη δοσομετρησης
       
       RobSend ("!out 16 - 1:")
       Form1.Caption = "Αρχη Συνταγης"

'Exit Sub
End Sub
Sub akyro_ferpot_iris() '  ;Μεταφορα ποτηριων απο τη θεση τους

RobSend ("!comexc0:")
anixe

RobSend_w ("!d,,(var93):!go,,1:")

var(25) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19
 If var(25) > 0 Then ' ;Αν υπαρχει εστω και 1 μπουκ. στις κεφαλες
      Form1 = "Υπαρχει μπουκαλα στις κεφαλες"
    Exit Sub
 End If
 
 var(49) = X_Pot95
 
 If var(30) = 1 Then
       var(44) = W_dex_kef_zyg65
      Else
         var(44) = var(61)
 End If
  
  
  'ΠΑΝΕ ΠΑΝΩ ΑΠΟ ΤΟ ΠΟΤΗΡΙ
   '95= Χ  ποτηριού     96=Υ ποτηριού   93=καθ.ανω Ζ
   RobSend ("!d " + Str(var(49)) + "," + Str(Y_Pot96) + "," + Str(Ano_Z_93) + "," + Str(var(44)) + ":")
  
  
RobSend_w ("!go1111:")


If var(12) > 19000 And var(12) < 19020 Then RobSend ("!out.6-1:")
'  if var(14)=1 then strofi
  
 RobSend_w ("!d,," + Str(Z_Pot97) + " :!go,,1:")
 
 klise
 RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")
 
 If var(30) = 1 Then  'Αν ειναι μπουκαλια από χειριστη για επιστροφη στη θεση τους
         Exit Sub      'Σταματα εδω
 End If
  
  
RobSend_w ("!d " + Str(X_Zyg_85) + "," + Str(Y_Zyg_86) + ":!go11:")


potzyg
 
  End Sub


Private Sub Command1_Click()
' πάνε μπουκάλι  σε κεφαλή

var(11) = 1 ' αρχισε η συνταγή
var(12) = 1000 + Val(Text2.Text)
var(13) = 0

If Val(Text2.Text) < 1 Or Val(Text1.Text) > Ar_Mpoykal Then
    MsgBox "λαθος μπουκάλι"
    Exit Sub
End If

cmd$ = "!var12=" + Str(var(12)) + ":"
find_uesh (cmd$)

talos  '----------

End Sub


Private Sub Command2_Click()

' επιστροφη ποτηριού από ζυγαριά

var(12) = 4000 + Val(Text1.Text)
var(13) = 0

If Val(Text1.Text) < 1 Or Val(Text1.Text) > Ar_Pothr Then
    MsgBox "λαθος ποτήρι"
    Exit Sub
End If

cmd$ = "!var12=" + Str(var(12)) + ":"
find_uesh (cmd$)

talos  '----------

End Sub

Private Sub Command3_Click()

' πάνε ποτήρι σε ζυγαριά

var(12) = 3000 + Val(Text1.Text)
var(13) = 0

If Val(Text1.Text) < 1 Or Val(Text1.Text) > Ar_Pothr Then
    MsgBox "λαθος ποτήρι"
    Exit Sub
End If

cmd$ = "!var12=" + Str(var(12)) + ":"
find_uesh (cmd$)

talos  '----------

End Sub

Private Sub Command4_Click()
'μεταφορά μπουκαλιού var(12) σε θέση ποτηριού var(13)
' var(12):μπουκάλι   var(13):ποτήρι
    RobSend ("!v 8,12,10,1:")
var(12) = 6000 + Val(Text2.Text)
var(13) = Val(Text1.Text)
var(94) = 18
find_uesh ("!var12=" + Str(var(12)) + ":!var13=" + Str(var(13)) + ":")

talos
    
    
    
End Sub

Private Sub talos()
' απο χειριστή σε αρχ.θέση μπουκαλιού
 
var(31) = (var(74) - var(78)) / (W_ar_kef_zyg66 - W_dex_kef_zyg65) ' ;αCoef14


var(32) = var(78) - (W_dex_kef_zyg65 * var(31))    ';βCoef14

var(33) = W_dex_kef_zyg65
var(23) = W_ar_kef_zyg66

var(56) = var(33) + W_dex_kef_zyg65 - var(23)

var(56) = var(56) * var(31)
var(56) = var(56) + var(32)

var(79) = var(56) ' ;Συντεταγμενη αριστερης κεφαλης με την δεξιά στην ζυγαριά

var(94) = 18
 
' αποθηκεύω τις συντεταγμένες του ποτηριού σε άλλες μεταβλητές
   X_Pot95 = var(24)
   Y_Pot96 = var(77)
   Z_Pot97 = var(89)
    


Y_diadromos62 = var(75) + var(58) + var(58) ';Διαδρομος μπουκαλιων
  
 
 
'μεταφορά μπουκαλιού var(12) σε θέση ποτηριού var(13)
' var(12):μπουκάλι   var(13):ποτήρι
    RobSend ("!v 8,12,10,1:")


Ano_Z_93 = var(93) ' ανω Ζ universal


X_Zyg_85 = var(85) '    Συντεταγμενη ζυγαριας κατα Χ
Y_Zyg_86 = var(86) '     Συντεταγμενη ζυγαριας κατα Υ
Z_Zyg_87 = var(87) '     Συντεταγμενη ζυγαριας κατω Ζ
Z_Mpoyk_92 = var(92) ' συντ.κάτω Ζ  μπουκάλας

X_Anad_70 = var(70) ' Συντεταγμενη θεσης αναδευσης κατα Χ
Y_Anad_71 = var(71) ' Συντεταγμενη θεσης  αναδευσης κατα Υ
Z_Anad_72 = var(72) ' Συντεταγμενη θεσης  αναδευσης κατω Ζ
  


' TA ΕΠΟΜΕΝΑ 4  στελνονται από το FIND_COORDIN
X_Mpoykal_36 = var(36) ' ;/////VAR36     Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
Y_Mpoykal_37 = var(37) '/////VAR37     Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
X_2Mpoykal_38 = var(38) '   Αποστολγ συντεταγεμβνης Χ μπουκαλιιν
Y_2Mpoykal_39 = var(39) '    Αποστολγ συντεταγεμβνης Υ μπουκαλιιν
Y_diadromos62 = var(62)

W_Park_61 = var(661)  '     Num 5,Head W:θεση παρκαρισματος Α₯ονα 4



'/////VAR65     Num 1,Head W:Συντεταγμενη 1ης κεφαλης α₯ονα 4 κεντραρισμενου στη ζυγαρια
'/////VAR66     Num 2,Head W:Συντεταγμενη 2ης κεφαλης α₯ονα 4 κεντραρισμενου στη ζυγαρια
'/////VAR67     Num 3,Head W:Συντεταγμενη 3ης κεφαλης νεροη α₯ονα 4 κεντραρισμενης στη ζυγαρια
'/////VAR68     Num 4,Head W:Συντεταγμενη  κεφαλης σκονης α₯ονα 4 κεντραρισμενου στη ζυγαρια
W_dex_kef_zyg65 = var(65) '
W_ar_kef_zyg66 = var(66) '
W_Ner_kef_zyg67 = var(67) '
W_Skon_kef_zyg68 = var(68)
  
  
'X_Mpoyk_80  ';VAR80     Συντεταγμενη κατα Χ της προς ευρεση μπουκαλας
'Y_Mpoyk_81  ' ;VAR81     Συντεταγμενη κατα Υ της προς ευρεση μπουκαλας



' Απο Vb ερχονται τα var24->Χ ποτηριων var77 -> Yποτηριων var89-> Z ποτηριων
If var(12) = 100 Then '  ;Τοποθέτηση ΜΠΟΥΚΑΛΙΟΥ σε θεση αναδευσης
        var(1) = var(12)
      comexc0
        If var(11) = 0 Then
         Form1.Caption = "Δεν υπαρχει ποτηρι στη ζυγαρια"
         Exit Sub
         RobSend ("!out.16 - 1:")
        End If
         F_0epis_1seZyg35 = 0
             anixe
             Find_Axes
             z0 = Ano_Z_93
             RobSend_w ("!d" + Str(X_Zyg_85) + "," + Str(Y_Zyg_86) + "," + Str(Ano_Z_93) + "," + Str(var(61)) + ":!go,,1:")
             x0 = X_Zyg_85: y0 = Y_Zyg_86: w0 = var(61)
             RobSend_w ("!go1101:")
             z0 = Z_Zyg_87
             RobSend_w ("!d,," + Str(Z_Zyg_87) + ":!go,,1:")
             klise
             Find_Axes
             z0 = Ano_Z_93
             RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")
             z0 = Ano_Z_93
             RobSend_w ("!d" + Str(X_Anad_70) + "," + Str(Y_Anad_71) + "," + Str(Ano_Z_93) + ":!go,,1:")
             x0 = X_Anad_70: y0 = Y_Anad_71:
             RobSend_w ("!go11:")
             z0 = Z_Anad_72
             RobSend_w ("!d,," + Str(Z_Anad_72) + ":!go,,1:")
             anixe
             z0 = Ano_Z_93
             RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")
        
             var(1) = 0
             var(2) = 0
             var(30) = 0
             var(11) = 3
      
    RobSend ("!out.16 - 1:")
     
    Exit Sub
nif

If var(12) = 101 Then '  ;Επιστροφη ποτηριου απο αναδευση σε ζυγαρια
        var(1) = var(12)
     comexc0
        var35 = 1
        anixe
        RobSend ("!d" + Str(X_Anad_70) + "," + Str(Y_Anad_71) + "," + Str(Ano_Z_93) + ":")
        Find_Axes
        z0 = Ano_Z_93
        RobSend_w ("!go,,1:")
        Find_Axes
        x0 = X_Anad_70: y0 = Y_Anad_71
        RobSend_w ("!go11:")
        RobSend ("!d,," + Str(Z_Anad_72) + ":")
        Find_Axes
        z0 = Z_Anad_72
        RobSend_w ("!go,,1:")
        klise
        RobSend ("!d,," + Str(Ano_Z_93) + ":")
        Find_Axes
        z0 = Ano_Z_93
        RobSend ("!go,,1:")
        
        RobSend ("!d" + Str(X_Zyg_85) + "," + Str(Y_Zyg_86) + "," + Str(Ano_Z_93) + "," + Str(var(61)) + ":")
        x0 = X_Zyg_85: y0 = Y_Zyg_86: z0 = Ano_Z_93: w0 = var(61)
        RobSend_w ("!go1111:")
        var11 = 2
        topzyg
        RobSend ("!d,,," + Str(W_Ner_kef_zyg67) + ":")
        RobSend ("!go,,,1:")
        var(1) = 0: var(2) = 0: var(30) = 0
        RobSend ("!out.16 - 1:")
     
        Exit Sub
nif

If var(12) > 200 And var(12) < 221 Then ';Επιστροφη ποτηριου απο αναδευση στη θεση του
        var(1) = var(12) - 200
     comexc0
        var(11) = 2: F_0epis_1seZyg35 = 0: var(30) = 0
        anixe
        
        RobSend ("!d " + Str(X_Anad_70) + "," + Str(Y_Anad_71) + "," + Str(Ano_Z_93) + ":")
        Find_Axes
        z0 = Ano_Z_93
        RobSend_w ("! go,,1:")
        Find_Axes
        x0 = X_Anad_70: y0 = Y_Anad_71
        RobSend_w ("!go11:")
        Find_Axes
        z0 = Z_Anad_72
        RobSend ("!d,," + Str(Z_Anad_72) + ":")
        
        RobSend_w ("!go,,1:")
        klise
        RobSend ("!d,," + Str(Ano_Z_93) + ":")
        Find_Axes
        z0 = Ano_Z_93
        RobSend_w ("!go,,1:")
       
        mixpot

        var(1) = 0: var(2) = 0
        Form1.Caption = "Τέλος Παρασκευής Διαλύματος"
        RobSend ("!out.16 - 1:")
     
        Exit Sub
End If





' ;Επαναφορά μπουκαλιων στη θέση τους  από χειριστη
If var(12) > 12000 And var(12) < 18000 Then
       apo_pot_se_buk
End If

' ποτήρι  από χειριστη σε ζυγό
If var(12) > 3000 And var(12) < 4000 Then
   pane_pothri_se_zygo
End If
' ;Επαναφορά ποτηριού απο ζυγό στην θέση του
If var(12) > 4000 And var(12) < 5000 Then
   apo_zygo_se_pothri
End If


var(99) = 1000 + Ar_Mpoykal + 1
If var(12) > 1000 And var(12) < 1300 Then ' ;Τοποθέτηση μπουκάλας σε κεφαλή δοσομέτρησης
    var(1) = var(12)
    var(1) = var(1) - 1000
    var(44) = ArMp_ArKef21 + ArMp_DexKef22
    bukkef
    Form1.Caption = "Μπουκαλα σε κεφαλες"
    var(1) = 0: var(2) = 0: var(30) = 0
    RobSend ("!OUT.14 - 1:")
End If


If var(12) > 19000 And var(12) < 19020 Then ' Μεταφορά ποτηριων στη ζυγαρια για σκονη
        var(1) = var(12)
        RobSend ("!out.6 - 1:")
        var(1) = var(1) - 19000
        F_0epis_1seZyg35 = 1
        var(30) = 0
        var(99) = var(75)
        var(75) = var(73) ' ;Βαλε συντ.Κεφαλης σκονης κατα Υ
        Ferpot
        RobSend ("!out.6 - 0:")
        MilSec 500   't1
        Find_Axes
        w0 = W_Skon_kef_zyg68
        RobSend_w ("!d,,," + Str(W_Skon_kef_zyg68) + ":!go,,,1:")
        
        var(1) = 0
        var(6) = 0
        var(2) = 0
        var(11) = 2
        Form1.Caption = "Εναρξη παρασκευης διαλυματος"
        RobSend ("!out.16 - 1:")
        Exit Sub
End If

If var(12) > 19999 And var(12) < 100000 Then ';Ντρουβανισμα ,Z σε mm πχ 20030 =30mm=30*200=6000  var1=var1-20000 :var11=2
           'Αν δεν εισαι ετοιμος druvanisma ready
           If var(6) = 0 Then
               druvrd ' ανέβα στο ύψος του βολάν και μετά FLAG ( VAR(6)=1 )
               Form1.Caption = "Εναρξη Παρασκευής Διαλύματος..."
           Else
               var(1) = var(12) - 20000
               druvan
           End If
           RobSend ("!out .16 - 1:")
           Exit Sub
End If



' ;Μεταφορά μπουκαλιων στο χειριστή
If var(12) > 6000 And var(12) < 12000 Then
        var(30) = 1 ' Απο μπουκαλα σε ποτηρι για  χειριστη
        var(1) = var(12) - 6000
        var(2) = var(13) ' αριθμός ποτηριού
        If var(13) = 0 Or var(13) > var(94) Then
          Form1.Caption = "Λαθος ποτηρι"
          Exit Sub
        End If
        
        var(25) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19   'αριθμος μπουκαλιών σε κεφαλές & μιξερ
        If var(25) > 0 Then
         Form1.Caption = "Υπαρχει μπουκαλα στις κεφαλες"
         Exit Sub
        End If
        RobSend ("!OUT.14 - 0:")
        bukpot
        var(1) = 0
        var(2) = 0
        RobSend ("!OUT.14 - 1:")
        var(12) = 0
        var(13) = 0
        Form1.Caption = "Μεταφορα μπουκαλιων απο τη θέση τους στο χειριστη"
 End If


If var(12) > 2000 And var(12) < 2280 Then
'   ;Επαναφορά μπουκάλας στη βαση της
        var(1) = var(12)
    var(1) = var(1) - 2000
    kefbuk
    var(1) = 0: var(2) = 0: var(30) = 0
    RobSend ("!out .16 - 1:")
           If ArMp_ArKef21 < 2 And ArMp_DexKef22 = 0 And var(11) = 1 Then '  ;Αν συνεχιζει η δοσομετρηση και υπαρχει το πολυ μια μπουκαλα
            var(4) = 1
           End If
     Form1.Caption = "Επιστροφη Μπουκαλας"
     Exit Sub
End If



If var(12) > 100000 And var(12) < 100280 Then 'Πανε μπουκ και φερε μπουκ ταυτοχρονα
        var(1) = var(12) - 100000
        var(2) = var(13)
        var(30) = 0
      If var(13) = 0 Or var(13) > 280 Then
          Form1.Caption = "Λαθος φερε μπουκαλι"
          Exit Sub
      End If
      RobSend ("!out.16 - 0:")
      kefbuk
      Form1.Caption = "Επιστροφη Μπουκαλας"
        
        var(1) = var(2)
        X_Mpoykal_36 = X_2Mpoykal_38
        Y_Mpoykal_37 = Y_2Mpoykal_39      ' Κύκλος των var για Β' μπουκάλι
        var(44) = ArMp_ArKef21 + ArMp_DexKef22
        bukkef
        var(1) = 0: var(2) = 0
        RobSend ("!out.16 - 1:")
        var(12) = 0
        var(13) = 0
        Form1.Caption = "Μπουκάλα σε κεφαλες"
        Exit Sub
 End

If var(12) > 100000 And var(12) < 100280 Then '  ;Πανε μπουκ και φερε μπουκ ταυτοχρονα
        var(1) = var(12) - 100000: var(2) = var(13): var(30) = 0
      If var(13) = 0 Or var(13) > 280 Then
        Form1.Caption "Λαθος φερε μπουκαλι"
         Exit Sub
      End If
     
      RobSend ("!out16 - 0:")

    kefbuk
    Form1.Caption = "Επιστροφη Μπουκαλας"
        var(1) = var(2)
        X_Mpoykal_36 = X_2Mpoykal_38
        Y_Mpoykal_37 = Y_2Mpoykal_39 'Κύκλος των var για Β' μπουκάλι
        ' WRVAR1
        var(44) = ArMp_ArKef21 + ArMp_DexKef22
        bukkef
    var(1) = 0: var(2) = 0
    RobSend ("!out.16 - 1:")
    var(12) = 0
    var(13) = 0
       ' wrvar21
        'wrvar22
    Form1.Caption = "Μπουκάλα σε κεφαλες"
    Exit Sub
 End If







End Sub


Sub chakef()
' ;Αλλαζει τις παραμετρους των κεφαλων κατα την μετακινηση  του 4

var(44) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19
  If var(44) = 1 And F_0epis_1seZyg35 = 0 Then '  ;Αν θα παρεις την τελευταια μπουκαλα
      RobSend ("!comexc0:")
      Find_Axes
      w0 = W_Ner_kef_zyg67
      RobSend_w ("!d,,," + Str(W_Ner_kef_zyg67) + ":!go,,,1:") '   ;Πανε στα χημικα
      Exit Sub
  End If
 If var(44) = 0 And F_0epis_1seZyg35 = 1 Then ' ;Αν δεν υπαρχ.μπουκ.και ειναι να φερει
        If var(11) = 2 Then '    ;Αν ειναι για σκονη
              Find_Axes
              w0 = W_Skon_kef_zyg68
              RobSend ("!d , , , " + Str(W_Skon_kef_zyg68) + ":")
        Else
           var(44) = axe(4)
           If var(44) <> W_dex_kef_zyg65 Then
                Find_Axes
                w0 = W_dex_kef_zyg65
                RobSend ("!d,,," + Str(W_dex_kef_zyg65) + ":!comexc1:!go,,,1:")
           End If
           Exit Sub
        End If
   End If
 If var(44) = 2 And F_0epis_1seZyg35 = 0 Then
        var(44) = axe(4) 'pm
        
        Find_Axes
        If var(44) = W_dex_kef_zyg65 Then '           ;Αν τελευταια τοποθετησες αριστερα
              w0 = W_ar_kef_zyg66
              RobSend ("!d,,," + Str(W_ar_kef_zyg66) + ":") '   ;Τραβα δεξια
         Else
              w0 = W_dex_kef_zyg65
              RobSend ("!d,,," + Str(W_dex_kef_zyg65) + ":") '     ;Τραβα αριστερα
         End If
         RobSend_w ("!go,,,1:")     ' ;κανε κλικ
 End If
End Sub


Sub vrskef()

'def vrskef  ;ArMp_ArKef21=Μπουκαλα στην αριστερη κεφαλη,ArMp_DexKef22= μπουκαλα στην δεξια κεφαλη

var(16) = ArMp_ArKef21: var(17) = ArMp_DexKef22: var(33) = axe(4) 'Για να παρει = 0 , για να φερει =1
  
  
  If var(11) > 1 And F_0epis_1seZyg35 = 1 Then
        If ArMp_Mixer19 = 1 Then
            Form1.Caption = "Η κεφαλη σκονης ειναι κατηλειμενη"
            RobSend ("!ps:")
            Exit Sub
         End If
        
        var(23) = W_Skon_kef_zyg68: ArMp_Mixer19 = 1
        var(56) = var(33) + W_dex_kef_zyg65 - var(23)
       var(56) = var(56) * var(31): var(56) = var(56) + var(32)
       Exit Sub
  End If
 
 If var(11) > 1 And F_0epis_1seZyg35 = 0 Then
        If ArMp_Mixer19 = 0 Then
            Form1.Caption = "Η κεφαλη σκονης ειναι αδεια"
            RobSend ("!ps:")
            Exit Sub
         End If
         var(23) = W_Skon_kef_zyg68: ArMp_Mixer19 = 0
        var(56) = var(33) + W_dex_kef_zyg65 - var(23)
       var(56) = var(56) * var(31): var(56) = var(56) + var(32)
       Exit Sub
  End If
          ' η κεφαλή στην θέση που ρίχνει το νερό χωρίς μπουκάλια επάνω
          If var(16) = 0 And var(17) = 0 Then
             If F_0epis_1seZyg35 = 1 Then ' ;Αν ειναι να φερει
                  var(23) = W_dex_kef_zyg65
                  var(16) = 1
                  var(56) = var(33)  'axe(4)
                  var(56) = var(56) * var(31)  '
                  var(56) = var(56) + var(32)
                  ' var(56) = ( axe(4)-var(65))*var(31)  +  var(78)
                  Exit Sub
                  ' όταν ρίχνει νερό πού είναι η δεξια κεφ. κατά Χ
                  ' δηλ. var(56)= axe(4) * var(31)+ W_dex_kef_zyg65
             Else
                  RobSend ("!ps:")
             End If
          End If
          
          If var(16) = 0 And var(17) = 1 Then '   ) ; 2
            If F_0epis_1seZyg35 = 1 Then '   ;Αν ειναι να φερει
                var(23) = W_dex_kef_zyg65
                var(16) = 1
                var(56) = var(33)
                var(56) = var(56) * var(31)
                var(56) = var(56) + var(32)
                'var(56)=(axe(4)-var(65))*var(31) + var(78)
                Exit Sub
            Else
                var(23) = W_ar_kef_zyg66
                var(17) = 0
                var(56) = var(33) + W_dex_kef_zyg65 - var(23)
                var(56) = var(56) * var(31)
                var(56) = var(56) + var(32)
                'var(56)=(axe(4)-var(66))*var(31) +var(78)
                Exit Sub
            End If
        End If
          If var(16) = 1 And var(17) = 0 Then ' εχει μπουκάλι μόνο η αριστερή κεφαλή
            If F_0epis_1seZyg35 = 1 Then ')  ;Αν ειναι να φερει
               If var(11) = 2 Then
                  Form1.Caption = "Η κεφαλη σκονης ειναι κατηλειμενη"
                  RobSend ("!ps:")
                  Exit Sub
                End If
                var(23) = W_ar_kef_zyg66: var(17) = 1
                var(56) = var(33) + W_dex_kef_zyg65 - var(23)
                var(56) = var(56) * var(31)
                var(56) = var(56) + var(32)
                Exit Sub
            Else
                  If var(11) > 1 Then
                    var(23) = W_Skon_kef_zyg68
                    var(16) = 0
                    var(56) = var(33) + W_dex_kef_zyg65 - var(23)
                    var(56) = var(56) * var(31)
                    var(56) = var(56) + var(32)
                    Exit Sub
                  End If
                  var(23) = W_dex_kef_zyg65
                  var(16) = 0
                  var(56) = var(33) + W_dex_kef_zyg65 - var(23)
                  var(56) = var(56) * var(31)
                  var(56) = var(56) + var(32)
                  Exit Sub
            End If
          End If
          If var(16) = 1 And var(17) = 1 Then ' ) ; 4
             If F_0epis_1seZyg35 = 1 Then '  Αν ειναι να φερει
                RobSend ("!ps:")
            Else
                If var(33) = W_ar_kef_zyg66 Then
                    var(23) = W_dex_kef_zyg65
                    var(16) = 0
                 End If
                 If var(33) = W_dex_kef_zyg65 Then
                     var(23) = W_ar_kef_zyg66
                     var(17) = 0
                 End If
                 var(56) = var(33) + W_dex_kef_zyg65 - var(23)
                 var(56) = var(56) * var(31)
                 var(56) = var(56) + var(32)
                 Exit Sub
            End If
          End If
  
  End Sub



Sub bukkef() '  ;Μεταφορα Μπουκαλιων στις κεφαλες

F_0epis_1seZyg35 = 1

RobSend ("!COMEXC0:")
If (var(64) = 1) Then
  If (var(11) <> 2) Then
     chakef
  End If '; καταργήθηκε με την κίνηση ποτηριών με μπουκάλια στις κεφαλές
End If

talos4

ferbuk

vrskef

rota

If var(43) = 0 Then
   Find_Axes
   x0 = var(56): y0 = var(75)
   RobSend_w ("!d" + Str(var(56)) + "," + Str(var(75)) + ":!go11:")
End If

If var(43) = 1 Then RobSend_w ("!d" + Str(var(56)) + "," + Str(Y_diadromos62) + ":!go11:")
If var(43) = 2 Then
    RobSend ("!gowhen(2pm>" + Str(var(75)) + "):")
    Find_Axes
    x0 = var(56): y0 = Y_diadromos62
   RobSend_w ("!d" + Str(var(56)) + "," + Str(Y_diadromos62) + ":!go11:")
    RobSend_w ("!go1:")
End If

topkef

End Sub

Sub topkef()  ' ;Τοποθετηση σε κεφαλη

     RobSend ("!COMEXC0:")
   Find_Axes
   z0 = Ano_Z_93
     RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")
   var40 = axe(1) 'pm:
     If (var40) > var(56) Then
           RobSend ("!gowhen(,1pm<=" + Str(var(56)) + "):")
      Else
           RobSend ("!gowhen(,1pm>=" + Str(var(56)) + "):")
     End If
     Find_Axes
     x0 = var(56): y0 = var(75)
      RobSend_w ("!d " + Str(var(56)) + "," + Str(var(75)) + ":!go11:")
     
     
     If F_0epis_1seZyg35 = 1 Then '      ;Αν ειναι να φερεις
         Find_Axes
         If var(11) = 2 Then
             z0 = var(69)
             RobSend ("!d,," + Str(var(69)))
         Else
             z0 = var(76)
             RobSend ("!d , ," + Str(var(76)) + ":")
         End If
         RobSend_w ("!go,,1:")
         anixe
         ArMp_ArKef21 = var(16): ArMp_DexKef22 = var(17)
         If var(11) = 2 Then
            RobSend_w ("!d,," + Str(var(88)) + ":!go,,1:")
            klise
             MilSec 1000   ' t1
              var6 = 1    '           ';Πιασε το χερουλι
             Form1.Caption = "Εναρξη δοσομέτρησης σκόνης."
         Else
            Find_Axes
            z0 = Ano_Z_93
            RobSend_w ("!d" + Str(var(78)) + "," + Str(Y_diadromos62) + "," + Str(Ano_Z_93) + ":!go,,1:")
            
             y0 = Y_diadromos62
            RobSend_w ("!go,1:")
         End If
    Else
         Find_Axes
         
         If var(11) = 2 Then
             z0 = var(69)
             RobSend ("!d , ," + Str(var(69)) + ":")
         Else
             z0 = var(76)
             RobSend ("!d , , " + Str(var(76)) + ":")
         End If
          RobSend_w ("!go,,1 :")
          klise
          z0 = Ano_Z_93
          RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")
          ArMp_ArKef21 = var(16)
          ArMp_DexKef22 = var(17)
    End If
End Sub


Sub apo_pot_se_buk()
       
        var(1) = var(12) - 12000
        var(2) = var(13)
        var(30) = 1
        If var(13) = 0 Or var(13) > var(94) Then
           Form1.Caption = "Λαθος ποτηρι"
           Exit Sub
        End If
        var(25) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19
        
        If var(25) > 0 Then
               Form1.Caption = "Υπαρχει μπουκαλα στις κεφαλες"
               Exit Sub
        End If
     
     RobSend ("!OUT.14 - 0:")
        
        var(3) = var(1)
        var(1) = var(2)
        var(2) = var(3) '  ;To μπουκαλι ποτηρι
        
        potbuk   ' -----------
        
        var(1) = 0
        var(2) = 0
        
        RobSend ("!OUT.14 - 1:")
        var(12) = 0
        var(13) = 0
        
        Form1.Caption = "Επαναφορά μπουκαλιων στη θέση τους  από χειριστη"
        
End Sub

Private Sub Command5_Click()

' επιστροφή μπουκαλιού από χειριστή στην θέση του

var(12) = 12000 + Val(Text2.Text)
var(13) = Val(Text1.Text)

cmd$ = "!var12=" + Str(var(12)) + ":!var13=" + Str(var(13)) + ":"

find_uesh (cmd$)

talos  '----------


End Sub

Private Sub Command6_Click()
' φέρε  μπουκάλι  σε ζυγαριά

var(11) = 1 ' αρχισε η συνταγή
var(12) = 2000 + Val(Text2.Text)
var(13) = 0

If Val(Text2.Text) < 1 Or Val(Text1.Text) > Ar_Mpoykal Then
    MsgBox "λαθος μπουκάλι"
    Exit Sub
End If

cmd$ = "!var12=" + Str(var(12)) + ":"
find_uesh (cmd$)

talos  '----------
End Sub

Private Sub Command7_Click()
' πάνε μπουκάλι  στην θέση του και φέρε ένα άλλο

'Erxetai = a_Fere(Move)
'Feygei = a_Pane(Move)
'frmSYNT.Caption = a_Tkinhsh(Move) + Str(a_Pane(Move)) + " ΦΕΡΕ ΤΟ " + Str(a_Fere(Move))
'kod6400 = 100000 + a_Pane(Move)
'cmd$ = "!var12=" + Str(kod6400) +;
'":!var13=" + Str(a_Fere(Move)) + ":"


'dum = wait_to_understand(cmd$)

'If Typos_Kinhshs = 1 Then
 '            find_Coordin = "!var36=" + LTrim(Str(SYN_X)) + ":!var37=" + LTrim(Str(SYN_Y)) + ":"
  '            X_Mpoykal_36 = SYN_X
   '           Y_Mpoykal_37 = SYN_Y
    '         Exit Function
     '   Else
      '       find_Coordin = "!var38=" + LTrim(Str(SYN_X)) + ":!var39=" + LTrim(Str(SYN_Y)) + ":"
       '      X_2Mpoykal_38 = SYN_X
        '     Y_2Mpoykal_39 = SYN_Y
         '
          '   Exit Function




var(11) = 1 ' αρχισε η συνταγή


var(12) = 100000 + Val(Text2.Text)  'μπουκάλι που φεύγει από κεφαλή
' X_Mpoykal_36=X   Y_Mpoykal_37=Y

var(13) = Val(Text3.Text) ' μπουκάλι που έρχεται σε κεφαλή
' X_2Mpoykal_38=X   Y_2Mpoykal_39=Y


If Val(Text2.Text) < 1 Or Val(Text1.Text) > Ar_Mpoykal Then
    MsgBox "λαθος μπουκάλι"
    Exit Sub
End If

cmd$ = "!var12=" + Str(var(12)) + ":"
find_uesh (cmd$)

talos  '----------
End Sub

Private Sub Form_Load()

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
End Sub

Private Sub Timer1_Timer()

     Label1.Caption = IIf(FindInput(Gripper_Input) = 1, "ανοικτος", "κλειστός")
     Label2.Caption = Str(Bit_Axona(1, 1)) + Str(Bit_Axona(2, 1)) + Str(Bit_Axona(3, 1))
End Sub



Sub druvrd() '   ;druvanisma ready

Find_Axes
w0 = W_Skon_kef_zyg68
RobSend_w ("!d,,," + Str(W_Skon_kef_zyg68) + ":!go,,,1:")

If var(10) = 1 Then anixe
Find_Axes
z0 = var(88) ' υψος βολάν
RobSend_w ("!d,," + Str(var(88)) + ":!go,,1:")
klise
MilSec 300 ' t1:
var(6) = 1   ';Πιασε το χερουλι
End Sub


Sub druvan()
'
If var(8) = 0 Then
   var(8) = velocity(3) '3v  'ΑΠΟΘΗΚΕΥΩ την ταχύτητα του 3ου άξονα
   RobSend ("!v , , 0.1:")
End If

If var(1) = 0 Then
   druvfn
   Exit Sub
End If

RobSend ("!ma0000:")

' ανέβασμα ντρουβανίσματος
Find_Axes
dz = var(1)
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

' επαναφέρω την αρχική ταχύτητα
RobSend ("!v , ," + Str(var(8)) + ":")
anixe ' αφήνει το ποτήρι της σκόνης
var(8) = 0
F_0epis_1seZyg35 = 0
RobSend ("!ma1111:")

Find_Axes
z0 = var(69) ' κατέβα να πιάσεις το μπουκάλι της σκόνης
RobSend_w ("!d,," + Str(var(69)) + ":!go,,1:")

klise

Find_Axes
z0 = Ano_Z_93
RobSend_w ("!d,," + Str(Ano_Z_93) + ":!go,,1:")

var(6) = 0
RobSend_w ("!d " + Str(X_Mpoyk_80) + "," + Str(Y_Mpoyk_81) + "," + Str(Ano_Z_93) + "," + Str(W_Ner_kef_zyg67) + ":!go1111:")
ArMp_ArKef21 = 0
panbuk
ArMp_Mixer19 = 0
var(75) = var(99)
var(11) = 1 ';Ξαναβαλε συντ.Κεφαλων κατα Υ την παλια
End Sub
Sub rota()

var(40) = axe(1) 'pm ;Τρεχουσες συντεταγμενες.Ισχυει για την περιπτ.που ο gripper  ειναι δεξια του βραχιονα
var(41) = axe(2)
var(41) = var(41) - var(58) ';Περιθωριο Υ ενα βημα

var(44) = ArMp_ArKef21 + ArMp_DexKef22 + ArMp_Mixer19 ' ;var(21 μπουκάλα σε αρ.κεφαλή var(22 σε δεξ.κεφ  var(19 σε mixer

';X_Mpoyk_80συντ.Χ της προς εύρεση μπουκάλας
';var(81 συντ.Υ της προς εύρεση μπουκάλας


If F_0epis_1seZyg35 = 1 And var(10) = 0 Then
' ; Περ 1,αν ειναι να παρει μπουκαλα απο την θεση της για τις κεφαλες-ζυγαρια
' ;var(10)=0 gripper ανοιχτός  var(35=1 προς ζυγαριά ή κεφαλές
  
  If (var(44) = 0) Then 'δεν έχω μπουκάλια στις κεφαλές
     var(43) = 0  'ισια πορεία
     Exit Sub
  End If ';Αν οι κεφαλες ειναι αδειες πανε ισια

  
  If (var(41) > Y_diadromos62 And Y_Mpoyk_81 > Y_diadromos62) Then
     var(43) = 0
     Exit Sub
  End If ';Αν εισαι και πας μακρυα πανε ισια


  If (var(40) < var(79) And X_Mpoyk_80 < var(79)) Then
      var(43) = 0
      Exit Sub
  End If ' Αν εισαι και πας ολο αριστερα πανε ισια

  If (var(40) > var(74) And X_Mpoyk_80 > var(74)) Then
      var(43) = 0
      Exit Sub
  End If ' Αν εισαι και πας ολο δεξια πανε ισια

  If (var(41) < Y_diadromos62 And Y_Mpoyk_81 > Y_diadromos62) Then
      var(43) = 1
      Exit Sub
  End If ' Αν εισαι κοντα και πας μακρυα κανε κλικ μονο-διπλο

  If (var(41) > Y_diadromos62 And Y_Mpoyk_81 < Y_diadromos62) Then
     var(43) = 2
     Exit Sub
  End If  ';Αν εισαι μακρυα και πας κοντα κανε κλικ διπλο-μονο

  If (var(41) < Y_diadromos62 And Y_Mpoyk_81 < Y_diadromos62) Then
      var(43) = 3
      Exit Sub
  End If ';Αν εισαι πριν κοντα και μετα κοντα κανε δυο κλικ

End If

If (F_0epis_1seZyg35 = 0 And var(10) = 0) Then ' ; Περ 2,αν ειναι να παρει μπουκαλα απο τις κεφαλες-ζυγαρια για τη θεση της

 If (var(44) = 0) Then
    var(43) = 0
    Exit Sub
 End If ' Αν οι κεφαλες ειναι αδειες πανε ισια

 If (var(40) >= var(79) And var(40) <= var(74)) Then
     var(43) = 0
     Exit Sub
 End If ' Αν εισαι μπροστα στις κεφαλες πανε ισια

 If (var(41) < Y_diadromos62) Then
     var(43) = 1
     Exit Sub
 End If  ';Αν εισαι κοντα κανε μονο κλικ

 var(43) = 1
   Exit Sub

End If

If (F_0epis_1seZyg35 = 0 And var(10) = 1) Then ' ; Περ 3,αν ειναι να παει μπουκαλα που πηρε απο τις κεφαλες-ζυγαρια για τη θεση της

 If (var(44) = 0) Then
     var(43) = 0
     Exit Sub
 End If ';Αν οι κεφαλες ειναι αδειες πανε ισια

 If (X_Mpoyk_80 <= var(56) And var(56) < var(78)) Then
    var(43) = 0
    Exit Sub
 End If
 
 If (X_Mpoyk_80 >= var(56) And var(56) > var(78)) Then
     var(43) = 0
     Exit Sub
 End If

 If (Y_Mpoyk_81 >= Y_diadromos62) Then
     var(43) = 1
     RobSend ("!gowhen(2pm>=" + Str(Y_diadromos62) + "):")
     Exit Sub
 End If ';Αν ειναι να πας μακρυα πανε ισια
 
 
 If (Y_Mpoyk_81 < Y_diadromos62) Then
     var(43) = 2
     Exit Sub
 End If  ';Αν εισαι κοντα κανε διπλο κλικ

 var(43) = 2
 Exit Sub
End If

If (F_0epis_1seZyg35 = 1 And var(10) = 1) Then ' ; Περ 4,αν ειναι να παει μπουκαλα στις κεφαλες-ζυγαρια
  If (var(44) = 0) Then
      var(43) = 0
      Exit Sub
  End If ';Αν οι κεφαλες ειναι αδειες πανε ισια
  
  If (var(40) < var(56)) Then
    If (ArMp_DexKef22 = 0) Then
       var(43) = 0
       Exit Sub
    End If
  End If
  
  If (var(40) > var(56)) Then
    If ArMp_ArKef21 = 0 Then
       var(43) = 0
       Exit Sub
    End If
  End If
  var(45) = 2 * var(58)
  var(45) = var(45) + Y_diadromos62 ';ειναι απεναντι αλλα τουλαχιστον 2 βηματα μακρυα
  
  If (var(41) > var(45) And var(40) > var(79) And var(40) < var(74)) Then
      var(43) = 0
      Exit Sub
  End If  ' ;Αν ερχεσαι απο ακριβως απεναντι ελα ισια
  
  If (Y_Mpoyk_81 >= Y_diadromos62) Then
      var(43) = 1
      Exit Sub
  End If ';Αν ερχεσαι απο μακρυα κανε κλικ διπλο-μονο
  
  If (var(41) < Y_diadromos62) Then
      var(43) = 2
      Exit Sub
  End If  ';Αν εισαι κοντα κανε μονο κλικ
  var(43) = 2
End If

End Sub

