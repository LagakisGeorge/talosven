VERSION 4.00
Begin VB.Form Form0 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   1155
   ClientTop       =   1575
   ClientWidth     =   8505
   Height          =   6345
   Left            =   1095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   8505
   Top             =   1230
   Width           =   8625
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   3135
      TabIndex        =   14
      Top             =   3720
      Width           =   1125
   End
   Begin VB.CommandButton Command7 
      Caption         =   "πάρε ένα μπουκάλι από κεφαλ΄ξ-φέρε άλλο μπουκάλι σε κεφαλή"
      Height          =   540
      Left            =   7140
      TabIndex        =   13
      Top             =   5145
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ΕΠΕΣΤΡΕΨΕ ΜΠΟΥΚΑΛΙ ΑΠΟ ΚΕΦΑΛΗ ΣΤΗΝ ΘΕΣΗ ΤΟΥ"
      Height          =   1005
      Left            =   4500
      TabIndex        =   12
      Top             =   3960
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "φερε μπουκάλι σε ζυγαριά "
      Height          =   645
      Left            =   4455
      TabIndex        =   11
      Top             =   3195
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1485
      TabIndex        =   9
      Text            =   "201"
      Top             =   3735
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1485
      TabIndex        =   7
      Text            =   "1"
      Top             =   3195
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eπιστροφή μπουκαλιού από χειριστή"
      Height          =   690
      Left            =   6345
      TabIndex        =   6
      Top             =   3960
      Width           =   1425
   End
   Begin VB.CommandButton Command4 
      Caption         =   "φερε μουκάλι σε χειριστή"
      Height          =   735
      Left            =   6345
      TabIndex        =   5
      Top             =   3150
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "φερε ποτηρι σε ζυγαριά "
      Height          =   645
      Left            =   1935
      TabIndex        =   3
      Top             =   4545
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3150
      Top             =   5805
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ΕΠΙΣΤΡΟΦΗ ΠΟΤΗΡΙΟΥ ΑΠΟ ΖΥΓΑΡΙΑ "
      Height          =   600
      Left            =   180
      TabIndex        =   1
      Top             =   4590
      Width           =   1500
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
   Begin VB.Label Label4 
      Caption         =   "μπουκάλι"
      Height          =   330
      Left            =   225
      TabIndex        =   10
      Top             =   3735
      Width           =   1230
   End
   Begin VB.Label Label3 
      Caption         =   "ποτήρι"
      Height          =   330
      Left            =   225
      TabIndex        =   8
      Top             =   3195
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Top             =   5265
      Width           =   6765
   End
   Begin VB.Label Label1 
      Caption         =   "gripper"
      Height          =   285
      Left            =   3195
      TabIndex        =   2
      Top             =   3015
      Width           =   780
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "R_FORM.frx":0000
      Height          =   2835
      Left            =   240
      OleObjectBlob   =   "R_FORM.frx":000E
      TabIndex        =   0
      Top             =   90
      Width           =   7905
   End
End
Attribute VB_Name = "Form0"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Sub kefbuk() '  ;Επιστροφη μπουκαλας απο κεφαλες
     var(35) = 0 ' ;Για επιστροφη
     anixe
     chakef
     'trace0
     vrskef
     rota
     If var(43) = 1 Then
         Find_Axes
         y0 = var(62)
         RobSend_w ("!d" + Str(var(56)) + "," + Str(var(62)) + ":!go,1:")
         x0 = var(56)
         RobSend ("!go1:")
       End If
       talos4
  If var(14) = 1 Then strofi
     topkef
     Find_Axes
     x0 = var(80): y0 = var(81): z0 = var(93)
     RobSend ("!d " + Str(var(80)) + "," + Str(var(81)) + "," + Str(var(93)) + ":")
     panbuk

End Sub





Sub metato()
' βλεπει αν χρειάζεται μετακίνηση ο αγκώνας της ΙΡΙΣ
       var(7) = 0
        var(44) = axe(3) '    ;Αν εισαι αλλού
        If var(44) <> var(93) Then
            Find_Axes
            z0 = var(93)
            RobSend ("!d,," + Str(var(93)) + ":") '   ; Πάνε στο πανω Υ μπουκαλιων
            RobSend_w ("!go,,1:")
            'wait(3pm=var93)
        End If

        var(44) = var(21) + var(22) + var(19)
        If var(44) > 0 Then   '      ' ;Αν εχεις εστω και μια μπουκαλα στις κεφαλες
            var(44) = axe(2) 'pm
           If var(44) < var(62) Then
               var(44) = axe(1) ' pm
              If var(44) > -70000 Then
                 RobSend ("!d," + Str(var(62)) + ":") '     ;Πανε στριψε στο διαδρομο
                 Find_Axes
                 y0 = var(62)
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
    z0 = var(93)
    RobSend ("!d,," + Str(var(93)) + ":")
    RobSend_w ("!go,,1:")
     If var(14) <> var(15) Then strofi
End If


rota

If var(43) = 0 Then ';Κατ'ευθειαν σκυψιμο ισιας πορειας
   
   var(44) = axe(1) 'pm
    If var(44) > var(80) Then
      var(44) = var(44) - var(80)
    Else
      var(44) = var(80) - var(44)
    End If

   var(45) = axe(2) 'pm
   If var(45) > var(81) Then
      var(45) = var(45) - var(81)
   Else
      var(45) = var(81) - var(45)
   End If
  If var(44) > var(45) Then
   If var(80) > axe(1) Then 'pm)
      var(44) = var(80) - var(55)
      RobSend ("!gowhen(,,1pm>" + Str(var(44)) + "):")
   End If
   If var(80) < axe(1) Then ' pm)
      var(44) = var(80) + var(55)
      RobSend ("!gowhen(,,1pm<" + Str(var(44)) + "):")
   End If
  Else
   
     If var(81) > axe(2) Then 'pm)
        var(45) = var(81) - var(55)
        RobSend ("!gowhen(,,2pm>" + Str(var(45)) + "):")
     End If
     If var(81) < axe(2) Then 'pm)
        var(45) = var(81) + var(55)
        RobSend ("!gowhen(,,2pm<" + Str(var(45)) + "):")
     End If
  End If
  


Find_Axes
x0 = var(80): y0 = var(81): z0 = var(92)
RobSend ("!d" + Str(var(80)) + "," + Str(var(81)) + "," + Str(var(92)) + ":") ' ;για ισια

 
 If ERGOSTASIO = "iris" Then
    If (var(21) = 0) Then
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
   If var(44) > var(80) Then
      var(44) = var(44) - var(80)
   Else
      var(44) = var(80) - var(44)
   End If
   
   var(45) = var(62)
   If var(45) > var(81) Then
      var(45) = var(45) - var(81)
   Else
      var(45) = var(81) - var(45)
   End If
  
  
  If var(44) > var(45) Then
     If var(80) > axe(1) Then
        var(44) = var(80) - var(55)
        RobSend ("!gowhen(2pm>" + Str(var(62)) + ",,1pm>" + Str(var(44)) + "):")
     End If
     If var(80) < axe(1) Then
        var(44) = var(80) + var(55)
        RobSend ("!gowhen(2pm>" + Str(var(62)) + ",,1pm<" + Str(var(44)) + "):")
     End If
  Else
     var(45) = var(81) - var(55)
     RobSend ("!gowhen(2pm>" + Str(var(62) + ",,2pm>" + Str(var(45))) + "):")
  End If
   RobSend_w ("!d" + Str(var(80)) + "," + Str(var(81)) + "," + Str(var(92)) + ":!go111:")
End If



If var(43) = 2 Then '   ;για κλικ διπλο-μονο
  var(44) = axe(1) 'pm
  If var(44) > var(80) Then
     var(44) = var(80) + var(55)
     var(45) = var(81) + var(55)
     RobSend ("!gowhen(,1pm<" + Str(var(44)) + ",2pm<" + Str(var(45)) + "):")
  Else
     var(44) = var(80) - var(55)
     var(45) = var(81) + var(55)
     RobSend ("!gowhen(,1pm>" + Str(var(44)) + ",2pm<" + Str(var(45)) + "):")
 End If
   Find_Axes
   x0 = var(80): y0 = var(81): z0 = var(92)
   RobSend_w ("!d " + Str(var(80)) + "," + Str(var(81)) + "," + Str(var(92)) + ":!go111")
End If


If var(43) = 3 Then   ' ;για δυο κλικ
  Find_Axes
  y0 = var(62)
  RobSend_w ("!d," + Str(var(62)) + ":!go,1:")
    var(44) = axe(1) 'pm
  If var(44) > var(80) Then
     var(44) = var(79) - var(80)
     var(45) = var(62) - var(81)
    If var(44) > var(45) Then
       var(44) = var(80) + var(55)
       RobSend ("!gowhen(,1pm<" + Str(var(79)) + ",1pm<" + Str(var(44)) + "):")
    Else
      var(45) = var(81) + var(55)
      RobSend ("!gowhen(,1pm<" + Str(var(79)) + ",2pm<" + Str(var(45)) + "):")
    End If
  Else
     var(44) = var(80) - var(74)
     var(45) = var(62) - var(81)
    If var(44) > var(45) Then
       var(44) = var(80) - var(55)
       RobSend ("!gowhen(,1pm>" + Str(var(74)) + ",1pm>" + Str(var(44)) + "):")
    Else
       var(45) = var(81) + var(55)
       RobSend ("!gowhen(,1pm>" + Str(var(74)) + ",2pm<" + Str(var(45)) + "):")
    End If
 End If

      Find_Axes
     x0 = var(80): y0 = var(81): z0 = var(92)
     RobSend_w ("!d " + Str(var(80)) + "," + Str(var(81)) + "," + Str(var(92)) + ":!go111:")
End If

 RobSend ("!comexc0:")
 Find_Axes
  z0 = var(92)

  RobSend_w ("!d,," + Str(var(92)) + ":!go,,1:")
  klise
  
  Find_Axes
  z0 = var(93)
  RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")

End Sub




Sub bukpot()
'     Mεταφορα μπουκαλιων στο χρηστη


var(30) = 1
var(3) = var(1)
talos4
RobSend ("!COMEXC0:")
Find_Axes
z0 = var(93): w0 = var(65)
RobSend_w ("!d , ," + Str(var(93)) + "," + Str(var(65)) + ":!go0011:")
ferbuk
var(1) = var(2) ';Αλλαγη μεταβλητης
var(49) = var(95)

Find_Axes
x0 = var(49): y0 = var(96)
RobSend_w ("!d" + Str(var(49)) + "," + Str(var(96)) + ":!go11:")

Find_Axes
z0 = var(97)
RobSend_w ("!d,," + Str(var(97)) + ":!go,,1:")

anixe

Find_Axes
z0 = var(93)
RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")
var(44) = var(51) - var(1)
End Sub
Sub talos4()
'  Υπολογισμός στήλης και γραμμης που βρίσκεται η μπουκαλα
var(34) = 0
    var(80) = var(36) '                  ;Συντεταγμενη Χ
    var(81) = var(37) '                 ;Συντεταγμενη Y
 End Sub



Sub potbuk() ' Mεταφορα μπουκαλιων απο χρηστη στη θεση τους

RobSend ("!COMEXC0:")
var(30) = 1
Find_Axes
z0 = var(93): w0 = var(65)
RobSend_w ("!d , ," + Str(var(93)) + "," + Str(var(65)) + ":!go0011:")
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
    y0 = var(62)
    RobSend_w ("!d," + Str(var(62)) + ":!go,1:")
End If
Find_Axes
x0 = var(80): y0 = var(81)
RobSend ("!d " + Str(var(80)) + "," + Str(var(81)) + ":")

'if(var14<>var15):strofi:nif
'if(var43=0 or var30=1): go11:nif ;Αν ειναι χειριστη για χειριστη ή πας μακρυα
'****** dokkyr.txt
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
  If var(80) < var(79) Then RobSend_w ("!gowhen(,1pm <=" + Str(var(79)) + "):!go11:")
  If var(80) > var(74) Then RobSend ("!gowhen(,1pm >=" + Str(var(74)) + "):!go11:")
End If
RobSend ("!comexc0:")

If ERGOSTASIO = "iris" Then
      Find_Axes
      x0 = var(80): y0 = var(81)
      RobSend_w ("!d " + Str(var(80)) + "," + Str(var(81)) + ":!go11:")
End If
Find_Axes
z0 = var(92)
RobSend ("!d,," + Str(var(92)) + ":") '            ; συντ. μπουκάλας κάτω Z
RobSend_w ("!go,,1:") ' κατεβαινει

anixe

Find_Axes
z0 = var(93)
RobSend ("!d , , " + Str(var(93)) + ":")
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
z0 = var(93)

RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")

var(25) = var(21) + var(22) + var(19)
 
If ERGOSTASIO = "iris" Then
 If var(25) > 0 Then ' ;Αν υπαρχει εστω και 1 μπουκ. στις κεφαλες
   Form1.Caption = "Υπαρχει μπουκαλα στις κεφαλες"
   Exit Sub
 End If
End If

   var(49) = var(95) ' ΣΥΝΤΕΤΑΓΜΕΝΗ ΠΡΩΤΟΥ ΠΟΤΗΡΙΟΥ ΚΑΤΑ Χ
 If var(30) = 1 Then ' ΜΕΤΑΦΟΡΑ ΣΕ ΧΕΙΡΙΣΤΗ=1
    var(44) = var(65) ' θέση 4ου : 1η κεφ σε ζυγ
  Else
    var(44) = var(61) 'θέση παρκαρίσματος 4 αξ
 End If
  
  'ΠΑΝΕ ΠΑΝΩ ΑΠΟ ΤΟ ΠΟΤΗΡΙ
   '95= Χ  ποτηριού     96=Υ ποτηριού   93=καθ.ανω Ζ
x0 = var(49): y0 = var(96): z0 = var(93): w0 = var(44)
RobSend ("!d " + Str(var(49)) + "," + Str(var(96)) + "," + Str(var(93)) + "," + Str(var(44)) + ":")
 
 If ERGOSTASIO = "iris" Then
        RobSend_w ("!go1111:")
        If var(14) = 1 Then strofi
 Else
        z0 = var(93): Find_Axes: RobSend_w ("!go,,1:")
 End If
  
   RobSend ("!d,," + Str(var(97)) + ":")
 
 If ERGOSTASIO = "iris" Then
   Find_Axes
   z0 = var(97): RobSend_w ("!go,,1:")
 Else
  
     var(44) = axe(1)  ' pm
     If var(44) > var(49) Then
        var(44) = var(44) - var(49)
     Else
        var(44) = var(49) - var(44)
     End If
     
     var(45) = axe(2) 'pm
     If var(45) > var(96) Then
        var(45) = var(45) - var(96)
     Else
        var(45) = var(96) - var(45)
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
         If var(96) > axe(2) Then 'pm)
            var(45) = var(96) - var(55)
            RobSend ("!gowhen(,,2pm>" + Str(var(45)) + ":")
         End If
         If var(96) < 2 Then 'pm)
            var(45) = var(96) + var(55)
            RobSend ("!gowhen(,,2pm<" + Str(var(45)) + ":")
         End If
     End If
     
     Find_Axes
     x0 = var(49): y0 = var(96)
'
     z0 = var(97)
     RobSend_w ("!go1111:")
     Find_Axes
     z0 = var(97)
     RobSend_w ("!d,," + Str(var(97)) + ":!go,,1:")
 End If
 
 klise
 
 Find_Axes
 z0 = var(93)
 RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")
 
 If var(30) = 1 Then '              ;Αν ειναι μπουκαλια από χειριστη για επιστροφη στη θεση τους
         Exit Sub                     ';Σταματα εδω
 End If
  Find_Axes
  x0 = var(85): y0 = var(86)
  RobSend_w ("!d " + Str(var(85)) + "," + Str(var(86)) + ":!go11:")
  potzyg
 End Sub




Sub anixe()

If FindInput(Gripper_Input) = 1 Then  ' kleistos
     RobSend ("!" + Gripper_On + ":")
     Do While FindInput(Gripper_Input) = 1
       DoEvents
    Loop
    var(10) = 0
End If

End Sub
Sub potzyg() 'Τοποθετηση σε ζυγαρια

  RobSend ("comexc0")
                    ' Χ ζυγού            Υ  ζυγού             Ζ ανω                  W parking
    
        
    x0 = var(85): y0 = var(86): z0 = var(93): w0 = var(61)
    RobSend ("!d" + Str(var(85)) + "," + Str(var(86)) + "," + Str(var(93)) + "," + Str(var(61)) + ":") ' ΧΥ  πανω Ζ ζυγαρ
    RobSend_w ("!go1111:") '              ;Ελα από πανω
    
    z0 = var(87)
    RobSend ("!d,," + Str(var(87)) + ":") ' κατω Ζ ζυγ
    RobSend_w ("!go,,1:") '               ;Κατεβα
  
  
  If var(10) = 0 Then     '  Αν εισαι ανοικτος
     klise
  Else
    anixe
  End If
    
' ανέβα πάνω Ζ
    z0 = var(93)
    RobSend ("!d , ," + Str(var(93)) + ":")
    RobSend_w ("go,,1:")
End Sub


Sub panpot() ' ;Επιστροφη ποτηριου

' αριθμός μπουκαλιών σε κεφαλές & μιξερ
var(25) = var(21) + var(22) + var(19)

var(44) = var(61) ' θέση παρκαρίσματος
Find_Axes
w0 = var(44)
' παρκάρισε τον 4ο άξονα
RobSend_w ("!d,,," + Str(var(44)) + ":go,,,1:")

var(35) = 0

anixe

potzyg 'πιάσε το ποτήρι και ανέβα στο πάνω Ζ

var(49) = var(95)  ' συντεταγμένη ποτηριού κατά Χ

Find_Axes
x0 = var(49): y0 = var(96)
                    'Χ ποτηριού         Υ ποτηριού
RobSend ("!d " + Str(var(49)) + "," + Str(var(96)) + ":")
RobSend_w ("!go11:")

Find_Axes
z0 = var(97)
                    ' κάτω Ζ ποτηριού
  RobSend_w ("!d,," + Str(var(97)) + ":go,,1:")
  anixe
  


Find_Axes
z0 = var(93)
        
        ' άνω Ζ
  RobSend_w ("!d,," + Str(var(93)) + ":go,,1:")

End Sub








Sub klise()

If FindInput(Gripper_Input) = 0 Then  ' anoiktos
dz = -200
     z0 = z0 + dz
    RobSend_w ("!ma0000:!d,," + Str(dz) + ":!go,,1:")
    RobSend ("!WAIT (mov = bxx0x):")
    RobSend ("!ma1111:")  '  ;Κατεβαινει 1 mm για να αποφυγει διαφορες Ζ
    RobSend ("!" + Gripper_Off + ":")
   Do While FindInput(Gripper_Input) = 0 '  kleistos
      DoEvents
   Loop
   var(10) = 1
  ' RobSend ("!t0 0.5:")
  MilSec 10
End If

End Sub



Sub akyro_fer_pot() ' ;Μεταφορα ποτηριων απο τη θεση τους

RobSend ("!comexc0:")
anixe

var(25) = var(21) + var(22) + var(19)    ' ΑΡΙΘΜΟΣ ΜΠΟΥΚΑΛΙΩΝ ΣΕ ΚΑΦΑΛΕΣ & ΜΙΞΕΡ

var(49) = var(95)  ' ΣΥΝΤΕΤΑΓΜΕΝΗ ΠΡΩΤΟΥ ΠΟΤΗΡΙΟΥ ΚΑΤΑ Χ

 If var(30) = 1 Then  ' ΜΕΤΑΦΟΡΑ ΣΕ ΧΕΙΡΙΣΤΗ=1
      var(44) = var(65) ' θέση 4ου : 1η κεφ σε ζυγ
 Else
      var(44) = var(61) 'θέση παρκαρίσματος 4 αξ
 End If
 
 
 'ΚΙΝΗΣΗ ΣΤΟ ΑΝΩ Ζ   ( Χ ποτ           Υ ποτ                 άνω Ζ
 RobSend ("!d " + Str(var(49)) + "," + Str(var(96)) + "," + Str(var(93)) + "," + Str(var(44)) + ":go,,1:")
 
 ' κάτω Ζ
 RobSend ("!d , , " + Str(var(97)) + ":")

    
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
   If var(45) > var(96) Then
      var(45) = var(45) - var(96)
  Else
      var(45) = var(96) - var(45)
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
       
       If var(96) > axe(2) Then  'pm)
           var(45) = var(96) - var(55)
            RobSend ("gowhen(,,2pm>" + Str(var(45)) + "):")
        End If
        If var(96) < axe(2) Then  ' pm)
              var(45) = var(96) + var(55)
              RobSend ("!gowhen(,,2pm<" + Str(var(45)) + "):")
         End If
  End If

                                ' κατω Z ποτηριού
  RobSend ("!go1111 :  d,," + Str((var(97))) + " :  !go,,1:")
  
 klise
 RobSend ("!d,,(var(93) :  !go,,1:")
 
 If var(30) = 1 Then  'Αν ειναι μπουκαλια από χειριστη για επιστροφη στη θεση τους
         Exit Sub                     'Σταματα εδω
 End If
 
                 ' Χ ζυγαριάς          Υ ζυγαριάς
 RobSend ("!d " + Str(var(85)) + "," + Str((var(86))) + ":!go11:")
 
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
        var(35) = 1 ' ΣΕ ΖΥΓΑΡΙΑ - ΚΕΦΑΛΕΣ          0=ΕΠΙΣΤΡΟΦΗ ΣΤΗΝ ΑΡΧΙΚΗ ΘΕΣΗ
        var(30) = 0  ' ΜΕΤΑΦΟΡΑ ΟΧΙ ΣΕ ΧΕΙΡΙΣΤΗ
        'WRVAR1
        Ferpot
        RobSend ("d,,," + Str(var(67)) + ":go,,,1:")
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
    
    var(35) = 0 ' ΣΕ ΖΥΓΑΡΙΑ - ΚΕΦΑΛΕΣ          0=ΕΠΙΣΤΡΟΦΗ ΣΤΗΝ ΑΡΧΙΚΗ ΘΕΣΗ
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
        var(35) = 1 ' ΣΕ ΖΥΓΑΡΙΑ - ΚΕΦΑΛΕΣ          0=ΕΠΙΣΤΡΟΦΗ ΣΤΗΝ ΑΡΧΙΚΗ ΘΕΣΗ
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

var(25) = var(21) + var(22) + var(19)
 If var(25) > 0 Then ' ;Αν υπαρχει εστω και 1 μπουκ. στις κεφαλες
      Form1 = "Υπαρχει μπουκαλα στις κεφαλες"
    Exit Sub
 End If
 
 var(49) = var(95)
 
 If var(30) = 1 Then
       var(44) = var(65)
      Else
         var(44) = var(61)
 End If
  
  
  'ΠΑΝΕ ΠΑΝΩ ΑΠΟ ΤΟ ΠΟΤΗΡΙ
   '95= Χ  ποτηριού     96=Υ ποτηριού   93=καθ.ανω Ζ
   RobSend ("!d " + Str(var(49)) + "," + Str(var(96)) + "," + Str(var(93)) + "," + Str(var(44)) + ":")
  
  
RobSend_w ("!go1111:")


If var(12) > 19000 And var(12) < 19020 Then RobSend ("!out.6-1:")
'  if var(14)=1 then strofi
  
 RobSend_w ("!d,," + Str(var(97)) + " :!go,,1:")
 
 klise
 RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")
 
 If var(30) = 1 Then  'Αν ειναι μπουκαλια από χειριστη για επιστροφη στη θεση τους
         Exit Sub      'Σταματα εδω
 End If
  
  
RobSend_w ("!d " + Str(var(85)) + "," + Str(var(86)) + ":!go11:")


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
    RobSend ("!v 32,32,35,1.5:")
var(12) = 6000 + Val(Text2.Text)
var(13) = Val(Text1.Text)
var(94) = 18
find_uesh ("!var12=" + Str(var(12)) + ":!var13=" + Str(var(13)) + ":")

talos
    
    
    
End Sub

Private Sub talos()
' απο χειριστή σε αρχ.θέση μπουκαλιού
 
var(31) = (var(74) - var(78)) / (var(66) - var(65)) ' ;αCoef14
var(32) = var(78) - (var(65) * var(31))    ';βCoef14
var(33) = var(65): var(23) = var(66)
var(56) = var(33) + var(65) - var(23): var(56) = var(56) * var(31)
var(56) = var(56) + var(32)
var(79) = var(56) ' ;Συντεταγμενη αριστερης κεφαλης
var(94) = 18
 
' αποθηκεύω τις συντεταγμένες του ποτηριού σε άλλες μεταβλητές
   var(95) = var(24)
   var(96) = var(77)
   var(97) = var(89)
    
    
 
 
'μεταφορά μπουκαλιού var(12) σε θέση ποτηριού var(13)
' var(12):μπουκάλι   var(13):ποτήρι
'    RobSend ("!v 8,12,10,1:")
  RobSend ("!v 32,32,35,1.5:")


' Απο Vb ερχονται τα var24->Χ ποτηριων var77 -> Yποτηριων var89-> Z ποτηριων
If var(12) = 100 Then '  ;Τοποθέτηση ΜΠΟΥΚΑΛΙΟΥ σε θεση αναδευσης
        var(1) = var(12)
       RobSend ("!COMEXC0:")
        If var(11) = 0 Then
         Form1.Caption = "Δεν υπαρχει ποτηρι στη ζυγαρια"
         Exit Sub
         RobSend ("!out.16 - 1:")
        End If
         var(35) = 0
             anixe
             Find_Axes
             z0 = var(93)
             RobSend_w ("!d" + Str(var(85)) + "," + Str(var(86)) + "," + Str(var(93)) + "," + Str(var(61)) + ":!go,,1:")
             x0 = var(85): y0 = var(86): w0 = var(61)
             RobSend_w ("!go1101:")
             z0 = var(87)
             RobSend_w ("!d,," + Str(var(87)) + ":!go,,1:")
             klise
             Find_Axes
             z0 = var(93)
             RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")
             z0 = var(93)
             RobSend_w ("!d" + Str(var(70)) + "," + Str(var(71)) + "," + Str(var(93)) + ":!go,,1:")
             x0 = var(70): y0 = var(71):
             RobSend_w ("!go11:")
             z0 = var(72)
             RobSend_w ("!d,," + Str(var(72)) + ":!go,,1:")
             anixe
             z0 = var(93)
             RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")
        
             var(1) = 0
             var(2) = 0
             var(30) = 0
             var(11) = 3
      
    RobSend ("!out.16 - 1:")
     
    Exit Sub
End If

If var(12) = 101 Then '  ;Επιστροφη ποτηριου απο αναδευση σε ζυγαρια
        var(1) = var(12)
     RobSend ("!COMEXC0:")
        var35 = 1
        anixe
        RobSend ("!d" + Str(var(70)) + "," + Str(var(71)) + "," + Str(var(93)) + ":")
        Find_Axes
        z0 = var(93)
        RobSend_w ("!go,,1:")
        Find_Axes
        x0 = var(70): y0 = var(71)
        RobSend_w ("!go11:")
        RobSend ("!d,," + Str(var(72)) + ":")
        Find_Axes
        z0 = var(72)
        RobSend_w ("!go,,1:")
        klise
        RobSend ("!d,," + Str(var(93)) + ":")
        Find_Axes
        z0 = var(93)
        RobSend ("!go,,1:")
        
        RobSend ("!d" + Str(var(85)) + "," + Str(var(86)) + "," + Str(var(93)) + "," + Str(var(61)) + ":")
        x0 = var(85): y0 = var(86): z0 = var(93): w0 = var(61)
        RobSend_w ("!go1111:")
        var11 = 2
       ' topzyg
        RobSend ("!d,,," + Str(var(67)) + ":")
        RobSend ("!go,,,1:")
        var(1) = 0: var(2) = 0: var(30) = 0
        RobSend ("!out.16 - 1:")
     
        Exit Sub
End If

If var(12) > 200 And var(12) < 221 Then ';Επιστροφη ποτηριου απο αναδευση στη θεση του
        var(1) = var(12) - 200
     RobSend ("!COMEXC0:")
        var(11) = 2: var(35) = 0: var(30) = 0
        anixe
        
        RobSend ("!d " + Str(var(70)) + "," + Str(var(71)) + "," + Str(var(93)) + ":")
        Find_Axes
        z0 = var(93)
        RobSend_w ("! go,,1:")
        Find_Axes
        x0 = var(70): y0 = var(71)
        RobSend_w ("!go11:")
        Find_Axes
        z0 = var(72)
        RobSend ("!d,," + Str(var(72)) + ":")
        
        RobSend_w ("!go,,1:")
        klise
        RobSend ("!d,," + Str(var(93)) + ":")
        Find_Axes
        z0 = var(93)
        RobSend_w ("!go,,1:")
       
        ' mixpot

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
If var(12) > 1000 And var(12) < 1280 Then ' ;Τοποθέτηση μπουκάλας σε κεφαλή δοσομέτρησης
        var(1) = var(12)
    var(1) = var(1) - 1000
        var(44) = var(21) + var(22)
    bukkef
    Form1.Caption = "Μπουκαλα σε κεφαλες"
    var(1) = 0: var(2) = 0: var(30) = 0
    RobSend ("!OUT.14 - 1:")
End If


If var(12) > 19000 And var(12) < 19020 Then ' Μεταφορά ποτηριων στη ζυγαρια για σκονη
        var(1) = var(12)
        RobSend ("!out.6 - 1:")
        var(1) = var(1) - 19000
        var(35) = 1
        var(30) = 0
        var(99) = var(75)
        var(75) = var(73) ' ;Βαλε συντ.Κεφαλης σκονης κατα Υ
        Ferpot
        RobSend ("!out.6 - 0:")
        MilSec 500   't1
        Find_Axes
        w0 = var(68)
        RobSend_w ("!d,,," + Str(var(68)) + ":!go,,,1:")
        
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
        
        var(25) = var(21) + var(22) + var(19)   'αριθμος μπουκαλιών σε κεφαλές & μιξερ
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


If var(12) > 2000 And var(12) < 2280 Then '   ;Επαναφορά μπουκάλας στη βαση της
        var(1) = var(12)
    var(1) = var(1) - 2000
    kefbuk
    var(1) = 0: var(2) = 0: var(30) = 0
    RobSend ("!out .16 - 1:")
           If var(21) < 2 And var(22) = 0 And var(11) = 1 Then '  ;Αν συνεχιζει η δοσομετρηση και υπαρχει το πολυ μια μπουκαλα
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
        
        var(1) = var2
        var(36) = var(38)
        var(37) = var(39)      ' Κύκλος των var για Β' μπουκάλι
        var(44) = var(21) + var(22)
        bukkef
        var(1) = 0: var(2) = 0
        RobSend ("!out.16 - 1:")
        var(12) = 0
        var(13) = 0
        Form1.Caption = "Μπουκάλα σε κεφαλες"
        Exit Sub
 End If

If var(12) > 100000 And var(12) < 100280 Then '  ;Πανε μπουκ και φερε μπουκ ταυτοχρονα
        var(1) = var(12) - 100000: var(2) = var(13): var(30) = 0
      If var(13) = 0 Or var(13) > 280 Then
        Form1.Caption = "Λαθος φερε μπουκαλι"
         Exit Sub
      End If
     
      RobSend ("!out16 - 0:")

    kefbuk
    Form1.Caption = "Επιστροφη Μπουκαλας"
        var(1) = var(2)
        var(36) = var(38)
        var(37) = var(39) 'Κύκλος των var για Β' μπουκάλι
        ' WRVAR1
        var(44) = var(21) + var(22)
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

var(44) = var(21) + var(22) + var(19)
  If var(44) = 1 And var(35) = 0 Then '  ;Αν θα παρεις την τελευταια μπουκαλα
      RobSend ("!comexc0:")
      Find_Axes
      w0 = var(67)
      RobSend_w ("!d,,," + Str(var(67)) + ":!go,,,1:") '   ;Πανε στα χημικα
      Exit Sub
  End If
 If var(44) = 0 And var(35) = 1 Then ' ;Αν δεν υπαρχ.μπουκ.και ειναι να φερει
        If var(11) = 2 Then '    ;Αν ειναι για σκονη
              Find_Axes
              w0 = var(68)
              RobSend ("!d , , , " + Str(var(68)) + ":")
        Else
           var(44) = axe(4)
           If var(44) <> var(65) Then
                Find_Axes
                w0 = var(65)
                RobSend ("!d,,," + Str(var(65)) + ":!comexc1:!go,,,1:")
           End If
           Exit Sub
        End If
   End If
 If var(44) = 2 And var(35) = 0 Then
        var(44) = axe(4) 'pm
        
        Find_Axes
        If var(44) = var(65) Then '           ;Αν τελευταια τοποθετησες αριστερα
              w0 = var(66)
              RobSend ("!d,,," + Str(var(66)) + ":") '   ;Τραβα δεξια
         Else
              w0 = var(65)
              RobSend ("!d,,," + Str(var(65)) + ":") '     ;Τραβα αριστερα
         End If
         RobSend_w ("!go,,,1:")     ' ;κανε κλικ
 End If
End Sub


Sub vrskef()
'def vrskef  ;var(21)=Μπουκαλα στην αριστερη κεφαλη,var(22)= μπουκαλα στην δεξια κεφαλη

var(16) = var(21): var(17) = var(22): var(33) = axe(4) 'Για να παρει = 0 , για να φερει =1
  If var(11) > 1 And var(35) = 1 Then
        If var(19) = 1 Then
            Form1.Caption = "Η κεφαλη σκονης ειναι κατηλειμενη"
            RobSend ("!ps:")
            Exit Sub
         End If
        var(23) = var(68): var(19) = 1
        var(56) = var(33) + var(65) - var(23)
       var(56) = var(56) * var(31): var(56) = var(56) + var(32)
       Exit Sub
  End If
 If var(11) > 1 And var(35) = 0 Then
        If var(19) = 0 Then
            Form1.Caption = "Η κεφαλη σκονης ειναι αδεια"
            RobSend ("!ps:")
            Exit Sub
         End If
         var(23) = var(68): var(19) = 0
        var(56) = var(33) + var(65) - var(23)
       var(56) = var(56) * var(31): var(56) = var(56) + var(32)
       Exit Sub
  End If
          If var(16) = 0 And var(17) = 0 Then
             If var(35) = 1 Then ' ;Αν ειναι να φερει
                  var(23) = var(65)
                  var(16) = 1
                  var(56) = var(33) + var(65) - var(23)
                  var(56) = var(56) * var(31): var(56) = var(56) + var(32)
                  Exit Sub
             Else
                  RobSend ("!ps:")
             End If
          End If
          If var(16) = 0 And var(17) = 1 Then '   ) ; 2
            If var(35) = 1 Then '   ;Αν ειναι να φερει
          
                var(23) = var(65)
                var(16) = 1
                var(56) = var(33) + var(65) - var(23)
                var(56) = var(56) * var(31)
                var(56) = var(56) + var(32)
                Exit Sub
            Else
                var(23) = var(66)
                var(17) = 0
                var(56) = var(33) + var(65) - var(23)
                var(56) = var(56) * var(31)
                var(56) = var(56) + var(32)
                Exit Sub
            End If
        End If
          If var(16) = 1 And var(17) = 0 Then ' εχει μπουκάλι μόνο η αριστερή κεφαλή
            If var(35) = 1 Then ')  ;Αν ειναι να φερει
               If var(11) = 2 Then
                  Form1.Caption = "Η κεφαλη σκονης ειναι κατηλειμενη"
                  RobSend ("!ps:")
                  Exit Sub
                End If
                var(23) = var(66): var(17) = 1: var(56) = var(33) + var(65) - var(23): var(56) = var(56) * var(31): var(56) = var(56) + var(32)
                Exit Sub
            Else
                  If var(11) > 1 Then
                    var(23) = var(68)
                    var(16) = 0
                    var(56) = var(33) + var(65) - var(23)
                    var(56) = var(56) * var(31)
                    var(56) = var(56) + var(32)
                    Exit Sub
                  End If
                  var(23) = var(65)
                  var(16) = 0
                  var(56) = var(33) + var(65) - var(23)
                  var(56) = var(56) * var(31)
                  var(56) = var(56) + var(32)
                  Exit Sub
            End If
          End If
          If var(16) = 1 And var(17) = 1 Then ' ) ; 4
             If var(35) = 1 Then '  Αν ειναι να φερει
                RobSend ("!ps:")
            Else
                If var(33) = var(66) Then
                    var(23) = var(65)
                    var(16) = 0
                 End If
                 If var(33) = var(65) Then
                     var(23) = var(66)
                     var(17) = 0
                 End If
                 var(56) = var(33) + var(65) - var(23)
                 var(56) = var(56) * var(31)
                 var(56) = var(56) + var(32)
                 Exit Sub
            End If
          End If
  
  End Sub



Sub bukkef() '  ;Μεταφορα Μπουκαλιων στις κεφαλες
var(35) = 1
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

If var(43) = 1 Then RobSend_w ("!d" + Str(var(56)) + "," + Str(var(62)) + ":!go11:")
If var(43) = 2 Then
    RobSend ("!gowhen(2pm>" + Str(var(75)) + "):")
    Find_Axes
    x0 = var(56): y0 = var(62)
   RobSend_w ("!d" + Str(var(56)) + "," + Str(var(62)) + ":!go11:")
    RobSend_w ("!go1:")
End If

topkef

End Sub

Sub topkef()  ' ;Τοποθετηση σε κεφαλη

     RobSend ("!COMEXC0:")
   Find_Axes
   z0 = var(93)
     RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")
   var40 = axe(1) 'pm:
     If (var40) > var(56) Then
           RobSend ("!gowhen(,1pm<=" + Str(var(56)) + "):")
      Else
           RobSend ("!gowhen(,1pm>=" + Str(var(56)) + "):")
     End If
     Find_Axes
     x0 = var(56): y0 = var(75)
      RobSend_w ("!d " + Str(var(56)) + "," + Str(var(75)) + ":!go11:")
     
     
     If var(35) = 1 Then '      ;Αν ειναι να φερεις
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
         var(21) = var(16): var(22) = var(17)
         If var(11) = 2 Then
            RobSend_w ("!d,," + Str(var(88)) + ":!go,,1:")
            klise
             MilSec 1000   ' t1
              var6 = 1    '           ';Πιασε το χερουλι
             Form1.Caption = "Εναρξη δοσομέτρησης σκόνης."
         Else
            Find_Axes
            z0 = var(93)
            RobSend_w ("!d" + Str(var(78)) + "," + Str(var(62)) + "," + Str(var(93)) + ":!go,,1:")
            
             y0 = var(62)
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
          z0 = var(93)
          RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")
          var(21) = var(16)
          var(22) = var(17)
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
        var(25) = var(21) + var(22) + var(19)
        
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
  '            var(36) = SYN_X
   '           var(37) = SYN_Y
    '         Exit Function
     '   Else
      '       find_Coordin = "!var38=" + LTrim(Str(SYN_X)) + ":!var39=" + LTrim(Str(SYN_Y)) + ":"
       '      var(38) = SYN_X
        '     var(39) = SYN_Y
         '
          '   Exit Function




var(11) = 1 ' αρχισε η συνταγή


var(12) = 100000 + Val(Text2.Text)  'μπουκάλι που φεύγει από κεφαλή
' var(36)=X   var(37)=Y

var(13) = Val(Text3.Text) ' μπουκάλι που έρχεται σε κεφαλή
' var(38)=X   var(39)=Y


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
w0 = var(68)
RobSend_w ("!d,,," + Str(var(68)) + ":!go,,,1:")

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
var(35) = 0
RobSend ("!ma1111:")

Find_Axes
z0 = var(69) ' κατέβα να πιάσεις το μπουκάλι της σκόνης
RobSend_w ("!d,," + Str(var(69)) + ":!go,,1:")

klise

Find_Axes
z0 = var(93)
RobSend_w ("!d,," + Str(var(93)) + ":!go,,1:")

var(6) = 0
RobSend_w ("!d " + Str(var(80)) + "," + Str(var(81)) + "," + Str(var(93)) + "," + Str(var(67)) + ":!go1111:")
var(21) = 0
panbuk
var(19) = 0
var(75) = var(99)
var(11) = 1 ';Ξαναβαλε συντ.Κεφαλων κατα Υ την παλια
End Sub
Sub rota()

var(40) = axe(1) 'pm ;Τρεχουσες συντεταγμενες.Ισχυει για την περιπτ.που ο gripper  ειναι δεξια του βραχιονα
var(41) = axe(2): var(41) = var(41) - var(58) ';Περιθωριο Υ ενα βημα
var(44) = var(21) + var(22) + var(19) ' ;var(21 μπουκάλα σε αρ.κεφαλή var(22 σε δεξ.κεφ  var(19 σε mixer

';var(80) συντ.Χ της προς εύρεση μπουκάλας
';var(81 συντ.Υ της προς εύρεση μπουκάλας

If var(35) = 1 And var(10) = 0 Then ' ; Περ 1,αν ειναι να παρει μπουκαλα απο την θεση της για τις κεφαλες-ζυγαρια
        ' ;var(10)=0 gripper ανοιχτός  var(35=1 προς ζυγαριά ή κεφαλές
  If (var(44) = 0) Then
     var(43) = 0
     Exit Sub
  End If ';Αν οι κεφαλες ειναι αδειες πανε ισια

  If (var(41) > var(62) And var(81) > var(62)) Then
     var(43) = 0
     Exit Sub
  End If ';Αν εισαι και πας μακρυα πανε ισια

  If (var(40) < var(79) And var(80) < var(79)) Then
      var(43) = 0
      Exit Sub
  End If ' Αν εισαι και πας ολο αριστερα πανε ισια

  If (var(40) > var(74) And var(80) > var(74)) Then
      var(43) = 0
      Exit Sub
  End If ' Αν εισαι και πας ολο δεξια πανε ισια

  If (var(41) < var(62) And var(81) > var(62)) Then
      var(43) = 1
      Exit Sub
  End If ' Αν εισαι κοντα και πας μακρυα κανε κλικ μονο-διπλο

  If (var(41) > var(62) And var(81) < var(62)) Then
     var(43) = 2
     Exit Sub
  End If  ';Αν εισαι μακρυα και πας κοντα κανε κλικ διπλο-μονο

  If (var(41) < var(62) And var(81) < var(62)) Then
      var(43) = 3
      Exit Sub
  End If ';Αν εισαι πριν κοντα και μετα κοντα κανε δυο κλικ

End If

If (var(35) = 0 And var(10) = 0) Then ' ; Περ 2,αν ειναι να παρει μπουκαλα απο τις κεφαλες-ζυγαρια για τη θεση της

 If (var(44) = 0) Then
    var(43) = 0
    Exit Sub
 End If ' Αν οι κεφαλες ειναι αδειες πανε ισια

 If (var(40) >= var(79) And var(40) <= var(74)) Then
     var(43) = 0
     Exit Sub
 End If ' Αν εισαι μπροστα στις κεφαλες πανε ισια

 If (var(41) < var(62)) Then
     var(43) = 1
     Exit Sub
 End If  ';Αν εισαι κοντα κανε μονο κλικ

 var(43) = 1
   Exit Sub

End If

If (var(35) = 0 And var(10) = 1) Then ' ; Περ 3,αν ειναι να παει μπουκαλα που πηρε απο τις κεφαλες-ζυγαρια για τη θεση της

 If (var(44) = 0) Then
     var(43) = 0
     Exit Sub
 End If ';Αν οι κεφαλες ειναι αδειες πανε ισια

 If (var(80) <= var(56) And var(56) < var(78)) Then
    var(43) = 0
    Exit Sub
 End If
 
 If (var(80) >= var(56) And var(56) > var(78)) Then
     var(43) = 0
     Exit Sub
 End If

 If (var(81) >= var(62)) Then
     var(43) = 1
     RobSend ("!gowhen(2pm>=" + Str(var(62)) + "):")
     Exit Sub
 End If ';Αν ειναι να πας μακρυα πανε ισια
 
 
 If (var(81) < var(62)) Then
     var(43) = 2
     Exit Sub
 End If  ';Αν εισαι κοντα κανε διπλο κλικ

 var(43) = 2
 Exit Sub
End If

If (var(35) = 1 And var(10) = 1) Then ' ; Περ 4,αν ειναι να παει μπουκαλα στις κεφαλες-ζυγαρια
  If (var(44) = 0) Then
      var(43) = 0
      Exit Sub
  End If ';Αν οι κεφαλες ειναι αδειες πανε ισια
  
  If (var(40) < var(56)) Then
    If (var(22) = 0) Then
       var(43) = 0
       Exit Sub
    End If
  End If
  
  If (var(40) > var(56)) Then
    If var(21) = 0 Then
       var(43) = 0
       Exit Sub
    End If
  End If
  var(45) = 2 * var(58)
  var(45) = var(45) + var(62) ';ειναι απεναντι αλλα τουλαχιστον 2 βηματα μακρυα
  
  If (var(41) > var(45) And var(40) > var(79) And var(40) < var(74)) Then
      var(43) = 0
      Exit Sub
  End If  ' ;Αν ερχεσαι απο ακριβως απεναντι ελα ισια
  
  If (var(81) >= var(62)) Then
      var(43) = 1
      Exit Sub
  End If ';Αν ερχεσαι απο μακρυα κανε κλικ διπλο-μονο
  
  If (var(41) < var(62)) Then
      var(43) = 2
      Exit Sub
  End If  ';Αν εισαι κοντα κανε μονο κλικ
  var(43) = 2
End If

End Sub

