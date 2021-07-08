VERSION 4.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5940
   ClientLeft      =   2715
   ClientTop       =   1785
   ClientWidth     =   6690
   ClipControls    =   0   'False
   Height          =   6345
   Left            =   2655
   LinkTopic       =   "Form2"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1440
   Width           =   6810
   Begin VB.Label Label19 
      BackColor       =   &H000080FF&
      Caption         =   $"NEA_SKON.frx":0000
      Height          =   450
      Left            =   3960
      TabIndex        =   18
      Top             =   7800
      Width           =   5085
   End
   Begin VB.Label Label18 
      BackColor       =   &H00404080&
      Caption         =   $"NEA_SKON.frx":0057
      Height          =   705
      Left            =   1425
      TabIndex        =   17
      Top             =   7545
      Width           =   2475
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FF00FF&
      Caption         =   "Παίρνει 3ο ποτήρι ΑΠΟ ζυγο ΣΕ ΑΝΑΔΕΥΣΗ ΚΑΙ .."
      Height          =   270
      Left            =   1425
      TabIndex        =   16
      Top             =   7260
      Width           =   7665
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Caption         =   "Μεταφέρω το ποτήρι 2 απο  ανάδευση στην θέση του //ΤΕΛΟΣ"
      Height          =   315
      Left            =   1425
      TabIndex        =   15
      Top             =   6885
      Width           =   7650
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF00FF&
      Caption         =   "Παίρνει 3ο ποτήρι σε ζυγο,ρίχνει σκόνη,ρίχνει νερό πάστας "
      Height          =   600
      Left            =   3975
      TabIndex        =   14
      Top             =   6225
      Width           =   5115
   End
   Begin VB.Label Label14 
      BackColor       =   &H00404080&
      Caption         =   $"NEA_SKON.frx":00A5
      Height          =   705
      Left            =   1440
      TabIndex        =   13
      Top             =   6135
      Width           =   2475
   End
   Begin VB.Label Label13 
      BackColor       =   &H000080FF&
      Caption         =   "Μεταφέρω το ποτήρι 2 απο ζυγό σε ανάδευση "
      Height          =   315
      Left            =   1455
      TabIndex        =   12
      Top             =   5790
      Width           =   7635
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF00FF&
      Caption         =   "Μεταφέρω το ποτήρι 1 από ανάδευση στην θέση του//ΤΕΛΟΣ"
      Height          =   600
      Left            =   1455
      TabIndex        =   11
      Top             =   5100
      Width           =   2460
   End
   Begin VB.Label Label11 
      BackColor       =   &H000080FF&
      Caption         =   $"NEA_SKON.frx":00F3
      Height          =   915
      Left            =   3975
      TabIndex        =   10
      Top             =   4815
      Width           =   5130
   End
   Begin VB.Label Label10 
      BackColor       =   &H00404080&
      Caption         =   $"NEA_SKON.frx":013E
      Height          =   600
      Left            =   1440
      TabIndex        =   9
      Top             =   4485
      Width           =   2475
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF00FF&
      Caption         =   "Παίρνει 1ο ποτήρι από  ζυγο σε ανάδευση για τελευταία μίξη"
      Height          =   600
      Left            =   1440
      TabIndex        =   8
      Top             =   3825
      Width           =   7650
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Caption         =   "Μεταφέρω το ποτήρι 2 απο ανάδευση στην θέση του "
      Height          =   465
      Left            =   1440
      TabIndex        =   7
      Top             =   3255
      Width           =   7650
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "Μεταφέρω το ποτήρι 2 από την θέση του στην ανάδευση και ..."
      Height          =   465
      Left            =   3945
      TabIndex        =   6
      Top             =   2280
      Width           =   5130
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF00FF&
      Caption         =   "Μεταφέρω το ποτήρι 1 σε ζυγό και ρίχνω νερό μεχρι να γεμίσει"
      Height          =   1080
      Left            =   1425
      TabIndex        =   5
      Top             =   2115
      Width           =   2385
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404080&
      Caption         =   $"NEA_SKON.frx":018C
      Height          =   405
      Left            =   3975
      TabIndex        =   4
      Top             =   2775
      Width           =   5100
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Caption         =   "Σταματά το  ανεξάρτητο πρόγραμμα ανάδευσης "
      Height          =   255
      Left            =   1395
      TabIndex        =   3
      Top             =   1740
      Width           =   7635
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   $"NEA_SKON.frx":01DA
      Height          =   390
      Left            =   1410
      TabIndex        =   2
      Top             =   1245
      Width           =   5085
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   $"NEA_SKON.frx":0231
      Height          =   705
      Left            =   6540
      TabIndex        =   1
      Top             =   945
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
      Caption         =   $"NEA_SKON.frx":027F
      Height          =   600
      Left            =   1380
      TabIndex        =   0
      Top             =   300
      Width           =   7620
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
