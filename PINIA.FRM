VERSION 4.00
Begin VB.Form pinia 
   BackColor       =   &H00FF8080&
   Caption         =   "Γιάννης : Τήλιγμα Πηνίων"
   ClientHeight    =   5040
   ClientLeft      =   1545
   ClientTop       =   1020
   ClientWidth     =   6990
   Height          =   5445
   Icon            =   "PINIA.frx":0000
   Left            =   1485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6990
   Top             =   675
   Width           =   7110
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2565
      TabIndex        =   56
      Text            =   "1"
      Top             =   2175
      Width           =   390
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Tαχύτητα Joystick"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   765
      TabIndex        =   55
      Top             =   2175
      Width           =   1770
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   53
      Top             =   2250
      Width           =   1000
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   6600
      TabIndex        =   52
      Text            =   "2"
      Top             =   4650
      Width           =   360
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Φαρδύτερα"
      Height          =   330
      Left            =   4800
      TabIndex        =   51
      Top             =   4650
      Width           =   1110
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Στενότερα"
      Height          =   330
      Left            =   3600
      TabIndex        =   50
      Top             =   4650
      Width           =   1155
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6600
      TabIndex        =   49
      Text            =   "3"
      Top             =   4200
      Width           =   390
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Αλλαγή Επιτάχ./Επιβράδ."
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3615
      TabIndex        =   48
      Top             =   4200
      Width           =   2520
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6165
      TabIndex        =   47
      Text            =   "3"
      Top             =   4200
      Width           =   390
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Εικονική Λειτουργία"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1560
      TabIndex        =   46
      Top             =   1050
      Value           =   1  'Checked
      Width           =   1920
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5505
      TabIndex        =   44
      Text            =   "3"
      Top             =   3795
      Width           =   390
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Αλλαγή Ταχύτητας"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3660
      TabIndex        =   43
      Top             =   3795
      Width           =   1875
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5850
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   2820
      Width           =   1000
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   39
      Top             =   870
      Width           =   1000
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   37
      Text            =   "44.8"
      Top             =   585
      Width           =   1000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Υπολογισμός Πηνίου"
      Height          =   315
      Left            =   1605
      TabIndex        =   36
      Top             =   705
      Width           =   1900
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5850
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "4000"
      Top             =   1665
      Width           =   1000
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5850
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2535
      Width           =   1000
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   29
      Text            =   "255"
      Top             =   1965
      Width           =   1000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Σταμάτα"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5250
      TabIndex        =   28
      Top             =   3150
      Width           =   1605
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ξεκίνα"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3630
      TabIndex        =   27
      Top             =   3135
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Πάρε Τελική Θέση"
      Height          =   300
      Left            =   1605
      TabIndex        =   26
      Top             =   405
      Width           =   1900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Πάρε Αρχική Θέση"
      Height          =   300
      Left            =   1605
      TabIndex        =   25
      Top             =   105
      Width           =   1900
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   24
      Text            =   "30"
      Top             =   1395
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   23
      Text            =   "16"
      Top             =   1155
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   22
      Text            =   "-88745"
      Top             =   330
      Width           =   1000
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   7
      Left            =   2295
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4305
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   6
      Left            =   1080
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4305
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   1
      Left            =   2745
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":0CC6
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   15
      Top             =   2745
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   0
      Left            =   705
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":1108
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   14
      Top             =   2715
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   4
      Left            =   120
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":154A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4290
      Width           =   510
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   2
      Left            =   1695
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":198C
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3705
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   5
      Left            =   105
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":1DCE
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1545
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   3
      Left            =   1710
      MousePointer    =   10  'Up Arrow
      Picture         =   "PINIA.frx":2210
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1530
      Width           =   525
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   150
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5850
      TabIndex        =   0
      Text            =   "-101254"
      Top             =   75
      Width           =   1000
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Σπείρες"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5925
      TabIndex        =   54
      Top             =   4650
      Width           =   660
   End
   Begin VB.Label text12 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   3615
      TabIndex        =   45
      Top             =   3495
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Συν. βήματα Πηνίου "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   42
      Top             =   2850
      Width           =   2205
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Βήματα Χ "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   40
      Top             =   885
      Width           =   2205
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Μήκος Πηνίου (mm) "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   38
      Top             =   615
      Width           =   2205
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Σπείρες ανα στρώση (W) "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   35
      Top             =   1980
      Width           =   2205
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Βήματα ανα Στρώση "
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3615
      TabIndex        =   34
      Top             =   2265
      Width           =   2205
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Συν. Σπείρες Πηνίου "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   33
      Top             =   2565
      Width           =   2205
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Βήματα ανα Σπείρα"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   32
      Top             =   1695
      Width           =   2205
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Αρχ. Θέση άξονα 1 "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   21
      Top             =   105
      Width           =   2205
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Τελ. Θέση άξονα 1 "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   20
      Top             =   360
      Width           =   2205
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Πάχος Σύρματος (mm) "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   19
      Top             =   1425
      Width           =   2205
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Αριθ. Στρώσεων  Πηνίου "
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3615
      TabIndex        =   18
      Top             =   1170
      Width           =   2205
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   1635
      X2              =   2250
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   360
      X2              =   360
      Y1              =   2085
      Y2              =   4185
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   1305
      X2              =   2700
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   1950
      X2              =   1935
      Y1              =   2145
      Y2              =   3675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   3525
      Left            =   75
      Top             =   1455
      Width           =   3345
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis Position"
      ForeColor       =   &H00FFFF00&
      Height          =   465
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   795
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   900
      Width           =   795
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   1125
      Width           =   795
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   675
      Width           =   795
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   450
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 3"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   900
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 4"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   1125
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 2"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   660
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 1"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   645
   End
End
Attribute VB_Name = "pinia"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Sub endis(arg As Boolean)
Text1.Enabled = arg
'Text2.Enabled = arg
Text4.Enabled = arg
Text5.Enabled = arg
Text3.Enabled = arg
Command1.Enabled = arg
Command2.Enabled = arg
Command5.Enabled = arg
Command3.Enabled = arg
End Sub

Sub katis(arg)
For nt = 1 To arg: DoEvents: Next nt
End Sub
Public Sub RobSend(cmd$)
      On Error Resume Next
       temp% = SendAT6400Block(768, cmd$, 0)
End Sub

Private Sub Check1_Click()
If Check1 = 1 Then
      Check1.Caption = "Εικονική Λειτουργία "
      Text6 = 1
      Text5 = 30
      Command5_Click
    Else
      Check1.Caption = "Κανονική Λειτουργία"
      Text6 = 255
      Text5 = 0.178
      Command5_Click
End If
End Sub

Private Sub Command1_Click()
Text1 = position(0)
Text3 = Text1 - 44.8 * 213
Command5_Click
End Sub

Private Sub Command10_Click()
If Text16 = "" Or Val(Text16) = 0 Or Text2 = "" Then Exit Sub
asw = Text1
Text1 = asw + 213 * Text16
Command5_Click
End Sub

Private Sub Command11_Click()
If Text16 = "" Or Val(Text16) = 0 Or Text2 = "" Then Exit Sub
asw = Text6
Text6 = asw + Val(Text16)
asw = Text2
Text2 = asw + 4000 * Text16
'Command5_Click

End Sub

Private Sub Command2_Click()
Text3 = position(0)
Command5_Click
End Sub

Private Sub Command3_Click()
endis False
Text12 = "": Text12.Width = 0
Command4.Enabled = True
Command5_Click
Me.Refresh
cmd$ = "!ma1111:"
RobSend (cmd$)
katis 1000
cmd$ = "!d" + Text1 + ":!go1:"
RobSend (cmd$)
robwait Val(position(0)), Val(Text1)
Text12.Visible = True
If Text10 = "0" Then GoTo ex_t
cur_x = position(0)
For g = 1 To Val(Text4)
If System_Ready = 0 Then katis 5000
If Command4.Enabled = False Then Exit Sub
If g / 2 = Int(g / 2) Then metatop = -Text10 Else metatop = Text10
 If Check1 = 1 Then
    If g / 2 = Int(g / 2) Then axis4 = -Val(Text2) Else axis4 = Val(Text2)
Else
    axis4 = Val(Text2)
End If
If Val(Text13) = 0 Or Val(Text14) = 0 Or Val(Text15) = 0 Then
    MsgBox "Δηλώστε ταχύτητα λειτουργίας , Επιτάχυνση , Επιβράδυνση ...", , "Γιάννης"
    Command3.Enabled = True
    Text13.SetFocus
    Exit Sub
End If
cmd$ = "!pv" + Trim(Text13) + ":!pa" + Trim(Text14) + ":!pad" + Trim(Text15) + ":"

RobSend (cmd$)

cmd$ = "!var1=" + Str(metatop) + ":!var2=" + Str(-Abs(axis4)) + ":"
RobSend (cmd$)
cmd$ = "!pcomp pinia:"
RobSend (cmd$)
katis 500
cmd$ = "!prun pinia:"
RobSend (cmd$)
DoEvents
 If g / 2 = Int(g / 2) Then Target = Val(cur_x) Else Target = Val(cur_x) + Val(metatop)
robwait Val(position(0)), Val(Target)
Text12.Width = 3225 * g / Val(Text4)
If Text12.Width > 500 Then Text12 = Str(Int(g / Val(Text4) * 100)) + "%"
Text12.Refresh
Next g
ex_t:
Text12 = "100%"
endis True
End Sub

Private Sub Command4_Click()
cmd$ = "!k:"
RobSend (cmd$)
g = 0
endis True
End Sub

Private Sub Command5_Click()
On Error GoTo er_sc
If Text9 <> "" Then Text3 = Text1 - Text9 * 213
Text10 = Int(Text3 - Text1)
If Text9 = "" Then Text9 = Text10 / 213
If Text6 = "" Then Text6 = Int(Text9 / Val(Text5))
Text2 = Int(Text6 * Text8)
Text7 = Int(Text4 * Text6)
Text11 = Int(Text2 * Text7)
Exit Sub
er_sc:
MsgBox "Συμπλήρωσε ότι λείπει ...", , "Γιάννης"
Resume Next
End Sub

Sub robwait(CurrentPosition As Variant, Target As Variant)
  If System_Ready = 0 Then
     Exit Sub
  End If
  Do
     DoEvents
   Loop Until Val(position(0)) = Val(Target)
End Sub
Sub sendveloc(arg)
'tax = 1,  epit = 3,  epib = 5
cmd$ = "!var3=" + Str(tax * Val(arg)) + ":!var4=" + Str(epit * Val(arg)) + ":!var5=" + Str(epib * Val(arg)) + ":"
RobSend (cmd$)
cmd$ = "!v" + Str(tax_j * Val(arg)) + "," + Str(tax_j * Val(arg)) + "," + Str(tax_j * Val(arg)) + "," + Str(tax_jw * Val(arg)) + ":"
RobSend (cmd$)
cmd$ = "!a" + Str(tax_j * Val(arg)) + "," + Str(tax_j * Val(arg)) + "," + Str(tax_j * Val(arg)) + "," + Str(tax_jw * Val(arg)) + ":"
RobSend (cmd$)
cmd$ = "!ad" + Str(tax_j * Val(arg)) + "," + Str(tax_j * Val(arg)) + "," + Str(tax_j * Val(arg)) + "," + Str(tax_jw * Val(arg)) + ":"
RobSend (cmd$)

End Sub

Private Sub Command6_Click()
If Text13 = "" Or Val(Text13) = 0 Then
  MsgBox "Βάλε σωστά την ταχύτητα" + Chr(13) + " πχ 1.5 ή 0.2 ή 5.1 κλπ", , "Γιάννης"
    Text13 = 1
    Text13.SetFocus
    Exit Sub
End If
cmd$ = "!pv" + Trim(Text13) + ":"
RobSend (cmd$)
End Sub

Private Sub Command7_Click()
On Error Resume Next
'Global Const epit = 0.3
'Global Const epib = 0.3
cmd$ = "!pa" + Trim(Text14) + ":!pad" + Trim(Text15) + ":"
RobSend (cmd$)
End Sub

Private Sub Command8_Click()
If Text16 = "" Or Val(Text16) = 0 Or Text2 = "" Then Exit Sub
asw = Text6
Text6 = asw - Val(Text16)
asw = Text2
Text2 = asw - 4000 * Text16
'Command5_Click

End Sub

Private Sub Command9_Click()
If Text17 = "" Or Val(Text17) = 0 Then
  MsgBox "Βάλε σωστά την ταχύτητα" + Chr(13) + " πχ 1.5 ή 0.2 ή 5.1 κλπ", , "Γιάννης"
    Text17 = 1
    Text17.SetFocus
    Exit Sub
End If
sendveloc (Text17)
End Sub

Private Sub Form_Load()
On Error GoTo er_ef
epib = 0.3: epit = 0.3
findfile = Dir("c:\talos\pinia.prg")
If findfile = "" Then
    ans = MsgBox("Το αρχείο κινήσεων του Robot  ( pinia.prg ) δεν υπάρχει εγκατεστημένο στον κατάλογο c:\talos , οπότε  το πρόγραμμα δεν θα λειτουργήσει ..." + Chr(13) + "Θέλετε να συνεχίσετε ;", 20, "Γιάννης : Τηλ 37540 , 36312 (εσ. 400 ή 401)")
    If ans = 7 Then End
End If
System_Ready = IsOSLoaded(768)
If System_Ready = 1 Then
cmd$ = "!k:"
For i = 1 To 1000: DoEvents: Next i
RobSend (cmd$)
tmp% = SendAT6400File(0, 768, 0, cmd$, 0)
If tmp% = -3 Then MsgBox "Το αρχείο c:\Talos\ pinia.prg δεν φορτώθηκε στην κάρτα . " + Chr(13) + "Γυρίστε στο Terminal και επιλέξτε :" + Chr(13) + "Trasfers ---> Send Motion Programs ----> c:\Talos\pinia.prg", , "Γιάννης"
Else
    ans = MsgBox("Το λειτουργικό σύστημα  του Robot δεν υπάρχει . Ξεκινείστε το  Robot από την αρχή." + Chr(13) + "Θέλετε να συνεχίσετε ;", 20, "Γιάννης : Τηλ 37540 , 36312 (εσ. 400 ή 401)")
    If ans = 7 Then End
End If
sendveloc 1
Exit Sub
er_ef:
MsgBox Err.Description, , "Γιάννης"
Resume Next
End Sub

Private Sub Grip_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub term_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
 Case 0
     cmd$ = "!d" + Str(-5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go1:"
      RobSend (cmd$)
Case 1
     cmd$ = "!d" + Str(5000000) + ":"
       RobSend (cmd$)
       cmd$ = "!go1:"
       RobSend (cmd$)
Case 2
    cmd$ = "!d," + Str(-5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
Case 3
    cmd$ = "!d," + Str(5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
Case 4
    cmd$ = "!d,," + Str(-5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,,1:"
      RobSend (cmd$)
Case 5
     cmd$ = "!d,," + Str(5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,,1:"
      RobSend (cmd$)
Case 6
    cmd$ = "!d,,," + Str(-5000000) + ":"
    RobSend (cmd$)
    cmd$ = "!go,,,1:"
    RobSend (cmd$)
Case 7
    cmd$ = "!d,,," + Str(5000000) + ":"
    RobSend (cmd$)
    cmd$ = "!go,,,1:"
    RobSend (cmd$)
End Select
   If term(Index).BackColor = Green Then term(Index).BackColor = orange
End Sub
Private Sub term_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If term(Index).BackColor = Red Then
    Else
        term(Index).BackColor = Green
    End If
    cmd$ = "!s:"
   RobSend (cmd$)
End Sub

Private Sub Timer1_Timer()
If System_Ready = 0 Then
    For i = 0 To 3
    position(i) = "-1"
    Next i
    Exit Sub
End If
   Result% = Request_Status(768)
    ' ******* Update Motor Position ****************
    Set_Pointer 768, AXIS1_MOTOR
    For i = 0 To 3
        Read_Status 768, status_high%, status_low%, status&
        position(i).Caption = status&
    Next i

End Sub
