VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   1620
   ClientTop       =   1530
   ClientWidth     =   8145
   Height          =   6345
   Left            =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8145
   Top             =   1185
   Width           =   8265
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1485
      TabIndex        =   10
      Top             =   3840
      Width           =   2010
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Coordinates of machine only"
      Height          =   270
      Left            =   500
      TabIndex        =   9
      Top             =   3390
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Data Files(Recipies/Solutions) only"
      Height          =   270
      Left            =   500
      TabIndex        =   8
      Top             =   3015
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Only the files to service(email)"
      Height          =   255
      Left            =   500
      TabIndex        =   6
      Top             =   975
      Width           =   3000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      Height          =   975
      Left            =   4725
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eπισκευή Βάσης δεδομένων"
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESTORE"
      Height          =   615
      Left            =   500
      TabIndex        =   2
      Top             =   2280
      Width           =   3000
   End
   Begin VB.FileListBox File1 
      Height          =   1035
      Left            =   4680
      Pattern         =   "*.zip"
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACKUP"
      Height          =   495
      Left            =   500
      TabIndex        =   0
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      Height          =   255
      Left            =   540
      TabIndex        =   11
      Top             =   3855
      Width           =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver. 2.01"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Option Explicit
Private Sub Check_directory()
' δημιουργεί το directory \backup αν δεν υπάρχει
Dim mdir
On Error GoTo create
  mdir = Dir(Text1.Text + "\*.zip")
If mdir = "" Then
  MkDir Text1.Text
End If
Exit Sub

create:
  MkDir Text1.Text

End Sub

Function ACCESS_TO_ACCESS(file, MDB1, MDB2)
'ΜΕΤΑΦΟΡΑ ΑΡΧΕΙΟΥ FILE ΑΠΟ MDB1 ΣΕ MDB2 DATABASE
Dim Errors, mydb
'On Error GoTo ERR
Errors = 0
    Set mydb = Workspaces(0).OpenDatabase(MDB2)
    mydb.Execute "drop table " + file



    Set mydb = Workspaces(0).OpenDatabase(MDB1)
    mydb.Execute "SELECT " + file + ".* INTO " + file + " in '" + MDB2 + "'  from " + file, dbFailOnError
    ACCESS_TO_ACCESS = Errors
    Exit Function
ERR:
  Errors = 1
  Resume Next
  MsgBox "Problem occurer during the transfer " + file
End Function


Private Sub Command1_Click()
Dim mday, a, k, mystamp, ans

'Check_directory
mday = LTrim(Str(Day(Date)))

On Error Resume Next
mystamp = FileDateTime(Text1.Text + "\" + mday + ".zip") ' Returns "2/12/93 4:35:47 PM".

If Check1 Then
   MsgBox "Attention ! This is not a normal backup." + Chr(13) + "It will create the file service.zip"
   a = Shell("\TALOS\pkzip.exe  " + Text1.Text + "\service   \talos\recipies.mdb", 1)
Else

    If DateDiff("d", mystamp, Date) > 0 Then
       a = Shell("\TALOS\pkzip.exe  " + Text1.Text + "\" + mday + " \talos\*.mdb \talos\*.dbf", 1)
    Else
        ans = MsgBox("Allready exists.Overwrite the previous backup?", vbYesNo)
        If ans = vbYes Then
           a = Shell("\TALOS\pkzip.exe  " + Text1.Text + "\" + mday + " \talos\*.mdb \talos\*.dbf", 1)
        End If
    End If

End If

File1.Visible = False
For k = 1 To 1000
   DoEvents
Next

'File1.Path = "c:\backup"

File1.Refresh
File1.Visible = True
End Sub


Private Sub Command2_Click()
Dim ans, myvalue, mdir, mystamp, a, Recipies_stamp
If UCase(Right(File1.filename, 3)) <> "ZIP" Then
      MsgBox "You must click the file from which I will restore."   'Δεν δείξατε από ποιό αρχείο θα γίνει η επαναφορά"
      Exit Sub
End If


If UCase(Left(File1.filename, 3)) = "SER" Then
      MsgBox "This file is for service. The RESTORE procedure will stop."   'Δεν δείξατε από ποιό αρχείο θα γίνει η επαναφορά"
      Exit Sub
End If





ans = MsgBox("The current files will be replaced with BACKUP files." + Chr(13) + "Are you sure?", vbYesNo + vbDefaultButton2)

If ans = vbNo Then
   Exit Sub
End If

myvalue = InputBox("Give the password 1234567890", , Space(10))

If myvalue <> "1234567890" Then Exit Sub
ChDir "\talos"

If CurDir = "C:\TALOS" Then
   
   
   
   
   'αντιγραφει το σημερινό recipies.mdb στο    "c:\"
    mdir = Dir("c:\rec.mdb")
    If mdir <> "" Then
       mystamp = FileDateTime("\rec.mdb") ' Returns "2/12/93 4:35:47 PM".
    Else
       mystamp = Date - 1
    End If
   'εάν αυτό που είναι στο root ειναι παλιό πάρε το καινούριο
   If DateDiff("d", mystamp, Date) > 0 Then
         FileCopy "recipies.mdb", "\rec.mdb"
   End If
   
   ' Recipies_stamp = FileDateTime("recipies.mdb")
    
   If Check3 Then
       a = Shell("\TALOS\pkUNzip.exe  -o " + Text1.Text + "\" + File1.filename + " Coord_tb.mdb", 1)
   ElseIf Check2 Then
       a = Shell("\TALOS\pkUNzip.exe  -o " + Text1.Text + "\" + File1.filename + " -xCoord_tb.mdb", 1)
   Else
     a = Shell("\TALOS\pkUNzip.exe  -o " + Text1.Text + "\" + File1.filename, 1)
   End If
     
     If Recipies_stamp = FileDateTime("recipies.mdb") Then
        MsgBox "The restore procedure was not completed." + Chr(13) + "!!!!"
     Else
        MsgBox "The restore procedure was successful."
     End If
    
     ans = MsgBox("Would you like to Try to save the file with today solutions ? ", vbYesNo, "Keep the solutions,otherwise make all the today solutions.")
   
If ans = vbYes Then
     a = ACCESS_TO_ACCESS("ximitech", "c:\rec.mdb", "c:\talos\recipies.mdb")
     
     If a = 0 Then
         MsgBox "The transfer completed."
     Else
         a = Shell("\TALOS\pkUNzip.exe  -o " + Text1.Text + "\" + File1.filename + " recipies.mdb", 1)
        MsgBox "The file with solutions was not saved."
        MsgBox "The today solutions are not in the memory. Take care of it."
     End If
End If
   
End If

End Sub


Private Sub Command3_Click()
On Error Resume Next
   Kill "c:\talos\rec.mdb"
On Error GoTo 0
   DBEngine.CompactDatabase "c:\talos\recipies.mdb", "c:\talos\rec.mdb", dbLangGreek
   FileCopy "rec.mdb", "recies.mdb"    ' Copy source to target.
   'name  rec.mdb as recipies.mdb
End Sub

Private Sub Command4_Click()
Dim a
    a = ACCESS_TO_ACCESS("ximitech", "c:\rec.mdb", "c:\talos\recipies.mdb")
    
    If a = 0 Then
       MsgBox "Ολοκληρώθηκε με επιτυχία"
    Else
       MsgBox "Δεν ολοκληρώθηκε η μεταφορά."
    End If
    
End Sub

Private Sub Command5_Click()
Dim b, c, a
    b = ACCESS_TO_ACCESS("prospau2", "c:\rec.mdb", "c:\talos\recipies.mdb")
    c = ACCESS_TO_ACCESS("joblist", "c:\rec.mdb", "c:\talos\recipies.mdb")
    a = b + c
    If a = 0 Then
       MsgBox "Ολοκληρώθηκε με επιτυχία"
    Else
       MsgBox "Δεν ολοκληρώθηκε η μεταφορά."
    End If

End Sub

Private Sub Command6_Click()
   Unload Me
End Sub

Private Sub File1_Click()
  Label1.Caption = FileDateTime(File1.Path + "\" + File1.filename)

End Sub

Private Sub File1_PathChange()
'On Error Resume Next
 ' Label1.Caption = File1.filename + " "
'  If File1.filename <> "" Then
  
'  Form1.Caption = (File1.Path + "\" + File1.filename)
 ' End If
End Sub


Private Sub Form_Load()
' Check_directory
' δημιουργεί το directory \backup αν δεν υπάρχει
Dim mdir


'ελεγχω αν υπάρχει το file backup.dat
 mdir = Dir("c:\backup.dat")
 If mdir = "" Then
    Text1.Text = "c:\backup"
    savefile
 End If
openfile


On Error GoTo create
  mdir = Dir(Text1.Text, vbDirectory)
If mdir = "" Then
  MkDir Text1.Text
End If

File1.Path = Text1.Text
Check1.BackColor = Form1.BackColor



Exit Sub

create:
  MkDir Text1.Text





 File1.Path = Text1.Text
End Sub



Public Sub savefile()
Dim file
file = FreeFile
Open "c:\backup.dat" For Output As #file
Print #file, Text1.Text
Close #file



End Sub

Public Sub openfile()
Dim a$, file
file = FreeFile
Open "c:\backup.dat" For Input As #file
Input #file, a$
Close #file
Text1.Text = a$
End Sub

Private Sub Text1_LostFocus()
savefile



End Sub


