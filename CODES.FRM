VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   2805
   ClientTop       =   1590
   ClientWidth     =   6690
   Height          =   6345
   Left            =   2745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1245
   Width           =   6810
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\TALOS\RECIPIES.MDB"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM RECIP"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CODES.frx":0000
      Height          =   1815
      Left            =   840
      OleObjectBlob   =   "CODES.frx":000E
      TabIndex        =   1
      Top             =   3240
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Check_passw()
' GENERATE RECORDS
Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\RECIPIES.mdb")
             Set r = dbsJet.OpenRecordset("recip")
             
             For K = 1 To 100
             r.AddNew
               If K < 10 Then
                   r("HME") = Now
               End If
               r("K") = 5 * K ^ 1 + 6 * K ^ 2 + 7 * K ^ 3 + 8 * K ^ 4
             r.Update
             Next
           r.Close
             
End Sub


Private Sub Command2_Click()
Dim CODE_NEXT
'амоицы то аявеио
Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\RECIPIES.mdb")
             Set r = dbsJet.OpenRecordset("RECIP")
             
        Do While Not r.EOF
              
              If IsNull(r("HME")) Then  'паы сто текеутаио ле глеяолгмиа
                 CODE_NEXT = r("K") 'PIANO TON KODIKO TOY EPOMENOY RECORD
                 r.MovePrevious 'GYRIZO PISO GIA NA PIASO THN HMEROMHNIA THS TELEYTAIAS XRHSHS
                 Exit Do
               End If
               r.MoveNext
        Loop
       ' Form1.Caption = CODE_NEXT 'R("K")
'        R.Close
        
        
        
        
        If Month(Now) <> Month(r("HME")) Then
            'фгтаеи йыдийо
             w = InputBox("Give The Password please.")
             If CODE_NEXT = w Then 'йыдийо поу пкгйтяокоцгсе о вягстгс
                r.MoveNext
                r.Edit
                r("HME") = Now
                r.Update
             Else
                 MsgBox "Sorry , no good Password"
                 
                 End
             End If
         Else
           'CONTINUE
         End If
         
         
         
        

End Sub


