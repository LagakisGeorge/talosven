VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   1290
   ClientTop       =   1785
   ClientWidth     =   7455
   Height          =   6345
   Left            =   1230
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7455
   Top             =   1440
   Width           =   7575
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   720
      Left            =   4200
      TabIndex        =   2
      Top             =   1365
      Width           =   1920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Μεταφορά αρχείoυ γλωσσών σε dbaseIV αρχείο"
      Height          =   705
      Left            =   4095
      TabIndex        =   1
      Top             =   3705
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Μεταφορά αρχείων από Dbase IV σε Access"
      Height          =   705
      Left            =   4095
      TabIndex        =   0
      Top             =   2835
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\recipies.mdb")
'mydb.Execute "drop table dokprosp"
'mydb.Execute "drop table xhmedit"

'mydb.Execute "drop table pelates"
'mydb.Execute "drop table ximitech"
'mydb.Execute "drop table joblist"
'mydb.Execute "drop table prospau2"
'mydb.Execute "drop table pania"
'mydb.Execute "drop table eidh_baf"

Set mydb = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")



mydb.Execute "SELECT pelates.* INTO pelates in 'c:\talos\recipies.mdb'  from pelates", dbFailOnError

End


Set mydb = Workspaces(0).OpenDatabase("c:\TALOSsos", False, False, "dBASE IV;")
'On Error Resume Next



mydb.Execute "SELECT pelates.* INTO pelates in 'c:\talos\recipies.mdb'  from pelates", dbFailOnError
mydb.Execute "SELECT ximitech.* INTO ximitech in 'c:\talos\recipies.mdb'  from ximitech", dbFailOnError
mydb.Execute "SELECT joblist.* INTO joblist in 'c:\talos\recipies.mdb'  from joblist", dbFailOnError
mydb.Execute "SELECT prospau2.* INTO prospau2 in 'c:\talos\recipies.mdb'  from prospau2", dbFailOnError
mydb.Execute "SELECT TOP 1 XIMITECH.* INTO DOKX in 'c:\talos\recipies.mdb'  from XIMITECH IN 'c:\talos\recipies.mdb' ", dbFailOnError
mydb.Execute "SELECT mpoyk.* INTO mpoyk in 'c:\talos\recipies.mdb'  from mpoyk", dbFailOnError

'mydb.Execute "SELECT joblist3.* INTO joblist3 in 'c:\talos\recipies.mdb'  from joblist3", dbFailOnError
'mydb.Execute "SELECT prospau3.* INTO prospau3 in 'c:\talos\recipies.mdb'  from prospau3", dbFailOnError

mydb.Execute "SELECT pania.* INTO pania in 'c:\talos\recipies.mdb'  from pania", dbFailOnError
mydb.Execute "SELECT eidh_baf.* INTO eidh_baf in 'c:\talos\recipies.mdb'  from eidh_baf", dbFailOnError
mydb.Execute "SELECT xhmedit.* INTO xhmedit in 'c:\talos\recipies.mdb'  from xhmedit", dbFailOnError
mydb.Execute "SELECT NERA.* INTO NERA in 'c:\talos\recipies.mdb'  from NERA", dbFailOnError
mydb.Execute "SELECT EIDDIAL.* INTO EIDDIAL in 'c:\talos\recipies.mdb'  from EIDDIAL", dbFailOnError

'αν δεν υπάρχει το joblist3.dbf,prospau3.dbf
      mydb.Execute "SELECT top 1 prospau2.* INTO prospau3 in 'c:\talos\recipies.mdb'  from prospau2", dbFailOnError
      mydb.Execute "SELECT top 1 joblist.* INTO joblist3 in 'c:\talos\recipies.mdb'  from joblist", dbFailOnError
'-------------------------------
'δεν δουλεύει σωστά mydb.Execute "delete  joblist3.* from joblist3 where sxesh_mpan>0"
'δεν δουλεύει σωστά mydb.Execute "delete *from prospau3"

End Sub

Private Sub Command2_Click()

Set mydb = Workspaces(0).OpenDatabase("c:\TALOS_se\coord_tb.mdb")

mydb.Execute "SELECT Language.* INTO Language in 'c:\talos' 'dBase IV;'  from Language", dbFailOnError
'mydb.Execute "SELECT Labels.* INTO Labels in 'c:\talos' 'dBase IV;'  from Labels", dbFailOnError


'Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\coord_tb.mdb")
'mydb.Execute "drop table language"
'mydb.Execute "drop table labels"



Set mydb = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")

mydb.Execute "SELECT Language.* INTO language in 'c:\talos\coord_tb.mdb'  from language", dbFailOnError
mydb.Execute "SELECT Labels.* INTO labels in 'c:\talos\coord_tb.mdb'  from labels", dbFailOnError



End Sub

Private Sub Command3_Click()
Set mydb = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
On Error Resume Next

 

End Sub

Private Sub Command4_Click()

' σβηνω όλες τις εγγραφές της γλώσσας
Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\coord_tb.mdb")
mydb.Execute "delete *from language  where id_label>=0"

Set mydb = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
mydb.Execute "insert into language  in 'c:\talos\coord_tb.mdb' select language.*   from language", dbFailOnError


'INSERT INTO Customers
'SELECT [New Customers].*
'FROM [New Customers];   Selects all records in a hypothetical New Customers table and adds them to the Customers table. (When individual columns are not designated, the SELECT table column

'Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\recipies.mdb")
'mydb.Execute "update ximitech set perigrafh=mperigrafh where addr_prot>=0", dbFailOnError

End Sub


