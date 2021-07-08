VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   3255
   ClientTop       =   1590
   ClientWidth     =   6690
   Height          =   6345
   Left            =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1245
   Width           =   6810
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM XIMITECH"
      Top             =   5340
      Width           =   1395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1665
      Left            =   2010
      TabIndex        =   3
      Top             =   3615
      Width           =   2805
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   420
      Left            =   2880
      TabIndex        =   2
      Top             =   525
      Width           =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   1950
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1515
      Left            =   2445
      TabIndex        =   0
      Top             =   1665
      Width           =   3540
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2145
      Top             =   3810
      _extentx        =   741
      _extenty        =   741
      _stockprops     =   0
      reportfilename  =   "c:\talos\vis4-99.rpt"
      destination     =   0
      windowleft      =   100
      windowtop       =   100
      windowwidth     =   490
      windowheight    =   300
      windowtitle     =   ""
      windowborderstyle=   2
      windowcontrolbox=   -1  'True
      windowmaxbutton =   -1  'True
      windowminbutton =   -1  'True
      copiestoprinter =   1
      printfilename   =   ""
      printfiletype   =   0
      selectionformula=   ""
      groupselectionformula=   ""
      connect         =   ""
      username        =   ""
      reportsource    =   0
      boundreportheading=   ""
      boundreportfooter=   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error Resume Next
'   Kill "C:\TALOS\dokpros2.dbf"
Kill "C:\TALOS\dokprosp.dbf"

On Error GoTo 0
p_katal = "2"
Set MyDB = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
MyDB.Execute "SELECT prospau" + p_katal + ".*,XIMITECH.PERIGRAFH INTO DOKPROSp from prospau" + p_katal + " INNER JOIN XIMITECH ON prospau" + p_katal + ".KOD=XIMITECH.KOD where entolh='000001'"

 
'MyDB.Execute "SELECT dokpros2.*,joblist.* INTO DOKPROSp from dokpros2 INNER JOIN joblist ON dokpros2.entolh=joblist" + j_katal + ".entolh"
'Job.Index = "ENTOLH"
MyDB.Close
CrystalReport1.Action = 1
End Sub

Private Sub Command2_Click()
EPANEKTYP.Show
End Sub

Private Sub Command3_Click()

Set MyDB = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
'MyDB.Execute "SELECT pelates.* INTO pelates in 'c:\talos\recipies.mdb'  from pelates"
'MyDB.Execute "SELECT ximitech.* INTO ximitech in 'c:\talos\recipies.mdb'  from ximitech"
'MyDB.Execute "SELECT joblist.* INTO joblist in 'c:\talos\recipies.mdb'  from joblist"
'MyDB.Execute "SELECT prospau2.* INTO prospau2 in 'c:\talos\recipies.mdb'  from prospau2"

'MyDB.Execute "SELECT TOP 1 XIMITECH.* INTO DOKX in 'c:\talos\recipies.mdb'  from XIMITECH IN 'c:\talos\recipies.mdb' "


'MyDB.Execute "SELECT joblist3.* INTO joblist3 in 'c:\talos\recipies.mdb'  from joblist3"
'MyDB.Execute "SELECT prospau3.* INTO prospau3 in 'c:\talos\recipies.mdb'  from prospau3"

'MyDB.Execute "SELECT pania.* INTO pania in 'c:\talos\recipies.mdb'  from pania"
'MyDB.Execute "SELECT eidh_baf.* INTO eidh_baf in 'c:\talos\recipies.mdb'  from eidh_baf"
'MyDB.Execute "SELECT xhmedit.* INTO xhmedit in 'c:\talos\recipies.mdb'  from xhmedit"
 'MyDB.Execute "SELECT NERA.* INTO NERA in 'c:\talos\recipies.mdb'  from NERA"
' MyDB.Execute "SELECT EIDDIAL.* INTO EIDDIAL in 'c:\talos\recipies.mdb'  from EIDDIAL"

''''''''''Set MyDB = Workspaces(0).OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
''''''''''MyDB.Execute "DROP TABLE DOKXHM ;"


End Sub

Private Sub Command4_Click()
 Data1.RecordSource = "SELECT *FROM XIMITECH WHERE Val(KOD)=1;"
           Data1.Refresh
End Sub
