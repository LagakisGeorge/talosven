Attribute VB_Name = "Module1"
' ΜΟΝΟ ΓΙΑ ΤΑ ΧΡ
'ΓΙΑ ΤΑ WINDOWS 95-98   _SYNTAGE.BAS

' EINAI IΔIO ΜΕ ΤΟ ΜΟΝΟ ΠΟΥ ΔΙΑΒΑΖΕΙ ΑΛΛΗ ΒΙΒΛΙΟΘΗΚΗ  NT6400.DLL
' ΓΙΑ ΛΟΓΟΥΣ ΣΥΜΒΑΤΟΤΗΤΑΣ ΕΧΩ ΔΗΜΙΟΥΡΓΗΣΕΙ ΤΗΝ ΣΥΝΑΡΤΗΣΗ SET_DEVICE(X,Y)


'VB-ASM DLL declarations
Type MODEl
AIM As Integer
End Type

Type VBREGS
    AX As Integer
    BX As Integer
    CX As Integer
    DX As Integer
    BP As Integer
    SI As Integer
    DI As Integer
    Flags As Integer
    DS As Integer
    ES As Integer
End Type


Global mUndelete_SYNTAGH ' για να μην κάνει δύο φορές έλεγχο για μισοσβησμένη συνταγή
Global Menu
Global m_entolh
Global numbers_on_labels
Global aritmos_ximikon
Global Const ho = 372
Global Const m0 = 600

Global User_ID

Global m_Cmd$
Global Const ERGOSTASIO = "iris"

Global Xim1_on, xim1_Off
Global Xim2_on, xim2_Off
Global Xim3_on, xim3_Off
Global Xim4_on, xim4_Off
Global XimNer_on, XimNer_Off
Global XimNer

Global gXimikaDial
Global dejia, aristerh










     
     '***********************  I R I S *****************************
  
  
  Global Const Ar_Mpoykal = 300
  
  Global ForNexts_Milsec  ' ποσα for-next kanei se ena milisecond
  Global Ar_1_mpoykal
  Global Ar_tel_mpoykal
  Global Lim1
  Global Lim2
  Global DOYLEYO_ALATI, UESH_ALATIOY
  Global vis_talos
  Global x0, y0, z0, w0
  Global YperXeilish
  
 Global Const DemoVersion = 0
  Global Const Ar_Pothr = 30
  Global ar_real_pothr
  Global Const K_AR = 0   '7=iris  ' farbetex=4
  Global Const bit_ok = 14  ' farbetex=24
  Global Const Open_Valve = ""             '   farb = "out.3-0:"
  Global Const Close_Valve = ""             ' farb="out.3-1:"
  Global Const Porta_Ar = 17               '  farb="2"
  Global Const Porta_Dex = "18"     '   farb="1"
  Global Const Porta_Ner = "4"              ' farb="4"
  
  
  Global Const grip_o = 11    '       IRIS= 5       ' farb=11
  Global Const grip_c = 10   'IRIS= 4     'farb=12
  Global Const Serial_Zyg = 1
  Global Const coordin = 1    ' Αν ειναι 0 οι συντεταγμένες ειναι στανταρ ,αν ειναι 1 τις δινει ο χρηστης

     '*****************************************************************

     '*********************** FARBETEX  *****************************
'    Global Const Ar_Mpoykal = 260
'    Global Const K_AR = 4     ' farbetex=4
'    Global Const bit_ok = 24   ' farbetex=24
'    Global Const Open_Valve = "out.3-0:"
'    Global Const Close_Valve = "out.3-1:"
'    Global Const Porta_Ar = "2"
'    Global Const Porta_Dex = "1"
'    Global Const Porta_Ner = "4"
'    Global Const grip_o = 11
'    Global Const grip_c = 12
'    Global Const Serial_Zyg = 2
'    Global Const coordin = 1    ' Αν ειναι 0 οι συντεταγμένες ειναι στανταρ ,αν ειναι 1 τις δινει ο χρηστης
     '*****************************************************************

'Global xhm_c
'Global xhmedit_c
'Global joblist_c
'Global synmaker_c
'Global wait_c
'Global syntages_c
'Global fere_c
'Global feresk_c
'Global pane_c
'Global zygpani_c

Global dokimastikh_syntagh
Global allazo_xromata
Global allazo_fonts
Global prozygismena
Global Fornext_small, Fornext_big
Global Default_baros_panioy

Global MaxTimeWater

Global Avance_Online_Water

Global AXRHSIMOPOIHTES_ORES
Global xymeno As Single
Global baros As Single
Global xronos
Global valve_closed
Global gINDEX
Global kod6400
Global kodTimer
Global EPIU
Global Flag_Show
Global Flag_Kefal
Global porta As String
Global a As Single
Global Zyg

Global flag_paint


Global var(1 To 100)
Global X_arxh, Y_arxh, Z_arxh, Xt, Yt, Zt




Global Exo_entolh


Global cmd2$
Global Ano_Z_93
Global var65_1kef
Global var66_2kef
Global var67_3kef
Global var68_4kef


Global AA_SYNTAG
Global Const VBFLAGS_CARRY = &H1
Global Const VBFLAGS_PARITY = &H4
Global Const VBFLAGS_AUX = &H10
Global Const VBFLAGS_ZERO = &H40
Global Const VBFLAGS_SIGN = &H80

'vbGetDriveType return values
Global Const VBDRIVE_REMOVABLE = &H1002
Global Const VBDRIVE_FIXED = &H1003
Global Const VBDRIVE_REMOTE = &H1004
Global Const VBDRIVE_CDROM = &H1005
Global Const VBDRIVE_FLOPPY = &H1006
Global Const VBDRIVE_RAMDISK = &H1007
Global Const VBDRIVE_UNKNOWN = &H1010
Global Const VBDRIVE_INVALID = &H0
Global flag As Integer
Global gram As Single
'xp = True
#If Win32 Then
Declare Function SendAT6400Block% Lib "nt6400.dll" (ByVal ADDRESS%, ByVal cmd$, ByVal irqnum%)
Declare Function RecvAT6400Block% Lib "nt6400.dll" (ByVal ADDRESS%, ByVal resp$, ByVal irqnum%)
Declare Function OSLoad% Lib "nt6400.dll" Alias "osload" (ByVal ADDRESS%, ByVal options$, ByVal handle%)
Declare Function IsOSLoaded% Lib "nt6400.dll" (ByVal ADDRESS%)
Declare Function Request_Status% Lib "nt6400.dll" Alias "request_status" (ByVal ADDRESS%)
Declare Sub Set_Pointer Lib "nt6400.dll" Alias "set_pointer" (ByVal ADDRESS%, ByVal status_offset%)
Declare Sub Read_Status Lib "nt6400.dll" Alias "read_status" (ByVal ADDRESS%, status_high%, status_low%, status&)
Declare Function IsAT6400Ready% Lib "nt6400.dll" (ByVal ADDRESS%)
Declare Sub SetTimeout Lib "nt6400.dll" (ByVal timeout&)
Declare Sub Delay Lib "nt6400.dll" (ByVal timedelay&)
Declare Function SendAT6400File% Lib "nt6400.dll" (ByVal handle%, ByVal ADDRESS%, ByVal irqnum%, ByVal Filename$, ByVal options$)
Declare Function SetNTParam% Lib "nt6400.dll" (ByVal ADDRESS&, ByVal irqnum&)
'Declare Function GetFastStatus% Lib "nt6400.dll" (ByVal ADDRESS%, faststatus As AT6400INFO)
'Declare Function GetExFastStatus% Lib "nt6400.dll" (ByVal ADDRESS%, ByVal block7%, ByVal block8%, faststatus As AT6400INFO)

'
'32 BIT
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function PEPrintReport Lib "CRPE.DLL" (ByVal reportFilePath As String, ByVal toPrinter As Integer, ByVal toWindow As Integer, ByVal title As String, ByVal left As Integer, ByVal top As Integer, ByVal width As Integer, ByVal height As Integer, ByVal style As Long, ByVal parentWindow As Integer) As Integer
Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long

'Declare Function SendAT6400Block% Lib "c:\talos\wn956000.dll" (ByVal ADDRESS%, ByVal cmd$)
'Declare Function RecvAT6400Block% Lib "c:\talos\wn956000.dll" (ByVal ADDRESS%, ByVal resp$)
'Declare Function OSLoad% Lib "c:\talos\wn956000.dll" Alias "osload" (ByVal ADDRESS%, ByVal rsvdsrt$, ByVal irqnum%)
'Declare Function IsOSLoaded% Lib "c:\talos\wn956000.dll" (ByVal ADDRESS%)
'Declare Function Request_Status% Lib "c:\talos\wn956000.dll" Alias "request_status" (ByVal ADDRESS%)
'Declare Sub Set_Pointer Lib "c:\talos\wn956000.dll" Alias "set_pointer" (ByVal ADDRESS%, ByVal status_offset%)
'Declare Sub Read_Status Lib "c:\talos\wn956000.dll" Alias "read_status" (ByVal ADDRESS%, status_high%, status_low%, status&)
'Declare Function IsAT6400Ready% Lib "c:\talos\wn956000.dll" (ByVal ADDRESS%)
'Declare Sub SetTimeout Lib "c:\talos\wn956000.dll" (ByVal timeout&)
'Declare Sub Delay Lib "c:\talos\wn956000.dll" (ByVal timedelay&)
'Declare Function SendAT6400File% Lib "c:\talos\wn956000.dll" (ByVal handle%, ByVal ADDRESS%, ByVal irqnum%, ByVal filename$, ByVal options$)
'Declare Function SetDevice Lib "c:\talos\wn956000.dll" (ByVal devnum&, ByVal ADDRESS%) As Boolean
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

#Else
'  16BIT
Declare Function SendAT6400Block Lib "win6400.dll" (ByVal ADDRESS%, ByVal cmd$, ByVal irqnum%) As Integer
Declare Function RecvAT6400Block Lib "win6400.dll" (ByVal ADDRESS%, ByVal response$, ByVal irqnum%) As Integer
Declare Function OSLoad Lib "win6400.dll" (ByVal ADDRESS%, ByVal rsvdstr$, ByVal rsvd%) As Integer
Declare Function IsOSLoaded Lib "c:\talos\win6400.dll" (ByVal ADDRESS%) As Integer
Declare Function SetNTParam Lib "win6400.dll" (ByVal ADDRESS%, ByVal irqnum%) As Integer
Declare Function Request_Status Lib "win6400.dll" (ByVal ADDRESS%) As Integer
Declare Sub Set_Pointer Lib "win6400.dll" (ByVal ADDRESS%, ByVal status_offset%)
Declare Sub Read_Status Lib "win6400.dll" (ByVal ADDRESS%, status_high%, status_low%, status&)
Declare Function PEPrintReport Lib "CRPE.DLL" (ByVal reportFilePath As String, ByVal toPrinter As Integer, ByVal toWindow As Integer, ByVal title As String, ByVal left As Integer, ByVal top As Integer, ByVal width As Integer, ByVal height As Integer, ByVal style As Long, ByVal parentWindow As Integer) As Integer
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetCurrentTime Lib "user" () As Long

Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer




'Public Declare Function GetWindowsDirectory Lib "kernel" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


'Declare Function GetWindowsDirectory Lib "kernel" () As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If



'To call the Sleep function, you could use code like the following:

'Call Sleep(1000)
'----------------------------------------------------------------------------
' Global Constants
'----------------------------------------------------------------------------
Global Const BASEP = 768            'AT6400 Base Address
Global Const MIN_BASEP = &H100      'Minimum address setting
Global Const Max_BASEP = &H3F8      'Maximum address setting
Global Const ENTER = 13
Global Const LIGHTBLUE = &HFFC0C0
Global Const WHITE = &HFFFFFF
Global Const LIMIT_OPEN = 1
Global Const LIMIT_CLOSED = 0
Global Const HOME_INACTIVE = 0
Global Const HOME_ACTIVE = 1

Global Const WF_WINNT = &H4000

'----------------------------------------------------------------------------
' AT6400 Fast Status Area Constants
' Used by WIN6400.DLL SUB ---> set_pointer
'----------------------------------------------------------------------------
Global Const AXIS1_MOTOR = &H0
Global Const AXIS2_MOTOR = &H2
Global Const AXIS3_MOTOR = &H4
Global Const AXIS4_MOTOR = &H6
Global Const AXIS1_ENCODER = &H8
Global Const AXIS2_ENCODER = &HA
Global Const AXIS3_ENCODER = &HC
Global Const AXIS4_ENCODER = &HE
Global Const AXIS1_VELOCITY = &H10
Global Const AXIS2_VELOCITY = &H12
Global Const AXIS3_VELOCITY = &H14
Global Const AXIS4_VELOCITY = &H16
Global Const AXIS1_STATUS = &H18
Global Const AXIS2_STATUS = &H1A
Global Const AXIS3_STATUS = &H1C
Global Const AXIS4_STATUS = &H1E
Global Const INPUT_STATUS = &H20
Global Const OUTPUT_STATUS = &H22
Global Const LIMIT_STATUS = &H24
Global Const INO_STATUS = &H25
Global Const ANALOG_STATUS = &H26
Global Const INT_STATUS = &H28
Global Const SYSTEM_STATUS = &H2A
Global Const USER_STATUS = &H2C
Global Const TFRM_STATUS = &H2D
Global Const TIMER_STATUS = &H2E

'----------------------------------------------------------------------------
' AT6400 Fast Status Response Masks
'----------------------------------------------------------------------------
Global Const CW1_MASK = &H1
Global Const CCW1_MASK = &H2
Global Const CW2_MASK = &H4
Global Const CCW2_MASK = &H8
Global Const CW3_MASK = &H10
Global Const CCW3_MASK = &H20
Global Const CW4_MASK = &H40
Global Const CCW4_MASK = &H80
Global Const HOME1_MASK = &H100
Global Const HOME2_MASK = &H200
Global Const HOME3_MASK = &H400
Global Const HOME4_MASK = &H800
Global Const PCUT_MASK = &H20
Global Const SYSTEM_READY_MASK = &H1


'----------------- giannhs
Global system_ready As Integer
Global Pause As Integer
Global Air_Input As Integer
Global Gripper As Integer  '1-->ανοικτός , 0--> κλειστός
Global Gripper_input As Integer
Global Door_input As Integer
Global Pause_input As Integer
Global Const Green = &HFF00&
Global Const Red = &HFF&
Global Const Grey = &H808080
Global Const orange = &HFFFF&
Global Balance_Asking As String
Global Balance_Settings As String
Global Balance_2_Digits As String
Global Balance_3_Digits As String
Global Balance_Clear As String
Global Balance_Zero As String
Global Gripper_On As String
Global Gripper_Off As String
Global Ready_Bit As String

Global mixer_ON As String
Global mixer_OFF As String

Global mixer2_ON As String
Global mixer2_OFF As String



Global skon_motor_on As String
Global skon_motor_off As String
Global Cold_Water_On As String
Global Cold_Water_Off As String
Global Hot_Water_On As String
Global Hot_Water_Off As String
Global Alarm1_on As String
Global Alarm1_off As String
Global Balance_Port As Integer
Global valve_mixed_on As String
Global valve_mixed_off As String
Global Central_Water_On As String
Global Central_Water_Off As String
Global valve_1_on As String
Global valve_1_off As String
Global valve_2_on As String
Global valve_2_off As String
Global Motor_Input As String
Global Motor_Output_on As String
Global Motor_Output_off As String
Global Motor_Preheating As Long
Global Motor_Preheated As Integer
Global Motor_Output_Number As Integer
Global Water_Bypass As String
Global Balance_Type As String
Global Balance_Mode As String
Global Beaker_Bottle As String ' Αν είναι <> ""  επιτρέπει μεταφορά ποτηριών με μπουκάλια στις κεφαλές
Global Program1 As String
Global Program2 As String
Global Program3 As String
Global Program4 As String

Global Prouermansh_On
Global Prouermansh_Off

Function SETDEVICE(X, Y)
' ΓΙΑ ΛΟΓΟΥΣ ΣΥΜΒΑΤΟΤΗΤΑΣ
   SET_DEVICE = 0
End Function



Sub PAO_ARXH()
    
' ΑΝΕΒΑΣΜΑ
   ANIXE
    
    
    
    
    ' ανέβα πάνω Ζ
    z0 = Ano_Z_93
    
    Find_Axes
    z0 = Ano_Z_93
    cmd2$ = cmd2$ + "d,," + LTrim(str(Ano_Z_93)) + ":"
    cmd2$ = cmd2$ + "go,,1:"


RobSend (cmd2$)
'cm_1 = cmd2$
cmd2$ = ""




' πήγαινε πάνω από το αντικείμενο

'If dejia + aristerh = 0 Then
   
    Find_Axes
                                        'kai den eisai pros kefales meria
If z0 - Zt > 11000 And (Not (X_arxh < 10000 + var(70) And Y_arxh < var(62))) Then ' psilotera kata  5cm toylaxiston
  '    RobSend ("gowhen ,,2pm>" + LTrim(var(62)) + ":")
        
        
    ' einai pano stis kefales   na kanei ligo piso kai na ksekinisei
       If x0 < 10000 + var(70) And y0 < var(62) Then
            Find_Axes
            y0 = var(62)
             cmd2$ = cmd2$ + "d," + LTrim(str(y0)) + ":"
             cmd2$ = cmd2$ + "go,1:"
           Find_Axes
       End If
'If var(12) = 110 Then
'Else
    cmd2$ = cmd2$ + "comexc1:"
'End If

    If var(12) > 1500 And var(12) < 1800 Then
       x0 = X_arxh: y0 = Y_arxh: w0 = var(68): z0 = Z_arxh + 25000   ' normal 11000 karaso 25000
         cmd2$ = cmd2$ + "d" + LTrim(str(x0)) + "," + LTrim(str(y0)) + "," + LTrim(str(z0)) + "," + LTrim(str(w0)) + ":"
         cmd2$ = cmd2$ + "go1,1,1,1:"
    Else
      x0 = X_arxh: y0 = Y_arxh: z0 = Z_arxh + 25000   ' normal 11000 karaso 25000
        cmd2$ = cmd2$ + "d" + LTrim(str(x0)) + "," + LTrim(str(y0)) + "," + LTrim(str(z0)) + ":"
        cmd2$ = cmd2$ + "go1,1,1:"
    End If
    
    
    
Else

    cmd2$ = cmd2$ + "comexc1:"
    If var(12) > 1500 And var(12) < 1800 Then
       x0 = X_arxh: y0 = Y_arxh: w0 = var(68)
         cmd2$ = cmd2$ + "d" + str(X_arxh) + "," + str(Y_arxh) + ",," + str(w0) + ":"
         cmd2$ = cmd2$ + "go1,1,,1:"
    Else
      x0 = X_arxh: y0 = Y_arxh
      cmd2$ = cmd2$ + "d" + str(X_arxh) + "," + str(Y_arxh) + ":"
      cmd2$ = cmd2$ + "go1,1:"
    End If
   
End If
     
     
     
     
     cmd2$ = cmd2$ + "comexc0:"
'  RobSend (cmd2$)
  

'Else
    



' πηγαινε τον W σε θέση παρκαρίσματος  όταν είναι
' ;Επαναφορά ποτηριού απο ζυγό στην θέση του 4000-5000
' ή  ποτήρι  από χειριστη σε ζυγό 3000-4000
' ή ποτήρι σε ζυγαριά για ΔΙΑΛΥΜΑ  19000-19020
' ή μπουκάλι από ανάδευση σε ζυγό  101
'ή μπουκάλι από ζυγό σε ανάδευση  100
If (var(12) > 3000 And var(12) < 5000) Or (var(12) >= 19000 And var(12) <= 19030) Or var(12) = 100 Or (var(12) >= 101 And var(12) < 130) Then
    ' παρκάρισε τον W
    'Find_Axes
    w0 = var(61)
      cmd2$ = cmd2$ + "d ,,," + str(w0) + ":"
     cmd2$ = cmd2$ + "go,,,1:"
End If

' κατέβα κάτω Ζ
'    Find_Axes
    z0 = Z_arxh
    cmd2$ = cmd2$ + "d , ," + str(Z_arxh) + ":"
cmd2$ = cmd2$ + "go,,1:"
'RobSend (cmd2$)
' mexris edo

End Sub

Sub metak_fortomenos()

' ΑΝΕΒΑΣΜΑ
    
    ' ανέβα πάνω Ζ
    z0 = Ano_Z_93
    
    Find_Axes
    z0 = Ano_Z_93
    cmd2$ = cmd2$ + "d,," + str(Ano_Z_93) + ":"
    cmd2$ = cmd2$ + "go,,1:"



 
 If var(12) > 2000 And var(12) < 2500 Then   '      ;Επαναφορά μπουκάλας στη βαση της
      ' πρώτα πήγαινε μέχρι το var62    MONO TON Y
      Find_Axes
      y0 = var(62)
      cmd2$ = cmd2$ + "d," + str(y0) + ":"
      cmd2$ = cmd2$ + "go,1:"
  End If
     
If var(12) > 1000 And var(12) < 1900 Then   ' ;Τοποθέτηση μπουκάλας σε κεφαλή δοσομέτρησης
   ' πρώτα πήγαινε μέχρι το ΜΕΧΡΙ ΤΟ var62 ΚΑΙ ΣΥΓΧΡΟΝΩΣ ΤΟΝ Χ
    Find_Axes
    x0 = Xt: y0 = var(62)
    cmd2$ = cmd2$ + "d" + str(x0) + "," + str(y0) + ":"
    cmd2$ = cmd2$ + "go1,1:"
End If

    
' πηγαινε στο τελικο σημείο
'    Find_Axes
    x0 = Xt: y0 = Yt
    cmd2$ = cmd2$ + "d" + str(x0) + "," + str(y0) + ":"
    cmd2$ = cmd2$ + "go1,1:"



    ' κατέβα κάτω Ζ
'    Find_Axes
    z0 = Zt
    cmd2$ = cmd2$ + "d,," + str(Zt) + ":"
    cmd2$ = cmd2$ + "go,,1:"

End Sub

Function zyg_neo(porta)
    'milsec 50   6 metrhseis / sec me SI
    'milsec 30   4 metrhseis / sec me SI
    'milsec 40   5 metrhseis / sec me SI
    'milsec 60   6.3 metrhseis / sec me SI
    'milsec 70   7.5 metrhseis / sec me SI
    'milsec 80   6.8 metrhseis / sec me SI
    'milsec 90   7.6- 7.2- 8.5 metrhseis / sec me SI



Dim FromModem$, buf As String
Dim D As Integer, s As Integer, tel As Integer
Dim k As Integer, Zyg As String

Zyg = "-999999"
     porta.InBufferCount = 0
     FromModem$ = ""
     porta.Output = "SI" + Chr$(13) + Chr$(10)
     MilSec (70)
       
For k = 1 To 10
     
 If porta.InBufferCount Then
           buf = porta.InBufferCount
           FromModem$ = FromModem$ + porta.Input
           D = InStr(FromModem$, "D")
           s = InStr(2, FromModem$, "S")
           tel = InStr(2, FromModem$, ".")
           c13 = InStr(FromModem$, Chr(13))
     If D > 0 And c13 > 0 And tel > 0 And tel > D And c13 > tel Then
            Zyg = Mid$(FromModem$, D + 1, 12)
            Exit For
     Else
        If s > 0 And c13 > 0 And tel > 0 And tel > s And c13 > tel Then
            Zyg = Mid$(FromModem$, s + 3, 10)
            Exit For
        End If
     End If
          
  End If
   
Next
    
' Me.Caption = FromModem$
zyg_neo = Val(Zyg)

End Function


Sub Gwrite_zygish(f, Z)
 On Error Resume Next
    Set mydb = Workspaces(0).OpenDatabase("c:\talos\syntages.MDB")
    Set r = mydb.OpenRecordset("metler")
    r.AddNew
    r("FromModem") = left(f, 50)
    r("datet") = Now
    r("Zygis0") = Z
    r.update
    r.Close
    
End Sub

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = gstrSEP_DIR Then
        strPathName = left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)

    Close intFileNum

    Err = 0
End Function


Public Sub ch_comdlg(value)
Dim Mycom, mycomset

'mydate = Date
'mytime = Time

'Date = CDate("12/11/1996")
'Time = #12:12:00 PM#

MilSec 500




  Set Mycom = Workspaces(0).OpenDatabase(GetWindowsDir() + "\SYSTEM", False, False, "dBASE IV;")
  Set mycomset = Mycom.OpenRecordset("COMDLG16.DLL")
    mycomset.EDIT
    mycomset("k") = mycomset("k") + value
    mycomset.update
    mycomset.Close

'Date = mydate
'Time = mytime



End Sub

Sub G_Balance_Digits(porta, DIG)

On Error Resume Next
If system_ready = 0 Then Exit Sub
 
 If Balance_Type = "ADAMEQ" Then
    If DIG = 2 Then
        porta.Output = Balance_2_Digits + Chr$(13)
    Else
        porta.Output = Balance_3_Digits + Chr$(13)
    End If
    MilSec 500
 End If
          
End Sub
 
Sub Load_Parametroi()
'=====================================================
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
Set r = dbsjet.OpenRecordset("PARAMETROI")
  
  allazo_xromata = r("allazo_xromata")
  allazo_fonts = r("allazo_fonts")
  prozygismena = r("prozygismena")
   dokimastikh_syntagh = r("dokimastikh_syntagh")
  AXRHSIMOPOIHTES_ORES = r("AXRHSIMOPOIHTES_ORES")
  Default_baros_panioy = r("default_baros_panioy")
  Ar_tel_mpoykal = r("ar_tel_mpoykal")
  Ar_1_mpoykal = r("ar_1_mpoykal")
  DUM = r("MPANIO")
  If IsNull(r("ar_pothr")) Then
     ar_real_pothr = 24
  Else
     ar_real_pothr = r("ar_pothr")
  End If
If IsNull(r("maxtimewater")) Then
   r.EDIT
   r("maxtimewater") = MaxTimeWater
   r.update
Else
   MaxTimeWater = r("maxtimewater")
End If
  
  DUM = r("ColormatchDir")
  DUM = r("Colormatchfilename")
  
   
   
   
If IsNull(r("Fornext_small")) Then
   r.EDIT
   r("Fornext_small") = 40000
   r("Fornext_big") = 120000
   r.update
End If
   
   Fornext_small = r("Fornext_small")
   Fornext_big = r("Fornext_big")
  
  
  If IsNull(r("doyleyo_alati")) Then
     DOYLEYO_ALATI = 0
 Else
     DOYLEYO_ALATI = r("doyleyo_alati")
  End If
  UESH_ALATIOY = r("UESH_ALATIOY")
  
  If IsNull(r("vis_talos")) Then
     vis_talos = 0
 Else
    vis_talos = r("VIS_TALOS")
  End If
 DUM = r("METATOPISH_W")
 
 
 
 
 r.MoveNext

If IsNull(r("vis_talos")) Then
   gXimikaDial = False
Else
   If r("vis_talos") = 1 Then gXimikaDial = True Else gXimikaDial = False
End If

If IsNull(r("maxtimewater")) Then
   Avance_Online_Water = 5000
Else
   Avance_Online_Water = r("maxtimewater")
End If
 
  
  
If IsNull(r("default_baros_panioy")) Then
   YperXeilish = 300
   r.EDIT
   r("default_baros_panioy") = 300
   r.update
Else
   YperXeilish = r("default_baros_panioy")
End If
r.Close
End Sub

'-----------------------------------------------------------
' FUNCTION: UCase16
'
' Returns the upper-case conversion of a string
'   under 16 bits, or else returns an unmodified
'   copy of the string under 32 bits.
'
' IN: [str] - String to copy/upper-case
'
'-----------------------------------------------------------
'
Function UCase16(ByVal str As String)
#If Win16 Then
    UCase16 = UCase$(str)
#Else
    UCase16 = str
#End If
End Function

'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'
Sub AddDirSep(strPathName As String)
    If Right$(RTrim$(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub


'-----------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'-----------------------------------------------------------
'
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
'-----------------------------------------------------------
' FUNCTION: GetWindowsDir
'
' Calls the windows API to get the windows directory and
' ensures that a trailing dir separator is present
'
' Returns: The windows directory
'-----------------------------------------------------------
'
Function GetWindowsDir() As String
    Dim strBuf As String, gintMAX_SIZE
     gintMAX_SIZE = 255
    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf

        GetWindowsDir = UCase16(strBuf)
    Else
        GetWindowsDir = gstrNULL
    End If
  End Function


Function Zygis0_adamEQ(g_PORTa, X As Single)
  Dim counter, START1, FromModem$, dummy, Zyg
  Dim Zygis0_adam
    ' Flush the input buffer.
       g_PORTa.Output = "SEND" + Chr$(13)
       g_PORTa.InBufferCount = 0
       counter = 0
    ' Wait for "OK" to come back from the modem.
    START1 = GetCurrentTime()
    Do
       MilSec 40
       dummy = DoEvents()
       ' If there is data in the buffer, then read it.
       If g_PORTa.InBufferCount Then
          FromModem$ = FromModem$ + g_PORTa.Input
          ' Check for "OK".
          If InStr(FromModem$, Chr$(13)) Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                    Exit Do
                End If
          End If
       Else
            counter = counter + 1
            If counter >= 500 Then  ' 500
                counter = 0
                g_PORTa.Output = "SEND" + Chr$(13)
                g_PORTa.InBufferCount = 0
                FromModem$ = ""
             End If
       End If
        If GetCurrentTime() - START1 > 1000 Then
            Zygis0_adamEQ = -999000999
            Exit Function
       End If
    Loop
    
' FIERATEX
         Zygis0_adam = Mid$(FromModem$, InStr(FromModem$, ".") - 4, 8)
         
         If InStr(FromModem$, "-") > 0 Or InStr(FromModem$, "?") > 0 Then
             Zygis0_adamEQ = -Val(LTrim(Zygis0_adam)) * 1000
         Else
              Zygis0_adamEQ = Val(LTrim(Zygis0_adam)) * 1000
         End If
    
    
    
    
'PALIO   2/9/2002   3 SEIRES
'    ZYG = Val(left$(FromModem$, 9))
 '   Zygis0_adamEQ = ZYG * 1000
End Function


Sub Find_Axes()
   x0 = axe(1)
   y0 = axe(2)
   z0 = axe(3)
   w0 = axe(4)
End Sub


'----------------- giannhs

'
Sub Elegxo_Katastash()


'11-12-2002 If vis_talos = 0 Then
     palaio_Elegxo_Katastash   ' ΠΛΗΡΗΣ ΕΛΕΓΧΟΣ
     Exit Sub
' 11-12-2002 End If
     




' δινω σημα στο άλλο πρόγραμμα να σταματήσει    r("ektelesthke") = 9
        Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
        Set r = dbsjet.OpenRecordset("commun")
  
        If r("ektelesthke") = 9 Then
             r.Close
             Stamathma_Talos "     "
             End
        End If
        r.Close


End Sub

Public Function IsLoaded(FormName)
    Dim I
    IsLoaded = False
    For I = 0 To Forms.Count - 1
        If Forms(I).Name = FormName Then
            IsLoaded = True
            Exit Function
        End If
    Next
End Function
Function oldmL_Res(etik)

End Function

Public Sub RobSend_w(cmd$)
Dim a
If system_ready = 0 Then Exit Sub
       m_Cmd$ = cmd$
       TEMP% = SendATBlock(768, cmd$, 0)
       MilSec 500
  
  Do While Bit_Axona(1, 1) + Bit_Axona(2, 1) + Bit_Axona(3, 1) + Bit_Axona(4, 1) > 0
       DoEvents
  Loop
  
  Do While True
     
    If Abs(axe(1) - x0) < 1 And Abs(axe(2) - y0) < 1 And Abs(axe(3) - z0) < 1 And Abs(axe(4) - w0) < 1 Then
          Exit Do
    End If
     DoEvents
  Loop
  
   a = 1
End Sub

Sub G_REZERO(porta)
Dim Zero


On Error Resume Next
If system_ready = 0 Then Exit Sub
 
 If Balance_Type = "ADAMEQ" Then
      Zero = "ZERO" + Chr(13)
 Else
      Zero = "Z" + Chr$(13) + Chr(10)
 End If

porta.Output = Zero
'secwait (1)
'MSComm1.Output = Zero
'secwait (1)'

'secwait (1)
'asw1 = ZYGIS3()
'secwait (1)
'asw2 = ZYGIS3()
'If Abs(asw1) + Abs(asw2) > 10 Then
'   MilSec 1000
'   GoTo 17
'End If
  
  
End Sub
Function mL_Res(etik)
' αριθμός + Label
Dim db As Database, r As Recordset, cret As String
 On Error Resume Next
  Set db = OpenDatabase("c:\talos\coord_tb.mdb")
  Set r = db.OpenRecordset("Language", dbOpenTable)
  mL_Res = "    "
  r.Index = "PrimaryKey"
  r.Seek "=", etik + 1000
  If Not r.NoMatch And (Not IsNull(r("LABEL"))) Then mL_Res = LTrim(str(1000 + etik)) + " " + r("label")
  r.Close
End Function
Function mLResnew(etik, label0, num)
' αριθμός + Label  num=1 grafei num  num=0 den grafei num
Dim db As Database, r As Recordset, cret As String
 On Error Resume Next
  Set db = OpenDatabase("c:\talos\coord_tb.mdb")
  Set r = db.OpenRecordset("Language", dbOpenTable)
  mLResnew = "    "
  r.Index = "PrimaryKey"
  r.Seek "=", etik + 1000
  If r.NoMatch Then
     
     r.AddNew
     r("id_label") = etik + 1000
     r("english") = label0
     r("greek") = label0
     r("turkey") = lebel0
     r("label") = label0
     r.update
     mLResnew = LTrim(str(1000 + etik)) + " " + label0
  Else
     If (Not IsNull(r("LABEL"))) Then
         If num = 1 Then
            mLResnew = LTrim(str(1000 + etik)) + " " + r("label")
         Else
            mLResnew = r("label")
         End If
     End If
  End If
  r.Close
End Function

Function mL_CapRes(etik)
'mL_Res = "    "
'On Error Resume Next
'   mL_Res = LoadResString(etik)
'On Error GoTo 0
' αριθμός + Label
Dim db As Database, r As Recordset, cret As String
 On Error Resume Next
  Set db = OpenDatabase("c:\talos\coord_tb.mdb")
  Set r = db.OpenRecordset("Language", dbOpenTable)
  mL_CapRes = "    "
  r.Index = "PrimaryKey"
  r.Seek "=", etik + 1000
  If Not r.NoMatch And (Not IsNull(r("LABEL"))) Then
     If numbers_on_labels = 1 Then
        mL_CapRes = str(etik) + "-" + r("label")
     Else
        mL_CapRes = r("label")
     End If
  End If
  r.Close

End Function

Function SendATBlock(n, cmd$, DUM)
Dim a
#If Win32 Then
   a = SendAT6400Block(768, cmd$, 0)
#Else
   a = SendAT6400Block(768, cmd$, 0)
#End If
End Function

Sub ZAP_MOVES()
'*********************
On Error Resume Next '12/99

Set mydb = Workspaces(0).OpenDatabase("c:\WATCH.mdb")
mydb.Execute "DELETE MOVES.* FROM MOVES"
mydb.Close

On Error GoTo 0 '12/99
End Sub

Function zyg_decimals(asw)
'asw = zygisi4(0)
    poin = InStr(asw, ".")
    For I = 1 To 20
       If IsNumeric(Mid(asw, poin + I, 1)) Then dec = dec + 1 Else Exit For
     Next I
 zyg_decimals = dec
 
Exit Function
End Function
Public Function new_wait_to_understand(cmd$)

Dim Runner, TEMP%, synolo
Dim dbsjet As Database
Dim r As Recordset, MHome, timing1
'   Dim dbsJet As Database
cmd2$ = ""
On Error Resume Next
'καταγραφη εντολης στο αρχείο Commun
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
             Set r = dbsjet.OpenRecordset("commun")
             r.EDIT
               r("cmd") = left$(cmd$, 150)
               r("diabasthke") = 0
               r("ektelesthke") = 0


           r.update

           timing1 = r("ora")
           r.Close
On Error GoTo 0
synolo = axe(1) + axe(2) + axe(3) + axe(4)

On Error GoTo run_again
userinput.Text1.LinkMode = 1 ' AUTOMATIC
userinput.Text1.text = cmd$
userinput.Text1.LinkPoke
On Error GoTo 0

'εδω και πολύ ώρα το πρόγραμμα δεν δουλεύει
          '  If Not IsNull(Timing1) Then
          '          If Abs(Val(Timing1) - GetCurrentTime()) > 180000 Then
          '               Stamathma_Talos "The robot is not working." ' Μη λειτουργία προγράμματος Robot"
          '               MsgBox mL_Res(636)
          '               '636 "Το πρόγραμμα του robot δεν ανταποκρίνεται.Kάντε Home και ξανατρέξτε το πρόγραμμα συνταγών."
          '               End
          '          End If
          '   End If
             
          m_Cmd$ = cmd$
          
          If system_ready = 0 Then Exit Function

Begin$ = Time$
 
 While out_bit(bit_ok) = 1 'And Abs(synolo - (axe(1) + axe(2) + axe(3) + axe(4))) < 10 ' περιμένω μέχρι να καταλάβει την εντολή
                  DoEvents
                
                If Runner Mod 1000 = 0 Then
                             ' RobSend (cmd$)
                             timdif = DateDiff("s", Begin$, Time$)
                             If timdif > 30 Then
                                      If timdif > 120 Then
                                                Stamathma_Talos "It is imbossible to send commands to robot" '"Αδυναμία αποστολής εντολών σε  Robot."
                                      Else
                                                
                                      End If
                            End If

                End If
                Runner = Runner + 1
                
Wend

Exit Function





'Do While True
'διαβάζω την  εντολη cmd$ από το αρχείο Commun
'        Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
 '       Set r = dbsJet.OpenRecordset("commun")
  '       If r("diabasthke") = 1 Then
   '         r.Close
    '        Exit Do
     '    End If
      '   r.Close
       '  MilSec 100
'Loop

run_again:
 'retVal = Shell("C:\talos\movement.EXE", 3)
' MilSec 1000
 
 
 ' xanastelno thn entolh
 
 On Error Resume Next
'καταγραφη εντολης στο αρχείο Commun
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
             Set r = dbsjet.OpenRecordset("commun")
             r.EDIT
               r("cmd") = left$(cmd$, 150)
               r("diabasthke") = 0
               r("ektelesthke") = 0


           r.update

           timing1 = r("ora")
           r.Close
On Error GoTo 0
 
 
 
 
 
 
 
 
 
 
 
 
 
 Resume
Return
End Function
Public Function wait_to_understand(cmd$)


Dim Runner, TEMP%
Dim dbsjet As Database
Dim r As Recordset, MHome


  Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r = dbsjet.OpenRecordset("var")
            For k = 1 To 99: var(k) = 0: Next
            Do Until r.eof
              V = r("field_name")
              t = r("field_val")
              If IsNull(t) Then
                t = 0
              End If
                k = Val(Mid$(V, 4, 2))
                If k > 0 And k < 100 Then var(k) = t
              
              'cmd$ = "" + V + "=" + Str$(t) + ":"
              
              'temp% = SendAT6400Block(768, cmd$, 0)
             r.MoveNext
        Loop
      r.Close








'If vis_talos = 1 Then
'     wait_to_understand = new_wait_to_understand(cmd$)
'     Exit Function
'End If
     

'Dim dbsJet As Database
'Dim r As Recordset, MHome
Dim tim

    
    var(13) = 0
         
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
   
   
   RobSend ("out.14-0:")
   
   
 If var(12) = 90 Then ' δεξιά κεφαλή σε ζυγαριά
        Find_Axes
        w0 = var(65)
        RobSend ("d ,,," + str(w0) + ":")
        RobSend ("go,,,1:")
ElseIf var(12) >= 20000 And var(12) < 100000 Then
        druvfn
ElseIf var(12) = 91 Then ' αριστερή κεφαλή σε ζυγαριά
        Find_Axes
        w0 = var(66)
        RobSend ("d ,,," + str(w0) + ":")
        RobSend ("go,,,1:")
ElseIf var(12) = 93 Then ' αριστερή κεφαλή σε ζυγαριά
        Find_Axes
        w0 = var(68)
        RobSend ("d ,,," + str(w0) + ":")
        RobSend ("go,,,1:")




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
   'Form2.Druv
   RobSend ("out.14-1:")
       'ενημερώνω το αρχείο Commun οτι εκτελέστηκε η εντολή
'        Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
 '       Set r = dbsjet.OpenRecordset("commun")
  '       r.EDIT
   '      r("ektelesthke") = 1
    '     r.update
     '    r.Close

MilSec 100
' Loop

End Function

Sub Movement()

Dim var93, var88, Skacimo, cm_1, cm_2, cm_3


    var93 = var(93)
    var88 = var(88)
   Skacimo = var(99)

 
 







Ano_Z_93 = var93

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

RobSend (cmd2$)
'Printer.Print cmd2$
cm_1 = cmd2$
cmd2$ = ""


If var(12) = 94 Then '  ΜΟΝΟ ΝΑ ΠΑΩ ΜΕΧΡΙ ΤΟ ΒΟΛΑΝ
     Exit Sub
End If


'5. Mετακινούμαι με το αντικείμενο προς το τέρμα
'RobSend (cmd2$)

metak_fortomenos

RobSend (cmd2$)
'Printer.Print cmd2$
cm_2 = cmd2$
cmd2$ = ""


If var(12) = 100 Or var(12) = 112 Or var(12) = 122 Then ' απο ζυγό=>ανάδευση  ( δεν αφήνει το μπουκάλι μόνο του στην ανάδευση)

ElseIf var(12) > 1500 And var(12) <= 1999 Then    '   And axe(4) = var(68) Then
    'ποτηρι για διάλυμα σε κεφαλή ΣΚΟΝΗΣ
    ANIXE


  If var(13) = 0 Then ' διαλυματα
    Find_Axes
    z0 = var(88) ' υψος βολάν
    cmd2$ = cmd2$ + "d,," + str(var(88)) + ":go,,1:"
    klise
  End If
Else
      
      
      
      '6. Ανοίγω τον gripper
       ANIXE
       ' 7 ano Z
        Aneba_Pano_Z
End If

 
'8 Κινησεις τακτοποίησης
 kinhseis_taktopoihs



RobSend (cmd2$)
'Printer.Print cmd2$
cm_3 = cmd2$
cmd2$ = ""

End Sub



Public Function FindInput(InputBit As Integer) As Integer
 If system_ready = 0 Then FindInput = 0: Exit Function
    result% = Request_Status(768)
    Set_Pointer 768, INPUT_STATUS
    Read_Status 768, status_high%, status_low%, status&
     If (status_low% And 2 ^ (InputBit - 1)) = 0 Then
          FindInput = 0
      Else
          FindInput = 1
     End If
End Function
Public Function FindOutput(Bit_Num)
    If system_ready = 0 Then
       FindOutput = 1
       Exit Function
    End If
result% = Request_Status(768)
  Set_Pointer 768, OUTPUT_STATUS
    Read_Status 768, status_high%, status_low%, status&
     If (status& And 2 ^ (Bit_Num - 1)) = 0 Then
           FindOutput = 0
     Else
           FindOutput = 1
     End If
End Function

Function find_uesh(cmd$) As String

Dim MyPos, kod, kod2 As Long
   ' A textual comparison starting at position 4. Returns 6.
  
    MyPos = InStr(1, cmd$, "=")
     kod = Val(Mid$(cmd$, MyPos + 1, 10))
          If kod > 1000 And kod < 1399 Then ' μπουκάλι σε κεφαλή
             find_uesh = cmd$ + find_Coordin(1, kod - 1000, 1)
             
          ElseIf kod > 1500 And kod < 1899 Then ' μπουκάλι σε κεφαλή
             find_uesh = cmd$ + find_Coordin(1, kod - 1500, 1)
   
          ElseIf kod > 2000 And kod < 2999 Then    ' μπουκάλι  από κεφαλή σε θέση
             find_uesh = cmd$ + find_Coordin(1, kod - 2000, 1)
          ElseIf kod > 3000 And kod < 3999 Then ' φερε ποτήρι σε ζυγαριά
             find_uesh = cmd$ + find_Coordin(2, kod - 3000, 1)
          ElseIf kod > 4000 And kod < 4999 Then ' φερε ποτήρι απο ζυγαριά στην θέση του
             find_uesh = cmd$ + find_Coordin(2, kod - 4000, 1)
          
          ElseIf kod > 6000 And kod < 12000 Then ''Μεταφορά μπουκαλιων στο χειριστή
             find_uesh = cmd$ + find_Coordin(1, kod - 6000, 1) ' μπουκαλι
              
              MyPos2 = InStr(MyPos + 1, cmd$, "=")
              kod2 = Val(Mid$(cmd$, MyPos2 + 1, 10))
              find_uesh = find_uesh + find_Coordin(2, kod2, 1) '  pothri
          
          ElseIf kod > 12000 And kod < 18000 Then   '  Επαναφορά μπουκαλιων στη θέση τους  από χειριστη
             find_uesh = cmd$ + find_Coordin(1, kod - 12000, 1)
          
              MyPos2 = InStr(MyPos + 1, cmd$, "=")
              kod2 = Val(Mid$(cmd$, MyPos2 + 1, 10))
              find_uesh = find_uesh + find_Coordin(2, kod2, 1) '  pothri
          ElseIf kod > 19000 And kod < 19030 Then    '
             find_uesh = cmd$ + find_Coordin(2, kod - 19000, 1)
          
          ElseIf kod > 100000 And kod < 100300 Then    '
               ' μπουκάλι κεφαλής ( που θα φύγει απο κεφαλή )
                find_uesh = cmd$ + find_Coordin(1, kod - 100000, 1) ' μπουκάλι - μπουκάλι
                 
                ' μπουκάλι που θα ρθεί στην κεφαλή
              MyPos2 = InStr(MyPos + 1, cmd$, "=")
              kod2 = Val(Mid$(cmd$, MyPos2 + 1, 10))
              find_uesh = find_uesh + find_Coordin(1, kod2, 2) '  pothri
         ElseIf kod > 200 And kod <= 230 Then
             find_uesh = cmd$ + find_Coordin(2, kod - 200, 1)
          
          ElseIf kod < 100 Then
              find_uesh = cmd$
   End If


End Function

'
Sub palaio_Elegxo_Katastash()
        ' ελεγχος αν πατήθηκε το CutOff
        If isCutOff(0) = 1 Then
                     Stamathma_Talos "Cutoff"
                    '  "Διακοπή λειτουργίας προγράμματος λόγω ενεργοποίησης Αισθητήρα Πρόσκρουσης (Crush Sensor) ή Emergency Stop"
                      MsgBox "612-613." + mL_CapRes(612) + Chr$(13) + mL_CapRes(613), , "Talos" '  "Διακοπή λειτουργίας προγράμματος λόγω ενεργοποίησης Αισθητήρα Πρόσκρουσης (Crush Sensor) ή Emergency Stop"
                    End
        End If
        
         ' If is_pause = 1 Then
            '  MsgBox mL_Res(637)
              ' 637 "Διακοπή λειτουργίας λόγω ενεργοποίησης διακόπτη PAUSE"
         ' End If
          
          
          
          
        
          If PTOSH_PIESHS = 1 Then
             cmd$ = "" + Central_Water_Off + ":" + "" + Cold_Water_Off + ":" + "" + Hot_Water_Off + ":" + "" + valve_1_off + ":" + "" + valve_2_off + ":"
             RobSend (cmd$)
             MsgBox mL_Res(638)
             '638 "Διακοπή λειτουργίας λόγω ΠΤΩΣΗΣ ΠΙΕΣΗΣ."
             
             Stamathma_Talos "The programm was interrupted because the pressure is low" '"Διακοπή λειτουργίας λόγω ΠΤΩΣΗΣ ΠΙΕΣΗΣ."
          End If
        
If Lim1 < 17 And Lim2 < 17 Then ' αν hard limits
   'ελεγχος τερματικών
    result% = Request_Status(768)
    Set_Pointer 768, LIMIT_STATUS
    Read_Status 768, status_high%, status_low%, status&
        
    For k = 1 To 4
        If (status_high% And 2 ^ (2 * k - 2)) = 0 Then
            Stamathma_Talos mL_Res(639) + str(k) '639 "Τερματικό άξονα "
            MsgBox mL_Res(639) + str(k)
            End
        End If
        If (status_high% And 2 ^ (2 * k - 1)) = 0 Then
            Stamathma_Talos mL_Res(639) + str(k)
            MsgBox mL_Res(639) + str(k)
            End
        End If
    Next
        
End If
        
        
         ' ΕΛΕΓΧΟΣ ΚΙΝΗΤΗΡΩΝ
        If error(13, 0) = 1 Or error(14, 0) = 1 Then
                Stamathma_Talos "Terminal 13-14" '"ενεργοποίηση τερματικού 13-14"
                MsgBox mL_Res(614), , "Talos"     ' "Διακοπή λειτουργίας προγράμματος λόγω ενεργοποίησης τερματικού"
                End
        End If
        
        ' ΕΛΕΓΧΟΣ ΚΙΝΗΤΗΡΩΝ
       If error(Lim1, 0) = 1 Or error(Lim2, 0) = 1 Then
                Stamathma_Talos "ενεργοποίηση τερματικού 15-16"
                '    "Διακοπή λειτουργίας προγράμματος λόγω ενεργοποίησης τερματικού"
                MsgBox mL_Res(615), , "Talos"
                End
                '
        End If
        
        
        
        
           ' ΕΛΕΓΧΟΣ encoder
'     If K_AR = 7 Then
        If error(12, 0) = 1 Then
                Stamathma_Talos "ενεργοποίηση encoder"
                '  "Διακοπή λειτουργίας  λόγω ενεργοποίησης encoder "
                MsgBox mL_Res(616), , "Talos"   '"Διακοπή λειτουργίας  λόγω ενεργοποίησης encoder " '
                End
         End If
'    End If
    
    
    Do While Is_Porta() = "α"
       cmd$ = "" + Central_Water_Off + ":" + "" + Cold_Water_Off + ":" + "" + Hot_Water_Off + ":" + "" + valve_1_off + ":" + "" + valve_2_off + ":"
        RobSend (cmd$)
       
       '  wait_to_understand (cmd$)
         MsgBox mL_Res(640)
         ' 640 "Κλείστε την πόρτα αμέσως για να συνεχιστεί η διαδικασία."
     Loop
End Sub
Public Function LoadLabel(Number As Long) As String
On Error GoTo er_ck
Dim db As Database, r As Recordset, cret As String
  Set db = OpenDatabase("c:\talos\coord_tb.mdb")
  Set r = db.OpenRecordset("Language", dbOpenDynaset)
    cret = "[Id_Label]=" & Number
    r.FindFirst cret
   LoadLabel = r("Label")
    If r.NoMatch Then LoadLabel = "Error : Label " + str(Number) + "  missing"
    r.Close
Exit Function
er_ck:
LoadLabel = "Error Label :" + str(Number) + "  Missing"
End Function
Public Sub PreHeating()

If Len(Motor_Input) > 0 And FindInput(Val(Motor_Input)) = 0 Then
  If Motor_Preheating > 0 Then
       Msg = "465-467." + mL_CapRes(465) + Chr$(13) + mL_CapRes(466) + Chr$(13) + mL_CapRes(467)
      style = vbYesNo + vbCritical + vbDefaultButton2
      title = mL_Res(490)
      response = MsgBox(Msg, style, title)
            If response = 7 Then Screen.MousePointer = 1: Exit Sub
             cmd$ = "" + Motor_Output_on + ":"
              RobSend (cmd$)
              MilSec 200
            makehome = Shell("C:\talos\chek_hom.EXE", 1)
            MilSec 20000
            Do: DoEvents: Loop Until FindOutput(Ready_Bit) = 1
             cmd$ = "" + Motor_Output_on + ":"
          Motor_Preheated = 1
          RobSend (cmd$)
     End If
End If
If Motor_Input <> "" And FindOutput(Motor_Output_Number) = 0 Then
        cmd$ = "" + Motor_Output_on + ":"
        Motor_Preheated = 1
        RobSend (cmd$)
End If

ValveReset
End Sub

Function begin_talos(X As Integer) As Integer
'******* κάνω εκκίνηση το τάλος
Dim counter
counter = 0

If vis_talos = 1 Then
     begin_talos = new_begin_talos(0)
     Exit Function
End If
     



ValveReset


'begin_talos = 1
'Exit Function




  If isCutOff(0) = 1 Then
  '   "Ελευθερώστε  το Emergency Stop.."
         MsgBox mL_Res(617), , "Talos "
         begin_talos = 0
         Exit Function
  End If



Do While True
     counter = 0
     Do While True   ' ελεγχος αν όντως έκανε το ok_bit = 0
           cmd$ = "halt:miden:out." + LTrim(str(bit_ok)) + "-0:"
           TEMP% = SendATBlock(768, cmd$, 0)
           MilSec 50
           If out_bit(bit_ok) = 0 Then
                  Exit Do
           End If
           MilSec 500
        '   For k = 1 To 50
        '       DoEvents
        '   Next
           counter = counter + 1
           
           If counter = 5 Then   ' δεν μπορεί να μηδενίσει το βιτ_οκ
                 begin_talos = 0
                 MsgBox mL_Res(618) + str(bit_ok), , "Talos "  ' "Πρόβλημα  στην Εξοδο της CompuMotor  Νο "
                 Exit Function
           End If
           
     Loop
     MilSec 500
     'For k = 1 To 50
    '    DoEvents
    ' Next
     cmd$ = "var1=1:talos:"
     TEMP% = SendATBlock(768, cmd$, 0)
   MilSec 500
'     For k = 1 To 50
'         DoEvents
'     Next
     If out_bit(bit_ok) = 1 Then
         begin_talos = 1
         Exit Do
      End If
      counter = counter + 1
      If counter = 3 Then
          begin_talos = 0
          Exit Do
      End If
  
Loop


End Function
Public Function Bit_Axona(axe, No_bit)
Dim stat, status_array(1 To 5), Msg

Msg = " "
sfalma = 0
 I = axe
    Select Case I
       Case 1
        
      result% = Request_Status(768)
       Set_Pointer 768, AXIS1_STATUS
       Case 2
            
             result% = Request_Status(768)
              Set_Pointer 768, AXIS2_STATUS
       Case 3
              
                 result% = Request_Status(768)
                Set_Pointer 768, AXIS3_STATUS
       Case 4
               
               result% = Request_Status(768)
              Set_Pointer 768, AXIS4_STATUS
End Select

 Read_Status 768, status_high%, status_low%, status&
  
  status_array(I) = status&
  
'If tas_bit(No_bit, status&) <> EPIU Then

      If (status& And 2 ^ (No_bit - 1)) = 0 Then
           Bit_Axona = 0
     Else
           Bit_Axona = 1
     End If


End Function

Function new_begin_talos(X As Integer) As Integer
'******* κάνω εκκίνηση το τάλος
Dim counter
counter = 0

ValveReset


new_begin_talos = 1
Exit Function




  If isCutOff(0) = 1 Then
  '   "Ελευθερώστε  το Emergency Stop.."
         MsgBox mL_Res(617), , "Talos "
         new_begin_talos = 0
         Exit Function
  End If



Do While True
     counter = 0
     Do While True   ' ελεγχος αν όντως έκανε το ok_bit = 0
           cmd$ = "halt:miden:out." + LTrim(str(bit_ok)) + "-0:"
           TEMP% = SendATBlock(768, cmd$, 0)
           MilSec 50
           If out_bit(bit_ok) = 0 Then
                  Exit Do
           End If
           MilSec 500
        '   For k = 1 To 50
        '       DoEvents
        '   Next
           counter = counter + 1
           
           If counter = 5 Then   ' δεν μπορεί να μηδενίσει το βιτ_οκ
                 new_begin_talos = 0
                 MsgBox mL_Res(618) + str(bit_ok), , "Talos "  ' "Πρόβλημα  στην Εξοδο της CompuMotor  Νο "
                 Exit Function
           End If
           
     Loop
     MilSec 500
     'For k = 1 To 50
    '    DoEvents
    ' Next
     cmd$ = "var1=1:talos:"
     TEMP% = SendATBlock(768, cmd$, 0)
   MilSec 500
'     For k = 1 To 50
'         DoEvents
'     Next
     If out_bit(bit_ok) = 1 Then
         new_begin_talos = 1
         Exit Do
      End If
      counter = counter + 1
      If counter = 3 Then
          new_begin_talos = 0
          Exit Do
      End If
  
Loop


End Function

Public Function Wait_Robot(DUM As Integer) As Integer
  Dim Begin$, Pause
    comp_Wait_Robot = 1
    If system_ready = 0 Then Exit Function
         
    Begin$ = Time$
     While out_bit(bit_ok) = 0
           DoEvents
           timdif = DateDiff("s", Begin$, Time$)
           If timdif > 100 Then
                If Is_Porta = "κ" And is_pause = 1 And PTOSH_PIESHS = 0 Then
                   Stamathma_Talos "the robot cannot move."
                   '   "Αδυναμία Κίνησης Robot. Για να δουλεύει ανεμπόδιστα ΠΡΕΠΕΙ:"
                   MsgBox "619-627" + mL_CapRes(619) + Chr$(13) + mL_CapRes(620) + Chr$(13) + mL_CapRes(621) + Chr$(13) + mL_CapRes(622) + Chr$(13) + mL_CapRes(623) + Chr$(13) + mL_CapRes(624) + Chr$(13) + mL_CapRes(625) + Chr$(13) + mL_CapRes(626) + Chr$(13) + mL_CapRes(627), , "Talos"
                   End
                   Wait_Robot = 0
                 End If
                While Is_Porta = "α"
                       '   "Αδυναμία Κίνησης λόγω ανοικτής πόρτας "
                         MsgBox "628-629" + mL_CapRes(628) + Chr$(13) + mL_CapRes(629), , "Talos"
                         Begin$ = Time$
                Wend
                While Pause = 1
                   '    "Αδυναμία Κίνησης λόγω διακόπτη PAUSE" , "Γυρίστε το PAUSE στο 1 καί κάντε κλικ στο ΟΚ"
                         MsgBox "630-631" + mL_CapRes(630) + Chr$(13) + mL_CapRes(631), , "Talos"
                         Begin$ = Time$
                Wend
                While PTOSH_PIESHS = 1
                   '"Αδυναμία Κίνησης λόγω πτώσης Πίεσης " , "Μόλις η πίεση επανέρθει στα 3bar  κάντε κλικ στο ΟΚ"
                         MsgBox "632-633" + mL_CapRes(632) + Chr$(13) + mL_CapRes(633), , "Talos"
                         Begin$ = Time$
                 Wend
                 If DateDiff("s", Begin$, Time$) > 180 Then
                     '  "Δεν παίρνει σήμα το πρόγραμμα συνταγών ότι ολοκληρώθηκε η κίνηση (Wait_robot)"
                     MsgBox mL_Res(634), , "Talos"
                     End
                 End If
          End If
          
'
     Wend
     
     Write_No_home 0
     Wait_Robot = 1

End Function
Public Function new_Wait_Robot(DUM As Integer) As Integer
  Dim Begin$, Pause
Dim dbsjet As Database
Dim r As Recordset, MHome
If system_ready = 0 Then Exit Function

'Do While True
'διαβάζω την  εντολη cmd$ από το αρχείο Commun
 '       Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
  '      Set r = dbsJet.OpenRecordset("commun")
   '      If r("ektelesthke") = 1 Then
    '        r.Close
     '       Exit Do
      '   End If
'         r.Close
 '        MilSec 100
'Loop
         
  Begin$ = Time$
  While out_bit(bit_ok) = 0
           DoEvents
           timdif = DateDiff("s", Begin$, Time$)
           If timdif > 90 Then
                If Is_Porta = "κ" And is_pause = 1 And PTOSH_PIESHS = 0 Then
                   Stamathma_Talos "the robot cannot move." ' "Αδυναμία Κίνησης Robot."
                   '   "Αδυναμία Κίνησης Robot. Για να δουλεύει ανεμπόδιστα ΠΡΕΠΕΙ:"
                   MsgBox "619-627" + mL_CapRes(619) + Chr$(13) + mL_CapRes(620) + Chr$(13) + mL_CapRes(621) + Chr$(13) + mL_CapRes(622) + Chr$(13) + mL_CapRes(623) + Chr$(13) + mL_CapRes(624) + Chr$(13) + mL_CapRes(625) + Chr$(13) + mL_CapRes(626) + Chr$(13) + mL_CapRes(627), , "Talos"
                   End
                   new_Wait_Robot = 0
                 End If
                While Is_Porta = "α"
                       '   "Αδυναμία Κίνησης λόγω ανοικτής πόρτας "
                         MsgBox "628-629" + mL_CapRes(628) + Chr$(13) + mL_CapRes(629), , "Talos"
                         Begin$ = Time$
                Wend
                While Pause = 1
                   '    "Αδυναμία Κίνησης λόγω διακόπτη PAUSE" , "Γυρίστε το PAUSE στο 1 καί κάντε κλικ στο ΟΚ"
                         MsgBox "630-631" + mL_CapRes(630) + Chr$(13) + mL_CapRes(631), , "Talos"
                         Begin$ = Time$
                Wend
                While PTOSH_PIESHS = 1
                   '"Αδυναμία Κίνησης λόγω πτώσης Πίεσης " , "Μόλις η πίεση επανέρθει στα 3bar  κάντε κλικ στο ΟΚ"
                         MsgBox "632-633" + mL_CapRes(632) + Chr$(13) + mL_CapRes(633), , "Talos"
                         Begin$ = Time$
                 Wend
                 If DateDiff("s", Begin$, Time$) > 90 Then
                     '  "Δεν παίρνει σήμα το πρόγραμμα συνταγών ότι ολοκληρώθηκε η κίνηση (Wait_robot)"
                     MsgBox mL_Res(634), , "Talos"
                     End
                 End If
          End If
     Wend
  
  
  new_Wait_Robot = 1
End Function

Public Sub Stamathma_Talos(aitia)

Dim dbsjet As Database
Dim r As Recordset, MHome

'stamathma  KATAΓΡΑΦΗ ΔΙΑΚΟΠΗΣ
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
             Set r = dbsjet.OpenRecordset("stamathma")
             r.AddNew
               r("aitia") = left$(aitia, 150)
               r("last_command") = left$(m_Cmd$, 150)
               r("datetime") = Now
               r("x") = axe(1)
               r("y") = axe(2)
               r("z") = axe(3)
               r("w") = axe(4)
           r.update
             r.Close
  

ValveReset

Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
      Set r = dbsjet.OpenRecordset("mem3")
       r.MoveFirst
       r.EDIT
       r("HOME") = "NO HOME"
       r("XRHSTHS") = "NO BUSY"
       r.update
       r.Close
           
           
           
           cmd$ = "k:miden:out." + LTrim(str(bit_ok)) + "-0:"
           TEMP% = SendATBlock(768, cmd$, 0)
           
           
End Sub
Function tas_bit(ar_bit, stat)
     If (stat And 2 ^ (ar_bit - 1)) = 0 Then
           tas_bit = 0
     Else
           tas_bit = 1
     End If
End Function


Public Sub Err_Ck(ERR_NUM As Long)
If ERR_NUM = 53 Then MsgBox mL_Res(25), , "Talos": Exit Sub
If ERR_NUM = 3011 Then MsgBox mL_Res(26), , "Talos": Exit Sub
MsgBox "Error number " + str(Err) + Chr(13) + Err.Description, , "Talos"
End Sub
Public Sub ValveReset() ' Κάνει Reset σ'όλες τις βαλβίδες κατά την έναρξη προγραμμάτων
FindValves
cmd$ = "" + mixer_OFF + ":"
cmd$ = cmd$ + "" + Cold_Water_Off + ":"
cmd$ = cmd$ + "" + Hot_Water_Off + ":"
If Central_Water_Off <> "" Then cmd$ = cmd$ + "" + Central_Water_Off + ":"
cmd$ = cmd$ + "" + valve_1_off + ":"
cmd$ = cmd$ + "" + valve_2_off + ":"
cmd$ = cmd$ + "" + skon_motor_off + ":"
RobSend (cmd$)
End Sub

Public Sub Reindexing(DUM)
Dim M_AP, M_KO, M_HM
M_AP = 0: M_KO = 0: M_HM = 0

EXCL = True

Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
'-------------------------prospau2 ---------------------------------------------------
Set r = mydb.OpenRecordset("prospau2")
ok = 0
For k = 0 To r.Fields.Count - 1
     If UCase(r.Fields(k).Name) = "APOXRVSH" Then M_AP = 1
     If UCase(r.Fields(k).Name) = "KOD_PEL" Then M_KO = 1
     If UCase(r.Fields(k).Name) = "HME_PARAD" Then M_HM = 1
Next
r.Close

On Error Resume Next

If M_AP = 0 Then mydb.Execute "ALTER TABLE PROSPAU2 ADD COLUMN APOXRVSH TEXT(20)"
If M_KO = 0 Then mydb.Execute "ALTER TABLE PROSPAU2 ADD COLUMN KOD_PEL  TEXT(20)"
If M_HM = 0 Then mydb.Execute "ALTER TABLE PROSPAU2 ADD COLUMN HME_PARAD DATETIME"

'Kill "C:\TALOS\PROSPAU2.MDX"
 SQL = "DROP  INDEX  ENTOLH  ON   prospau2;"
 mydb.Execute SQL




SQL = "CREATE  INDEX  ENTOLH  ON   prospau2(ENTOLH,AA_PROSPAU);"
mydb.Execute SQL



'----------------------------------------  prospau 3
Set r = mydb.OpenRecordset("prospau3")
ok = 0
M_AP = 0: M_KO = 0: M_HM = 0
For k = 0 To r.Fields.Count - 1
     If UCase(r.Fields(k).Name) = "APOXRVSH" Then M_AP = 1
     If UCase(r.Fields(k).Name) = "KOD_PEL" Then M_KO = 1
     If UCase(r.Fields(k).Name) = "HME_PARAD" Then M_HM = 1
Next
r.Close

On Error Resume Next

If M_AP = 0 Then mydb.Execute "ALTER TABLE PROSPAU3 ADD COLUMN APOXRVSH TEXT(20)"
If M_KO = 0 Then mydb.Execute "ALTER TABLE PROSPAU3 ADD COLUMN KOD_PEL  TEXT(20)"
If M_HM = 0 Then mydb.Execute "ALTER TABLE PROSPAU3 ADD COLUMN HME_PARAD DATETIME"

'Kill "C:\TALOS\PROSPAU2.MDX"
 SQL = "DROP  INDEX  ENTOLH  ON   prospau3;"
 mydb.Execute SQL




SQL = "CREATE  INDEX  ENTOLH  ON   prospau3(ENTOLH,AA_PROSPAU);"
mydb.Execute SQL
'---------------------------------












 
 SQL = "DROP  INDEX  ENTOLH  ON   joblist;"
 mydb.Execute SQL

SQL = "CREATE  INDEX  ENTOLH  ON  JOBLIST(ENTOLH);"
mydb.Execute SQL


 SQL = "DROP  INDEX  HME_PARAD  ON   joblist;"
 mydb.Execute SQL
 SQL = "CREATE  INDEX  HME_PARAD  ON   JOBLIST(HME_PARAD);"
 mydb.Execute SQL

SQL = "CREATE  INDEX  PARAT  ON   JOBLIST(PARAT);"
mydb.Execute SQL

SQL = "CREATE  INDEX  APOXRVSH  ON   JOBLIST(APOXRVSH);"
mydb.Execute SQL


'On Error Resume Next
'Kill "C:\TALOS\PELATES.MDX"
'On Error GoTo 0

 SQL = "DROP  INDEX  ONOMA  ON   PELATES;"
 mydb.Execute SQL

SQL = "CREATE  INDEX  ONOMA  ON   PELATES(ONOMA);"
mydb.Execute SQL


'On Error Resume Next
'Kill "C:\TALOS\PANIA.MDX"
'On Error GoTo 0
 
 SQL = "DROP  INDEX  PANI  ON   PANIA;"
 mydb.Execute SQL


SQL = "CREATE  INDEX  PANI  ON   PANIA(PANI);"
mydb.Execute SQL

'On Error Resume Next
'Kill "C:\TALOS\XIMITECH.MDX"
'On Error GoTo 0

'Kill "dok.dbf"
'MyDB.Execute "select * into dok from ximitech"
'Kill "ximitech.dbf"
'Name "dok.dbf" As "ximitech.dbf"

'On Error GoTo 0

 SQL = "DROP  INDEX  KOD  ON  XIMITECH;"
 mydb.Execute SQL

 SQL = "DROP  INDEX  PER  ON  XIMITECH;"
 mydb.Execute SQL


SQL = "CREATE  INDEX  KOD  ON   XIMITECH(KOD);"
mydb.Execute SQL

SQL = "CREATE  INDEX  PER  ON   XIMITECH(PERIGRAFH);"
mydb.Execute SQL

SQL = "CREATE  INDEX  mPER  ON   XIMITECH(mPERIGRAFH);"
mydb.Execute SQL

mydb.Close



'RetVal = Shell("\talos\reindex.exe", 1)    ' Run Calculator.


End Sub

Public Function error(No_bit, EPIU)
Dim stat, status_array(1 To 5), Msg

Msg = " "
sfalma = 0
For I = 1 To 4
    Select Case I
       Case 1
        
      result% = Request_Status(768)
       Set_Pointer 768, AXIS1_STATUS
       Case 2
            
             result% = Request_Status(768)
              Set_Pointer 768, AXIS2_STATUS
       Case 3
              
                 result% = Request_Status(768)
                Set_Pointer 768, AXIS3_STATUS
       Case 4
               
               result% = Request_Status(768)
              Set_Pointer 768, AXIS4_STATUS
End Select

 Read_Status 768, status_high%, status_low%, status&
  
  status_array(I) = status&
  
If tas_bit(No_bit, status&) <> EPIU Then
Msg = "    ...."
 Select Case No_bit
  Case 13
      ' MsgBox "Θερμικό ή έλλειψη τάσης  στον κινητήρα του άξονα  Νο" + Str(I) + " Ελέγξτε την συνδεσμολογία του πίνακα."
       MsgBox "600-601" + mL_CapRes(600) + str(I) + mL_CapRes(601), , "Talos"
   Case 14
       'MsgBox "Δεν υπάρχει επικοινωνία μεταξύ κάρτας CompuMotor  και κινητήρα  στον άξονα Νο " + Str(I) + ".   Ελέγξτε την συνδεσμολογία του πίνακα."
       MsgBox "602-603" + mL_CapRes(602) + str(I) + mL_CapRes(603), , "Talos"
   Case 5
        ' MsgBox "Δεν έγινε σωστό Home   στον άξονα Νο" + Str(I)
       MsgBox mL_Res(604) + str(I), , "Talos"
   Case 15
        ' MsgBox "Ενεργοποιήθηκε τερματικό  του άξονα  Νο" + Str(I)
       MsgBox mL_Res(605) + str(I) + " .", , "Talos"
   Case 16
         'MsgBox "Ενεργοποιήθηκε τερματικό  του άξονα  Νο" + Str(I)
       MsgBox mL_Res(606) + str(I) + " .", , "Talos"
  Case 12
      '  MsgBox "Χάθηκαν βήματα στον άξονα  Νο" + Str(I)
       MsgBox mL_Res(607) + str(I) + " .", , ""

 End Select
       
       
       Stamathma_Talos Msg
       GoSub err_msg
       '   "Διακοπή λειτουργίας προγράμματος ."
       MsgBox mL_Res(635), , "Talos"
       End
  End If
Next
    
    If sfalma = 1 Then
        error = 1
    Else
       error = 0
   End If
Exit Function
   
   
   
err_msg:
sfalma = 1 'Σφάλμα κατάστασης
 
 
 Select Case No_bit
 
 Case 13
      ' MsgBox "Θερμικό ή έλλειψη τάσης  στον κινητήρα του άξονα  Νο" + Str(I) + " Ελέγξτε την συνδεσμολογία του πίνακα."
       MsgBox "600-601" + mL_CapRes(600) + str(I) + mL_CapRes(601), , "Talos"
   Case 14
       'MsgBox "Δεν υπάρχει επικοινωνία μεταξύ κάρτας CompuMotor  και κινητήρα  στον άξονα Νο " + Str(I) + ".   Ελέγξτε την συνδεσμολογία του πίνακα."
       MsgBox "602-603" + mL_CapRes(602) + str(I) + mL_CapRes(603), , "Talos"
   Case 5
        ' MsgBox "Δεν έγινε σωστό Home   στον άξονα Νο" + Str(I)
       MsgBox mL_Res(604) + str(I), , "Talos"
   Case 15
        ' MsgBox "Ενεργοποιήθηκε τερματικό  του άξονα  Νο" + Str(I)
       MsgBox mL_Res(605) + str(I) + " .", , "Talos"
   Case 16
         'MsgBox "Ενεργοποιήθηκε τερματικό  του άξονα  Νο" + Str(I)
       MsgBox mL_Res(606) + str(I) + " .", , "Talos"
  Case 12
      '  MsgBox "Χάθηκαν βήματα στον άξονα  Νο" + Str(I)
       MsgBox mL_Res(607) + str(I) + " .", , ""
 
End Select
  Return
End Function





Function Is_Porta() As String
   Dim Z
      Z = FindInput(Door_input)
      If Z >= 1 Then
          Is_Porta = "α"
      Else
          Is_Porta = "κ"
     End If
End Function

Function PTOSH_PIESHS() As Integer
   Dim Z
  PTOSH_PIESHS = FindInput(Air_Input)
End Function


Public Function out_bit(ByVal ar_bit As Integer)   ' ελέγχω το βιτ out_bit
 
    On Error Resume Next
    If flag = 1 Then Exit Function
    result% = Request_Status(768)
    ' ******* Update Motor Position ****************

    Set_Pointer 768, OUTPUT_STATUS
    Read_Status 768, status_high%, status_low%, status&
    
     If (status& And 2 ^ (ar_bit - 1)) = 0 Then
           out_bit = 0
     Else
           out_bit = 1
     End If
Exit Function
End Function


Public Sub MilSec(WAIT As Long)
Dim start, c_tim, k, k0
     If WAIT < 1 Then
        k0 = WAIT * ForNexts_Milsec
        For k = 1 To k0: Next
        Exit Sub
     End If
     start = GetCurrentTime()
     Do
        c_tim = GetCurrentTime()
        DoEvents
     Loop Until c_tim >= start + WAIT
End Sub
Function miliseconds(WAIT As Long)
    Dim start, c_tim, k, k0
     start = GetCurrentTime()
     Do
        c_tim = GetCurrentTime()
        DoEvents
     Loop Until c_tim >= start + WAIT
End Function

Public Function FIND_COLOR(MERES)
                          Select Case MERES
                                    Case 0
                                         FIND_COLOR = &H80FF&           '  πορτοκαλί
                                     Case Is > 0
                                         FIND_COLOR = &H8000& '  πράσινο
                                      Case Else
                                          FIND_COLOR = &HFF&         'κοκκινο
                            End Select

End Function

Public Sub FindValves()

Dim mydb As Database, r As Recordset

On Error GoTo er_c:
Set mydb = OpenDatabase("c:\talos\coord_tb.mdb")
Set r = mydb.OpenRecordset("Βαλβιδες")
Beaker_Bottle = ""
Do Until r.eof
  If r("valve_kod") = 1 Then
     Gripper_On = r("valve_on")
     
     Gripper_Off = r("valve_off")
   
   End If
   If r("valve_kod") = 2 Then
      Ready_Bit = Val(Mid(r("valve_on"), 5, 2))
   End If
    If r("valve_kod") = 3 Then
   Braxionas = r("valve_on")
   End If
  If r("valve_kod") = 4 Then
    skon_motor_on = r("valve_on")
     skon_motor_off = r("valve_off")
   End If
   If r("valve_kod") = 5 Then
     Cold_Water_On = r("valve_on")
     Cold_Water_Off = r("valve_off")
   End If
   If r("valve_kod") = 6 Then
     Hot_Water_On = r("valve_on")
     Hot_Water_Off = r("valve_off")
   End If
   If r("valve_kod") = 7 Then
     mixer_ON = r("valve_on")
     mixer_OFF = r("valve_off")
   End If
   If r("valve_kod") = 8 Then
     Balance_Port = r("valve_on")
  End If
  If r("valve_kod") = 9 Then
     valve_1_on = r("valve_on")
     valve_1_off = r("valve_off")
   End If
      If r("valve_kod") = 10 Then
     valve_2_on = r("valve_on")
     valve_2_off = r("valve_off")
   End If
   If r("valve_kod") = 11 Then
     Alarm1_on = r("valve_on")
     Alarm1_off = r("valve_off")
  End If
   If r("valve_kod") = 12 Then
     Pause_input = r("valve_on")
     Air_Input = r("valve_off")
  End If
    If r("valve_kod") = 13 Then
     Door_input = r("valve_on")
  End If
  If r("valve_kod") = 14 Then
    Gripper_input = r("valve_on")
  End If
   If r("valve_kod") = 15 Then
    Talos_Fonts = r("valve_on")
    Talos_Language = r("valve_off")
  End If
  If r("valve_kod") = 16 Then
    Motor_Input = r("valve_on")
    Motor_Preheating = r("valve_off")
  End If
   If r("valve_kod") = 17 Then
    Motor_Output_on = r("valve_on")
    Motor_Output_off = r("valve_off")
    Motor_Output_Number = Val(Mid(Motor_Output_on, 5, 1))
  End If
    If r("valve_kod") = 18 Then
    Alarm2_on = r("valve_on")
    Alarm2_off = r("valve_off")
  End If
    If r("valve_kod") = 19 Then
    Alarm3_On = r("valve_on")
    Alarm3_Off = r("valve_off")
  End If
   If r("valve_kod") = 20 Then
    Central_Water_On = r("valve_on")
    Central_Water_Off = r("valve_off")
  End If
  If r("valve_kod") = 21 Then
    Balance_Settings = r("valve_on")
    Balance_Asking = r("valve_off")
 End If
  If r("valve_kod") = 22 Then
    Balance_2_Digits = r("valve_on")
    Balance_3_Digits = r("valve_off")
 End If
    If r("valve_kod") = 23 Then
    Balance_Zero = r("valve_on")
    Balance_Clear = r("valve_off")
 End If
    If r("valve_kod") = 24 Then
    Water_Bypass = r("valve_on")
  End If
  If r("valve_kod") = 25 Then
    Balance_Type = r("valve_on")
    Balance_Mode = r("valve_off")
 End If
  If r("valve_kod") = 26 Then
    Program1 = r("valve_on")
    Program2 = r("valve_off")
 End If
   If r("valve_kod") = 27 Then
    Program3 = r("valve_on")
  If IsNull(r("valve_off")) Then Program4 = "" Else Program4 = r("valve_off")
 End If
If r("valve_kod") = 28 Then
    Beaker_Bottle = r("valve_on")
End If

If r("valve_kod") = 29 Then
    Prouermansh_On = r("valve_on")
    Prouermansh_Off = r("valve_off")
End If

If r("valve_kod") = 30 Then
    mixer2_ON = r("valve_on")
    mixer2_OFF = r("valve_off")
End If


If r("valve_kod") = 31 Then
    XimNer_on = r("valve_on")
    XimNer_Off = r("valve_off")
    If Len(XimNer_on) > 2 Then XimNer = 1
End If

If r("valve_kod") = 32 Then
    Xim1_on = r("valve_on")
    xim1_Off = r("valve_off")
    If Len(Xim1_on) > 2 Then aritmos_ximikon = 1
End If
If r("valve_kod") = 33 And Not IsNull(r("valve_on")) Then
    Xim2_on = r("valve_on")
    xim2_Off = r("valve_off")
    If Len(Xim2_on) > 2 Then aritmos_ximikon = 2
End If

If r("valve_kod") = 34 And Not IsNull(r("valve_on")) Then
    Xim3_on = r("valve_on")
    xim3_Off = r("valve_off")
    If Len(Xim3_on) > 2 Then aritmos_ximikon = 3
End If

If r("valve_kod") = 35 And Not IsNull(r("valve_on")) Then
    Xim4_on = r("valve_on")
    xim4_Off = r("valve_off")
    If Len(Xim4_on) > 2 Then aritmos_ximikon = 4
End If


r.MoveNext
Loop
r.Close

 aritmos_ximikon = 0
 XimNer = 0

Exit Sub
er_c:
If Err = 13 Then Resume Next
If Err = 94 Then Resume Next
Err_Ck (Err)
Resume Next
End Sub


Public Sub RobSend(cmd$)
If system_ready = 0 Then Exit Sub
        '    Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
        '     Set r = dbsJet.OpenRecordset("moves")
        '     r.AddNew
        '       r("movement") = left$(cmd$, 150)
        '       r("datetime") = Now
        '     r.update
        '     r.Close

        m_Cmd$ = cmd$
       TEMP% = SendATBlock(768, cmd$, 0)
End Sub

Public Sub Begin_Work()
   Dim dbsjet As Database
   Dim r As Recordset, MHome
    
      Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
      Set r = dbsjet.OpenRecordset("mem3")

     r.MoveFirst
      
      If r("XRHSTHS") = "JOYSTICK" Then
           MsgBox mL_Res(73), , "Talos"
           End
       End If
      
      If r("HOME") = "HOME=OK" Then
        '
      Else
             MsgBox mL_Res(74), , "Talos"
             End
      End If
      
      
      
      If r("XRHSTHS") = "BUSY" Then
             MsgBox mL_Res(75), , "Talos"
             End
       End If
      
      r.EDIT
       r("HOME") = "NO HOME"
       r("XRHSTHS") = "BUSY"
       r.update
       r.Close
       
End Sub




Function find_colors(forma)
 '  find colors of screen
Dim FI(1 To 15), r, k
   
          Select Case UCase(forma.Name)
          Case "FRMXHM"
             color_field = "F0"
          Case "FRMXHMEDIT"
             color_field = "F1"
          Case "MJOBLIST"
             color_field = "F2"
          Case "FRMSYNT2"
             color_field = "F3"
          Case "WAIT"
             color_field = "F4"
          Case "SYNTAGES"
             color_field = "F5"
          Case "FERE"
             color_field = "F6"
          Case "FERESK"
             color_field = "F7"
          Case "PANE"
             color_field = "F8"
          Case "PELATES"
             color_field = "F9"
         End Select
   
   Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\syntages.mdb")
   Set r = dbsjet.OpenRecordset("form_colors")
          
      On Error Resume Next
   
          'If UCase(obj) = "COLORS" Then
             r.MoveFirst
             'find_colors = Val(R(color_field))
            If allazo_xromata = -1 Then
               forma.BackColor = Val(r(color_field))
            End If
         ' ElseIf UCase(obj) = "FONTS" Then
                 r.MoveNext
                 s = r(color_field)
                 If Not IsNull(r(color_field)) Then
                    If allazo_fonts = -1 Then
                       forma.FontBold = IIf(Mid$(r(color_field), 1, 1) = "1", True, False)
                       forma.FontItalic = IIf(Mid$(r(color_field), 2, 1) = "1", True, False)
                       forma.FontStrikethru = IIf(Mid$(r(color_field), 3, 1) = "1", True, False)
                       forma.FontUnderline = IIf(Mid$(r(color_field), 4, 1) = "1", True, False)
                       forma.FontSize = Val(Mid$(r(color_field), 5, 7))
                       forma.FontName = Mid$(r(color_field), 12, 30)
                    End If
                  End If
          ' End If
           r.Close
                 For k = 1 To forma.Controls.Count
                     forma.Controls(k).FontSize = forma.FontSize
                 Next
      On Error GoTo 0
      
End Function

Function save_colors(forma, New_Color As Long)
Dim a, xx
 '  find colors of screen
If allazo_xromata = 0 Then
   Exit Function
End If
         xx = forma.BackColor
         forma.CommonDialog1.ShowColor
        
        a = forma.CommonDialog1.Color
        
        #If Win32 Then
             forma.BackColor = a
             New_Color = forma.CommonDialog1.Color
        #Else
           If forma.CommonDialog1.Color <> 0 Then
             forma.BackColor = a
             New_Color = forma.CommonDialog1.Color
           Else
             New_Color = xx
           End If
        #End If
        
   Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\syntages.mdb")
     Set r = dbsjet.OpenRecordset("form_colors")
          
      On Error Resume Next
          r.MoveFirst
          r.EDIT
          Select Case UCase(forma.Name)
          Case "FRMXHM"
             r("F0") = New_Color
          Case "FRMXHMEDIT"
             r("F1") = New_Color
          Case "MJOBLIST"
             r("F2") = New_Color
          Case "FRMSYNT2"
            r("F3") = New_Color
          Case "WAIT"
           r("F4") = New_Color
          Case "SYNTAGES"
             r("F5") = New_Color
          Case "FERE"
            r("F6") = New_Color
          Case "FERESK"
            r("F7") = New_Color
          Case "PANE"
             r("F8") = New_Color
          Case "PELATES"
             r("F9") = New_Color
         End Select
      On Error GoTo 0
        r.update
      r.Close
End Function

Function axe(ajonas As Integer)
         
 result% = Request_Status(768)
  Set_Pointer 768, AXIS1_MOTOR
      For I = 0 To 3
      Read_Status 768, status_high%, status_low%, status&
       If ajonas = I + 1 Then
         axe = status&
        Exit Function
       End If
     Next I
End Function

Function VELOCITY(AXONAS)
 result% = Request_Status(768)
  Set_Pointer 768, AXIS2_VELOCITY
      For I = 0 To 3
      Read_Status 768, status_high%, status_low%, status&
       If AXONAS = I + 1 Then
          VELOCITY = status&
        Exit Function
       End If
     Next I
End Function


Function isCutOff(DUM As Integer)
     Dim X
     
      result% = Request_Status(768)
   ' ******* Update Pulse Cutoff Status ****************
    Set_Pointer 768, INO_STATUS
    Read_Status 768, status_high%, status_low%, status&
    If (status_high% And PCUT_MASK) > 0 Then
         isCutOff = 0  ' οχι πατημένο
    Else
         isCutOff = 1 ' πατημένο
    End If
End Function





Function Pause_Check() As Integer
   ' ελέγχω το βιτ  2o   ( porta )
   '    pause = 1   σημαίνει ότι δουλεύει κανονικά
   '    Pause = 1   σημαίνει ότι  σταμάτησε
    
    Pause_Check = FindInput(Pause_input)
    
End Function


Public Function find_Coordin(Eidos As Integer, Number As Long, Typos_Kinhshs As Integer) As String
'*************************************
   Dim dbsjet As Database
   Dim r As Recordset
' Typos_kinhshs = 1  όλες εκτός μπουκάλι-μπουκάλι         2=  μπουκάλι-μπουκάλι
' Εidos=1  μπουκάλι       2=ποτήρι
    Set mydb = Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = mydb.OpenRecordset("Συντεταγμενες", dbOpenSnapshot)
    
    
'#If Win32 Then
 '   On Error GoTo create
 '      r.Index = "aa"
 '      r.Seek "=", Eidos, Number
 '      If r.NoMatch Then
 '        SYN_X = 0: SYN_Y = 0: syn_z = 0
 '      Else
 '         SYN_X = r("Συντεταγμένη_Χ_user")
 '         SYN_Y = r("Συντεταγμένη_Υ_user")
 '         syn_z = r("Συντεταγμένη_Ζ_user")
 '      End If
       
 '   On Error GoTo 0
'#Else


Do Until r.eof

If r("ΑΑ_Item") = Eidos Then
 If r("Νο_Τεμαχιου") = Number Then
   SYN_X = r("Συντεταγμένη_Χ_user")
   SYN_Y = r("Συντεταγμένη_Υ_user")
   syn_z = r("Συντεταγμένη_Ζ_user")
   Exit Do
   End If
 End If
  r.MoveNext
 Loop
  
  
'#End If
    
    
    

'Do Until r.EOF
'If r("ΑΑ_Item") = Eidos Then
' If r("Νο_Τεμαχιου") = Number Then
'   SYN_X = r("Συντεταγμένη_Χ_user")
'   SYN_Y = r("Συντεταγμένη_Υ_user")
'   syn_z = r("Συντεταγμένη_Ζ_user")
'   Exit Do
'   End If
' End If
'  r.MoveNext
' Loop
  
 If SYN_X = 0 And SYN_Y = 0 And syn_z = 0 Then
     MsgBox mL_Res(608), , "Talos(608)" 'sΤΑΜΑΤΗΜΑ ΛΟΓΩ ΛΑΘΟΥΣ ΣΥΝΤΕΤΑΓΜΕΝΩΝ"
     Stamathma_Talos "wrong coordinates"
     End
  End If
  Select Case Eidos
    Case 1   ' μπουκάλι
        If Typos_Kinhshs = 1 Then
             find_Coordin = "var36=" + LTrim(str(SYN_X)) + ":var37=" + LTrim(str(SYN_Y)) + ":"
              var(35) = syn_z
              var(36) = SYN_X
              var(37) = SYN_Y
             Exit Function
        Else
             find_Coordin = "var38=" + LTrim(str(SYN_X)) + ":var39=" + LTrim(str(SYN_Y)) + ":"
             var(38) = SYN_X
             var(39) = SYN_Y
             var(40) = syn_z
             Exit Function
        End If
        
     Case 2    ' ποτηρι
        If Typos_Kinhshs = 1 Then
             find_Coordin = "var24=" + LTrim(str(SYN_X)) + ":var77=" + LTrim(str(SYN_Y)) + ":" + "var89=" + LTrim(str(syn_z)) + ":"
             
                 var(24) = SYN_X
                 var(77) = SYN_Y
                 var(89) = syn_z
             Exit Function
        End If
End Select

Exit Function
create:
'If r("ΑΑ_Item") = Eidos Then
' If r("Νο_Τεμαχιου") = Number Then
r.Close

     mydb.Execute "CREATE  INDEX  aa ON   [Συντεταγμενες]([ΑΑ_Item],[Νο_Τεμαχιου]);"
     
     Set r = mydb.OpenRecordset("Συντεταγμενες", dbOpenTable)
     r.Index = "aa"
     Resume Next
     Return












End Function
Sub find_arxh()
'--------------------------------------------
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


Sub klise()
a = InStr(Gripper_Off, "!")
If a > 0 Then
   Gripper_Off = Mid$(Gripper_Off, 1, a - 1) + Mid$(Gripper_Off, a + 1, Len(Gripper_Off) - a)
End If
' If FindInput(Gripper_input) = gripper_anoiktos Then  ' anoiktos
dz = -var(99)
     z0 = z0 + dz
    cmd2$ = cmd2$ + "ma0000:d,," + str(dz) + ":go,,1:"
    cmd2$ = cmd2$ + "WAIT (mov = bxx0x):"
    cmd2$ = cmd2$ + "ma1111:" '  ;Κατεβαινει 1 mm για να αποφυγει διαφορες Ζ
   cmd2$ = cmd2$ + "" + Gripper_Off + ":wait(in.1=b1):t0.3:" ' ":t.03:t0.3:"  ' wait(in.1=b1)
   
  

'   Do While FindInput(Gripper_input) = gripper_anoiktos '  kleistos
 '     DoEvents
 '  Loop
 '  var(10) = 1
  ' RobSend ("t0 0.5:")
'  MilSec 500
'End If

End Sub
Sub ANIXE()
Dim a

a = InStr(Gripper_On, "!")
If a > 0 Then
   Gripper_On = Mid$(Gripper_On, 1, a - 1) + Mid$(Gripper_On, a + 1, Len(Gripper_On) - a)
End If
'If FindInput(Gripper_input) = gripper_kleistos Then  ' kleistos
  ' oso to bit 1 einai iso me to 0( diladi o gripper einai klistos) perimene
  ' bit 1 =1  shmainei o gripper anoiktos
     cmd2$ = cmd2$ + "" + Gripper_On + ":wait(in.1=b0):t0.3:"
'     Do While FindInput(Gripper_input) = gripper_kleistos  ' sigk gripper_anoiktos
'       DoEvents
'    Loop
'    var(10) = 0
'End If

End Sub
Sub Write_No_home(ByVal Yes_No As Integer)
' 1 γραφει no home    0 : Home=OK

     Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r = dbsjet.OpenRecordset("mem3")

     If Yes_No = 1 Then
        r.EDIT
        r("HOME") = "NO HOME"
        r("XRHSTHS") = "BUSY"
        r.update
     Else
        r.EDIT
        r("XRHSTHS") = "NO BUSY"
        r("HOME") = "HOME=OK"
        r.update
     End If
     
     
     
     
       r.Close

End Sub

Sub Aneba_Pano_Z()

  ' ανέβα πάνω Ζ
    z0 = Ano_Z_93
   
    Find_Axes
    z0 = Ano_Z_93
    cmd2$ = cmd2$ + "d , ," + str(Ano_Z_93) + ":"
    cmd2$ = cmd2$ + "go,,1:"
    
End Sub
Sub kinhseis_taktopoihs()

' ετοιμάζω την κεφαλή νερού για δοσομέτρηση
If var(12) > 3000 And var(12) < 4000 Then
    Find_Axes
    w0 = var(67)
     cmd2$ = cmd2$ + "d ,,," + str(w0) + ":"
    cmd2$ = cmd2$ + "go,,,1:"
     
     ' ετοιμάζω την κεφαλή σκόνης για δοσομέτρηση
     
ElseIf (var(12) >= 2000 And var(12) <= 2499 And var(13) = 0) Then 'pane mpoykali kai oxi pane-fere(var(13)>0)
          
    Find_Axes
    x0 = var(85): y0 = var(86) ' pano apo zygaria
     cmd2$ = cmd2$ + "d " + str(x0) + "," + str(y0) + ":"
     cmd2$ = cmd2$ + "go 1,1:"
   
     
     
ElseIf (var(12) >= 19000 And var(12) <= 19030) Then
 'sigk   Find_Axes
 'sigk   w0 = var(68)
 'sigk   RobSend ("d ,,," + Str(w0) + ":")
 'sigk   RobSend_w ("go,,,1:")
End If


End Sub
Sub get_uesh(aritmos)
 
    Set db = OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = db.OpenRecordset("Συντεταγμενες")
       Do Until r.eof
           If r("ΑΑ_Item") = 1 Then
           If r("Νο_Τεμαχιου") = aritmos Then
                x0 = r("Συντεταγμένη_Χ_user")
                y0 = r("Συντεταγμένη_Υ_user")
                z0 = r("Συντεταγμένη_Ζ_user")
                Exit Do
            End If
           End If
           r.MoveNext
       Loop
    r.Close

End Sub

Sub load_taxythtes()

RobSend ("!INFNC2-A:")
RobSend ("!out.7-0:")
RobSend ("!out.7-1:")
RobSend ("!out.14-1:")  ' readybit ok

 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = dbsjet.OpenRecordset("Ταχυτητες")
     
    Do Until r.eof
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




TaxyThtes = "!v " + str(vx) + "," + str(vy) + "," + str(vz) + "," + str(vw) + ":"
Epitaxynseis = "!a " + str(a_x) + "," + str(a_y) + "," + str(a_z) + "," + str(a_w) + ":"
epibradynseis = "!ad " + str(e_x) + "," + str(e_y) + "," + str(e_z) + "," + str(e_w) + ":"
RobSend (TaxyThtes)
RobSend (Epitaxynseis)
RobSend (epibradynseis)
End Sub

Sub druvfn() '   ;druvanisma finish

Dim aa, x1, y1, Z1
Dim dddb, r, vim
cmd2$ = ""
'    Set dddb = Workspaces(0).OpenDatabase("c:\talos\recipies.MDB")
'    Set r = dddb.OpenRecordset("parametroi")
'      vim = r("metatopish_W")
'    dddb.Close



' επαναφέρω την αρχική ταχύτητα
cmd2$ = cmd2$ + TaxyThtes
ANIXE ' αφήνει το ποτήρι της σκόνης
cmd2$ = cmd2$ + "ma1111:"

get_uesh 400
Z1 = z0
Find_Axes
z0 = Z1 ' var(69) ' κατέβα να πιάσεις το μπουκάλι της σκόνης
cmd2$ = cmd2$ + "d,," + str(z0) + ":go,,1:"

klise

Find_Axes
z0 = Ano_Z_93
cmd2$ = cmd2$ + "d,," + str(Ano_Z_93) + ":go,,1:"


  y0 = var(62)
  cmd2$ = cmd2$ + "d," + LTrim(str(y0)) + ":"
  cmd2$ = cmd2$ + "go,1:"




get_uesh var(12) - 20000

'x0 = var(36)
'y0 = var(37)
'z0 = Ano_Z_93
' w0 = var(67) + Vim


  


cmd2$ = cmd2$ + "d " + str(x0) + "," + str(y0) + "," + str(Ano_Z_93) + ":go111:"


' κατεβάζει το μπουκάλι
'Find_Axes
'z0 = var(35)    '  var(92)
'w0 = var(67)
cmd2$ = cmd2$ + "d,," + str(z0) + ":"
cmd2$ = cmd2$ + "go,,1:"


RobSend (cmd2$)
cmd2$ = ""
ANIXE

'Aneba_Pano_Z

    z0 = Ano_Z_93
    Find_Axes
    z0 = Ano_Z_93
cmd2$ = cmd2$ + "d , ," + str(Ano_Z_93) + ":"
cmd2$ = cmd2$ + "go,,1:"







' ερχεται πάνω απο την ζυγαριά
Find_Axes
' συντ.ζυγού
    x0 = var(85)
    y0 = var(86)

cmd2$ = cmd2$ + "d " + str(x0) + "," + str(y0) + ":"
cmd2$ = cmd2$ + "go11:"

RobSend (cmd2$)


End Sub



Sub park()
     '  ΠΑΡΚΑΡΙΣΜΑ
     RobSend ("d ,," + str(var(93)) + ":")
     
     z0 = var(93)
     RobSend ("go,,1:")
     
     
     RobSend ("d " + str(0) + "," + str(0) + ":")
     RobSend ("go11:")
     
     w0 = var(61)
     RobSend ("d,,," + str(w0) + ":")
     RobSend ("go,,,1:")
  
End Sub
