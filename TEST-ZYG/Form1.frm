VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdZYGISEYTALOS 
      Caption         =   "ZYGISEYTALOS"
      Height          =   360
      Left            =   7920
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCommand4 
      Caption         =   "REZERO_TALOS"
      Height          =   600
      Left            =   7800
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REZERO T"
      Height          =   735
      Left            =   7800
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REZERO ON/OFF"
      Height          =   615
      Left            =   7680
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim system_ready ' = 0
Dim Balance_Asking As String
Dim Balance_Type '= "ADAM"
Private Sub cmdCommand4_Click()
  G_REZERO MSComm1
  MsgBox "OK"
  
  
End Sub

Private Sub cmdZYGISEYTALOS_Click()

Dim cou As Integer
cou = 0
Do While cou < 3
   cou = cou + 1
   Label1.BackColor = vbRed
    Label1.Caption = zyg2_ADAM(0)
    Label1.BackColor = vbGreen
    
    
    
    'ам дем апамтаеи тгм нейимаеи лем акка тгм лгдемифеи
   ' отам еиами ожж евеи йаккыс
   'акка отам паеи сто лемоу????
   
   
    If Len(Label1.Caption) < 2 Then
       ' MSComm1.Output = "ON" + Chr(13)
       ' MilSec 3000  '4 OK
    Else
       Exit Do
    End If
        
Loop

End Sub

'16.3.1Control commands
'PLUxx       Call-up         PLU from data memory
'T           Tare placed weighing vessel
'T123.456    Numeric tare valueZZeroing
'P           Printing
'M+          Add and print weighing data in the summation memory
'MR          Call-up data from memory
'MC          Delete memory
'U123.456    Save the average piece weight 123.456 [g] or [lb]
'S123        Input number of pieces e.g. 123 pieces
'SL          Switch over to reference balance
'SR          Switch over to bulk material scales

Private Sub Command1_Click()
Dim cou As Integer
cou = 0
Do While cou < 3
   cou = cou + 1
   Label1.BackColor = vbRed
    Label1.Caption = zyg_ADAM(0)
    Label1.BackColor = vbGreen
    
    
    
    'ам дем апамтаеи тгм нейимаеи лем акка тгм лгдемифеи
   ' отам еиами ожж евеи йаккыс
   'акка отам паеи сто лемоу????
   
   
    If Len(Label1.Caption) < 2 Then
        MSComm1.Output = "ON" + Chr(13)
        MilSec 3000  '4 OK
    Else
       Exit Do
    End If
        
Loop

        
        
End Sub
Function zyg_ADAM(Zygis_Kind As Variant)
On Error GoTo er_det
   Dim counter
  Dim tot_counter
  Dim start
  Dim FromModem As String
  Dim dummy
  Dim BUF
  
   counter = 0: tot_counter = 0: Label10 = 0
' If system_ready = 0 Then zyg_ADAM = "System not Working": Exit Function
  start = GetCurrentTime()
  
  
    Do
    If GetCurrentTime() - start > 3001 Then zyg_ADAM = " ": Exit Function
     MSComm1.InBufferCount = 0
      FromModem = ""
           MSComm1.Output = Balance_Asking + Chr$(13)
          MilSec (50)
          dummy = DoEvents()
          If MSComm1.InBufferCount Then
             BUF = MSComm1.InBufferCount
             FromModem$ = FromModem$ + MSComm1.Input
             '  List1.AddItem FromModem$, 0
           
              If InStr(UCase(FromModem$), "OVER") > 0 Then
                'MsgBox "Scale Overload ...", , "Talos"
                zyg_ADAM = "OL"
                Exit Do
              End If
              If InStr(UCase(FromModem$), "UNDER") > 0 Then
                 'MsgBox "Scale Overload ...", , "Talos"
                 zyg_ADAM = "UL"
                 Exit Do
              End If
            If GetCurrentTime() - start > 5000 And Zygis_Kind = "OK" Then Zygis_Kind = 0: tot_counter = 0
          
          
          
          
          
    If InStr(FromModem$, ".") > 0 Then
          
            Dim NN As Integer
            NN = InStr(FromModem$, ".")
            Dim CC3 As String
            Dim DD As String
            
           If NN > 5 Then
            
            
            '    CC = Mid(FromModem$, NN - 3, 7)
             '   DD = Mid(FromModem$, NN, 4)
              '  If InStr(DD, "?") > 0 Then
               '    DD = Replace(DD, "?", "0")
               ' End If
                
                
                CC3 = Mid(FromModem$, NN - 5, 9)
                
                If InStr(CC3, "?") > 0 Then
                 
                   
                Else
                
                  '  CC = Replace(CC, "?", "0")
                    zyg_ADAM = CC3
                    Exit Do
                End If
            Else
               zyg_ADAM = ""
            End If
            
               
          
     End If

          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          If InStr(FromModem$, Chr$(13)) >= 5 Then
                 If InStr(FromModem$, ".") >= 5 Then
                       Exit Do
                End If
          End If
       
       
       
       
       counter = counter + 1
       If counter >= 15 Then tot_counter = tot_counter + 1
            Label10 = counter
          If tot_counter >= 20 And Zygis_Kind <> "OK" Then
          counter = 0: tot_counter = tot_counter + 1
               zyg_ADAM = " "
               Exit Function
           End If
          If counter >= 50 Then  ' 500
             counter = 0: tot_counter = tot_counter + 1
               MSComm1.InBufferCount = 0
              MSComm1.Output = Balance_Asking + Chr$(13)
              End If
     End If
     If Zygis_Kind = "OK" Then
             'Me.Zyg_Show.Width = (GetCurrentTime() - start) / 1000 * 650
              'Me.Zyg_Show.Caption = Int(((GetCurrentTime() - start) / 1000) + 0.5) & " Sec"
             'Me.Zyg_Show.Refresh
      End If
      
    Loop
         
       '  zyg_ADAM = Mid$(FromModem$, InStr(FromModem$, ".") - 4, 8)
        ' If InStr(FromModem$, "-") > 0 Or InStr(FromModem$, "?") > 0 Then
         '    zyg_ADAM = Str(-Val(LTrim(zyg_ADAM)))
        ' End If
er_ex:
Exit Function
er_det:
 zyg_ADAM = " "
Resume er_ex
End Function

Private Sub Command2_Click()

    'If Len(Label1.Caption) < 2 Then
        MSComm1.Output = "ON" + Chr(13)
        MilSec 3000  '4 OK
         MSComm1.Output = "T" + Chr(13)
    'End If


End Sub

Private Sub Command3_Click()
     MSComm1.Output = "T " '+ Chr(13) '+ Chr(10)
        MilSec 3000
       '  MSComm1.Output = "T" + Chr(13)
End Sub

Private Sub Form_Load()
  Balance_Asking = "E"
Balance_Type = "ADAM"

  MSComm1.PortOpen = True
  system_ready = 1

End Sub

Sub G_REZERO(porta)

    Dim Zero

    On Error Resume Next

    If system_ready = 0 Then Exit Sub
 
    If Balance_Type = "ADAM" Then
 
        porta.Output = "T" + Chr(13)

        If Abs(zyg2_ADAM(0)) < 2 Then
            ' OK
        Else
            porta.Output = "ON" + Chr(13)
            MilSec 4000  '4 OK
            porta.Output = "T" + Chr(13)

            If Abs(zyg2_ADAM(0)) < 2 Then
                ' OK
            Else
                ' гтам сбгстг йаи тгм амоицеи
                porta.Output = "ON" + Chr(13)
                MilSec 4000  '4 OK
                porta.Output = "T" + Chr(13)
            End If
       
        End If
      
    Else
        Zero = "Z" + Chr$(13) + Chr(10)
        porta.Output = Zero
    End If

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



Function zyg2_ADAM(Zygis_Kind As Variant) As Long  'TALOS
On Error GoTo er_det
   Dim counter
  Dim tot_counter
  Dim start
  Dim FromModem As String
  Dim dummy
  Dim BUF
  
   counter = 0: tot_counter = 0: Label10 = 0
' If system_ready = 0 Then zyg_ADAM = "System not Working": Exit Function
  start = GetCurrentTime()
  
  
    Do
    If GetCurrentTime() - start > 10001 Then zyg2_ADAM = -999000999: Exit Function
     MSComm1.InBufferCount = 0
      FromModem$ = ""
           MSComm1.Output = Balance_Asking + Chr$(13)
          MilSec (50)
          dummy = DoEvents()
          If MSComm1.InBufferCount Then
           BUF = MSComm1.InBufferCount
          FromModem$ = FromModem$ + MSComm1.Input
        '  List1.AddItem FromModem$, 0
        
        
        
         
              If InStr(UCase(FromModem$), "OVER") > 0 Then
                'MsgBox "Scale Overload ...", , "Talos"
                zyg2_ADAM = -7900000  ' "OL"
                Exit Do
              End If
              If InStr(UCase(FromModem$), "UNDER") > 0 Then
                 'MsgBox "Scale Overload ...", , "Talos"
                 zyg2_ADAM = -9900000 ' "UL"
                 Exit Do
              End If
        
        
        
           
'           If InStr(FromModem$, "OL") > 0 Then
'                'MsgBox "Scale Overload ...", , "Talos"
'                zyg_ADAM = "OL"
'                Exit Do
'            End If
'           If InStr(FromModem$, "UL") > 0 Then
'                'MsgBox "Scale Overload ...", , "Talos"
'                zyg_ADAM = "UL"
'                Exit Do
'            End If
            If GetCurrentTime() - start > 5000 And Zygis_Kind = "OK" Then Zygis_Kind = 0: tot_counter = 0
          
          
          
          
          
    If InStr(FromModem$, ".") > 0 Then
          
            Dim NN As Integer
            NN = InStr(FromModem$, ".")
            Dim CC3 As String
            Dim DD As String
            
           If NN > 3 Then
            
            
            '    CC = Mid(FromModem$, NN - 3, 7)
             '   DD = Mid(FromModem$, NN, 4)
              '  If InStr(DD, "?") > 0 Then
               '    DD = Replace(DD, "?", "0")
               ' End If
                
                
                CC3 = Mid(FromModem$, NN - 3, 7)
                
                If InStr(CC3, "?") > 0 Then
                 
                   
                Else
                
                  '  CC = Replace(CC, "?", "0")
                    zyg2_ADAM = Val(CC3) * 1000
                    Exit Do
                End If
            Else
               zyg2_ADAM = -999000999
            End If
            
               
          
     End If

          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          If InStr(FromModem$, Chr$(13)) >= 5 Then
                 If InStr(FromModem$, ".") >= 5 Then
                       Exit Do
                End If
          End If
       
       
       
       
       counter = counter + 1
       If counter >= 15 Then tot_counter = tot_counter + 1
            Label10 = counter
          If tot_counter >= 20 And Zygis_Kind <> "OK" Then
          counter = 0: tot_counter = tot_counter + 1
               zyg2_ADAM = " "
               Exit Function
           End If
          If counter >= 50 Then  ' 500
             counter = 0: tot_counter = tot_counter + 1
               MSComm1.InBufferCount = 0
              MSComm1.Output = Balance_Asking + Chr$(13)
              End If
     End If
     If Zygis_Kind = "OK" Then
             'Me.Zyg_Show.Width = (GetCurrentTime() - start) / 1000 * 650
              'Me.Zyg_Show.Caption = Int(((GetCurrentTime() - start) / 1000) + 0.5) & " Sec"
             'Me.Zyg_Show.Refresh
      End If
      
    Loop
         
       '  zyg_ADAM = Mid$(FromModem$, InStr(FromModem$, ".") - 4, 8)
        ' If InStr(FromModem$, "-") > 0 Or InStr(FromModem$, "?") > 0 Then
         '    zyg_ADAM = Str(-Val(LTrim(zyg_ADAM)))
        ' End If
er_ex:
Exit Function
er_det:
 zyg2_ADAM = " "
Resume er_ex
End Function



