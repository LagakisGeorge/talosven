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


Dim Balance_Asking As String



Private Sub Command1_Click()
   Label1.BackColor = vbRed
    Label1.Caption = zyg_ADAM(0)
    Label1.BackColor = vbGreen
    If Len(Label1.Caption) < 2 Then
        MSComm1.Output = "ON" + Chr(13)
    End If
    
  
End Sub
Function zyg_ADAM(Zygis_Kind As Variant)
On Error GoTo er_det
   Dim counter
  Dim tot_counter
  Dim START
  Dim FromModem As String
  Dim DUMMY
  Dim BUF
  
   counter = 0: tot_counter = 0: Label10 = 0
' If system_ready = 0 Then zyg_ADAM = "System not Working": Exit Function
  START = GetCurrentTime()
  
  
    Do
    If GetCurrentTime() - START > 10001 Then zyg_ADAM = " ": Exit Function
     MSComm1.InBufferCount = 0
      FromModem = ""
           MSComm1.Output = Balance_Asking + Chr$(13)
          MilSec (50)
          DUMMY = DoEvents()
          If MSComm1.InBufferCount Then
           BUF = MSComm1.InBufferCount
          FromModem$ = FromModem$ + MSComm1.Input
        '  List1.AddItem FromModem$, 0
           
           If InStr(FromModem$, "OL") > 0 Then
                'MsgBox "Scale Overload ...", , "Talos"
                zyg_ADAM = "OL"
                Exit Do
            End If
           If InStr(FromModem$, "UL") > 0 Then
                'MsgBox "Scale Overload ...", , "Talos"
                zyg_ADAM = "UL"
                Exit Do
            End If
            If GetCurrentTime() - START > 5000 And Zygis_Kind = "OK" Then Zygis_Kind = 0: tot_counter = 0
          
          
          
          
          
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
     MSComm1.Output = "T" + Chr(13)
        MilSec 3000
       '  MSComm1.Output = "T" + Chr(13)
End Sub

Private Sub Form_Load()
  Balance_Asking = "E"

  MSComm1.PortOpen = True
  

End Sub
