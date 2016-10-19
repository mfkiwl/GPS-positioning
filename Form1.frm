VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "GPS Pseudorange Positioning"
   ClientHeight    =   6465
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8595
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   6135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu File 
      Caption         =   "Files"
      Begin VB.Menu OpenFile 
         Caption         =   "Open Files"
      End
      Begin VB.Menu SaveFile 
         Caption         =   "Save Files"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Calculate 
      Caption         =   "Computation"
      Begin VB.Menu Cal 
         Caption         =   "Compute positions "
      End
   End
   Begin VB.Menu About 
      Caption         =   "about"
      Begin VB.Menu FileExplain 
         Caption         =   "Explanation"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub Cal_Click()
 Dim i As Integer, j As Integer, k As Integer
 Dim ReO() As Integer, ReN() As Integer                              '分别记录观测文件历元中的卫星编号和导航文件中的编号
 Dim p As Integer, q As Integer

 Dim LL() As Double, mm() As Double, nn() As Double, G() As Double, B1() As Double
 Dim P0() As Double, P1() As Double, Q1() As Double                  'Q1为卫星钟误差
 Dim GT() As Double, GTG() As Double, XX() As Double, d1() As Double, d2() As Double
 Dim apX As Double, apY As Double, apZ As Double                     '改正数地迭代计算
 Dim tt() As Double                                                  '卫星信号传播时间迭代计算
 Dim f1 As Double, f2 As Double                                      'L1,L2两个波段的载波频率
 
 f1 = 1575420: f2 = 1227600
 
 Nm = 0
 For i = 1 To Obs_LY                                                 'cycle of epoch
    Nm1 = 0
    For j = 1 To Obs_data(i).SateSum                                 ' cycle of the number of satellites in each epoch
       For k = 1 To Nav_LY                                           ' cycle of ephemeris
          If Obs_data(i).prn(j) = Nav_data(k).prn Then
             If Abs(Obs_data(i).Gtime - Nav_data(k).Gtime) <= 3600 Then
                Nm1 = Nm1 + 1
                ReDim Preserve ReN(Nm1)
                ReDim Preserve ReO(Nm1)
                ReN(Nm1) = k: ReO(Nm1) = j                           ' record the number of successfully matching data in O and N files
             End If
          End If
       Next k
    Next j
    
    If Nm1 >= 4 Then
       Nm = Nm + 1
       ReDim Preserve Pos(Nm): ReDim Preserve Poi_X(Nm)              'redefine arrays
       ReDim Preserve Poi_Y(Nm): ReDim Preserve Poi_Z(Nm):
       ReDim Pos(Nm).Sate(Nm1)
       ReDim Preserve PDOP(Nm): ReDim Preserve TDOP(Nm): ReDim Preserve GDOP(Nm)
       
       ReDim Q1(Nm1), P0(Nm1), P1(Nm1), B1(Nm1, 1), G(Nm1, 4)
       ReDim GT(4, Nm1), GTG(4, 4), XX(4, 1), tt(Nm1)
        
       For p = 1 To Nm1
         Call CalSatelliteP(Obs_data(i).Gtime - 0.07, ReN(p))              '卫星传播初始时间按照0.07秒来算
         Pos(Nm).Sate(p).X1 = Xk * Cos(We * 0.07) + Yk * Sin(We * 0.07)    '地球自转改正
         Pos(Nm).Sate(p).Y1 = -Xk * Sin(We * 0.07) + Cos(We * 0.07) * Yk
         Pos(Nm).Sate(p).Z1 = Zk
         
         P0(p) = Sqr((Pos(Nm).Sate(p).X1 - AppX) ^ 2 + (Pos(Nm).Sate(p).Y1 - AppY) ^ 2 + (Pos(Nm).Sate(p).Z1 - AppZ) ^ 2)
         
         Dim pp1 As Double, pp2 As Double
         pp1 = Obs_data(i).Data(ReO(p)).P1: pp2 = Obs_data(i).Data(ReO(p)).P2:
         P1(p) = (pp1 * f1 * f1 - pp2 * f2 * f2) / (f1 * f1 - f2 * f2)                               '伪距电离层改正
         
         
         G(p, 1) = -(Pos(Nm).Sate(p).X1 - AppX) / P0(p): G(p, 2) = -(Pos(Nm).Sate(p).Y1 - AppY) / P0(p)
         G(p, 3) = -(Pos(Nm).Sate(p).Z1 - AppZ) / P0(p): G(p, 4) = 1:
         B1(p, 1) = P1(p) - P0(p) + Qc * cc
 
       Next p
 
     ReDim d1(4, 4), d2(4, 1)
     Call MatrixTranspose(G(), GT())
     Call MatrixMultip(GT(), G(), GTG())
     Call MatrixInver(GTG(), d1())
     Call MatrixMultip(GT(), B1(), d2())
     Call MatrixMultip(d1(), d2(), XX())
        
     apX = AppX + XX(1, 1)
     apY = AppY + XX(2, 1)
     apZ = AppZ + XX(3, 1)
                                 
'     Do
'       Poi_X(Nm) = apX
'       Poi_Y(Nm) = apY
'       Poi_Z(Nm) = apZ
'
'      For p = 1 To Nm1
'        P0(p) = Sqr((Pos(Nm).Sate(p).X1 - apX) ^ 2 + (Pos(Nm).Sate(p).Y1 - apY) ^ 2 + (Pos(Nm).Sate(p).Z1 - apZ) ^ 2)
'        tt(p) = P0(p) / cc
'        Call CalSatelliteP(Obs_data(i).Gtime - tt(p), ReN(p))            '卫星传播时间进行迭代
'
'        Pos(Nm).Sate(p).X1 = Xk * Cos(We * tt(p)) + Yk * Sin(We * tt(p))
'        Pos(Nm).Sate(p).Y1 = -Xk * Sin(We * tt(p)) + Cos(We * tt(p)) * Yk
'        Pos(Nm).Sate(p).Z1 = Zk
'        P0(p) = Sqr((Pos(Nm).Sate(p).X1 - apX) ^ 2 + (Pos(Nm).Sate(p).Y1 - apY) ^ 2 + (Pos(Nm).Sate(p).Z1 - apZ) ^ 2)
'
'        G(p, 1) = -(Pos(Nm).Sate(p).X1 - apX) / P0(p): G(p, 2) = -(Pos(Nm).Sate(p).Y1 - apY) / P0(p)
'        G(p, 3) = -(Pos(Nm).Sate(p).Z1 - apZ) / P0(p): G(p, 4) = 1:
'        B1(p, 1) = P1(p) - P0(p) + Qc * cc
'      Next p
'
'     Call MatrixTranspose(G(), GT())
'     Call MatrixMultip(GT(), G(), GTG())
'     Call MatrixInver(GTG(), d1())
'     Call MatrixMultip(GT(), B1(), d2())
'     Call MatrixMultip(d1(), d2(), XX())
'
'     apX = apX + XX(1, 1)
'     apY = apY + XX(2, 1)
'     apZ = apZ + XX(3, 1)
'
'     Loop While (Abs(Poi_X(Nm) - apX) > 0.01 And Abs(Poi_Y(Nm) - apY) > 0.01 And Abs(Poi_Z(Nm) - apZ) > 0.01)
     
     Poi_X(Nm) = apX
     Poi_Y(Nm) = apY
     Poi_Z(Nm) = apZ
                      
     PDOP(Nm) = Sqr(d1(1, 1) + d1(2, 2) + d1(3, 3))
     TDOP(Nm) = Sqr(d1(4, 4))
     GDOP(Nm) = Sqr(d1(1, 1) + d1(2, 2) + d1(3, 3) + d1(4, 4))
                      
    End If
 Next i
 

 Text1.Text = Text1.Text + "X     Y     Z    PDOP   TDOP   GDOP" & vbCrLf
 Text1.Text = Text1.Text + "Part of the result：" & vbCrLf
 
 For i = 1 To 100
  
     Text1.Text = Text1.Text + Format(Poi_X(i), "0.00") & "    " & Format(Poi_Y(i), "0.00") & "    " & Format(Poi_Z(i), "0.00") & "    " & Format(PDOP(i), "0.00") & "    " & Format(TDOP(i), "0.00") & "    " & Format(GDOP(i), "0.00") & vbCrLf
    
 Next i

  MsgBox "Computation succeed！", , "tip"
  OpenFile.Enabled = False
  SaveFile.Enabled = True
  Cal.Enabled = False
  
End Sub

Private Sub Exit_Click()
Dim Msg As Integer
Msg = MsgBox("you want to exit?", vbYesNo + vbQuestion + vbDefaultButton2, "tip")

If Msg = vbYes Then
  End
Else
  Exit Sub
End If

End Sub

Private Sub FileExplain_Click()
  MsgBox "open O file and N file at the same time!", , "Warn"
End Sub


Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  Text1.Locked = True
  OpenFile.Enabled = True
  SaveFile.Enabled = False
  Cal.Enabled = False
End Sub

Private Sub OpenFile_Click()
  Dim FileName As String
  Dim filenames As Variant
  
  CommonDialog1.CancelError = True
  On Error GoTo 11
  CommonDialog1.InitDir = App.Path
  'CommonDialog1.Filter = "Rinex（*.txt）|*.txt"
  CommonDialog1.DialogTitle = "open files"
  CommonDialog1.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly  '不明白，后面两项可以不要
  CommonDialog1.ShowOpen
  FileName = CommonDialog1.FileName
  
  If InStr(FileName, Chr(0)) = 0 Then                    '中间没有空格，表示只选择了一个文件
     MsgBox "files error！", , "warn！": Exit Sub
  End If
  
  filenames = Split(FileName, Chr(0))
'  Text1.Text = Filenames(0)
'  Text2.Text = Filenames(1)
'  Text3.Text = Filenames(2)
  
  Dim str As String
  Open filenames(2) For Input As #1    'read O file

  Do
    Line Input #1, str
  Loop While (InStr(str, "APPROX POSITION XYZ") = 0) ' read "Approximate marker position"
  AppX = Val(Mid(str, 1, 14))
  AppY = Val(Mid(str, 15, 14))
  AppZ = Val(Mid(str, 29, 14))
  
  Do
    Line Input #1, str
  Loop While (InStr(str, "# / TYPES OF OBSERV") = 0)
  Obs_Type = Val(Mid(str, 6, 1))                       'read types of observation
  
  Do
    Line Input #1, str
  Loop While (InStr(str, "INTERVAL") = 0)
  Obs_Interval = Val(Mid(str, 5, 7))
  
  Do
    Line Input #1, str
  Loop While (InStr(str, "END OF HEADER") = 0)
  
  Dim i As Integer, j As Integer
  i = 0: j = 0
  
  Do While Not EOF(1)
    i = i + 1
    ReDim Preserve Obs_data(i)
    Line Input #1, str
    Obs_data(i).Year = Val(Mid(str, 1, 3))             'read year
    If Obs_data(i).Year > 80 Then
       Obs_data(i).Year = Obs_data(i).Year + 1900
    Else
       Obs_data(i).Year = Obs_data(i).Year + 2000      'year in O file is in 2 digits, need to add completely
    End If
    Obs_data(i).Month = Val(Mid(str, 4, 3))            'read month
    Obs_data(i).Day = Val(Mid(str, 7, 3))
    Obs_data(i).Hour = Val(Mid(str, 10, 3))
    Obs_data(i).Minute = Val(Mid(str, 13, 3))
    Obs_data(i).Second = Val(Mid(str, 16, 11))
    Obs_data(i).SateQua = Val(Mid(str, 27, 3))
    Obs_data(i).SateSum = Val(Mid(str, 30, 3))
    Obs_data(i).Gtime = GPSTIME(Obs_data(i).Year, Obs_data(i).Month, Obs_data(i).Day, Obs_data(i).Hour, Obs_data(i).Minute, Obs_data(i).Second)
    
    ReDim Obs_data(i).prn(Obs_data(i).SateSum)               'redefine the satellite number in each epoch
    For j = 1 To Obs_data(i).SateSum
        Obs_data(i).prn(j) = Val(Mid(str, 31 + 3 * j, 2))    'read satellite num. in each epoch
    Next
                                        
    ReDim Obs_data(i).Data(Obs_data(i).SateSum)
    For j = 1 To Obs_data(i).SateSum
      Line Input #1, str
      Obs_data(i).Data(j).P1 = Val(Mid(str, 1, 16))           'read code or phase
      Obs_data(i).Data(j).P2 = Val(Mid(str, 17, 16))
      Obs_data(i).Data(j).L1 = Val(Mid(str, 33, 16))
      Obs_data(i).Data(j).L2 = Val(Mid(str, 49, 16))
    Next j
  Loop
    Obs_LY = i                                                'read the number of the epoch
    Close #1
    
    Open filenames(1) For Input As #2                         'read N file
    Do
      Line Input #2, str
    Loop While (InStr(str, "END OF HEADER") = 0)
    
    Dim k As Integer: k = 0
    Do While Not EOF(2)
       k = k + 1                                              'read the number of epoch
       ReDim Preserve Nav_data(k)                             'redefine message array
       Line Input #2, str
       Nav_data(k).prn = Val(Mid(str, 1, 3))
       Nav_data(k).Year = Val(Mid(str, 4, 3))
            If Nav_data(k).Year > 80 Then
              Nav_data(k).Year = Nav_data(k).Year + 1900
            Else
              Nav_data(k).Year = Nav_data(k).Year + 2000
            End If
       Nav_data(k).Month = Val(Mid(str, 7, 3))
       Nav_data(k).Day = Val(Mid(str, 10, 3))
       Nav_data(k).Hour = Val(Mid(str, 13, 3))
       Nav_data(k).Minute = Val(Mid(str, 16, 3))
       Nav_data(k).Second = Val(Mid(str, 19, 4))
       Nav_data(k).af0 = Val(Mid(str, 23, 19))
       Nav_data(k).af1 = Val(Mid(str, 42, 19))
       Nav_data(k).af2 = Val(Mid(str, 61, 19))                'read line 1 with date
       
       Line Input #2, str
       Nav_data(k).aode = Val(Mid(str, 1, 22))                'read line 2
       Nav_data(k).Crs = Val(Mid(str, 23, 19))
       Nav_data(k).n1 = Val(Mid(str, 42, 19))
       Nav_data(k).M0 = Val(Mid(str, 61, 19))
       
       Line Input #2, str                                     'read line 3
       Nav_data(k).Cuc = Val(Mid(str, 1, 22))
       Nav_data(k).e = Val(Mid(str, 23, 19))
       Nav_data(k).Cus = Val(Mid(str, 42, 19))
       Nav_data(k).A1 = Val(Mid(str, 61, 19))
       
       Line Input #2, str                                     'read line 4
       Nav_data(k).toe = Val(Mid(str, 1, 22))
       Nav_data(k).Cic = Val(Mid(str, 23, 19))
       Nav_data(k).Dr0 = Val(Mid(str, 42, 19))
       Nav_data(k).Cis = Val(Mid(str, 61, 19))
       
       Line Input #2, str                                     'read line 5
       Nav_data(k).i0 = Val(Mid(str, 1, 22))
       Nav_data(k).Crc = Val(Mid(str, 23, 19))
       Nav_data(k).w = Val(Mid(str, 42, 19))
       Nav_data(k).Dr = Val(Mid(str, 61, 19))
       
       Line Input #2, str                                     'read line 6
       Nav_data(k).i1 = Val(Mid(str, 1, 22))
       Nav_data(k).cflgl2 = Val(Mid(str, 23, 19))
       Nav_data(k).weekno = Val(Mid(str, 42, 19))
       Nav_data(k).pflgl2 = Val(Mid(str, 61, 19))
       
       Line Input #2, str                                     'read line 7
       Nav_data(k).svacc = Val(Mid(str, 1, 22))
       Nav_data(k).svhlth = Val(Mid(str, 23, 19))
       Nav_data(k).tgd = Val(Mid(str, 42, 19))
       Nav_data(k).aodc = Val(Mid(str, 61, 19))
       
       Line Input #2, str                                     'read line 8
       Nav_data(k).ttm = Val(Mid(str, 1, 22))

       Nav_data(k).Gtime = GPSTIME(Nav_data(k).Year, Nav_data(k).Month, Nav_data(k).Day, Nav_data(k).Hour, Nav_data(k).Minute, Nav_data(k).Second)
    
    Loop
    
    Nav_LY = k
    Close #2
    
    MsgBox "files read successfully！", , "tip"
    OpenFile.Enabled = False
    SaveFile.Enabled = False
    Cal.Enabled = True
    
11:

End Sub

Private Sub SaveFile_Click()
 Dim i As Integer
 Dim FileName As String
 
 CommonDialog1.CancelError = True
 On Error GoTo 111
 CommonDialog1.InitDir = App.Path

 CommonDialog1.Filter = "text（*.txt）|*.txt"
 CommonDialog1.DialogTitle = "save files"
 CommonDialog1.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer + cdlOFNFileMustExist + cdlOFNHideReadOnly  '不明白，后面两项可以不要
 CommonDialog1.ShowSave

 Open CommonDialog1.FileName For Output As #2
 Print #2, "  X" & "         " & "Y" & "         " & "Z" & "       " & "PDOP" & "   " & "TDOP" & "   " & "GDOP" & vbCrLf
 For i = 1 To Nm

     Print #2, Format(Poi_X(i), "0.00") & "    " & Format(Poi_Y(i), "0.00") & "    " & Format(Poi_Z(i), "0.00") & "    " & Format(PDOP(i), "0.00") & "    " & Format(TDOP(i), "0.00") & "    " & Format(GDOP(i), "0.00") & vbCrLf
   
 Next i
 Close #2
 MsgBox "saved successfully" & CommonDialog1.FileName, , "tip"
   
111:
  OpenFile.Enabled = False
  Cal.Enabled = False

End Sub


Rem   ///////////
Rem 实验数据
'Dim y As Integer, M As Integer, D As Integer, H As Integer, F As Integer, S As Integer, Z As Long, Miao As Long
'Dim b As Variant, A As Variant
'  b = Text1.Text
'  A = Split(b, " ")
'  y = A(0): M = A(1): D = A(2): H = A(3): F = A(4): S = A(5)
'  Call GPSTIME(y, M, D, H, F, S, Z, Miao)
'  Text2.Text = Z
'  Text3.Text = Miao
'Text2.Text = GPSTIME(2004, 5, 1, 10, 5, 15)
'Text2.Text = GPSTIME(2009, 1, 1, 2, 0, 0)


'Dim tt0 As Double
'ReDim Nav_data(1)
'Nav_data(1).af0 = -0.231899321079 * 10 ^ (-6): Nav_data(1).af1 = 0: Nav_data(1).af2 = 0
'Nav_data(1).toe = 7200#: Nav_data(1).A1 = 5153.65263176: Nav_data(1).e = 0.0067842121
'Nav_data(1).i0 = 0.958512: Nav_data(1).w = -2.584194: Nav_data(1).Dr0 = -1.3783598
'Nav_data(1).M0 = -0.290282: Nav_data(1).n1 = 0.45141166 * 10 ^ (-8): Nav_data(1).Dr = -0.81942699 * 10 ^ (-8)
'Nav_data(1).i1 = -0.25393914 * 10 ^ (-9)
'Nav_data(1).Cus = 0.91213733 * 10 ^ (-5): Nav_data(1).Cuc = 0.1899898 * 10 ^ (-6): Nav_data(1).Cis = 0.949949026 * 10 ^ (-7)
'Nav_data(1).Cic = 0.13038516 * 10 ^ (-7): Nav_data(1).Crs = 0.40625 * 10 ^ (1): Nav_data(1).Crc = 0.201875 * 10 ^ (3)
'tt0 = GPSTIME(1997, 11, 9, 2, 0, 0)
'Call CalSatelliteP(tt0 - 0.07, 1)
'Text1.Text = Xk: Text2.Text = Yk: Text3.Text = Zk

'Dim A(1 To 2, 1 To 2) As Double
'A(1, 1) = 4: A(1, 2) = 5: A(2, 1) = 7: A(2, 2) = 4
'Text1.Text = MatrixTranspose(A)
