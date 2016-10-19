Attribute VB_Name = "DataExplain"
Option Base 1
Option Explicit

Public Type Obs1
  P1 As Double
  P2 As Double
  L1 As Double
  L2 As Double
End Type

Public Type ObservationData
  Year As Integer
  Month As Integer
  Day As Integer
  Hour As Integer
  Minute As Integer
  Second As Double
  SateSum As Integer
  SateQua As Integer
  Gtime As Double
  prn() As Double
  Data() As Obs1          'O文件里的数据项，包括P1、P2、L1、L2
End Type
Public Obs_data() As ObservationData

Public Type NavData       'N文件中的数据参数
  prn As Integer
  Year As Integer
  Month As Integer
  Day As Integer
  Hour As Integer
  Minute As Integer
  Second As Double
  af0 As Double
  af1 As Double
  af2 As Double
  aode As Double
  Crs As Double
  n1 As Double
  M0 As Double
  Cuc As Double
  e As Double
  Cus As Double
  A1 As Double
  toe As Double
  Cic As Double
  Dr0 As Double
  Cis As Double
  i0 As Double
  Crc As Double
  w As Double
  Dr As Double
  i1 As Double
  cflgl2 As Double
  weekno As Double
  pflgl2 As Double
  svacc As Double
  svhlth As Double
  tgd As Double
  aodc As Double
  ttm As Double
  
  Gtime As Double
End Type
Public Nav_data() As NavData


Public Type PositionData
  X1 As Double
  Y1 As Double
  Z1 As Double
End Type

Public Type SatellitePosition
  Sate() As PositionData                     '表示卫星坐标
End Type
Public Pos() As SatellitePosition
