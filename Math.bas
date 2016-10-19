Attribute VB_Name = "Math"
'��������pi��ֵΪ3.14159265358979
Const PI = 3.14159265358979
'�������������귽λ��(��������)
Public Function Arctan(Xa, Ya, Xb, Yb) As Double
Dim dx As Double, dy As Double, D As Double
dx = Xb - Xa: dy = Yb - Ya
If Abs(dx) <= 0.000001 Then
   If dy > 0 Then A = PI / 2 Else A = PI * 3 / 2
Else
  D = dy / dx
  A = Atn(D)
  If dx > 0 And dy < 0 Then A = A + 2 * PI
  If dx < 0 Then A = A + PI
End If
Arctan = A
End Function
'�������������귽λ�ǣ���������
Public Function Arct(dx, dy) As Double
Dim D As Double
If Abs(dx) <= 0.000001 Then
   If dy > 0 Then A = PI / 2 Else A = PI * 3 / 2
Else
  D = dy / dx
  A = Atn(D)
  If dx > 0 And dy < 0 Then A = A + 2 * PI
  If dx < 0 Then A = A + PI
End If
Arct = A
End Function
Public Function ArcCos(x#)              '������
Dim y As Double
If x = 1 Then
  y = 0
  ElseIf x = -1 Then
  y = PI
  Else
  y = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)  '�����ķ����Ǻ���
End If
 ArcCos = y
End Function
'��������֪�����ε������ߣ������������ڽ�                                       ������ת��ΪSub����
Public Function FF3(Xa, Ya, Xb, Yb, Xc, Yc, A1, B1, c1) As Double
    Dim Sa#, Sb#, Sc#, cosa#, cosb#, cosc#
    Sa = Sqr((Xc - Xb) * (Xc - Xb) + (Yc - Yb) * (Yc - Yb))
    Sb = Sqr((Xc - Xa) * (Xc - Xa) + (Yc - Ya) * (Yc - Ya))
    Sc = Sqr((Xa - Xb) * (Xa - Xb) + (Ya - Yb) * (Ya - Yb))
    cosa = (Sb * Sb + Sc * Sc - Sa * Sa) / (2 * Sb * Sc)    ' �������Ҷ���
    cosb = (Sa * Sa + Sc * Sc - Sb * Sb) / (2 * Sa * Sc)
    cosc = (Sb * Sb + Sa * Sa - Sc * Sc) / (2 * Sa * Sb)
    
    A = ArcCos(cosa): B = ArcCos(cosb): C = ArcCos(cosc)    '���÷����Ǻ���
End Function

'�������Ƕ�ת��Ϊ����
Public Function DuHu(x) As Double
Dim y As Double
y = x * PI / 180
DuHu = y
End Function
'���������û�����ĽǶ���ʽת��Ϊ�Զ�Ϊ��λ����ʽ
Public Function DuTrans(D) As Single
Dim A As Integer, B As Integer, C As Integer, B1 As Single, DFM As Single, sig As Integer
If D >= 0 Then sig = 1 Else sig = -1
D = D * sig
A = Int(D)
B = Int((D - A) * 100)
B1 = B
C = (D - A - B1 / 100) * 10000
DFM = A + B1 / 60 + C / 3600
DuTrans = DFM * sig
End Function

