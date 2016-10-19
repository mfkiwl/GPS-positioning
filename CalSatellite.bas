Attribute VB_Name = "CalSatellite"
Option Explicit

Public Xk As Double, Yk As Double, Zk As Double, Qc As Double      '��ʱ��¼����λ�ü����������������Ӳ�Qc
Public Obs_Interval As Double                                      'O�ļ��б�ʾ�۲�ʱ����
Public Obs_Type As Integer                                         'O�ļ��б�ʾ���ݹ۲�����
Public Obs_LY As Integer                                           'O�ļ��б�ʾ�۲���Ԫ����
Public Nav_LY As Integer                                           'N�ļ��б�ʾ�������ݵ��ܸ���
Public AppX As Double, AppY As Double, AppZ As Double              'O�ļ��н��ջ���������
Public Nm As Integer, Nm1 As Integer                               '��������λ��ʱ�����õ�ѭ��������Nm��ʾƥ��ɹ�����Ŀ

Public Const PI As Double = 3.14159265358979                       '��ʾԲ���ʵĲ���
Public Const u As Double = 3.9860047 * 10 ^ 14                     '����u=GM
Public Const We As Double = 7.29211567 * 10 ^ (-5)                 '������ת������
Public Const cc As Double = 299792458                              '����

Public Poi_X() As Double, Poi_Y() As Double, Poi_Z() As Double     '���ջ�����
Public PDOP() As Double, TDOP() As Double, GDOP() As Double        '��������

'Public Const u As Double = 3.986004418 * 10 ^ 14      'WGS-84����ϵ�еĵ�������������u=GM
'Public Const PI As Double = 3.1415926535898           'Բ����
'Public Const We As Double = 7.2921151 * 10 ^ -5   '������ת������

Public Sub CalSatelliteP(Time As Double, bb As Integer)
Dim A As Double, N As Double, tk As Double, Mk As Double, E1 As Double, E2 As Double, Vk As Double
Dim W1 As Double, du As Double, drr As Double, di As Double, Uk As Double, Rk As Double, Ik As Double
Dim Drk As Double, X0 As Double, Y0 As Double

'step1:    Compute the average angular velocity of satellites.
 A = (Nav_data(bb).A1) ^ 2
 N = Nav_data(bb).n1 + Sqr(u) / (Nav_data(bb).A1) ^ 3
 
'setp2:    Compute the correction of satellite clock error and the naturalization time.
 Qc = Nav_data(bb).af0 + Nav_data(bb).af1 * (Time - Nav_data(bb).Gtime) + Nav_data(bb).af2 * (Time - Nav_data(bb).Gtime) ^ 2
 tk = Time - Qc - Nav_data(bb).toe
 
 If tk > 302400 Then
    tk = tk - 604800
 ElseIf tk < -302400 Then
    tk = tk + 604800
 End If
 
'step3:    Compute the mean anomaly.
 Mk = Nav_data(bb).M0 + N * tk
 
'step4;    Compute the eccentric anomaly.
 E1 = Mk
 Do
   E2 = E1
   E1 = Mk + Nav_data(bb).e * Sin(E2)
 Loop While (Abs(E1 - E2) > 10 ^ (-12))
 
'step5:    Compute the true anomaly.
 Vk = Arct(Cos(E1) - Nav_data(bb).e, Sin(E1) * Sqr(1 - Nav_data(bb).e ^ 2))

'step6:    Compute the argument of latitude
 W1 = Vk + Nav_data(bb).w

'step7:     Compute the perturbation corrections
 du = Nav_data(bb).Cuc * Cos(2 * W1) + Nav_data(bb).Cus * Sin(2 * W1)
 drr = Nav_data(bb).Crc * Cos(2 * W1) + Nav_data(bb).Crs * Sin(2 * W1)
 di = Nav_data(bb).Cic * Cos(2 * W1) + Nav_data(bb).Cis * Sin(2 * W1)

'step8:    Compute the argument of latitude after perturbation correction, the satellite vector and the orbit inclination.
 Uk = W1 + du
 Rk = A * (1 - Nav_data(bb).e * Cos(E1)) + drr
 Ik = Nav_data(bb).i0 + di + Nav_data(bb).i1 * tk

'step9:    Compute the coordinate of satellites in orbit plane coordinate system.
 X0 = Rk * Cos(Uk)
 Y0 = Rk * Sin(Uk)

'step10:    Compute the longitude of ascending node in observed time.
 Drk = Nav_data(bb).Dr0 + (Nav_data(bb).Dr - We) * tk - We * Nav_data(bb).toe
 
'step11:    Compute the rectangular coordinates of satellites in ECEF.
 Xk = X0 * Cos(Drk) - Y0 * Cos(Ik) * Sin(Drk)
 Yk = X0 * Sin(Drk) + Y0 * Cos(Ik) * Cos(Drk)
 Zk = Y0 * Sin(Ik)

End Sub

