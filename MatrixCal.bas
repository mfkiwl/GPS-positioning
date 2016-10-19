Attribute VB_Name = "MatrixCal"
'''MatrixPlus��ʾ������ӣ�MatrixSubtract��ʾ���������MatrixTranspose��ʾ����ת��
'''MatrixMultip��ʾ������ˣ�MatrixDet��ʾ������ʽ��MatrixInver��ʾ��������
Rem ���о����±��1��ʼ
Option Base 1
Option Explicit

Public Sub MatrixPlus(A() As Double, B() As Double, C() As Double)              'MatrixPlus��ʾ�������
Dim i As Integer, j As Integer, Row As Integer, Col As Integer
Row = UBound(A)
Col = UBound(A, 2)

If UBound(B) <> Row Or UBound(B, 2) <> Col Then
     MsgBox "�������Ͳ�ƥ�䣡", , "��Ϣ��ʾ": Exit Sub
 End If

For i = 1 To Row
  For j = 1 To Col
    C(i, j) = A(i, j) + B(i, j)
  Next
Next

End Sub
Public Sub MatrixSubtract(A() As Double, B() As Double, C() As Double)           'MatrixSubtract��ʾ�������
Dim i As Integer, j As Integer, Row As Integer, Col As Integer
Row = UBound(A)
Col = UBound(A, 2)

If UBound(B) <> Row Or UBound(B, 2) <> Col Then
     MsgBox "�������Ͳ�ƥ�䣡", , "��Ϣ��ʾ": Exit Sub
 End If

For i = 1 To Row
  For j = 1 To Col
    C(i, j) = A(i, j) - B(i, j)
  Next
Next

End Sub

Public Sub MatrixTranspose(A() As Double, C() As Double)                           'MatrixTranspose��ʾ����ת��
  Dim i As Integer, j As Integer
  Dim Row As Integer, Col As Integer
  
  Row = UBound(A)
  Col = UBound(A, 2)
  
  For i = 1 To Col   '�������������Ǵ�1��ʼ��
     For j = 1 To Row
        C(i, j) = A(j, i)
     Next
  Next
End Sub

Public Sub MatrixMultip(A() As Double, B() As Double, C() As Double)                'MatrixMultip��ʾ�������
  Dim i As Integer, j As Integer, k As Integer
  Dim Row As Integer, Col As Integer
  
  Row = UBound(A)
  Col = UBound(B, 2)
  
  If UBound(A, 2) <> UBound(B) Then
     MsgBox "��ע�������������������ϵ��", , "��Ϣ��ʾ": Exit Sub
  End If
     
  For i = 1 To Row   '�������������Ǵ�1��ʼ��
     For j = 1 To Col
        For k = 1 To UBound(A, 2)
          C(i, j) = C(i, j) + A(i, k) * B(k, j)
        Next
     Next
  Next
  
End Sub

Public Sub MatrixDet(A() As Double, det As Double)                       'MatrixDet��ʾ������ʽ
  Dim i As Integer, j As Integer, k As Integer, Row As Integer
  Dim LL() As Double, D() As Double                                      '����D��Ϊ�˼������Ҫ������ı�A��ֵ
  Row = UBound(A)
  If UBound(A, 2) <> Row Then
     MsgBox "����Ϊ�������������ʽ��", , "��Ϣ��ʾ": Exit Sub
  End If
  
  ReDim D(Row, Row): det = 1
  D = A
  For i = 1 To Row      '�������������Ǵ�0��ʼ��,�㵽ֻʣ���׾�ֹͣ
    
    '���жϵ�һ���Ƿ�Ϊ0�����Ϊ0������л��в������ò���Flag��ʾ,������FlagΪ0����ʾ��һ��ȫΪ0������ʽ����0
    Dim flag As Integer: flag = 0
    For j = i To Row
      If D(j, i) <> 0 Then flag = 1
    Next j
    If flag = 0 Then det = 0: GoTo 111
    
    
    '�����һ�в�ȫΪ0�����жϵ�һ���Ƿ�Ϊ0�����Ϊ0������л�������
    j = i + 1
    Do While (D(i, i) = 0)
      If j <= Row Then
        Call MatrixRow(D, i, j)
        det = det * (-1) ^ (i + j)
        j = j + 1
      End If
    Loop                   '��������ټӴ��룬���j=Row+1������ʽΪ0
    
    '��֤A(i)(i)��Ϊ0��������ȥ��������ʽ
    ReDim LL(Row):
   
    For j = i + 1 To Row    '���LL������Ȼ������ȥ������
      LL(j) = D(j, i) / D(i, i)
      For k = i To Row
        D(j, k) = D(j, k) - LL(j) * D(i, k)
      Next k
    Next j
  Next i
  
  For i = 1 To Row
    det = det * D(i, i)
  Next i
  
111:

End Sub

Public Sub MatrixInver(A() As Double, C() As Double)            'MatrixInver��ʾ��������
  '�����ж�����ʽ�Ƿ�Ϊ0�����Ϊ0����������
  Dim B() As Double, LL() As Double, D() As Double              'B���������LL�����ϵ����A��ֵ��D�����ı���ֵ
  Dim i As Integer, j As Integer, k As Integer
  Dim Row As Integer, Col As Integer
  Dim bii As Double, det As Double
  '���Ƚ��з��������ʽ���жϣ�������ǣ��޷�����
  Row = UBound(A)
  If UBound(A, 2) <> Row Then
    MsgBox "�����Ƿ����޷����棡", , "��Ϣ��ʾ"
    Exit Sub
  End If
  
  ReDim D(Row, Row): D = A
  Call MatrixDet(D(), det)
  If det = 0 Then
    MsgBox "��������ʽΪ�㣡�޷����棡", , "��Ϣ��ʾ"
    Exit Sub
  End If

  ReDim B(Row, 2 * Row)
  For i = 1 To Row                       '��B����ֵ
    For j = 1 To Row
         B(i, j) = D(i, j)
    Next j
    
    For j = Row + 1 To 2 * Row      '�������ֵ
      If j = i + Row Then
         B(i, j) = 1
       Else
         B(i, j) = 0
      End If
    Next j
  Next i
  
  For i = 1 To Row
    '���жϵ�һ��Ԫ���Ƿ�Ϊ0,���Ϊ0�����л���
     j = i + 1
     Do While (B(i, i) = 0)
       If j <= Row Then
         Call MatrixRow(B, i, j)
         j = j + 1
       End If
     Loop
    
    bii = B(i, i)
    For j = i To 2 * Row
       B(i, j) = B(i, j) / bii 'ʹ�õ�һ�еĵ�һ��Ԫ��Ϊ1
    Next
    
    ReDim LL(Row)
    For j = 1 To Row
       LL(j) = B(j, i)
       If j <> i Then
            For k = i To 2 * Row
              B(j, k) = B(j, k) - LL(j) * B(i, k)
            Next k
       End If
    Next j
  Next i
  
  For i = 1 To Row
    For j = 1 To Row
      C(i, j) = B(i, j + Row)  '��ȡ�����
    Next j
  Next i
  
End Sub

Rem  �������㳣�õ�С���ߣ�MatrixRow��ʾ�б任��MatrixCol��ʾ�б仯
Public Sub MatrixRow(A() As Double, M As Integer, N As Integer)      '�����б任,��M�к͵�N�н���
  Dim Row As Integer, Col As Integer, i As Integer
  Dim Temp() As Double
  
  Row = UBound(A): Col = UBound(A, 2)
  If M > Row Or N > Row Then
    MsgBox "�����������", , "��Ϣ��ʾ": Exit Sub
  End If
    
  ReDim Temp(Col)
  For i = 1 To Col
    Temp(i) = A(M, i)
    A(M, i) = A(N, i)
    A(N, i) = Temp(i)
  Next
End Sub

Public Sub MatrixCol(A() As Double, M As Integer, N As Integer)      '�����б任,��M�к͵�N�н���
  Dim Row As Integer, Col As Integer, i As Integer
  Dim Temp() As Double
  
  Row = UBound(A): Col = UBound(A, 2)
  If M > Col Or N > Col Then
    MsgBox "�����������", , "��Ϣ��ʾ": Exit Sub
  End If
    
  ReDim Temp(Row)
  For i = 1 To Row
    Temp(i) = A(i, M)
    A(i, M) = A(i, N)
    A(i, N) = Temp(i)
  Next
End Sub

