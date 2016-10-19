Attribute VB_Name = "MatrixCal"
'''MatrixPlus表示矩阵相加；MatrixSubtract表示矩阵相减；MatrixTranspose表示矩阵转置
'''MatrixMultip表示矩阵相乘；MatrixDet表示求行列式；MatrixInver表示矩阵求逆
Rem 所有矩阵下标从1开始
Option Base 1
Option Explicit

Public Sub MatrixPlus(A() As Double, B() As Double, C() As Double)              'MatrixPlus表示矩阵相加
Dim i As Integer, j As Integer, Row As Integer, Col As Integer
Row = UBound(A)
Col = UBound(A, 2)

If UBound(B) <> Row Or UBound(B, 2) <> Col Then
     MsgBox "矩阵类型不匹配！", , "消息提示": Exit Sub
 End If

For i = 1 To Row
  For j = 1 To Col
    C(i, j) = A(i, j) + B(i, j)
  Next
Next

End Sub
Public Sub MatrixSubtract(A() As Double, B() As Double, C() As Double)           'MatrixSubtract表示矩阵相减
Dim i As Integer, j As Integer, Row As Integer, Col As Integer
Row = UBound(A)
Col = UBound(A, 2)

If UBound(B) <> Row Or UBound(B, 2) <> Col Then
     MsgBox "矩阵类型不匹配！", , "消息提示": Exit Sub
 End If

For i = 1 To Row
  For j = 1 To Col
    C(i, j) = A(i, j) - B(i, j)
  Next
Next

End Sub

Public Sub MatrixTranspose(A() As Double, C() As Double)                           'MatrixTranspose表示矩阵转置
  Dim i As Integer, j As Integer
  Dim Row As Integer, Col As Integer
  
  Row = UBound(A)
  Col = UBound(A, 2)
  
  For i = 1 To Col   '行数和列数都是从1开始的
     For j = 1 To Row
        C(i, j) = A(j, i)
     Next
  Next
End Sub

Public Sub MatrixMultip(A() As Double, B() As Double, C() As Double)                'MatrixMultip表示矩阵相乘
  Dim i As Integer, j As Integer, k As Integer
  Dim Row As Integer, Col As Integer
  
  Row = UBound(A)
  Col = UBound(B, 2)
  
  If UBound(A, 2) <> UBound(B) Then
     MsgBox "请注意两个矩阵的行列数关系！", , "消息提示": Exit Sub
  End If
     
  For i = 1 To Row   '行数和列数都是从1开始的
     For j = 1 To Col
        For k = 1 To UBound(A, 2)
          C(i, j) = C(i, j) + A(i, k) * B(k, j)
        Next
     Next
  Next
  
End Sub

Public Sub MatrixDet(A() As Double, det As Double)                       'MatrixDet表示求行列式
  Dim i As Integer, j As Integer, k As Integer, Row As Integer
  Dim LL() As Double, D() As Double                                      '矩阵D是为了计算的需要，以免改变A的值
  Row = UBound(A)
  If UBound(A, 2) <> Row Then
     MsgBox "必须为方阵才能求行列式！", , "消息提示": Exit Sub
  End If
  
  ReDim D(Row, Row): det = 1
  D = A
  For i = 1 To Row      '行数和列数都是从0开始的,算到只剩二阶就停止
    
    '先判断第一个是否都为0，如果为0，则进行换行操作，用参数Flag表示,如果最后Flag为0，表示有一行全为0，行列式等于0
    Dim flag As Integer: flag = 0
    For j = i To Row
      If D(j, i) <> 0 Then flag = 1
    Next j
    If flag = 0 Then det = 0: GoTo 111
    
    
    '如果第一列不全为0，则判断第一个是否为0，如果为0，则进行换行运算
    j = i + 1
    Do While (D(i, i) = 0)
      If j <= Row Then
        Call MatrixRow(D, i, j)
        det = det * (-1) ^ (i + j)
        j = j + 1
      End If
    Loop                   '后面可以再加代码，如果j=Row+1，行列式为0
    
    '保证A(i)(i)不为0后，再用消去法求行列式
    ReDim LL(Row):
   
    For j = i + 1 To Row    '算出LL（），然后用消去法运算
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

Public Sub MatrixInver(A() As Double, C() As Double)            'MatrixInver表示矩阵求逆
  '首先判断行列式是否为0，如果为0，则不能求逆
  Dim B() As Double, LL() As Double, D() As Double              'B是增广矩阵，LL是相乘系数，A赋值给D，不改变其值
  Dim i As Integer, j As Integer, k As Integer
  Dim Row As Integer, Col As Integer
  Dim bii As Double, det As Double
  '首先进行方阵和行列式的判断，如果不是，无法求逆
  Row = UBound(A)
  If UBound(A, 2) <> Row Then
    MsgBox "矩阵不是方阵，无法求逆！", , "消息提示"
    Exit Sub
  End If
  
  ReDim D(Row, Row): D = A
  Call MatrixDet(D(), det)
  If det = 0 Then
    MsgBox "矩阵行列式为零！无法求逆！", , "消息提示"
    Exit Sub
  End If

  ReDim B(Row, 2 * Row)
  For i = 1 To Row                       '给B赋初值
    For j = 1 To Row
         B(i, j) = D(i, j)
    Next j
    
    For j = Row + 1 To 2 * Row      '增广矩阵赋值
      If j = i + Row Then
         B(i, j) = 1
       Else
         B(i, j) = 0
      End If
    Next j
  Next i
  
  For i = 1 To Row
    '先判断第一个元素是否为0,如果为0，进行换行
     j = i + 1
     Do While (B(i, i) = 0)
       If j <= Row Then
         Call MatrixRow(B, i, j)
         j = j + 1
       End If
     Loop
    
    bii = B(i, i)
    For j = i To 2 * Row
       B(i, j) = B(i, j) / bii '使得第一行的第一个元素为1
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
      C(i, j) = B(i, j + Row)  '获取逆矩阵
    Next j
  Next i
  
End Sub

Rem  矩阵运算常用的小工具；MatrixRow表示行变换，MatrixCol表示列变化
Public Sub MatrixRow(A() As Double, M As Integer, N As Integer)      '矩阵行变换,第M行和第N行交换
  Dim Row As Integer, Col As Integer, i As Integer
  Dim Temp() As Double
  
  Row = UBound(A): Col = UBound(A, 2)
  If M > Row Or N > Row Then
    MsgBox "参数输入错误！", , "消息提示": Exit Sub
  End If
    
  ReDim Temp(Col)
  For i = 1 To Col
    Temp(i) = A(M, i)
    A(M, i) = A(N, i)
    A(N, i) = Temp(i)
  Next
End Sub

Public Sub MatrixCol(A() As Double, M As Integer, N As Integer)      '矩阵列变换,第M列和第N列交换
  Dim Row As Integer, Col As Integer, i As Integer
  Dim Temp() As Double
  
  Row = UBound(A): Col = UBound(A, 2)
  If M > Col Or N > Col Then
    MsgBox "参数输入错误！", , "消息提示": Exit Sub
  End If
    
  ReDim Temp(Row)
  For i = 1 To Row
    Temp(i) = A(i, M)
    A(i, M) = A(i, N)
    A(i, N) = Temp(i)
  Next
End Sub

