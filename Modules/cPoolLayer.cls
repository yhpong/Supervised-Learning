VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPoolLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pinput_Lx As Long, pinput_Ly As Long, pinput_Lz As Long, pinput_L As Long
Private poutput_Lx As Long, poutput_Ly As Long, poutput_Lz As Long, poutput_L As Long
Private pLx As Long, pLy As Long, pstride As Long, pL As Long
Private ptype As String
Private pDelta As Variant
Private pMaxNodes As Variant 'pointer to selected nodes from previous layer


Sub Print_Model(vRng As Range, n As Long)
    With vRng
        .Offset(0, 0).Value = "POOLLAYER"
        .Offset(1, 0).Value = pinput_Lx
        .Offset(2, 0).Value = pinput_Ly
        .Offset(3, 0).Value = pinput_Lz
        .Offset(4, 0).Value = pLx
        .Offset(5, 0).Value = pLy
        .Offset(6, 0).Value = poutput_Lx
        .Offset(7, 0).Value = poutput_Ly
        .Offset(8, 0).Value = pstride
        .Offset(9, 0).Value = ptype
    End With
    n = 1
End Sub


Sub Read_Model(vRng As Range, n As Long)
    With vRng
        If .Offset(0, 0).Value <> "POOLLAYER" Then
            Debug.Print "Column is not labeled as POOLLAYER."
            Exit Sub
        End If
        pinput_Lx = .Offset(1, 0).Value
        pinput_Ly = .Offset(2, 0).Value
        pinput_Lz = .Offset(3, 0).Value
        pLx = .Offset(4, 0).Value
        pLy = .Offset(5, 0).Value
        poutput_Lx = .Offset(6, 0).Value
        poutput_Ly = .Offset(7, 0).Value
        pstride = .Offset(8, 0).Value
        ptype = .Offset(9, 0).Value
        poutput_Lz = pinput_Lz
        pinput_L = pinput_Lx * pinput_Ly
        pL = pLx * pLy
        poutput_L = poutput_Lx * poutput_Ly
    End With
    n = 1
End Sub


Sub Init_Train()
    If ptype = "MAX" Then ReDim pMaxNodes(1 To poutput_Lz)
End Sub

Sub Reset_Train()
    If ptype = "MAX" Then Erase pMaxNodes
End Sub


Sub Init(input_Lx As Long, input_Ly As Long, input_Lz As Long, _
            Lx As Long, Ly As Long, stride As Long, Optional ltype As String = "MAX")
    If (input_Lx - Lx) Mod stride <> 0 Or _
        (input_Ly - Ly) Mod stride <> 0 Then
        Debug.Print "Pooling Layer Init fail: Adjust pooling window and stride length"
        Exit Sub
    End If
    pinput_Lx = input_Lx
    pinput_Ly = input_Ly
    pinput_Lz = input_Lz
    ptype = VBA.UCase(ltype)
    pLx = Lx
    pLy = Ly
    pstride = stride
    poutput_Lx = (pinput_Lx - pLx) / pstride + 1
    poutput_Ly = (pinput_Ly - pLy) / pstride + 1
    poutput_Lz = pinput_Lz
    pinput_L = input_Lx * input_Ly
    pL = Lx * Ly
    poutput_L = poutput_Lx * poutput_Ly
End Sub


Sub InOut(x As Variant, v As Variant)
Dim i As Long, j As Long, k As Long, jj As Long
Dim i2 As Long, j2 As Long
Dim m As Long, n As Long
Dim v1() As Double, x1() As Double
    ReDim v(1 To poutput_Lz)
    ReDim v1(1 To poutput_L)
    If ptype = "MAX" Then
        For k = 1 To poutput_Lz
            x1 = x(k)
            jj = 0
            For j = 0 To (pinput_Ly - pLy) Step pstride
                For i = 0 To (pinput_Lx - pLx) Step pstride
                    jj = jj + 1
                    v1(jj) = Pool_Step(x1, i, j, m)
                Next i
            Next j
            v(k) = v1
        Next k
    End If
    Erase v1, x1
End Sub


Sub Fwd_Pass(x As Variant, v As Variant)
Dim i As Long, j As Long, k As Long, jj As Long
Dim i2 As Long, j2 As Long
Dim m As Long, n As Long
Dim v1() As Double, x1() As Double, max_ij() As Long
    ReDim v(1 To poutput_Lz)
    ReDim v1(1 To poutput_L)
    If ptype = "MAX" Then
        ReDim max_ij(1 To poutput_L)
        For k = 1 To poutput_Lz
            x1 = x(k)
            jj = 0
            For j = 0 To (pinput_Ly - pLy) Step pstride
                For i = 0 To (pinput_Lx - pLx) Step pstride
                    jj = jj + 1
                    v1(jj) = Pool_Step(x1, i, j, m)
                    max_ij(jj) = m
                Next i
            Next j
            v(k) = v1
            pMaxNodes(k) = max_ij
        Next k
    End If
    Erase x1, v1
End Sub

Private Function Pool_Step(x() As Double, i As Long, j As Long, i_max As Long) As Double
Dim p As Long, q As Long, jj As Long, k As Long
Dim tmp_x As Double
    Pool_Step = -999999
    i_max = 0
    jj = i + (j - 1) * pinput_Lx
    For q = 1 To pLy
        k = jj + q * pinput_Lx
        For p = 1 To pLx
            tmp_x = x(p + k)
            If tmp_x > Pool_Step Then
                Pool_Step = tmp_x
                i_max = p + k
            End If
        Next p
    Next q
End Function


Sub Calc_Delta(DeltaNext As Variant)
    pDelta = DeltaNext
End Sub


Sub Bkwd_Pass(prevLayer As cConvLayer)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim i2 As Long, j2 As Long
Dim DeltaPrev As Variant, delta1() As Double, delta2() As Double, max_ij() As Long
    ReDim DeltaPrev(1 To pinput_Lz)
    If ptype = "MAX" Then
        For k = 1 To poutput_Lz
            max_ij = pMaxNodes(k)
            delta2 = pDelta(k)
            ReDim delta1(1 To pinput_L)
            For j = 1 To poutput_Ly
                j2 = (j - 1) * poutput_Lx
                For i = 1 To poutput_Lx
                    i2 = max_ij(i + j2)
                    delta1(i2) = delta1(i2) + delta2(i + j2)
                Next i
            Next j
            DeltaPrev(k) = delta1
        Next k
        Erase max_ij
    End If
    Call prevLayer.Calc_Delta(DeltaPrev)
    Erase DeltaPrev, delta1, delta2
End Sub


Public Property Get input_dim(k As Long) As Long
    If k = 1 Then
        input_dim = pinput_Lx
    ElseIf k = 2 Then
        input_dim = pinput_Ly
    ElseIf k = 3 Then
        input_dim = pinput_Lz
    End If
End Property

Public Property Get output_dim(k As Long) As Long
    If k = 1 Then
        output_dim = poutput_Lx
    ElseIf k = 2 Then
        output_dim = poutput_Ly
    ElseIf k = 3 Then
        output_dim = poutput_Lz
    End If
End Property

Public Property Get filter_dim(k As Long) As Long
    If k = 1 Then
        filter_dim = pLx
    ElseIf k = 2 Then
        filter_dim = pLy
    End If
End Property

Public Property Get stride() As Long
    stride = pstride
End Property
