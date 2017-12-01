VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConvLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pinput_Lx As Long, pinput_Ly As Long, pinput_Lz As Long, pinput_L As Long
Private poutput_Lx As Long, poutput_Ly As Long, poutput_L As Long
Private pfilter_Lx As Long, pfilter_Ly As Long, pn_filter As Long, pfilter_L As Long
Private pstride As Long
Private pwgts As Variant, pwgts_chg As Variant
Private pbias() As Double, pbias_chg() As Double
Private pDelta As Variant
Private pActivation As String

Sub Print_Model(vRng As Range, n_col As Long)
Dim i As Long, n As Long, k As Long
Dim w1() As Double
    With vRng
        .Offset(0, 0).Value = "CONVLAYER"
        .Offset(1, 0).Value = pinput_Lx
        .Offset(2, 0).Value = pinput_Ly
        .Offset(3, 0).Value = pinput_Lz
        .Offset(4, 0).Value = pfilter_Lx
        .Offset(5, 0).Value = pfilter_Ly
        .Offset(6, 0).Value = pn_filter
        .Offset(7, 0).Value = poutput_Lx
        .Offset(8, 0).Value = poutput_Ly
        .Offset(9, 0).Value = pstride
        .Offset(10, 0).Value = pActivation
        i = 12
        For n = 1 To pn_filter
            .Offset(i, 0).Value = pbias(n)
            i = i + 1
        Next n
        n_col = 1
        For n = 1 To pn_filter
            For k = 1 To pinput_Lz
                w1 = pwgts(n)(k)
                .Offset(i, n_col - 1).Resize(pfilter_L, 1).Value = _
                    Application.WorksheetFunction.Transpose(w1)
                i = i + pfilter_L
                If (i + pfilter_L - 1) > 1048576 Then
                    n_col = n_col + 1
                    i = 12
                End If
            Next k
        Next n
        .Offset(11, 0).Value = n_col
    End With
End Sub


Sub Read_Model(vRng As Range, n_col As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim w As Variant, w1() As Double, vArr As Variant
    With vRng
        If .Offset(0, 0).Value <> "CONVLAYER" Then
            Debug.Print "Column is not labeled as CONVLAYER."
            Exit Sub
        End If
        pinput_Lx = .Offset(1, 0).Value
        pinput_Ly = .Offset(2, 0).Value
        pinput_Lz = .Offset(3, 0).Value
        pfilter_Lx = .Offset(4, 0).Value
        pfilter_Ly = .Offset(5, 0).Value
        pn_filter = .Offset(6, 0).Value
        poutput_Lx = .Offset(7, 0).Value
        poutput_Ly = .Offset(8, 0).Value
        pstride = .Offset(9, 0).Value
        pActivation = .Offset(10, 0).Value
        n_col = .Offset(11, 0).Value

        pinput_L = pinput_Lx * pinput_Ly
        poutput_L = poutput_Lx * poutput_Ly
        pfilter_L = pfilter_Lx * pfilter_Ly
        ReDim pbias(1 To pn_filter)
        ReDim pwgts(1 To pn_filter)
        ReDim w(1 To pinput_Lz)
        ReDim w1(1 To pfilter_L)

        i = 12
        For n = 1 To pn_filter
            pbias(n) = .Offset(i, 0).Value
            i = i + 1
        Next n
        m = 1

        For n = 1 To pn_filter
            For k = 1 To pinput_Lz
                vArr = .Offset(i, m - 1).Resize(pfilter_L, 1).Value
                If pfilter_L = 1 Then
                    w1(1) = vArr
                Else
                    For j = 1 To pfilter_L
                        w1(j) = vArr(j, 1)
                    Next j
                End If
                i = i + pfilter_L
                If (i + pfilter_L - 1) > 1048576 Then
                    m = m + 1
                    i = 12
                End If
                w(k) = w1
            Next k
            pwgts(n) = w
        Next n
 
        Erase w, w1
    End With
End Sub


Sub Init(input_Lx As Long, input_Ly As Long, input_Lz As Long, _
        n_filter As Long, filter_Lx As Long, filter_Ly As Long, stride As Long, _
        Optional strActivation As String = "RELU")
Dim i As Long, j As Long, k As Long, n As Long
Dim w As Variant, w1() As Double

    If (input_Lx - filter_Lx) Mod stride <> 0 Or _
        (input_Ly - filter_Ly) Mod stride <> 0 Then
        Debug.Print "Conv Layer Init fail: Adjust filter window and stride length"
        Exit Sub
    End If

    pinput_Lx = input_Lx
    pinput_Ly = input_Ly
    pinput_Lz = input_Lz
    pn_filter = n_filter
    pfilter_Lx = filter_Lx
    pfilter_Ly = filter_Ly
    pstride = stride
    poutput_Lx = (pinput_Lx - pfilter_Lx) / pstride + 1
    poutput_Ly = (pinput_Ly - pfilter_Ly) / pstride + 1
    pActivation = VBA.UCase(strActivation)
    
    pinput_L = input_Lx * input_Ly
    poutput_L = poutput_Lx * poutput_Ly
    pfilter_L = filter_Lx * filter_Ly
    
    Randomize
    ReDim pwgts(1 To pn_filter)
    For n = 1 To pn_filter
        ReDim w(1 To pinput_Lz)
        ReDim w1(1 To pfilter_L)
        For k = 1 To pinput_Lz
            For i = 1 To pfilter_L
                w1(i) = -0.1 + Rnd() * 0.2
            Next i
            w(k) = w1
        Next k
        pwgts(n) = w
    Next n
    
    ReDim pbias(1 To pn_filter)
End Sub

Sub Reset()
    Erase pwgts, pbias
End Sub

Sub Init_Train()
Dim n As Long, k As Long
Dim w As Variant, w1() As Double
    ReDim pbias_chg(1 To pn_filter)
    ReDim pwgts_chg(1 To pn_filter)
    ReDim w(1 To pinput_Lz)
    ReDim w1(1 To pfilter_L)
    For k = 1 To pinput_Lz
        w(k) = w1
    Next k
    For n = 1 To pn_filter
        pwgts_chg(n) = w
    Next n
End Sub

Sub Reset_Train()
    Erase pwgts_chg, pbias_chg
End Sub

Private Function Convol_Step(x() As Double, w() As Double, i As Long, j As Long) As Double
Dim p As Long, q As Long, jj As Long, k As Long, kk As Long
    Convol_Step = 0
    jj = i + (j - 1) * pinput_Lx
    For q = 1 To pfilter_Ly
        k = (q - 1) * pfilter_Lx
        kk = jj + q * pinput_Lx
        For p = 1 To pfilter_Lx
            Convol_Step = Convol_Step + w(p + k) * x(p + kk)
        Next p
    Next q
End Function

Sub InOut(x As Variant, v As Variant)
Dim i As Long, j As Long, k As Long, i2 As Long, j2 As Long, jj As Long
Dim p As Long, q As Long, r As Long
Dim tmp_x As Double, tmp_y As Double, bias_k As Double
Dim w1() As Double, x1() As Double, v1() As Double
    ReDim v(1 To pn_filter)
    tmp_y = 0
    For k = 1 To pn_filter
        bias_k = pbias(k)
        ReDim v1(1 To poutput_L)
        For r = 1 To pinput_Lz
            w1 = pwgts(k)(r) 'w(r)
            x1 = x(r)
            jj = 0
            For j = 0 To (pinput_Ly - pfilter_Ly) Step pstride
                For i = 0 To (pinput_Lx - pfilter_Lx) Step pstride
                    jj = jj + 1
                    v1(jj) = v1(jj) + Convol_Step(x1, w1, i, j)
                Next i
            Next
        Next r
        If pActivation = "RELU" Then
            For i = 1 To poutput_L
                tmp_x = v1(i) + bias_k
                If tmp_x >= 0 Then
                    v1(i) = tmp_x
                Else
                    v1(i) = 0
                End If
            Next i
        ElseIf pActivation = "SIGMOID" Then
            For i = 1 To poutput_L
                v1(i) = 1# / (1 + Exp(-(v1(i) + bias_k)))
            Next i
        ElseIf pActivation = "SOFTMAX" Then
            For i = 1 To poutput_Lx
                v1(i) = Exp(v1(i) + bias_k)
                tmp_y = tmp_y + v1(i)
            Next i
        End If
        v(k) = v1
    Next k
    'Normalize sofmax layer to 1
    If pActivation = "SOFTMAX" Then
        For k = 1 To pn_filter
            v1 = v(k)
            For i = 1 To poutput_L
                v1(i) = v1(i) / tmp_y
            Next i
            v(k) = v1
        Next k
    End If
End Sub


Sub Fwd_Pass(x As Variant, v As Variant, Optional y_tgt As Variant)
Dim i As Long, j As Long, k As Long, i2 As Long, j2 As Long, jj As Long
Dim p As Long, q As Long, r As Long
Dim tmp_x As Double, tmp_y As Double, bias_k As Double, softmax_sum As Double
Dim w1() As Double, x1() As Double, v1() As Double, delta1() As Double, y1() As Double
    ReDim v(1 To pn_filter)
    ReDim pDelta(1 To pn_filter)
    ReDim delta1(1 To poutput_L)
    tmp_y = 0
    For k = 1 To pn_filter
        bias_k = pbias(k)
        ReDim v1(1 To poutput_L)
        For r = 1 To pinput_Lz
            w1 = pwgts(k)(r)
            x1 = x(r)
            jj = 0
            For j = 0 To (pinput_Ly - pfilter_Ly) Step pstride
                For i = 0 To (pinput_Lx - pfilter_Lx) Step pstride
                    jj = jj + 1
                    v1(jj) = v1(jj) + Convol_Step(x1, w1, i, j)
                Next i
            Next j
        Next r
        If pActivation = "RELU" Then
            For i = 1 To poutput_L
                tmp_x = v1(i) + bias_k
                If tmp_x >= 0 Then
                    v1(i) = tmp_x
                    delta1(i) = 1
                Else
                    v1(i) = 0
                    delta1(i) = 0
                End If
            Next i
            pDelta(k) = delta1
        ElseIf pActivation = "SIGMOID" Then
            For i = 1 To poutput_L
                tmp_x = 1# / (1 + Exp(-(v1(i) + bias_k)))
                v1(i) = tmp_x
                delta1(i) = tmp_x * (1 - tmp_x)
            Next i
            pDelta(k) = delta1
        ElseIf pActivation = "SOFTMAX" Then
            For i = 1 To poutput_Lx
                v1(i) = Exp(v1(i) + bias_k)
                tmp_y = tmp_y + v1(i)
            Next i
        End If
        v(k) = v1
    Next k

    If pActivation = "SOFTMAX" Then
        'If softmax is used this has to be an output layer,
        'no need to calculate derivatives now
        For k = 1 To pn_filter
            v1 = v(k)
            For i = 1 To poutput_L
                v1(i) = v1(i) / tmp_y
            Next i
            v(k) = v1
        Next k
    End If

    'If target output is supplied this is the output layer, calculate delta
    If IsMissing(y_tgt) = False Then
        If pActivation = "SOFTMAX" Or pActivation = "SIGMOID" Then
            For k = 1 To pn_filter
                v1 = v(k)
                y1 = y_tgt(k)
                For i = 1 To poutput_L
                    delta1(i) = v1(i) - y1(i)
                Next i
                pDelta(k) = delta1
            Next k
        End If
    End If
    
End Sub


Sub Calc_Delta(DeltaNext As Variant)
Dim i As Long, j As Long, k As Long
Dim delta1() As Double, delta2() As Double
    For k = 1 To pn_filter
        delta1 = DeltaNext(k)
        delta2 = pDelta(k)
        For i = 1 To poutput_L
            delta2(i) = delta2(i) * delta1(i)
        Next i
        pDelta(k) = delta2
    Next k
End Sub

Sub Bkwd_Pass(x As Variant, grads As Variant, grads_bias() As Double, Optional prevLayer As Variant)
Dim i As Long, j As Long, k As Long, n As Long, jj As Long, kk As Long
Dim i2 As Long, j2 As Long, k2 As Long, q2 As Long
Dim p As Long, q As Long, r As Long
Dim tmp_x As Double, tmp_y As Double
Dim grad As Variant, grad1() As Double, w1() As Double, x1() As Double
Dim DeltaPrev As Variant, delta1() As Double, delta2() As Double

    'Calculate delta to be passed on to previous layer
    If IsMissing(prevLayer) = False Then
        ReDim DeltaPrev(1 To pinput_Lz)
        For r = 1 To pinput_Lz
            ReDim delta1(1 To pinput_L)
            For k = 1 To pn_filter
                w1 = pwgts(k)(r)
                delta2 = pDelta(k)
                jj = 0
                For j = 0 To (pinput_Ly - pfilter_Ly) Step pstride
                    For i = 0 To (pinput_Lx - pfilter_Lx) Step pstride
                        jj = jj + 1
                        tmp_x = delta2(jj)
                        kk = i + (j - 1) * pinput_Lx
                        For q = 1 To pfilter_Ly
                            kk = kk + pinput_Lx
                            q2 = (q - 1) * pfilter_Lx
                            For p = 1 To pfilter_Lx
                                delta1(p + kk) = delta1(p + kk) + w1(p + q2) * tmp_x
                            Next p
                        Next q
                    Next i
                Next j
            Next k
            DeltaPrev(r) = delta1
        Next r
        Call prevLayer.Calc_Delta(DeltaPrev)
        Erase DeltaPrev, delta1, delta2, w1
    End If

    'Calculate gradient
    ReDim grads_bias(1 To pn_filter)
    ReDim grads(1 To pn_filter)
    For n = 1 To pn_filter
        delta1 = pDelta(n)
        ReDim grad(1 To pinput_Lz)
        For r = 1 To pinput_Lz
            x1 = x(r)
            ReDim grad1(1 To pfilter_L)
            jj = 0
            For j = 0 To (pinput_Ly - pfilter_Ly) Step pstride
                For i = 0 To (pinput_Lx - pfilter_Lx) Step pstride
                    jj = jj + 1
                    tmp_x = delta1(jj)
                    kk = i + (j - 1) * pinput_Lx
                    For q = 1 To pfilter_Ly
                        kk = kk + pinput_Lx
                        q2 = (q - 1) * pfilter_Lx
                        For p = 1 To pfilter_Lx
                            grad1(p + q2) = grad1(p + q2) + x1(p + kk) * tmp_x
                        Next p
                    Next q
                Next i
            Next j
            grad(r) = grad1
        Next r
        grads(n) = grad
        tmp_x = 0
        For i = 1 To poutput_L
            tmp_x = tmp_x + delta1(i)
        Next i
        grads_bias(n) = tmp_x
    Next n
    Erase grad, grad1, x1
End Sub


Sub Update_Wgts(grads As Variant, grads_bias() As Double, learn_rate As Double, momentum As Double)
Dim i As Long, j As Long, k As Long, n As Long
Dim grad1() As Double, w1() As Double, wchg1() As Double
    For n = 1 To pn_filter
        For k = 1 To pinput_Lz
            grad1 = grads(n)(k)
            wchg1 = pwgts_chg(n)(k)
            w1 = pwgts(n)(k)
            For i = 1 To pfilter_L
                wchg1(i) = -learn_rate * grad1(i) + momentum * wchg1(i)
                w1(i) = w1(i) + wchg1(i)
            Next i
            pwgts(n)(k) = w1
            pwgts_chg(n)(k) = wchg1
        Next k
        pbias_chg(n) = -learn_rate * grads_bias(n) + momentum * pbias_chg(n)
        pbias(n) = pbias(n) + pbias_chg(n)
    Next n
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
        output_dim = pn_filter
    End If
End Property

Public Property Get filter_dim(k As Long) As Long
    If k = 1 Then
        filter_dim = pfilter_Lx
    ElseIf k = 2 Then
        filter_dim = pfilter_Ly
    End If
End Property

Public Property Get n_filter() As Long
    n_filter = pn_filter
End Property

Public Property Get stride() As Long
    stride = pstride
End Property

Public Property Get bias() As Double()
    bias = pbias
End Property

Public Property Get wgts() As Variant
    wgts = pwgts
End Property

Public Property Let bias(lbias() As Double)
    pbias = lbias
End Property

Public Property Let wgts(lwgts As Variant)
    pwgts = lwgts
End Property
