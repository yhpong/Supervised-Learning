VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cANN_Regression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================
'Single Layer Neural Network
'====================================================================================
'Implemented with momentum and adaptive gain on learning rate, L1, L2 and max norm regularization
'====================================================================================

Private pn_input As Long
Private pn_hidden As Long
Private pn_output As Long
Private pw_ih() As Double, pw_ih_chg() As Double, pw_ih_gain() As Double, pw_ih_grad() As Double
Private pw_ho() As Double, pw_ho_chg() As Double, pw_ho_gain() As Double, pw_ho_grad() As Double
Private pbias_h() As Double, pbias_h_chg() As Double, pbias_h_gain() As Double, pbias_h_grad() As Double
Private pbias_o() As Double, pbias_o_chg() As Double, pbias_o_gain() As Double, pbias_o_grad() As Double
Private pcost_function() As Double

Public Property Get w_ih() As Double()
    w_ih = pw_ih
End Property

Public Property Let w_ih(x() As Double)
    pw_ih = x
End Property

Public Property Get w_ho() As Double()
    w_ho = pw_ho
End Property

Public Property Let w_ho(x() As Double)
    pw_ho = x
End Property

Public Property Get bias_h() As Double()
    bias_h = pbias_h
End Property

Public Property Let bias_h(x() As Double)
    pbias_h = x
End Property

Public Property Get bias_o() As Double()
    bias_o = pbias_o
End Property

Public Property Let bias_o(x() As Double)
    pbias_o = x
End Property

Public Property Get cost_function() As Double()
    cost_function = pcost_function
End Property


Sub Print_Model(wksht As Worksheet)
Dim i As Long, j As Long, n As Long
With wksht
    .Range(.Cells(1, 1), .Cells(1000, 1000)).Clear
    .Cells(1, 1).Value = pn_input
    .Cells(2, 1).Value = pn_hidden
    .Cells(3, 1).Value = pn_output
    i = 3

    For j = 1 To pn_hidden
        .Cells(i + 1, j).Value = pbias_h(j)
    Next j
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = pw_ih
    For j = 1 To pn_output
        .Cells(i + 2 + pn_input, j).Value = pbias_o(j)
    Next j
    .Range(.Cells(i + 3 + pn_input, 1), .Cells(i + 2 + pn_input + pn_hidden, pn_output)).Value = pw_ho
End With
End Sub


Sub Read_Model(wksht As Worksheet)
Dim i As Long, j As Long, k As Long, n As Long
Dim vArr As Variant
With wksht
    pn_input = .Cells(1, 1).Value
    pn_hidden = .Cells(2, 1).Value
    pn_output = .Cells(3, 1).Value
    Call Init(.pn_input, pn_output, pn_hidden)
    i = 3
    For j = 1 To pn_hidden
        pbias_h(j) = .Cells(i + 1, j).Value
    Next j
    vArr = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    For j = 1 To pn_input
        For k = 1 To pn_hidden
            pw_ih(j, k) = vArr(j, k)
        Next k
    Next j
    For j = 1 To pn_output
        pbias_o(j) = .Cells(i + 2 + pn_input, j).Value
    Next j
    vArr = .Range(.Cells(i + 3 + pn_input, 1), .Cells(i + 2 + pn_input + pn_hidden, pn_output)).Value
    For j = 1 To pn_hidden
        For k = 1 To pn_output
            pw_ho(j, k) = vArr(j, k)
        Next k
    Next j
    Erase vArr
End With
End Sub


Sub Reset()
    pn_input = 0
    pn_output = 0
    pn_hidden = 0
    Erase pw_ih, pw_ho, pbias_h, pbias_o
End Sub

Sub Init(n_input As Long, n_output As Long, n_hidden As Long)
Dim i As Long, j As Long, k As Long
    pn_input = n_input
    pn_output = n_output
    pn_hidden = n_hidden
    
    ReDim pw_ih(1 To pn_input, 1 To pn_hidden)
    ReDim pw_ho(1 To pn_hidden, 1 To pn_output)
    ReDim pbias_h(1 To pn_hidden)
    ReDim pbias_o(1 To pn_output)
    
    '===== Set Initial Weights ======
    Randomize
    For j = 1 To pn_hidden
        pbias_h(j) = 0 '-0.01 + 0.02 * Rnd()
        For i = 1 To pn_input
            pw_ih(i, j) = (-1 + 2 * Rnd()) / pn_hidden
        Next i
    Next j
    
    For k = 1 To pn_output
        pbias_o(k) = 0 '-0.01 + 0.02 * Rnd()
        For j = 1 To pn_hidden
            pw_ho(j, k) = (-1 + 2 * Rnd()) / pn_hidden
        Next j
    Next k
    '===================================
End Sub


Private Sub Init_Train()
Dim i As Long, j As Long, k As Long
    ReDim pbias_h_chg(1 To pn_hidden)
    ReDim pbias_o_chg(1 To pn_output)
    ReDim pw_ih_chg(1 To pn_input, 1 To pn_hidden)
    ReDim pw_ho_chg(1 To pn_hidden, 1 To pn_output)
    
    ReDim pbias_h_gain(1 To pn_hidden)
    ReDim pbias_o_gain(1 To pn_output)
    ReDim pw_ih_gain(1 To pn_input, 1 To pn_hidden)
    ReDim pw_ho_gain(1 To pn_hidden, 1 To pn_output)
    
    For j = 1 To pn_hidden
        pbias_h_gain(j) = 1
        For i = 1 To pn_input
            pw_ih_gain(i, j) = 1
        Next i
    Next j
    For k = 1 To pn_output
        pbias_o_gain(k) = 1
        For j = 1 To pn_hidden
            pw_ho_gain(j, k) = 1
        Next j
    Next k
End Sub

Private Sub Reset_Train()
    Erase pbias_h_chg, pbias_h_gain
    Erase pbias_o_chg, pbias_o_gain
    Erase pw_ih_chg, pw_ih_gain
    Erase pw_ho_chg, pw_ho_gain
End Sub

Private Sub Init_MiniBatch()
    ReDim pbias_h_grad(1 To pn_hidden)
    ReDim pbias_o_grad(1 To pn_output)
    ReDim pw_ih_grad(1 To pn_input, 1 To pn_hidden)
    ReDim pw_ho_grad(1 To pn_hidden, 1 To pn_output)
End Sub

Private Sub Reset_MiniBatch()
    Erase pbias_h_grad, pw_ih_grad
    Erase pbias_o_grad, pw_ho_grad
End Sub

'Input: x_input(1 to n_raw, 1 to n_input)
'Input: target(1 to n_raw, 1 to n_output)
Sub Trainer(x() As Double, target() As Double, _
        Optional learn_rate As Double = 0.01, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 1000, Optional cost_min As Double = 0.01, _
        Optional L1 As Double = 0, Optional L2 As Double = 0, Optional LMAX As Double = 0, _
        Optional x_validate As Variant, Optional target_validate As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, epoch As Long
Dim n_raw As Long, batch_count As Long
Dim xn() As Double, y() As Double, x_v() As Double, tgt_v() As Double
Dim v() As Double, cost_function() As Double
Dim delta_o() As Double, delta_h() As Double
Dim tmp_x As Double, tmp_y As Double
Dim min_gain As Double, max_gain As Double
Dim iArr() As Long

n_raw = UBound(x, 1)
min_gain = 0.01
max_gain = 100 / learn_rate
Call Init_Train

'=== Pre allocate memory
ReDim delta_o(1 To pn_output)
ReDim delta_h(1 To pn_hidden)
ReDim xn(1 To pn_input)
ReDim cost_function(1 To epoch_max, 1 To 1)
'=======================================

If IsMissing(x_validate) = False Then
    x_v = x_validate
    tgt_v = target_validate
    ReDim cost_function(1 To epoch_max, 1 To 2)
End If

batch_count = 0
For epoch = 1 To epoch_max
    
    If epoch Mod 50 = 0 Then
        DoEvents
        Application.StatusBar = "cANN_Regression: Epoch: " & epoch & "/" & epoch_max
    End If
    
    iArr = modMath.index_array(1, n_raw)
    Call modMath.Shuffle(iArr)
    'iArr = Shuffle(1, n_raw)
    
    Call Init_MiniBatch
    ReDim y(1 To n_raw, 1 To pn_output)
    
    For m = 1 To n_raw

        n = iArr(m)
        For i = 1 To pn_input
            xn(i) = x(n, i)
        Next i
        
        '=== Input to hidden layer
        ReDim v(1 To pn_hidden)
        For j = 1 To pn_hidden
            tmp_x = pbias_h(j)
            For i = 1 To pn_input
                tmp_x = tmp_x + xn(i) * pw_ih(i, j)
            Next i
            v(j) = 1# / (1# + Exp(-tmp_x)) 'Sigmoid
        Next j
        '============================

        '=== Hidden layer to Output
        For k = 1 To pn_output
            tmp_x = pbias_o(k)
            For j = 1 To pn_hidden
                tmp_x = tmp_x + v(j) * pw_ho(j, k)
            Next j
            'Sigmoid
            tmp_x = 1# / (1# + Exp(-tmp_x))
            delta_o(k) = (tmp_x - target(n, k)) * tmp_x * (1 - tmp_x)
            'Tanh
            'tmp_x = Exp(2 * tmp_x)
            'tmp_x = (tmp_x - 1) / (tmp_x + 1)
            'delta_o(k) = (tmp_x - target(n, k)) * (1 - tmp_x ^ 2)
            'Identity
            'delta_o(k) = tmp_x - target(n, k)
        Next k
        '===========================================

        '=== Back Propagation
        For j = 1 To pn_hidden
            tmp_x = 0
            For k = 1 To pn_output
                tmp_x = tmp_x + pw_ho(j, k) * delta_o(k)
            Next k
            delta_h(j) = tmp_x * v(j) * (1 - v(j))  'Derivative of Sigmoid
        Next j
        
        Call Accumulate_Wgt_chg(delta_o, delta_h, v, xn)
        
        batch_count = batch_count + 1
        If batch_count = mini_batch Or m = n_raw Then
            If L1 > 0 Or L2 > 0 Then Call Normalize_Grad(L1, L2, batch_count)
            Call Learn_Rate_Gain(min_gain, max_gain)
            Call wgt_update(learn_rate / batch_count, momentum)
            If LMAX > 0 Then Call max_norm_reg(LMAX)
            Call Init_MiniBatch
            batch_count = 0
        End If
        '=====================================
        
    Next m

    Call Reset_MiniBatch
    
    Call InOut(x, y)
    cost_function(epoch, 1) = eval_SSE(target, y)
    If IsMissing(x_validate) = False Then
        Call InOut(x_v, y)
        cost_function(epoch, 2) = eval_SSE(tgt_v, y)
    End If
    
    'Stop if cost function is small enough
    If cost_function(epoch, 1) < cost_min Then
        ReDim v(1 To epoch, 1 To UBound(cost_function, 2))
        For i = 1 To epoch
            For j = 1 To UBound(cost_function, 2)
                v(i, j) = cost_function(i, j)
            Next j
        Next i
        cost_function = v
        Exit For
    End If
    
Next epoch

pcost_function = cost_function
Erase xn, cost_function, v, delta_o, delta_h
Call Reset_Train
Application.StatusBar = False
End Sub



'Input: x_input(1 to n_raw, 1 to pn_input)
'Output: y(1 to n_raw, 1 to pn_output)
Sub InOut(x() As Double, y() As Double)
Dim i As Long, j As Long, k As Long, n As Long
Dim n_raw As Long
Dim v() As Double
Dim tmp_x As Double
    n_raw = UBound(x, 1)
    ReDim v(1 To pn_hidden)
    ReDim y(1 To n_raw, 1 To pn_output)
    For n = 1 To n_raw
        '=== Input to hidden layer
        For j = 1 To pn_hidden
            tmp_x = pbias_h(j)
            For i = 1 To pn_input
                tmp_x = tmp_x + x(n, i) * pw_ih(i, j)
            Next i
            v(j) = 1# / (1 + Exp(-tmp_x))   'Sigmoid
        Next j
        '=== Hidden layer to Output
        For k = 1 To pn_output
            tmp_x = pbias_o(k)
            For j = 1 To pn_hidden
                tmp_x = tmp_x + v(j) * pw_ho(j, k)
            Next j
            y(n, k) = 1# / (1 + Exp(-tmp_x))   'Sigmoid
            'tmp_x = Exp(2 * tmp_x)
            'y(n, k) = (tmp_x - 1) / (tmp_x + 1) 'Tanh
            'y(n, k) = tmp_x 'Identity
        Next k
    
    Next n
    Erase v
End Sub


Private Sub Accumulate_Wgt_chg(delta_o() As Double, delta_h() As Double, v() As Double, x() As Double)
Dim i As Long, j As Long, k As Long

    For k = 1 To pn_output
        pbias_o_grad(k) = pbias_o_grad(k) + delta_o(k)
        For j = 1 To pn_hidden
            pw_ho_grad(j, k) = pw_ho_grad(j, k) + v(j) * delta_o(k)
        Next j
    Next k
    
    For j = 1 To pn_hidden
        pbias_h_grad(j) = pbias_h_grad(j) + delta_h(j)
        For i = 1 To pn_input
            pw_ih_grad(i, j) = pw_ih_grad(i, j) + x(i) * delta_h(j)
        Next i
    Next j

End Sub


Private Sub Learn_Rate_Gain(min_gain As Double, max_gain As Double)
Dim i As Long, j As Long, k As Long
    For j = 1 To pn_hidden
        Call calc_gain(pbias_h_gain(j), pbias_h_grad(j), pbias_h_chg(j), min_gain, max_gain)
        For i = 1 To pn_input
            Call calc_gain(pw_ih_gain(i, j), pw_ih_grad(i, j), pw_ih_chg(i, j), min_gain, max_gain)
        Next i
    Next j

    For k = 1 To pn_output
        Call calc_gain(pbias_o_gain(k), pbias_o_grad(k), pbias_o_chg(k), min_gain, max_gain)
        For j = 1 To pn_hidden
            Call calc_gain(pw_ho_gain(j, k), pw_ho_grad(j, k), pw_ho_chg(j, k), min_gain, max_gain)
        Next j
    Next k
End Sub

Private Sub calc_gain(gain As Double, gradient As Double, prev_chg As Double, min_gain As Double, max_gain As Double)
    If Sgn(gradient) <> Sgn(prev_chg) Then
        gain = gain * 1.1
    Else
        gain = gain * 0.9
    End If
    If gain < min_gain Then gain = min_gain
    If gain > max_gain Then gain = max_gain
End Sub


Private Sub Normalize_Grad(L1 As Double, L2 As Double, batch_count As Long)
Dim i As Long, j As Long, k As Long, m As Long, mm As Long
Dim tmp_x As Double, tmp_y As Double
    If L1 > 0 Then
        tmp_x = L1 * batch_count
        For j = 1 To pn_hidden
            For i = 1 To pn_input
                pw_ih_grad(i, j) = pw_ih_grad(i, j) + Sgn(pw_ih(i, j)) * tmp_x
            Next i
            For k = 1 To pn_output
                pw_ho_grad(j, k) = pw_ho_grad(j, k) + Sgn(pw_ho(j, k)) * tmp_x
            Next k
        Next j
    End If
    If L2 > 0 Then
        tmp_x = L2 * batch_count
        For j = 1 To pn_hidden
            For i = 1 To pn_input
                pw_ih_grad(i, j) = pw_ih_grad(i, j) + pw_ih(i, j) * tmp_x
            Next i
            For k = 1 To pn_output
                pw_ho_grad(j, k) = pw_ho_grad(j, k) + pw_ho(j, k) * tmp_x
            Next k
        Next j
    End If
End Sub

Private Sub wgt_update(learn_rate As Double, momentum As Double)
Dim i As Long, j As Long, k As Long

    For j = 1 To pn_hidden
        pbias_h_chg(j) = momentum * pbias_h_chg(j) - learn_rate * pbias_h_grad(j) * pbias_h_gain(j)
        pbias_h(j) = pbias_h(j) + pbias_h_chg(j)
        For i = 1 To pn_input
            pw_ih_chg(i, j) = momentum * pw_ih_chg(i, j) - learn_rate * pw_ih_grad(i, j) * pw_ih_gain(i, j)
            pw_ih(i, j) = pw_ih(i, j) + pw_ih_chg(i, j)
        Next i
    Next j
    
    For k = 1 To pn_output
        pbias_o_chg(k) = momentum * pbias_o_chg(k) - learn_rate * pbias_o_grad(k) * pbias_o_gain(k)
        pbias_o(k) = pbias_o(k) + pbias_o_chg(k)
        For j = 1 To pn_hidden
            pw_ho_chg(j, k) = momentum * pw_ho_chg(j, k) - learn_rate * pw_ho_grad(j, k) * pw_ho_gain(j, k)
            pw_ho(j, k) = pw_ho(j, k) + pw_ho_chg(j, k)
        Next j
    Next k
End Sub


Private Sub max_norm_reg(Optional max_norm As Double = 3)
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double
For j = 1 To pn_hidden
    tmp_x = 0
    For i = 1 To pn_input
        tmp_x = tmp_x + pw_ih(i, j) ^ 2
    Next i
    If tmp_x > max_norm Then
        tmp_x = Sqr(max_norm / tmp_x)
        For i = 1 To pn_input
            pw_ih(i, j) = pw_ih(i, j) * tmp_x
        Next i
    End If
    
    tmp_x = 0
    For k = 1 To pn_output
        tmp_x = tmp_x + pw_ho(j, k) ^ 2
    Next k
    If tmp_x > max_norm Then
        tmp_x = Sqr(max_norm / tmp_x)
        For k = 1 To pn_output
            pw_ho(j, k) = pw_ho(j, k) * tmp_x
        Next k
    End If
Next j
End Sub


Private Function eval_SSE(target() As Double, y() As Double) As Double
Dim i As Long, k As Long
Dim n_raw As Long, n_output As Long
    n_raw = UBound(target, 1)
    n_output = UBound(target, 2)
    eval_SSE = 0
    For i = 1 To n_raw
        For k = 1 To n_output
            eval_SSE = eval_SSE + (target(i, k) - y(i, k)) ^ 2
        Next k
    Next i
    eval_SSE = eval_SSE / (2 * n_raw * n_output)
End Function
