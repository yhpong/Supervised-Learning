VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLSTM_B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pn_input As Long
Private pn_hidden As Long
Private pn_output As Long
Private pw_ho() As Variant, pw_ho_chg() As Double, pw_ho_chg_batch() As Double
Private pw_ho_bkwd() As Variant, pw_ho_chg_bkwd() As Double, pw_ho_chg_batch_bkwd() As Double
Private py_hidden() As Double, py_hidden_bkwd() As Double
Private py_output() As Double, pderiv_output() As Double, pdelta_out() As Double
Private pcost_function() As Double
Private pLayerFwd As cLSTM_Layer
Private pLayerBkwd As cLSTM_Layer

Public Property Get cost_function() As Double()
    cost_function = pcost_function
End Property


Sub Print_Model(wksht As Worksheet)
Dim i As Long, n As Long
Dim bias As Variant, w_in As Variant, w_h As Variant, w_o As Variant, w_c As Variant
With wksht
    .Range(.Cells(1, 1), .Cells(10000, 1000)).Clear
    .Cells(1, 1).Value = "BLSTM"
    .Cells(2, 1).Value = pn_input
    .Cells(3, 1).Value = pn_hidden
    .Cells(4, 1).Value = pn_output
    i = 4

    n = pn_input + pn_hidden
    Call pLayerFwd.Get_Input_Gate(bias, w_in, w_h, w_c)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value = w_c
    i = i + 2 + n

    Call pLayerFwd.Get_Forget_Gate(bias, w_in, w_h, w_c)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value = w_c
    i = i + 2 + n

    Call pLayerFwd.Get_Cell(bias, w_in, w_h)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    i = i + 1 + n

    Call pLayerFwd.Get_Output_Gate(bias, w_in, w_h, w_c)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value = w_c
    i = i + 2 + n

    Call pLayerBkwd.Get_Input_Gate(bias, w_in, w_h, w_c)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value = w_c
    i = i + 2 + n

    Call pLayerBkwd.Get_Forget_Gate(bias, w_in, w_h, w_c)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value = w_c
    i = i + 2 + n

    Call pLayerBkwd.Get_Cell(bias, w_in, w_h)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    i = i + 1 + n

    Call pLayerBkwd.Get_Output_Gate(bias, w_in, w_h, w_c)
    .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value = bias
    .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value = w_in
    .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value = w_h
    .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value = w_c
    i = i + 2 + n

    .Range(.Cells(i + 1, 1), .Cells(i + pn_hidden, pn_output)).Value = pw_ho
End With
End Sub


Sub Read_Model(wksht As Worksheet)
Dim i As Long, j As Long, n As Long
Dim bias As Variant, w_in As Variant, w_h As Variant, w_o As Variant, w_c As Variant
Dim bias2 As Variant, w_c2 As Variant
With wksht
    Call Init(.Cells(2, 1).Value, .Cells(3, 1).Value, .Cells(4, 1).Value)
    i = 4

    n = pn_input + pn_hidden
    ReDim bias(1 To pn_hidden)
    ReDim w_c(1 To pn_hidden)

    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    w_c2 = .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
        w_c(j) = w_c2(1, j)
    Next j
    Call pLayerFwd.Assign_Input_Gate(bias, w_in, w_h, w_c)
    i = i + 2 + n

    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    w_c2 = .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
        w_c(j) = w_c2(1, j)
    Next j
    Call pLayerFwd.Assign_Forget_Gate(bias, w_in, w_h, w_c)
    i = i + 2 + n

    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
    Next j
    Call pLayerFwd.Assign_Cell(bias, w_in, w_h)
    i = i + 1 + n

    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    w_c2 = .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
        w_c(j) = w_c2(1, j)
    Next j
    Call pLayerFwd.Assign_Output_Gate(bias, w_in, w_h, w_c)
    i = i + 2 + n


    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    w_c2 = .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
        w_c(j) = w_c2(1, j)
    Next j
    Call pLayerBkwd.Assign_Input_Gate(bias, w_in, w_h, w_c)
    i = i + 2 + n

    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    w_c2 = .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
        w_c(j) = w_c2(1, j)
    Next j
    Call pLayerBkwd.Assign_Forget_Gate(bias, w_in, w_h, w_c)
    i = i + 2 + n

    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
    Next j
    Call pLayerBkwd.Assign_Cell(bias, w_in, w_h)
    i = i + 1 + n

    bias2 = .Range(.Cells(i + 1, 1), .Cells(i + 1, pn_hidden)).Value
    w_in = .Range(.Cells(i + 2, 1), .Cells(i + 1 + pn_input, pn_hidden)).Value
    w_h = .Range(.Cells(i + 2 + pn_input, 1), .Cells(i + 1 + n, pn_hidden)).Value
    w_c2 = .Range(.Cells(i + 2 + n, 1), .Cells(i + 2 + n, pn_hidden)).Value
    For j = 1 To pn_hidden
        bias(j) = bias2(1, j)
        w_c(j) = w_c2(1, j)
    Next j
    Call pLayerBkwd.Assign_Output_Gate(bias, w_in, w_h, w_c)
    i = i + 2 + n
    
    pw_ho = .Range(.Cells(i + 1, 1), .Cells(i + pn_hidden, pn_output)).Value
End With
End Sub


'*******************************************************
'*** Define loss function
'*******************************************************
Private Function Calc_Cost_Function(y() As Double, y_target() As Double) As Double
Dim i As Long, k As Long, t As Long
Dim tmp_x As Double
    For t = 1 To UBound(y_target, 1)
        For k = 1 To pn_output
            'Mean square error
            pdelta_out(t, k) = (y(t, k) - y_target(t, k)) * pderiv_output(t, k)
            tmp_x = tmp_x + (y(t, k) - y_target(t, k)) ^ 2
            
'            'Multiclass classification
'            'Cross Entropy, change output layer to sofmax.
'            pdelta_out(t, k) = y(t, k) - y_target(t, k)
'            tmp_x = tmp_x - y_target(t, k) * Log(y(t, k))
            
'            'Binary classification
'            'Cross Entropy, remember to change output layer to sigmoid as well.
'            pdelta_out(t, k) = y(t, k) - y_target(t, k)
'            tmp_x = tmp_x - y_target(t, k) * Log(y(t, k)) - (1 - y_target(t, k)) * Log(1 - y(t, k))
        Next k
    Next t
    Calc_Cost_Function = tmp_x / (2 * pn_output * UBound(y_target, 1))
End Function

'*******************************************************
'*** Define activation function of each unit
'*******************************************************
Private Function Sigmoid(x As Double) As Double
    Sigmoid = 1# / (1 + Exp(-x))
End Function

Private Function Sigmoid_Deriv(x As Double) As Double
Dim tmp_x As Double
    tmp_x = Sigmoid(x)
    Sigmoid_Deriv = tmp_x * (1 - tmp_x)
End Function

Private Function ReLU(x As Double) As Double
    If x > 0 Then
        ReLU = x
    Else
        ReLU = 0
    End If
End Function

Private Function ReLU_Deriv(x As Double) As Double
    If x > 0 Then
        ReLU_Deriv = 1
    Else
        ReLU_Deriv = 0
    End If
End Function

Private Function Activation_Output(x As Double) As Double
    Activation_Output = x
End Function
Private Function Gradient_Output(x As Double) As Double
    Gradient_Output = 1
End Function
'*******************************************************
'*******************************************************
 
Sub Reset()
    pn_input = 0
    pn_hidden = 0
    pn_output = 0
    Erase pw_ho, pw_ho_bkwd
    Call pLayerFwd.Reset
    Call pLayerBkwd.Reset
End Sub

Sub Init(n_input As Long, n_hidden As Long, n_output As Long, _
        Optional bias_i As Double = 0, _
        Optional bias_f As Double = 2, _
        Optional bias_c As Double = 0, _
        Optional bias_o As Double = 0)
Dim i As Long, j As Long
Dim tmp_x As Double
    pn_input = n_input
    pn_hidden = n_hidden
    pn_output = n_output
    
    Set pLayerFwd = New cLSTM_Layer
    Set pLayerBkwd = New cLSTM_Layer
    Call pLayerFwd.Init(pn_input, pn_hidden, pn_output, bias_i, bias_f, bias_c, bias_o)
    Call pLayerBkwd.Init(pn_input, pn_hidden, pn_output, bias_i, bias_f, bias_c, bias_o)
    
    Randomize
    tmp_x = 1# / (pn_hidden * (pn_input + pn_hidden + pn_output))
    ReDim pw_ho(1 To pn_hidden, 1 To pn_output)
    ReDim pw_ho_bkwd(1 To pn_hidden, 1 To pn_output)
    For i = 1 To pn_hidden
        For j = 1 To pn_output
            pw_ho(i, j) = -tmp_x + Rnd() * 2 * tmp_x
            pw_ho_bkwd(i, j) = -tmp_x + Rnd() * 2 * tmp_x
        Next j
    Next i
End Sub

Private Sub Train_Batch_Init()
    Call pLayerFwd.Train_Batch_Init
    Call pLayerBkwd.Train_Batch_Init
    ReDim pw_ho_chg(1 To pn_hidden, 1 To pn_output)
    ReDim pw_ho_chg_bkwd(1 To pn_hidden, 1 To pn_output)
End Sub

Private Sub Train_Batch_Reset()
    Call pLayerFwd.Train_Batch_Reset
    Call pLayerBkwd.Train_Batch_Reset
    Erase pw_ho_chg, pw_ho_chg_bkwd
End Sub

Private Sub Train_Mini_Batch_Init()
    Call pLayerFwd.Train_Mini_Batch_Init
    Call pLayerBkwd.Train_Mini_Batch_Init
    ReDim pw_ho_chg_batch(1 To pn_hidden, 1 To pn_output)
    ReDim pw_ho_chg_batch_bkwd(1 To pn_hidden, 1 To pn_output)
End Sub

Private Sub Train_Mini_Batch_Reset()
    Call pLayerFwd.Train_Mini_Batch_Reset
    Call pLayerBkwd.Train_Mini_Batch_Reset
    Erase pw_ho_chg_batch, pw_ho_chg_batch_bkwd
End Sub

Private Sub Train_Init(n_T As Long)
    Call pLayerFwd.Train_Init(n_T)
    Call pLayerBkwd.Train_Init(n_T)
    ReDim py_hidden(1 To n_T, 1 To pn_hidden)
    ReDim py_hidden_bkwd(1 To n_T, 1 To pn_hidden)
    ReDim py_output(1 To n_T, 1 To pn_output)
    ReDim pderiv_output(1 To n_T, 1 To pn_output)
    ReDim pdelta_out(1 To n_T, 1 To pn_output)
End Sub

Private Sub Train_Reset()
    Call pLayerFwd.Train_Reset
    Call pLayerBkwd.Train_Reset
    Erase py_hidden, py_hidden_bkwd, py_output, pderiv_output, pdelta_out
End Sub


Sub Train_Batch(xS As Variant, y_targets As Variant, Optional batch_size As Long = 5, _
        Optional learn_rate As Double = 0.001, Optional momentum As Double = 0.9, _
        Optional epoch_max As Long = 10, _
        Optional xs_validate As Variant, Optional y_targets_validate As Variant)
Dim i As Long, j As Long, k As Long, n As Long, epoch As Long, n_series As Long
Dim x() As Double, y_target() As Double
Dim xs_batch() As Variant, y_targets_batch() As Variant
Dim tmp_x As Double, tmp_y As Double
Dim iArr() As Long

Call Train_Batch_Init

n_series = UBound(xS, 1)
ReDim pcost_function(1 To epoch_max, 1 To 2)
For epoch = 1 To epoch_max
    
    'Shuffle training sequences
    iArr = Random_Array(1, n_series)
    
    tmp_y = 0
    For i = 1 To n_series Step batch_size
        
        If (i - 1) / batch_size Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = "LSTM_B Batch Training: " & epoch & "/" & epoch_max & ": " & i & "/" & n_series
        End If
        
        n = min2(batch_size, n_series - i + 1)
        ReDim xs_batch(1 To n)
        ReDim y_targets_batch(1 To n)
        For j = 1 To n
            xs_batch(j) = xS(iArr(i + j - 1))
            y_targets_batch(j) = y_targets(iArr(i + j - 1))
        Next j

        Call Train_MiniBatch(xs_batch, y_targets_batch, learn_rate, momentum, tmp_x)
        tmp_y = tmp_y + tmp_x
    Next i
    pcost_function(epoch, 1) = tmp_y / n_series

    If IsMissing(xs_validate) = False And IsMissing(y_targets_validate) = False Then
        tmp_y = 0
        For i = 1 To UBound(xs_validate, 1)
            x = xs_validate(i)
            y_target = y_targets_validate(i)
            x = Me.InOut(x)
            tmp_y = tmp_y + Calc_Cost_Function(x, y_target)
        Next i
        Call Train_Reset
        pcost_function(epoch, 2) = tmp_y / UBound(xs_validate, 1)
    End If
Next epoch

Call Train_Batch_Reset
Application.StatusBar = False
End Sub

Sub Train_MiniBatch(xS As Variant, y_targets As Variant, learn_rate As Double, momentum As Double, cost_function As Double)
Dim i As Long, n_series As Long
Dim x() As Double, y_target() As Double
Dim tmp_x As Double
    Call Train_Mini_Batch_Init
    cost_function = 0
    n_series = UBound(xS, 1)
    For i = 1 To n_series
        x = xS(i)
        y_target = y_targets(i)
        Call Train(x, y_target, learn_rate, momentum, tmp_x)
        cost_function = cost_function + tmp_x
    Next i
    Call Update_wgts(learn_rate / n_series, momentum)
    Call Train_Mini_Batch_Reset
End Sub

Sub Train(x() As Double, y_target() As Double, learn_rate As Double, momentum As Double, cost As Double)
Dim i As Long, t As Long, n_T As Long
Dim y() As Double
    n_T = UBound(x, 1)
    Call Train_Init(n_T)

    With pLayerFwd
        For t = 1 To n_T
            Call .FwdPass(x, t, t - 1)
        Next t
        py_hidden = .y_output
    End With
    
    With pLayerBkwd
        For t = n_T To 1 Step -1
            Call .FwdPass(x, t, t + 1)
        Next t
        py_hidden_bkwd = .y_output
    End With
    
    For t = 1 To n_T
        Call FwdPass(t)
    Next t

    cost = Calc_Cost_Function(py_output, y_target)
    
    With pLayerFwd
        For t = n_T To 1 Step -1
            Call .BkwdPass(t, t + 1, t - 1, pw_ho, pdelta_out)
        Next t
    End With

    With pLayerBkwd
        For t = 1 To n_T
            Call .BkwdPass(t, t - 1, t + 1, pw_ho_bkwd, pdelta_out)
        Next t
    End With

    Call Accumulate_Wgt_chg
    Call pLayerFwd.Accumulate_Wgt_chg(x, 1)
    Call pLayerBkwd.Accumulate_Wgt_chg(x, -1)
    Call Train_Reset
End Sub

Function InOut(x() As Double) As Double()
Dim i As Long, t As Long, n_T As Long
Dim y() As Double
    n_T = UBound(x, 1)
    Call Train_Init(n_T)
    
    With pLayerFwd
        For t = 1 To n_T
            Call .FwdPass(x, t, t - 1)
        Next t
        py_hidden = .y_output
    End With
    
    With pLayerBkwd
        For t = n_T To 1 Step -1
            Call .FwdPass(x, t, t + 1)
        Next t
        py_hidden_bkwd = .y_output
    End With
    
    For t = 1 To n_T
        Call FwdPass(t)
    Next t

    InOut = py_output
    Call pLayerFwd.Train_Reset
    Call pLayerBkwd.Train_Reset
    Erase py_hidden, py_output
End Function

Private Sub FwdPass(t As Long)
Dim i As Long, j As Long
Dim tmp_x As Double, tmp_y As Double
    For j = 1 To pn_output
        tmp_x = 0
        For i = 1 To pn_hidden
            tmp_x = tmp_x + py_hidden(t, i) * pw_ho(i, j) _
                        + py_hidden_bkwd(t, i) * pw_ho_bkwd(i, j)
        Next i
        py_output(t, j) = Activation_Output(tmp_x)
        pderiv_output(t, j) = Gradient_Output(tmp_x)
    Next j
    
'    'Softmax
'    tmp_y = 0
'    For j = 1 To pn_output
'        tmp_x = 0
'        For i = 1 To pn_hidden
'            tmp_x = tmp_x + py_hidden(t, i) * pw_ho(i, j) + py_hidden_bkwd(t, i) * pw_ho_bkwd(i, j)
'        Next i
'        py_output(t, j) = Exp(tmp_x)
'        tmp_y = tmp_y + py_output(t, j)
'    Next j
'    For j = 1 To pn_output
'        py_output(t, j) = py_output(t, j) / tmp_y
'    Next j
End Sub



Private Sub Accumulate_Wgt_chg()
Dim i As Long, j As Long, t As Long, n_T As Long
    n_T = UBound(py_hidden, 1)
    For i = 1 To pn_hidden
        For j = 1 To pn_output
            For t = 1 To n_T
                pw_ho_chg_batch(i, j) = pw_ho_chg_batch(i, j) + pdelta_out(t, j) * py_hidden(t, i)
                pw_ho_chg_batch_bkwd(i, j) = pw_ho_chg_batch_bkwd(i, j) + pdelta_out(t, j) * py_hidden_bkwd(t, i)
            Next t
        Next j
    Next i
End Sub

Private Sub Update_wgts(learn_rate As Double, momentum As Double)
Dim i As Long, j As Long
    Call pLayerFwd.Update_wgts(learn_rate, momentum)
    For i = 1 To pn_hidden
        For j = 1 To pn_output
            pw_ho_chg(i, j) = momentum * pw_ho_chg(i, j) _
                                    - learn_rate * pw_ho_chg_batch(i, j)
            pw_ho_chg_bkwd(i, j) = momentum * pw_ho_chg_bkwd(i, j) _
                                    - learn_rate * pw_ho_chg_batch_bkwd(i, j)
            pw_ho(i, j) = pw_ho(i, j) + pw_ho_chg(i, j)
            pw_ho_bkwd(i, j) = pw_ho_bkwd(i, j) + pw_ho_chg_bkwd(i, j)
        Next j
    Next i
    Call max_norm_reg(3)
End Sub

Private Sub max_norm_reg(Optional max_norm As Double = 3)
Dim i As Long, j As Long
Dim tmp_x As Double
    For i = 1 To pn_hidden
        tmp_x = 0
        For j = 1 To pn_output
            tmp_x = tmp_x + pw_ho(i, j) ^ 2
        Next j
        If tmp_x > max_norm Then
            tmp_x = Sqr(max_norm / tmp_x)
            For j = 1 To pn_output
                pw_ho(i, j) = pw_ho(i, j) * tmp_x
            Next j
        End If
        
        tmp_x = 0
        For j = 1 To pn_output
            tmp_x = tmp_x + pw_ho_bkwd(i, j) ^ 2
        Next j
        If tmp_x > max_norm Then
            tmp_x = Sqr(max_norm / tmp_x)
            For j = 1 To pn_output
                pw_ho_bkwd(i, j) = pw_ho_bkwd(i, j) * tmp_x
            Next j
        End If
        
    Next i
End Sub


'Randomly shuffle a base-1 integer array
Private Function Random_Array(i_min As Long, i_max As Long) As Long()
Dim i As Long, j As Long, n As Long
Dim k As Long
Dim y() As Long
n = i_max - i_min + 1
ReDim y(1 To n)
For i = 1 To n
    y(i) = i_min + i - 1
Next i
Randomize
For i = n To 2 Step -1
    j = Int(Rnd() * i) + 1
    k = y(j)
    y(j) = y(i)
    y(i) = k
Next i
Random_Array = y
Erase y
End Function


'Randomly shuffle a base-1 integer array
Private Function Shuffle(x() As Long) As Long()
Dim i As Long, j As Long, n As Long
Dim k As Long
Dim y() As Long
n = UBound(x)
ReDim y(1 To n)
y = x
Randomize
For i = n To 2 Step -1
    j = Int(Rnd() * i) + 1
    k = y(j)
    y(j) = y(i)
    y(i) = k
Next i
Shuffle = y
End Function


Private Function min2(i As Long, j As Long) As Long
    min2 = i
    If j < i Then min2 = j
End Function


'Generating whole sequence from seed
Function Generate(x_seed() As Double, n As Long, n_T As Long, Optional maxSelect As Boolean = True) As Double()
Dim i As Long, j As Long, k As Long, t As Long, m As Long, n_seed As Long
Dim x() As Double, y() As Double, z() As Double, tmp_vec() As Double
Dim tmp_x As Double
    n_seed = UBound(x_seed, 1)
    ReDim y(1 To n, 1 To pn_output)
    For t = 1 To n_seed
        For k = 1 To pn_output
            y(t, k) = x_seed(t, k)
        Next k
    Next t
    x = x_seed
    For t = n_seed + 1 To n
        If (t - n_seed - 1) Mod 50 = 0 Then
            DoEvents
            Application.StatusBar = "cLSTM_B: Generating sequence: " & t & "/" & n
        End If
        m = UBound(x, 1)
        z = InOut(x)
        If maxSelect = True Then
            tmp_x = -1
            j = 1
            For k = 1 To pn_output
                If z(m, k) > tmp_x Then
                    tmp_x = z(m, k)
                    j = k
                End If
            Next k
            For k = 1 To pn_output
                z(m, k) = 0
            Next k
            z(m, j) = 1
        End If
        For k = 1 To pn_output
            y(t, k) = z(m, k)
        Next k
        If m < n_T Then
            tmp_vec = x
            ReDim x(1 To m + 1, 1 To pn_output)
            For i = 1 To m
                For k = 1 To pn_output
                    x(i, k) = tmp_vec(i, k)
                Next k
            Next i
            For k = 1 To pn_output
                x(m + 1, k) = z(m, k)
            Next k
        ElseIf m >= n_T Then
            For i = 1 To m - 1
                For k = 1 To pn_output
                    x(i, k) = x(i + 1, k)
                Next k
            Next i
            For k = 1 To pn_output
                x(m, k) = z(m, k)
            Next k
        End If
    Next t
    Generate = y
    Erase x, y, z, tmp_vec
    Application.StatusBar = False
End Function


