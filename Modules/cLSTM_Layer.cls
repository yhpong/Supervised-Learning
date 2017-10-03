VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLSTM_Layer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pn_input As Long
Private pn_hidden As Long
Private pn_output As Long

Private pbias_i() As Variant, pbias_i_chg() As Variant, pbias_i_chg_batch() As Variant
Private pbias_f() As Variant, pbias_f_chg() As Variant, pbias_f_chg_batch() As Variant
Private pbias_o() As Variant, pbias_o_chg() As Variant, pbias_o_chg_batch() As Variant
Private pbias_c() As Variant, pbias_c_chg() As Variant, pbias_c_chg_batch() As Variant

Private pw_in_i() As Variant, pw_hidden_i() As Variant, pw_c_i() As Variant
Private pw_in_f() As Variant, pw_hidden_f() As Variant, pw_c_f() As Variant
Private pw_in_c() As Variant, pw_hidden_c() As Variant
Private pw_in_o() As Variant, pw_hidden_o() As Variant, pw_c_o() As Variant

Private pw_in_i_chg() As Variant, pw_hidden_i_chg() As Variant, pw_c_i_chg() As Variant
Private pw_in_f_chg() As Variant, pw_hidden_f_chg() As Variant, pw_c_f_chg() As Variant
Private pw_in_c_chg() As Variant, pw_hidden_c_chg() As Variant
Private pw_in_o_chg() As Variant, pw_hidden_o_chg() As Variant, pw_c_o_chg() As Variant

Private pw_in_i_chg_batch() As Variant, pw_hidden_i_chg_batch() As Variant, pw_c_i_chg_batch() As Variant
Private pw_in_f_chg_batch() As Variant, pw_hidden_f_chg_batch() As Variant, pw_c_f_chg_batch() As Variant
Private pw_in_c_chg_batch() As Variant, pw_hidden_c_chg_batch() As Variant
Private pw_in_o_chg_batch() As Variant, pw_hidden_o_chg_batch() As Variant, pw_c_o_chg_batch() As Variant

Private py_i() As Double, pderiv_i() As Double
Private py_f() As Double, pderiv_f() As Double
Private pstate() As Double, py_c() As Double, pderiv_c() As Double
Private py_o() As Double, pderiv_o() As Double
Private py_c_output() As Double, pderiv_c_output() As Double

Private pdelta_o() As Double
Private pdelta_c() As Double
Private pdelta_f() As Double
Private pdelta_i() As Double

Private perr_c_output() As Double
Private perr_state() As Double



Sub Get_Input_Gate(bias As Variant, w_in As Variant, w_hidden As Variant, w_c As Variant)
    bias = pbias_i
    w_in = pw_in_i
    w_hidden = pw_hidden_i
    w_c = pw_c_i
End Sub
Sub Assign_Input_Gate(bias As Variant, w_in As Variant, w_hidden As Variant, w_c As Variant)
    pbias_i = bias
    pw_in_i = w_in
    pw_hidden_i = w_hidden
    pw_c_i = w_c
End Sub

Sub Get_Forget_Gate(bias As Variant, w_in As Variant, w_hidden As Variant, w_c As Variant)
    bias = pbias_f
    w_in = pw_in_f
    w_hidden = pw_hidden_f
    w_c = pw_c_f
End Sub
Sub Assign_Forget_Gate(bias As Variant, w_in As Variant, w_hidden As Variant, w_c As Variant)
    pbias_f = bias
    pw_in_f = w_in
    pw_hidden_f = w_hidden
    pw_c_f = w_c
End Sub

Sub Get_Output_Gate(bias As Variant, w_in As Variant, w_hidden As Variant, w_c As Variant)
    bias = pbias_o
    w_in = pw_in_o
    w_hidden = pw_hidden_o
    w_c = pw_c_o
End Sub
Sub Assign_Output_Gate(bias As Variant, w_in As Variant, w_hidden As Variant, w_c As Variant)
    pbias_o = bias
    pw_in_o = w_in
    pw_hidden_o = w_hidden
    pw_c_o = w_c
End Sub

Sub Get_Cell(bias As Variant, w_in As Variant, w_hidden As Variant)
    bias = pbias_c
    w_in = pw_in_c
    w_hidden = pw_hidden_c
End Sub
Sub Assign_Cell(bias As Variant, w_in As Variant, w_hidden As Variant)
    pbias_c = bias
    pw_in_c = w_in
    pw_hidden_c = w_hidden
End Sub

Public Property Get y_output() As Double()
Dim i As Long, j As Long, m As Long, n As Long
Dim y() As Double
    m = UBound(py_c_output, 1) - 1
    n = UBound(py_c_output, 2)
    ReDim y(1 To m, 1 To n)
    For i = 1 To m
        For j = 1 To n
            y(i, j) = py_c_output(i, j)
        Next j
    Next i
    y_output = y
End Property


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

Private Function Activation_Gate(x As Double, x_deriv As Double) As Double
    Activation_Gate = Sigmoid(x)
    x_deriv = Activation_Gate * (1 - Activation_Gate)
End Function

Private Function Activation_Cell(x As Double, x_deriv As Double) As Double
'    Activation_Cell = Sigmoid(x)
'    x_deriv = Activation_Cell * (1 - Activation_Cell)
    Activation_Cell = 4 * Sigmoid(x) - 2
    x_deriv = 1 - (Activation_Cell ^ 2) / 4
End Function

Private Function Activation_Cell_Output(x As Double, x_deriv As Double) As Double
'    Activation_Cell_Output = Sigmoid(x)
'    x_deriv = Activation_Cell_Output * (1 - Activation_Cell_Output)
    Activation_Cell_Output = 4 * Sigmoid(x) - 2
    x_deriv = 1 - (Activation_Cell_Output ^ 2) / 4
End Function
'*******************************************************
'*******************************************************
 
 
Sub Reset()
    pn_input = 0
    pn_hidden = 0
    pn_output = 0
    Erase pbias_i, pbias_f, pbias_c, pbias_o
    Erase pw_in_i, pw_hidden_i, pw_c_i
    Erase pw_in_f, pw_hidden_f, pw_c_f
    Erase pw_in_c, pw_hidden_c
    Erase pw_in_o, pw_hidden_o, pw_c_o
End Sub

Sub Init(n_input As Long, n_hidden As Long, n_output As Long, _
        Optional bias_i As Double = 0, Optional bias_f As Double = 2, _
        Optional bias_c As Double = 0, Optional bias_o As Double = 0)
Dim i As Long, j As Long
Dim tmp_x As Double

    pn_input = n_input
    pn_hidden = n_hidden
    pn_output = n_output
    
    ReDim pbias_i(1 To n_hidden)
    ReDim pbias_f(1 To n_hidden)
    ReDim pbias_c(1 To n_hidden)
    ReDim pbias_o(1 To n_hidden)
    
    ReDim pw_in_i(1 To n_input, 1 To n_hidden)
    ReDim pw_in_f(1 To n_input, 1 To n_hidden)
    ReDim pw_in_c(1 To n_input, 1 To n_hidden)
    ReDim pw_in_o(1 To n_input, 1 To n_hidden)
    ReDim pw_hidden_i(1 To n_hidden, 1 To n_hidden)
    ReDim pw_hidden_f(1 To n_hidden, 1 To n_hidden)
    ReDim pw_hidden_c(1 To n_hidden, 1 To n_hidden)
    ReDim pw_hidden_o(1 To n_hidden, 1 To n_hidden)
    ReDim pw_c_i(1 To n_hidden)
    ReDim pw_c_f(1 To n_hidden)
    ReDim pw_c_o(1 To n_hidden)
    
    Randomize
    tmp_x = 1# / (pn_hidden * (pn_input + pn_hidden + pn_output))
    For i = 1 To pn_hidden
        
        pbias_i(i) = bias_i
        pbias_f(i) = bias_f
        pbias_c(i) = bias_c
        pbias_o(i) = bias_o
        
        pw_c_i(i) = -tmp_x + Rnd() * 2 * tmp_x
        pw_c_f(i) = -tmp_x + Rnd() * 2 * tmp_x
        pw_c_o(i) = -tmp_x + Rnd() * 2 * tmp_x
    
        For j = 1 To pn_input
            pw_in_i(j, i) = -tmp_x + Rnd() * 2 * tmp_x
            pw_in_f(j, i) = -tmp_x + Rnd() * 2 * tmp_x
            pw_in_c(j, i) = -tmp_x + Rnd() * 2 * tmp_x
            pw_in_o(j, i) = -tmp_x + Rnd() * 2 * tmp_x
        Next j
        
        For j = 1 To pn_hidden
            pw_hidden_i(i, j) = -tmp_x + Rnd() * 2 * tmp_x
            pw_hidden_f(i, j) = -tmp_x + Rnd() * 2 * tmp_x
            pw_hidden_c(i, j) = -tmp_x + Rnd() * 2 * tmp_x
            pw_hidden_o(i, j) = -tmp_x + Rnd() * 2 * tmp_x
        Next j
    Next i
    
End Sub

Sub Train_Batch_Init()
    ReDim pbias_i_chg(1 To pn_hidden)
    ReDim pbias_f_chg(1 To pn_hidden)
    ReDim pbias_c_chg(1 To pn_hidden)
    ReDim pbias_o_chg(1 To pn_hidden)
    ReDim pw_in_i_chg(1 To pn_input, 1 To pn_hidden)
    ReDim pw_in_f_chg(1 To pn_input, 1 To pn_hidden)
    ReDim pw_in_c_chg(1 To pn_input, 1 To pn_hidden)
    ReDim pw_in_o_chg(1 To pn_input, 1 To pn_hidden)
    ReDim pw_hidden_i_chg(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_hidden_f_chg(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_hidden_c_chg(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_hidden_o_chg(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_c_i_chg(1 To pn_hidden)
    ReDim pw_c_f_chg(1 To pn_hidden)
    ReDim pw_c_o_chg(1 To pn_hidden)
End Sub

Sub Train_Batch_Reset()
    Erase pbias_i_chg, pbias_f_chg, pbias_c_chg, pbias_o_chg
    Erase pw_in_i_chg, pw_hidden_i_chg, pw_c_i_chg
    Erase pw_in_f_chg, pw_hidden_f_chg, pw_c_f_chg
    Erase pw_in_c_chg, pw_hidden_c_chg
    Erase pw_in_o_chg, pw_hidden_o_chg, pw_c_o_chg
End Sub

Sub Train_Mini_Batch_Init()
    ReDim pbias_i_chg_batch(1 To pn_hidden)
    ReDim pbias_f_chg_batch(1 To pn_hidden)
    ReDim pbias_c_chg_batch(1 To pn_hidden)
    ReDim pbias_o_chg_batch(1 To pn_hidden)
    ReDim pw_in_i_chg_batch(1 To pn_input, 1 To pn_hidden)
    ReDim pw_in_f_chg_batch(1 To pn_input, 1 To pn_hidden)
    ReDim pw_in_c_chg_batch(1 To pn_input, 1 To pn_hidden)
    ReDim pw_in_o_chg_batch(1 To pn_input, 1 To pn_hidden)
    ReDim pw_hidden_i_chg_batch(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_hidden_f_chg_batch(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_hidden_c_chg_batch(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_hidden_o_chg_batch(1 To pn_hidden, 1 To pn_hidden)
    ReDim pw_c_i_chg_batch(1 To pn_hidden)
    ReDim pw_c_f_chg_batch(1 To pn_hidden)
    ReDim pw_c_o_chg_batch(1 To pn_hidden)
End Sub

Sub Train_Mini_Batch_Reset()
    Erase pbias_i_chg_batch, pbias_f_chg_batch, pbias_c_chg_batch, pbias_o_chg_batch
    Erase pw_in_i_chg_batch, pw_hidden_i_chg_batch, pw_c_i_chg_batch
    Erase pw_in_f_chg_batch, pw_hidden_f_chg_batch, pw_c_f_chg_batch
    Erase pw_in_c_chg_batch, pw_hidden_c_chg_batch
    Erase pw_in_o_chg_batch, pw_hidden_o_chg_batch, pw_c_o_chg_batch
End Sub

Sub Train_Init(n_T As Long)
    ReDim py_i(1 To n_T, 1 To pn_hidden)
    ReDim py_f(0 To n_T + 1, 1 To pn_hidden)
    ReDim py_c(1 To n_T, 1 To pn_hidden)
    ReDim py_o(1 To n_T, 1 To pn_hidden)
    ReDim py_c_output(0 To n_T + 1, 1 To pn_hidden)
    
    ReDim pderiv_i(1 To n_T, 1 To pn_hidden)
    ReDim pderiv_f(1 To n_T, 1 To pn_hidden)
    ReDim pstate(0 To n_T + 1, 1 To pn_hidden)
    ReDim pderiv_c(1 To n_T, 1 To pn_hidden)
    ReDim pderiv_o(1 To n_T, 1 To pn_hidden)
    ReDim pderiv_c_output(1 To n_T, 1 To pn_hidden)
    
    ReDim pdelta_out(1 To n_T, 1 To pn_output)
    ReDim pdelta_o(0 To n_T + 1, 1 To pn_hidden)
    ReDim pdelta_c(0 To n_T + 1, 1 To pn_hidden)
    ReDim pdelta_f(0 To n_T + 1, 1 To pn_hidden)
    ReDim pdelta_i(0 To n_T + 1, 1 To pn_hidden)

    ReDim perr_c_output(1 To n_T, 1 To pn_hidden)
    ReDim perr_state(0 To n_T + 1, 1 To pn_hidden)
End Sub

Sub Train_Reset()
    Erase py_i, pderiv_i, pdelta_i
    Erase py_f, pderiv_f, pdelta_f
    Erase py_c, pderiv_c, pdelta_c, pstate
    Erase py_o, pderiv_o, pdelta_o
    Erase py_c_output
    Erase perr_c_output, perr_state
End Sub



Sub FwdPass(x() As Double, t As Long, t_prev As Long)
Dim i As Long, j As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_z_deriv As Double
For j = 1 To pn_hidden
    tmp_x = pstate(t_prev, j) * pw_c_i(j) + pbias_i(j)
    tmp_y = pstate(t_prev, j) * pw_c_f(j) + pbias_f(j)
    tmp_z = pbias_c(j)
    For i = 1 To pn_input
        tmp_x = tmp_x + x(t, i) * pw_in_i(i, j)
        tmp_y = tmp_y + x(t, i) * pw_in_f(i, j)
        tmp_z = tmp_z + x(t, i) * pw_in_c(i, j)
    Next i
    For i = 1 To pn_hidden
        tmp_x = tmp_x + py_c_output(t_prev, i) * pw_hidden_i(i, j)
        tmp_y = tmp_y + py_c_output(t_prev, i) * pw_hidden_f(i, j)
        tmp_z = tmp_z + py_c_output(t_prev, i) * pw_hidden_c(i, j)
    Next i
    py_i(t, j) = Activation_Gate(tmp_x, pderiv_i(t, j))
    py_f(t, j) = Activation_Gate(tmp_y, pderiv_f(t, j))
    py_c(t, j) = Activation_Cell(tmp_z, tmp_z_deriv)
    pderiv_c(t, j) = py_i(t, j) * tmp_z_deriv
    pstate(t, j) = py_f(t, j) * pstate(t_prev, j) + py_i(t, j) * py_c(t, j)

    tmp_x = pstate(t, j) * pw_c_o(j) + pbias_o(j)
    For i = 1 To pn_input
        tmp_x = tmp_x + x(t, i) * pw_in_o(i, j)
    Next i
    For i = 1 To pn_hidden
        tmp_x = tmp_x + py_c_output(t_prev, i) * pw_hidden_o(i, j)
    Next i
    py_o(t, j) = Activation_Gate(tmp_x, pderiv_o(t, j))

    py_c_output(t, j) = py_o(t, j) * Activation_Cell_Output(pstate(t, j), tmp_x)
    pderiv_c_output(t, j) = py_o(t, j) * tmp_x
Next j
End Sub


Sub BkwdPass(t As Long, t_next As Long, t_prev As Long, w_cell_output_out As Variant, delta_out() As Double)
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double, tmp_y As Double
    For i = 1 To pn_hidden
        tmp_x = 0
        For k = 1 To pn_output
            tmp_x = tmp_x + w_cell_output_out(i, k) * delta_out(t, k)
        Next k
        For j = 1 To pn_hidden
            tmp_x = tmp_x + pw_hidden_o(i, j) * pdelta_o(t_next, j) _
                + pw_hidden_c(i, j) * pdelta_c(t_next, j) _
                + pw_hidden_f(i, j) * pdelta_f(t_next, j) _
                + pw_hidden_i(i, j) * pdelta_i(t_next, j)
        Next j
        perr_c_output(t, i) = tmp_x
        pdelta_o(t, i) = pderiv_o(t, i) * Activation_Cell_Output(pstate(t, i), tmp_y) * tmp_x
    
        perr_state(t, i) = pderiv_c_output(t, i) * perr_c_output(t, i) + _
            py_f(t_next, i) * perr_state(t_next, i) + _
            pw_c_i(i) * pdelta_i(t_next, i) + _
            pw_c_f(i) * pdelta_f(t_next, i) + _
            pw_c_o(i) * pdelta_o(t, i)
            
        pdelta_c(t, i) = pderiv_c(t, i) * perr_state(t, i)
        pdelta_f(t, i) = pderiv_f(t, i) * pstate(t_prev, i) * perr_state(t, i)
        pdelta_i(t, i) = pderiv_i(t, i) * py_c(t, i) * perr_state(t, i)
    Next i
End Sub



Sub Accumulate_Wgt_chg(x() As Double, Optional t_lag = 1)
Dim i As Long, j As Long, t As Long, n_T As Long
n_T = UBound(x, 1)
For j = 1 To pn_hidden
    For t = 1 To n_T
        pbias_i_chg_batch(j) = pbias_i_chg_batch(j) + pdelta_i(t, j)
        pbias_f_chg_batch(j) = pbias_f_chg_batch(j) + pdelta_f(t, j)
        pbias_o_chg_batch(j) = pbias_o_chg_batch(j) + pdelta_o(t, j)
        pbias_c_chg_batch(j) = pbias_c_chg_batch(j) + pdelta_c(t, j)
        pw_c_i_chg_batch(j) = pw_c_i_chg_batch(j) + pdelta_i(t, j) * pstate(t - t_lag, j)
        pw_c_f_chg_batch(j) = pw_c_f_chg_batch(j) + pdelta_f(t, j) * pstate(t - t_lag, j)
        pw_c_o_chg_batch(j) = pw_c_o_chg_batch(j) + pdelta_o(t, j) * pstate(t, j)
        For i = 1 To pn_input
            pw_in_i_chg_batch(i, j) = pw_in_i_chg_batch(i, j) + pdelta_i(t, j) * x(t, i)
            pw_in_f_chg_batch(i, j) = pw_in_f_chg_batch(i, j) + pdelta_f(t, j) * x(t, i)
            pw_in_o_chg_batch(i, j) = pw_in_o_chg_batch(i, j) + pdelta_o(t, j) * x(t, i)
            pw_in_c_chg_batch(i, j) = pw_in_c_chg_batch(i, j) + pdelta_c(t, j) * x(t, i)
        Next i
        For i = 1 To pn_hidden
            pw_hidden_i_chg_batch(i, j) = pw_hidden_i_chg_batch(i, j) + pdelta_i(t, j) * py_c_output(t - t_lag, i)
            pw_hidden_f_chg_batch(i, j) = pw_hidden_f_chg_batch(i, j) + pdelta_f(t, j) * py_c_output(t - t_lag, i)
            pw_hidden_o_chg_batch(i, j) = pw_hidden_o_chg_batch(i, j) + pdelta_o(t, j) * py_c_output(t - t_lag, i)
            pw_hidden_c_chg_batch(i, j) = pw_hidden_c_chg_batch(i, j) + pdelta_c(t, j) * py_c_output(t - t_lag, i)
        Next i
    Next t
Next j
End Sub


Sub Update_wgts(learn_rate As Double, momentum As Double)
Dim i As Long, j As Long
'wchg = momentum * wchg - learn_rate * gradient
For j = 1 To pn_hidden
    pbias_i_chg(j) = momentum * pbias_i_chg(j) - learn_rate * pbias_i_chg_batch(j)
    pbias_f_chg(j) = momentum * pbias_f_chg(j) - learn_rate * pbias_f_chg_batch(j)
    pbias_o_chg(j) = momentum * pbias_o_chg(j) - learn_rate * pbias_o_chg_batch(j)
    pbias_c_chg(j) = momentum * pbias_c_chg(j) - learn_rate * pbias_c_chg_batch(j)
    pw_c_i_chg(j) = momentum * pw_c_i_chg(j) - learn_rate * pw_c_i_chg_batch(j)
    pw_c_f_chg(j) = momentum * pw_c_f_chg(j) - learn_rate * pw_c_f_chg_batch(j)
    pw_c_o_chg(j) = momentum * pw_c_o_chg(j) - learn_rate * pw_c_o_chg_batch(j)
    For i = 1 To pn_input
        pw_in_i_chg(i, j) = momentum * pw_in_i_chg(i, j) - learn_rate * pw_in_i_chg_batch(i, j)
        pw_in_f_chg(i, j) = momentum * pw_in_f_chg(i, j) - learn_rate * pw_in_f_chg_batch(i, j)
        pw_in_o_chg(i, j) = momentum * pw_in_o_chg(i, j) - learn_rate * pw_in_o_chg_batch(i, j)
        pw_in_c_chg(i, j) = momentum * pw_in_c_chg(i, j) - learn_rate * pw_in_c_chg_batch(i, j)
    Next i
    For i = 1 To pn_hidden
        pw_hidden_i_chg(i, j) = momentum * pw_hidden_i_chg(i, j) - learn_rate * pw_hidden_i_chg_batch(i, j)
        pw_hidden_f_chg(i, j) = momentum * pw_hidden_f_chg(i, j) - learn_rate * pw_hidden_f_chg_batch(i, j)
        pw_hidden_o_chg(i, j) = momentum * pw_hidden_o_chg(i, j) - learn_rate * pw_hidden_o_chg_batch(i, j)
        pw_hidden_c_chg(i, j) = momentum * pw_hidden_c_chg(i, j) - learn_rate * pw_hidden_c_chg_batch(i, j)
    Next i
Next j

'w = w + w_chg
For i = 1 To pn_hidden
    pw_c_i(i) = pw_c_i(i) + pw_c_i_chg(i)
    pw_c_f(i) = pw_c_f(i) + pw_c_f_chg(i)
    pw_c_o(i) = pw_c_o(i) + pw_c_o_chg(i)
    pbias_i(i) = pbias_i(i) + pbias_i_chg(i)
    pbias_f(i) = pbias_f(i) + pbias_f_chg(i)
    pbias_o(i) = pbias_o(i) + pbias_o_chg(i)
    pbias_c(i) = pbias_c(i) + pbias_c_chg(i)
    For j = 1 To pn_input
        pw_in_i(j, i) = pw_in_i(j, i) + pw_in_i_chg(j, i)
        pw_in_f(j, i) = pw_in_f(j, i) + pw_in_f_chg(j, i)
        pw_in_c(j, i) = pw_in_c(j, i) + pw_in_c_chg(j, i)
        pw_in_o(j, i) = pw_in_o(j, i) + pw_in_o_chg(j, i)
    Next j
    For j = 1 To pn_hidden
        pw_hidden_i(i, j) = pw_hidden_i(i, j) + pw_hidden_i_chg(i, j)
        pw_hidden_f(i, j) = pw_hidden_f(i, j) + pw_hidden_f_chg(i, j)
        pw_hidden_c(i, j) = pw_hidden_c(i, j) + pw_hidden_c_chg(i, j)
        pw_hidden_o(i, j) = pw_hidden_o(i, j) + pw_hidden_o_chg(i, j)
    Next j
Next i

Call max_norm_reg(3)
End Sub


Private Sub max_norm_reg(Optional max_norm As Double = 3)
Dim i As Long, j As Long
Dim tmp_x As Double
    For i = 1 To pn_hidden
        If (pw_c_i(i) ^ 2) > max_norm Then pw_c_i(i) = Sgn(pw_c_i(i)) * Sqr(max_norm)
        If (pw_c_f(i) ^ 2) > max_norm Then pw_c_f(i) = Sgn(pw_c_f(i)) * Sqr(max_norm)
        If (pw_c_o(i) ^ 2) > max_norm Then pw_c_o(i) = Sgn(pw_c_o(i)) * Sqr(max_norm)
        Call max_norm_reg_w(pw_in_i, max_norm)
        Call max_norm_reg_w(pw_in_f, max_norm)
        Call max_norm_reg_w(pw_in_c, max_norm)
        Call max_norm_reg_w(pw_in_o, max_norm)
        Call max_norm_reg_w(pw_hidden_i, max_norm)
        Call max_norm_reg_w(pw_hidden_f, max_norm)
        Call max_norm_reg_w(pw_hidden_c, max_norm)
        Call max_norm_reg_w(pw_hidden_o, max_norm)
    Next i
End Sub



Private Sub max_norm_reg_w(w As Variant, max_norm)
Dim i As Long, j As Long, m As Long, n As Long
Dim tmp_x As Double
    m = UBound(w, 1)
    n = UBound(w, 2)
    For i = 1 To n
        tmp_x = 0
        For j = 1 To m
            tmp_x = tmp_x + w(j, i) ^ 2
        Next j
        If tmp_x > max_norm Then
            tmp_x = Sqr(max_norm / tmp_x)
            For j = 1 To m
                w(j, i) = w(j, i) * tmp_x
            Next j
        End If
    Next i
End Sub

