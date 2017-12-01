VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pinput_Lx As Long, pinput_Ly As Long, pinput_Lz As Long
Private pn_output As Long, pn_layer As Long
Private pisPool() As Long
Private pLayers As Collection


Sub Print_Model(vRng As Range)
Dim i As Long, k As Long, n As Long
    i = 0
    With vRng
        For k = 1 To pn_layer
            Call pLayers(k).Print_Model(.Offset(0, i), n)
            i = i + n
        Next k
    End With
End Sub


Sub Read_Model(vRng As Range)
Dim i As Long, k As Long, n As Long
Dim L1 As cConvLayer
Dim L2 As cPoolLayer
    n = Application.WorksheetFunction.CountA(vRng.Resize(1, 5000))
    pn_layer = 0
    Set pLayers = New Collection
    ReDim pisPool(1 To 1)
    With vRng
        i = 0
        For k = 1 To n
            If .Offset(0, i) = "POOLLAYER" Then
                Set L2 = New cPoolLayer
                Call L2.Read_Model(.Offset(0, i), n)
                pLayers.Add L2
                pn_layer = pn_layer + 1
                ReDim Preserve pisPool(1 To pn_layer)
                pisPool(pn_layer) = 1
            Else
                Set L1 = New cConvLayer
                Call L1.Read_Model(.Offset(0, i), n)
                pLayers.Add L1
                pn_layer = pn_layer + 1
                ReDim Preserve pisPool(1 To pn_layer)
            End If
            i = i + n
        Next k
    End With
    With pLayers(1)
        pinput_Lx = .input_dim(1)
        pinput_Ly = .input_dim(2)
        pinput_Lz = .input_dim(3)
    End With
    pn_output = pLayers(pn_layer).output_dim(3)
End Sub


Sub Init(input_Lx As Long, input_Ly As Long, input_Lz As Long, n_output As Long, _
    Optional n_filter As Long = 5, Optional Lx As Long = 5, Optional Ly As Long = 5, Optional stride As Long = 2, _
    Optional strType As String = "CONV", Optional strActivation As String = "RELU")
Dim i As Long, j As Long, k As Long
Dim L1 As cConvLayer
Dim L2 As cPoolLayer
    pinput_Lx = input_Lx
    pinput_Ly = input_Ly
    pinput_Lz = input_Lz
    pn_output = n_output
    ReDim pisPool(1 To 1)
    Set pLayers = New Collection
    If strType = "CONV" Then
        Set L1 = New cConvLayer
        Call L1.Init(input_Lx, input_Ly, input_Lz, n_filter, Lx, Ly, stride, strActivation)
        pLayers.Add L1
        pn_layer = 1
        Set L1 = Nothing
    ElseIf strType = "POOL" Then
        Set L2 = New cPoolLayer
        Call L2.Init(input_Lx, input_Ly, input_Lz, Lx, Ly, stride, strActivation)
        pLayers.Add L2
        pn_layer = 1
        pisPool(1) = 1
        Set L2 = Nothing
    ElseIf strType = "FULL" Then
        Set L1 = New cConvLayer
        Call L1.Init(input_Lx, input_Ly, input_Lz, n_filter, input_Lx, input_Ly, 1, strActivation)
        pLayers.Add L1
        pn_layer = 1
        Set L1 = Nothing
    End If
End Sub

Sub Add_Layer(Optional strType As String = "CONV", Optional n_filter As Long = 5, _
        Optional Lx As Long = 5, Optional Ly As Long = 5, Optional stride As Long = 2, _
        Optional strActivation As String = "RELU")
Dim i As Long, j As Long, k As Long
Dim L1 As cConvLayer
Dim L2 As cPoolLayer
    With pLayers(pn_layer)
        i = .output_dim(1)
        j = .output_dim(2)
        k = .output_dim(3)
    End With
    If strType = "CONV" Then
        Set L1 = New cConvLayer
        Call L1.Init(i, j, k, n_filter, Lx, Ly, stride, strActivation)
        pLayers.Add L1
        Set L1 = Nothing
    ElseIf strType = "POOL" Then
        Set L2 = New cPoolLayer
        Call L2.Init(i, j, k, Lx, Ly, stride, strActivation)
        pLayers.Add L2
        Set L2 = Nothing
    ElseIf strType = "FULL" Then
        Set L1 = New cConvLayer
        Call L1.Init(i, j, k, n_filter, i, j, 1, strActivation)
        pLayers.Add L1
        Set L1 = Nothing
    Else
        Debug.Print "cCNN: Add_Layer: Failed: Invalid layer type '" & strType & "'."
        End
    End If
    pn_layer = pn_layer + 1
    ReDim Preserve pisPool(1 To pn_layer)
    If strType = "POOL" Then pisPool(pn_layer) = 1
End Sub

Sub Reset()
Dim i As Long
    For i = 1 To pn_layer
        If pisPool(i) = 0 Then Call pLayers(i).Reset
    Next i
    Erase pisPool
    For i = pn_layer To 1 Step -1
        pLayers.Remove (i)
    Next i
End Sub

Private Sub Init_Train()
Dim i As Long
    For i = 1 To pn_layer
        Call pLayers(i).Init_Train
    Next i
End Sub

Private Sub Reset_Train()
Dim i As Long
    For i = 1 To pn_layer
        Call pLayers(i).Reset_Train
    Next i
End Sub

'For a single input x, returns a single output x_out
Private Sub InOut(x As Variant, x_out As Variant)
Dim i As Long
Dim u As Variant, v As Variant
    u = x
    For i = 1 To pn_layer - 1
        Call pLayers(i).InOut(u, v)
        u = v
    Next i
    Call pLayers(pn_layer).InOut(u, x_out)
    Erase v, u
End Sub

'For a single input x and target output, stores output of each layer in one forward pass
Private Sub Fwd_Pass(x As Variant, target As Variant, vs As Variant)
Dim i As Long
Dim u As Variant, v As Variant
    ReDim vs(1 To pn_layer - 1)
    u = x
    For i = 1 To pn_layer - 1
        Call pLayers(i).Fwd_Pass(u, v)
        vs(i) = v
        u = v
    Next i
    Call pLayers(pn_layer).Fwd_Pass(u, v, target)
End Sub

'After a forward pass, calculate gradients with back propagation
Private Sub Bkwd_Pass(x As Variant, vs As Variant, grads As Variant, grads_bias As Variant)
Dim i As Long
Dim grad_bias() As Double
    ReDim grads(1 To pn_layer)
    ReDim grads_bias(1 To pn_layer)
    For i = pn_layer To 2 Step -1
        If pisPool(i) = 0 Then
            Call pLayers(i).Bkwd_Pass(vs(i - 1), grads(i), grad_bias, pLayers(i - 1))
            grads_bias(i) = grad_bias
        Else    'Pooling layer has no gradient
            Call pLayers(i).Bkwd_Pass(pLayers(i - 1))
        End If
    Next i
    Call pLayers(1).Bkwd_Pass(x, grads(1), grad_bias)
    grads_bias(1) = grad_bias
End Sub

'Update weights of each lear
Private Sub Update_Wgts(grads As Variant, grads_bias As Variant, learn_rate As Double, momentum As Double)
Dim i As Long
Dim grad_bias() As Double
    For i = 1 To pn_layer
        If pisPool(i) = 0 Then  'Pooling layer has no gradient
            grad_bias = grads_bias(i)
            Call pLayers(i).Update_Wgts(grads(i), grad_bias, learn_rate, momentum)
        End If
    Next i
End Sub

'One single step in stochastic gradient descent
Private Sub Train_SinglePass(x As Variant, target As Variant, learn_rate As Double, momentum As Double)
Dim vs As Variant, grads As Variant, grads_bias As Variant
    Call Fwd_Pass(x, target, vs)
    Call Bkwd_Pass(x, vs, grads, grads_bias)
    Erase vs
    Call Update_Wgts(grads, grads_bias, learn_rate, momentum)
End Sub

'One single step in mini-batch gradient descent
Private Sub Train_BatchPass(xs As Variant, tgts As Variant, learn_rate As Double, momentum As Double)
Dim i As Long, n As Long, batch_size As Long
Dim vs As Variant, grads As Variant, grads_bias As Variant
Dim grads_total As Variant, grads_bias_total As Variant
    batch_size = UBound(xs, 1)
    For n = 1 To batch_size
        Call Fwd_Pass(xs(n), tgts(n), vs)
        Call Bkwd_Pass(xs(n), vs, grads, grads_bias)
        Erase vs
        If n = 1 Then
            grads_total = grads
            grads_bias_total = grads_bias
        Else
            For i = 1 To pn_layer
                If pisPool(i) = 0 Then
                    Call grad_Accumulate(grads_total(i), grads_bias_total(i), grads(i), grads_bias(i))
                End If
            Next i
        End If
    Next n
    Erase grads, grads_bias
    Call Update_Wgts(grads_total, grads_bias_total, learn_rate / batch_size, momentum)
End Sub

Private Sub grad_Accumulate(grad_total As Variant, grad_bias_total As Variant, grad As Variant, grad_bias As Variant)
Dim i As Long
Dim xArr As Variant
    For i = 1 To UBound(grad_total)
        grad_bias_total(i) = grad_bias_total(i) + grad_bias(i)
        xArr = grad_total(i)
        Call Tensor_Add(xArr, grad(i))
        grad_total(i) = xArr
    Next i
End Sub

Private Sub Tensor_Add(A As Variant, B As Variant)
Dim i As Long, k As Long
Dim A1() As Double, B1() As Double
    For k = 1 To UBound(A, 1)
        A1 = A(k)
        B1 = B(k)
        For i = 1 To UBound(A1, 1)
            A1(i) = A1(i) + B1(i)
        Next i
        A(k) = A1
    Next k
End Sub


'Main Routine to train the network
'Input: x(1:N), N sample of images, each picture is a jagged array of vectorize matrix
'           e.g. an RGB image of size 16x16, is store as variant v(1:3), where each element v(i)
'           is a vector of length 256 of one of the RGB-channels
'       target(1:N,1:D), class vector where target(i,j)=1 if sample i belongs to class j.
'       mini_batch, size of mini batch in gradient descent, value of 1 is equivalent to SGD.
'       max_epoch, number of epochs run.
'       learn rate & momentum, the usual neural network parameters.
'Output: train_out(1:N,1:D), prediction after training, same shape as target()
'        train_cost, train_accuracy, validate_cost and validate_accuracy, if supplied, these
'              will be evaluated after each epoch
Sub Train(x As Variant, target() As Double, train_out() As Double, _
        Optional learn_rate As Double = 0.02, Optional momentum As Double = 0.5, _
        Optional max_epoch As Long = 10, Optional mini_batch As Long = 1, _
        Optional train_cost As Variant, Optional train_accuracy As Variant, _
        Optional x_validate As Variant, Optional tgt_validate As Variant, Optional validate_out As Variant, _
        Optional validate_cost As Variant, Optional validate_accuracy As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim epoch As Long, batch_size As Long, n_raw As Long
Dim tmp_x As Double, tmp_y As Double
Dim x_in As Variant, x_tgt As Variant
Dim x_ins As Variant, x_tgts As Variant
Dim shuffle_index() As Long
    
    n_raw = UBound(x)
    If IsMissing(train_cost) = False Then ReDim train_cost(1 To max_epoch)
    If IsMissing(train_accuracy) = False Then ReDim train_accuracy(1 To max_epoch)
    If IsMissing(validate_cost) = False Then ReDim validate_cost(1 To max_epoch)
    If IsMissing(validate_accuracy) = False Then ReDim validate_accuracy(1 To max_epoch)
    
    For epoch = 1 To max_epoch
        DoEvents
        Application.StatusBar = "Epoch: " & epoch & "/" & max_epoch
        Call Init_Train
        'Shuffle learning sample
        shuffle_index = modMath.index_array(1, n_raw)
        Call modMath.Shuffle(shuffle_index)
        
        If mini_batch <= 1 Then 'Stochastic Gradient Descent
        
            Call Init_Jagged_Array(x_tgt, pn_output, 1, 1)
            For n = 1 To n_raw
                If n Mod 5 = 0 Then
                    DoEvents
                    Application.StatusBar = "Epoch: " & epoch & "/" & max_epoch & ", " & n & "/" & n_raw
                End If
                m = shuffle_index(n)
                x_in = x(m)
                For i = 1 To pn_output
                    x_tgt(i)(1) = target(m, i)
                Next i
                Call Train_SinglePass(x_in, x_tgt, learn_rate, momentum)
            Next n

        ElseIf mini_batch > 1 Then 'Mini-Batch Learning
        
            Call Init_Jagged_Array(x_tgt, pn_output, 1, 1)
            For n = 1 To n_raw Step mini_batch
                DoEvents
                Application.StatusBar = "Epoch: " & epoch & "/" & max_epoch & ", " & n & "/" & n_raw
                batch_size = mini_batch
                If (n + batch_size - 1) > n_raw Then batch_size = n_raw - n + 1
                ReDim x_ins(1 To batch_size)
                ReDim x_tgts(1 To batch_size)
                For m = 1 To batch_size
                    j = shuffle_index(n + m - 1)
                    x_ins(m) = x(j)
                    For i = 1 To pn_output
                        x_tgt(i)(1) = target(j, i)
                    Next i
                    x_tgts(m) = x_tgt
                Next m
                Call Train_BatchPass(x_ins, x_tgts, learn_rate, momentum)
            Next n
    
        End If
        
        'Evaluate performance of current epoch
        If IsMissing(x_validate) = False Then Call Predict(x_validate, validate_out, tgt_validate, tmp_x, tmp_y)
        If IsMissing(validate_cost) = False Then validate_cost(epoch) = tmp_x
        If IsMissing(validate_accuracy) = False Then validate_accuracy(epoch) = tmp_y
        
        If IsMissing(train_out) = False Then Call Predict(x, train_out, target, tmp_x, tmp_y)
        If IsMissing(train_cost) = False Then train_cost(epoch) = tmp_x
        If IsMissing(train_accuracy) = False Then train_accuracy(epoch) = tmp_y
        
        'Output network weigts of current epoch
        'Call Print_Model(ActiveWorkbook.Worksheets("TempNet").Range("B2"))
    Next epoch
    Call Reset_Train
    Application.StatusBar = False
End Sub

'Input: x_in(1:N), N samples
'Output: outputs(1:N,1:D), N samples of sofmax output, D is number of possible classes
'        if targets is provided then cost and accuracy will also be evaluated.
Sub Predict(x_in As Variant, outputs As Variant, _
    Optional targets As Variant, Optional cost As Variant, Optional accuracy As Variant)
Dim i As Long, n As Long, k As Long
Dim y As Variant
    n = UBound(x_in, 1)
    ReDim outputs(1 To n, 1 To pn_output)
    For i = 1 To n
        If i Mod 5 = 0 Then
            DoEvents
            Application.StatusBar = "Predicting..." & i & "/" & n
        End If
        Call InOut(x_in(i), y)
        For k = 1 To pn_output
            outputs(i, k) = y(k)(1)
        Next k
    Next i
    If IsMissing(targets) = False Then
        If IsMissing(accuracy) = False Then accuracy = Calc_Accuracy(outputs, targets)
        If IsMissing(cost) = False Then cost = Calc_Cross_Entropy(outputs, targets)
    End If
    Application.StatusBar = False
End Sub

'Calculate cross entropy of a mutliclass vector y(1:N,1:D) where
'N is number of samples, and D is number of possible classes
Private Function Calc_Cross_Entropy(y As Variant, target As Variant) As Double
Dim i As Long, k As Long, n As Long, n_output As Long
Dim tmp_x As Double, tmp_y As Double
    n = UBound(y, 1)
    n_output = UBound(y, 2)
    Calc_Cross_Entropy = 0
    For i = 1 To n
        For k = 1 To n_output
            tmp_y = y(i, k)
            If tmp_y > 0.99999 Then tmp_y = 0.99999
            If tmp_y < 0.00001 Then tmp_y = 0.00001
            Calc_Cross_Entropy = Calc_Cross_Entropy - target(i, k) * Log(tmp_y)
        Next k
    Next i
    Calc_Cross_Entropy = Calc_Cross_Entropy / n
End Function

'Calculate classification accuracy of a mutliclass vector y(1:N,1:D)
'N is number of samples, and D is number of possible classes
Private Function Calc_Accuracy(y As Variant, target As Variant) As Double
Dim i As Long, k As Long, n As Long, n_output As Long
Dim tmp_x As Double, tmp_y As Double, yi As Long, ti As Long
    n = UBound(y, 1)
    n_output = UBound(y, 2)
    Calc_Accuracy = 0
    For i = 1 To n
        tmp_x = -999
        tmp_y = -999
        yi = 0
        ti = 0
        For k = 1 To n_output
            If y(i, k) > tmp_x Then
                tmp_x = y(i, k)
                yi = k
            End If
            If target(i, k) > tmp_y Then
                tmp_y = target(i, k)
                ti = k
            End If
        Next k
        If yi = ti Then Calc_Accuracy = Calc_Accuracy + 1
    Next i
    Calc_Accuracy = Calc_Accuracy / n
End Function


'Create a jagged array x(1:N) where each element is a vector of length (1:ixj)
Private Sub Init_Jagged_Array(x As Variant, n As Long, i As Long, j As Long)
Dim k As Long
Dim tmp_vec() As Double
    ReDim x(1 To n)
    ReDim tmp_vec(1 To i * j)
    For k = 1 To n
        x(k) = tmp_vec
    Next k
End Sub
