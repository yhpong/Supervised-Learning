VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pChild_L As cCART, pChild_R As cCART
Private pSplit_Dim As Long, pSplit_Val As Double
Private pisLeaf As Boolean, pResponse As Double
Private Const pINFINITY As Double = 2E+30


'Output data to excel range that can be plot as a tree using scatter plots
Sub Chart_Tree(vRng As Range, Optional factor_labels As Variant)
Dim i As Long
Dim labels() As String
Dim vArr As Variant, uArr As Variant, iArr() As Long
    If IsMissing(factor_labels) = False Then
        labels = factor_labels
    Else
        ReDim labels(1 To 999)
        For i = 1 To 999
            labels(i) = VBA.Format(i, "000")
        Next i
    End If
    Call prv_chart_array(vArr, labels, 0, 0, 1, -1)
    Call prv_fill_pos(vArr, iArr, 1)
    Call prv_draw_lines(vArr, uArr)
    With vRng
        .Resize(UBound(vArr, 2), 3).Value = Application.WorksheetFunction.Transpose(vArr)
        .Offset(0, 4).Resize(UBound(uArr, 2), UBound(uArr, 1)).Value = Application.WorksheetFunction.Transpose(uArr)
    End With
    Erase iArr, uArr, vArr
End Sub


'Convert binary tree to array representation, set y-positions of all nodes and x-positions of leafs
Function prv_chart_array(vArr As Variant, labels() As String, i_pos As Long, depth As Long, leaf_order As Long, Optional i_parent As Long = -1) As Long
Dim i As Long, j As Long, k As Long, tmp_x As Double
    If i_pos = 0 Then
        i_pos = 1
        ReDim vArr(1 To 6, 1 To 1)
    Else
        i_pos = UBound(vArr, 2) + 1
        ReDim Preserve vArr(1 To 6, 1 To i_pos)
    End If
    vArr(3, i_pos) = -depth
    vArr(6, i_pos) = i_parent
    prv_chart_array = i_pos
    k = i_pos
    If pisLeaf = False Then
        vArr(1, i_pos) = labels(pSplit_Dim) & "<" & VBA.Format(pSplit_Val, "0.0")
        vArr(4, k) = pChild_L.prv_chart_array(vArr, labels, i_pos, depth + 1, leaf_order, k)
        vArr(5, k) = pChild_R.prv_chart_array(vArr, labels, i_pos, depth + 1, leaf_order, k)
    Else
        vArr(1, k) = "ret " & VBA.Format(pResponse, "0.0")
        vArr(2, k) = leaf_order
        vArr(4, k) = -1
        vArr(5, k) = -1
        leaf_order = leaf_order + 1
    End If
End Function

'Set x-positions of all nodes to be average of their left & right child
Private Sub prv_fill_pos(vArr As Variant, isFix() As Long, Optional i_pos As Long = 1)
Dim i As Long, j As Long, tmp_x As Double
    If i_pos = 1 Then ReDim isFix(1 To UBound(vArr, 2))
    If vArr(4, i_pos) = -1 Then
        tmp_x = vArr(2, i_pos)
        j = i_pos
        Do
            j = vArr(6, j)
            If isFix(j) < 2 Then
                vArr(2, j) = vArr(2, j) + tmp_x / 2
                isFix(j) = isFix(j) + 1
            End If
            If isFix(j) < 2 Then Exit Do
            tmp_x = vArr(2, j)
        Loop While j > 1
    Else
        i = vArr(4, i_pos)
        j = vArr(5, i_pos)
        Call prv_fill_pos(vArr, isFix, i)
        Call prv_fill_pos(vArr, isFix, j)
    End If
End Sub

'Draw connecting lines in tree
Private Sub prv_draw_lines(vArr As Variant, uArr As Variant, Optional i_pos As Long = 1)
Dim i As Long, j As Long, k As Long
    If i_pos = 1 Then
        k = 1
        ReDim uArr(1 To 2, 1 To 1)
    Else
        k = UBound(uArr, 2) + 1
        ReDim Preserve uArr(1 To 2, 1 To k)
    End If
    uArr(1, k) = vArr(2, i_pos)
    uArr(2, k) = vArr(3, i_pos)
    If vArr(4, i_pos) <> -1 Then
        i = vArr(4, i_pos)
        Call prv_draw_lines(vArr, uArr, i)
        
        k = UBound(uArr, 2) + 2
        ReDim Preserve uArr(1 To 2, 1 To k)
        uArr(1, k) = vArr(2, i_pos)
        uArr(2, k) = vArr(3, i_pos)
        
        j = vArr(5, i_pos)
        Call prv_draw_lines(vArr, uArr, j)
    End If
End Sub

'Print model to excel range using array representation
Sub Print_Model(vRng As Range)
Dim i As Long
Dim vArr As Variant
    i = prv_Build_Array(vArr, 0)
    With vRng
        .Value = UBound(vArr, 2)
        .Offset(1, 0).Resize(UBound(vArr, 2), UBound(vArr, 1)).Value = Application.WorksheetFunction.Transpose(vArr)
    End With
    Erase vArr
End Sub

'Read model from excel range
Sub Read_Model(vRng As Range)
Dim n As Long
Dim vArr As Variant
    With vRng
        n = .Value
        vArr = .Offset(1, 0).Resize(n, 4).Value
    End With
    Call prv_Read_Array(vArr, 1)
    Erase vArr
End Sub

'Convert array to binary tree
Sub prv_Read_Array(vArr As Variant, Optional i_pos As Long = 1)
Dim i As Long, j As Long
Dim dTL As cCART, dTR As cCART
    If vArr(i_pos, 1) = -1 Then
        pisLeaf = True
        pResponse = vArr(i_pos, 2)
    ElseIf vArr(i_pos, 1) > 0 Then
        pisLeaf = False
        pSplit_Dim = vArr(i_pos, 1)
        pSplit_Val = vArr(i_pos, 2)
        i = vArr(i_pos, 3): j = vArr(i_pos, 4)
        Set dTL = New cCART: Set dTR = New cCART
        Call dTL.prv_Read_Array(vArr, i): Call dTR.prv_Read_Array(vArr, j)
        Set pChild_L = dTL: Set pChild_R = dTR
    End If
End Sub

'Convert binary tree to array representation
Function prv_Build_Array(vArr As Variant, Optional i_pos As Long = 0) As Long
Dim k As Long
    If i_pos = 0 Then
        i_pos = 1
        ReDim vArr(1 To 4, 1 To 1)
    Else
        i_pos = UBound(vArr, 2) + 1
        ReDim Preserve vArr(1 To 4, 1 To i_pos)
    End If
    prv_Build_Array = i_pos
    k = i_pos
    If pisLeaf = False Then
        vArr(1, i_pos) = pSplit_Dim
        vArr(2, i_pos) = pSplit_Val
        vArr(3, k) = pChild_L.prv_Build_Array(vArr, i_pos)
        vArr(4, k) = pChild_R.prv_Build_Array(vArr, i_pos)
    Else
        vArr(1, i_pos) = -1
        vArr(2, i_pos) = pResponse
        vArr(3, i_pos) = -1
        vArr(4, i_pos) = -1
    End If
End Function

'Make predictions y(1:N) from inputs x(1:N,1:D)
Sub Predict(x As Variant, y As Variant)
Dim i As Long, j As Long, n As Long, n_dimension As Long
Dim xi As Variant, yi As Variant
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim y(1 To n)
    ReDim xi(1 To n_dimension)
    For i = 1 To n
        For j = 1 To n_dimension
            xi(j) = x(i, j)
        Next j
        Call Predict_Single(xi, yi)
        y(i) = yi
    Next i
    Erase xi
End Sub

'Make prediction y from input x(1:D)
Sub Predict_Single(x As Variant, y As Variant)
    If pisLeaf = True Then
        y = pResponse
        Exit Sub
    End If
    If x(pSplit_Dim) < pSplit_Val Then
        Call pChild_L.Predict_Single(x, y)
    Else
        Call pChild_R.Predict_Single(x, y)
    End If
End Sub

'Grow tree from training data with response y(1:N) and explanatory variables x(1:N,1:D)
Sub Fit(y As Variant, x As Variant, _
    Optional min_size As Long = 1, Optional max_depth As Long = 20, _
    Optional y_categorical As Boolean = False, Optional n_class As Long = 0, _
    Optional BagType As String = "NONE")
Dim i As Long, n As Long, max_feature As Long, n_dimension As Long
Dim leaf() As Long
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    Select Case VBA.UCase(BagType)
        Case "NONE"
            max_feature = n_dimension
        Case "SQRT"
            max_feature = Int(Sqr(n_dimension))
        Case "LOG"
            max_feature = Int(Log(n_dimension) / Log(2))
        Case Else
            Debug.Print "cCART: Fit:  Invalid BagType " & BagType
            Exit Sub
    End Select
    
    pisLeaf = False
    ReDim leaf(1 To n)
    For i = 1 To n
        leaf(i) = i
    Next i
    
    If y_categorical = False Then
        Call prv_Split_Recursive(y, x, leaf, min_size, max_depth, 0, max_feature, False)
    Else
        If n_class = 0 Then
            Debug.Print "cCART: Fit: number of target class must be supplied if y() is categorical."
            Exit Sub
        End If
        Call prv_Split_Recursive(y, x, leaf, min_size, max_depth, 0, max_feature, True, n_class)
    End If
End Sub

Sub prv_Split_Recursive(y As Variant, x As Variant, leaf() As Long, _
        min_size As Long, max_depth As Long, cur_depth As Long, _
        max_feature As Long, y_categorical As Boolean, Optional n_class As Long = 0)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, kk As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double
Dim best_dim As Long, n1 As Long, n2 As Long, n1_best As Long, n2_best As Long
Dim best_x_tmp As Double, err_tmp As Double, best_x As Double, err_best As Double
Dim xk() As Double, sort_idx() As Long, leaf_1() As Long, leaf_2() As Long, feature_set() As Long
Dim y_sum As Double, y2_sum As Double
Dim class_count() As Long

    n = UBound(leaf, 1)
    n_dimension = UBound(x, 2)
    
    If n <= min_size Or cur_depth = max_depth Then
        Call Set_As_Leaf(y, leaf, y_categorical, n_class)
        Exit Sub
    End If

    If y_categorical = True Then
        ReDim class_count(n_class)
        For i = 1 To n
            j = y(leaf(i))
            class_count(j) = class_count(j) + 1
        Next i
    Else
        y_sum = 0
        y2_sum = 0
        For i = 1 To n
            y_sum = y_sum + y(leaf(i))
            y2_sum = y2_sum + y(leaf(i)) ^ 2
        Next i
    End If

    n1_best = 0
    n2_best = 0
    err_best = pINFINITY
    ReDim xk(1 To n)
    
    feature_set = Random_Feature(max_feature, n_dimension)
    
    For kk = 1 To max_feature

        k = feature_set(kk)
        
        'Sort current dimension for faster calculation
        For i = 1 To n
            xk(i) = x(leaf(i), k)
        Next i
        Call modMath.Sort_Quick_A(xk, 1, n, sort_idx, 1)

        'Find best split point in current dimension
        If y_categorical = False Then
            Call Calc_Best_SSE(err_tmp, best_x_tmp, n1, n2, xk, y, leaf, sort_idx, y_sum, y2_sum, n)
        Else
            Call Calc_Best_GINI(err_tmp, best_x_tmp, n1, n2, xk, y, leaf, sort_idx, class_count, n_class, n)
        End If

        'See if this split point is best in all dimensions
        If err_tmp < err_best Then
            best_dim = k
            best_x = best_x_tmp
            err_best = err_tmp
            n1_best = n1
            n2_best = n2
        End If

    Next kk
    Erase xk, sort_idx
    
    If n1_best = 0 Or n2_best = 0 Then
        Call Set_As_Leaf(y, leaf, y_categorical, n_class)
        Exit Sub
    End If

    pSplit_Dim = best_dim
    pSplit_Val = best_x

    n1 = 0: ReDim leaf_1(1 To n1_best)
    n2 = 0: ReDim leaf_2(1 To n2_best)
    For i = 1 To n
        If x(leaf(i), best_dim) < best_x Then
            n1 = n1 + 1
            leaf_1(n1) = leaf(i)
        Else
            n2 = n2 + 1
            leaf_2(n2) = leaf(i)
        End If
    Next i

    Set pChild_L = New cCART
    Set pChild_R = New cCART
    If y_categorical = False Then
        Call pChild_L.prv_Split_Recursive(y, x, leaf_1, min_size, max_depth, cur_depth + 1, max_feature, False)
        Call pChild_R.prv_Split_Recursive(y, x, leaf_2, min_size, max_depth, cur_depth + 1, max_feature, False)
    Else
        Call pChild_L.prv_Split_Recursive(y, x, leaf_1, min_size, max_depth, cur_depth + 1, max_feature, True, n_class)
        Call pChild_R.prv_Split_Recursive(y, x, leaf_2, min_size, max_depth, cur_depth + 1, max_feature, True, n_class)
    End If

    Erase leaf, leaf_1, leaf_2
End Sub


Private Sub Set_As_Leaf(y As Variant, leaf() As Long, y_categorical As Boolean, Optional n_class As Long = 0)
Dim i As Long, k As Long, n As Long
Dim class_count() As Long
    n = UBound(leaf)
    pisLeaf = True
    If y_categorical = False Then
        pResponse = 0
        For i = 1 To n
            pResponse = pResponse + y(leaf(i))
        Next i
        pResponse = pResponse / n
    Else
        ReDim class_count(1 To n_class)
        For i = 1 To n
            k = y(leaf(i))
            class_count(k) = class_count(k) + 1
        Next i
        k = 0
        For i = 1 To UBound(class_count)
            If class_count(i) > k Then
                k = class_count(i)
                pResponse = i
            End If
        Next i
    End If
End Sub


Private Sub Calc_Best_SSE(err_split As Double, x_split As Double, n1_split As Long, n2_split As Long, _
    xk() As Double, y As Variant, leaf() As Long, sort_idx() As Long, _
    y_sum As Double, y2_sum As Double, n As Long)
Dim i As Long, j As Long, j_last As Long, n1 As Long, n2 As Long
Dim tmp_x As Double, tmp_y
Dim y_mean1 As Double, y_mean2 As Double
    y_mean1 = 0
    err_split = pINFINITY
    n1 = 0
    j_last = 1
    For i = 1 To n
        tmp_x = xk(i)
        For j = j_last To n
            If xk(j) < tmp_x Then
                n1 = n1 + 1
                y_mean1 = y_mean1 + y(leaf(sort_idx(j)))
            Else
                j_last = j
                Exit For
            End If
        Next j
        n2 = n - n1
        If n1 > 0 And n2 > 0 Then
            y_mean2 = y_sum - y_mean1
            tmp_y = y2_sum - (y_mean1 ^ 2) / n1 - (y_mean2 ^ 2) / n2
            If tmp_y < err_split Then
                err_split = tmp_y
                x_split = tmp_x
                n1_split = n1
                n2_split = n2
            End If
        End If
    Next i
End Sub


Private Sub Calc_Best_GINI(err_split As Double, x_split As Double, n1_split As Long, n2_split As Long, _
    xk() As Double, y As Variant, leaf() As Long, sort_idx() As Long, _
    class_count() As Long, n_class As Long, n As Long)
Dim i As Long, j As Long, j_last As Long, n1 As Long, n2 As Long, k As Long
Dim tmp_x As Double, tmp_y As Double
Dim class_count1() As Long, class_count2() As Long
    ReDim class_count1(1 To n_class)
    ReDim class_count2(1 To n_class)
    err_split = pINFINITY
    n1 = 0
    j_last = 1
    For i = 1 To n
        tmp_x = xk(i)
        For j = j_last To n
            If xk(j) < tmp_x Then
                n1 = n1 + 1
                k = y(leaf(sort_idx(j)))
                class_count1(k) = class_count1(k) + 1
            Else
                j_last = j
                Exit For
            End If
        Next j
        n2 = n - n1
        If n1 > 0 And n2 > 0 Then
            For j = 1 To n_class
                class_count2(j) = class_count(j) - class_count1(j)
            Next j
            tmp_y = n
            For j = 1 To n_class
                tmp_y = tmp_y - (class_count1(j) ^ 2) / n1 - (class_count2(j) ^ 2) / n2
            Next j
            If tmp_y < err_split Then
                err_split = tmp_y
                x_split = tmp_x
                n1_split = n1
                n2_split = n2
            End If
        End If
    Next i
End Sub


'Randomly pick k out of n features
Private Function Random_Feature(k As Long, n As Long) As Long()
Dim i As Long, j As Long
Dim y() As Long
    ReDim y(1 To k)
    For i = 1 To k
        y(i) = i
    Next i
    If k < n Then
        Randomize
        For i = k + 1 To n
            j = Int(Rnd() * i) + 1
            If j <= k Then y(j) = i
        Next i
    End If
    Random_Feature = y
End Function
