VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRForest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pn_tree As Long             'number of trees in forest
Private puseVote As Boolean         'use ensemble vote if output is categorical
Private pTrees As Collection        'collection of trees
Private pbest_tree As Long          'index of the tree that best fit training set

'Print a single tree from forest, if i is not specified then best-fitting tree is printed
Sub Print_Tree(vRng As Range, Optional factor_labels As Variant, Optional i As Long = -1)
Dim dT1 As cCART
    If i > 0 And i <= pn_tree Then
        Set dT1 = pTrees(i)
    ElseIf pbest_tree > 0 Then
        Set dT1 = pTrees(pbest_tree)
    Else
        Set dT1 = pTrees(1)
    End If
    If IsMissing(factor_labels) = False Then
        Call dT1.Chart_Tree(vRng, factor_labels)
    Else
        Call dT1.Chart_Tree(vRng)
    End If
End Sub

'Save model to excel range
Sub Print_Model(vRng As Range)
Dim i As Long
Dim dT1 As cCART
    i = 0
    vRng.Value = pTrees.Count
    vRng.Offset(0, 1).Value = puseVote
    vRng.Offset(0, 2).Value = pbest_tree
    For Each dT1 In pTrees
        i = i + 1
        Call dT1.Print_Model(vRng.Offset(1, (i - 1) * 4))
    Next dT1
End Sub

'Read model from excel range
Sub Read_Model(vRng As Range)
Dim i As Long
Dim dT1 As cCART
    pn_tree = vRng.Value
    puseVote = vRng.Offset(0, 1).Value
    pbest_tree = vRng.Offset(0, 2).Value
    Set pTrees = New Collection
    For i = 1 To pn_tree
        Set dT1 = New cCART
        Call dT1.Read_Model(vRng.Offset(1, (i - 1) * 4))
        pTrees.Add dT1
    Next i
End Sub

'Input : x(1:N,1:D), array storing N observations of D-dimensional predictors
'Output: y(1:N), vector storing N outputs
Sub Predict(x As Variant, y As Variant)
Dim i As Long, j As Long, k As Long, n As Long
Dim dT1 As cCART
Dim y1() As Double
Dim class_vote() As Long
    n = UBound(x, 1)
    ReDim y(1 To n)
    j = 0
    If puseVote = True Then ReDim class_vote(1 To n, 1 To 1)
    For Each dT1 In pTrees
        j = j + 1
        If j Mod 50 = 0 Then
            DoEvents
            Application.StatusBar = "cRForest: Predict: " & j & "/" & pn_tree
        End If
        Call dT1.Predict(x, y1)
        If puseVote = False Then
            For i = 1 To n
                y(i) = y(i) + y1(i) / pn_tree
            Next i
        Else
            For i = 1 To n
                k = y1(i)
                If k > UBound(class_vote, 2) Then ReDim Preserve class_vote(1 To n, 1 To k)
                class_vote(i, k) = class_vote(i, k) + 1
            Next i
        End If
    Next dT1
    
    If puseVote = True Then
        For i = 1 To n
            k = -1
            For j = 1 To UBound(class_vote, 2)
                If class_vote(i, j) > k Then
                    k = class_vote(i, j)
                    y(i) = j
                End If
            Next j
        Next i
    End If
    
    Application.StatusBar = False
End Sub


'Input: y(1:N), vector storing N observations target output
'       x(1:N,1:D), array storing N observations of D-dimensional predictors
'       n_tree, number of trees to include in forest
'       min_size, node smaller than min_size will not be split
'       max_depth, node deeper than max_depth will not be split
'       y_categorical, set to TRUE if y() is given as integer class labels
'       n_class, if y_categorical is set to TRUE, n_class must be provided as is the number of class labels
'       BagType, feature bagging type, number of featurs used in each split is given by
'           "NONE", no feature bagging
'           "LOG",  LOG(D,2)
'           "SQRT", SQRT(D)
'Output: oob_err, out-of-bagg error, given by mean-squared-error if y() is continuous,
'                 and 1-accuracy if y() is categorical
'        oob_output, Nx2 array comparing y vs y_oob  if y() is continuous
'                    and n_class x n_class confusion matrix if y() is categorical
Sub Fit(y As Variant, x As Variant, n_tree As Long, _
    Optional min_size As Long = 5, Optional max_depth As Long = 20, _
    Optional y_categorical As Boolean = False, Optional n_class As Long = 0, _
    Optional BagType As String = "LOG", _
    Optional oob_err As Variant, Optional oob_output As Variant)
Dim i As Long, j As Long, k As Long, n As Long, n_oob As Long, iterate As Long
Dim dT1 As cCART
Dim tmp_x As Double, best_err As Double
Dim y1 As Variant, x1 As Variant
Dim x_tmp As Variant, y_oob As Variant
Dim oob_list() As Long, oob_count() As Long, class_vote() As Long
Dim boolTestOOB As Boolean
    n = UBound(x, 1)
    pn_tree = n_tree
    puseVote = y_categorical
    Set pTrees = New Collection
    If IsMissing(oob_err) = False And IsMissing(oob_output) = False Then
        boolTestOOB = True
        ReDim oob_count(1 To n) 'number of times that i appears as out-of-bag
        ReDim y_oob(1 To n)
        ReDim class_vote(1 To n, 1 To n_class)
    End If
    For iterate = 1 To pn_tree
        If iterate Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = "cRForest: Fit: " & iterate & "/" & pn_tree
        End If
        
        If boolTestOOB = False Then
            Call Resampling(y, x, y1, x1)
        Else
            Call Resampling(y, x, y1, x1, x_tmp, oob_list)
        End If

        Set dT1 = New cCART      'Fit new tree with feature bagging
        Call dT1.Fit(y1, x1, min_size, max_depth, y_categorical, n_class, BagType)
        pTrees.Add dT1
        Erase x1, y1

        If boolTestOOB = True Then
            Call dT1.Predict(x_tmp, y1)
            For i = 1 To UBound(oob_list)
                j = oob_list(i)
                oob_count(j) = oob_count(j) + 1
            Next i
            If puseVote = False Then
                For i = 1 To UBound(oob_list)
                    j = oob_list(i)
                    y_oob(j) = y_oob(j) + y1(i)
                Next i
            Else
                For i = 1 To UBound(oob_list)
                    j = oob_list(i)
                    k = y1(i)
                    class_vote(j, k) = class_vote(j, k) + 1
                Next i
            End If
            Erase y1, oob_list
        End If
    Next iterate
    
    If boolTestOOB = True Then
        n_oob = 0
        For i = 1 To n
            If oob_count(i) > 0 Then
                n_oob = n_oob + 1
                If puseVote = False Then
                    y_oob(i) = y_oob(i) / oob_count(i)
                    oob_err = oob_err + (y_oob(i) - y(i)) ^ 2
                Else
                    k = -1
                    For j = 1 To UBound(class_vote, 2)
                        If class_vote(i, j) > k Then
                            k = class_vote(i, j)
                            y_oob(i) = j
                        End If
                    Next j
                    If y_oob(i) <> y(i) Then oob_err = oob_err + 1
                End If
            End If
        Next i
        oob_err = oob_err / n_oob
        
        j = 0
        ReDim oob_output(1 To n_oob, 1 To 2)
        For i = 1 To n
            If oob_count(i) > 0 Then
                j = j + 1
                oob_output(j, 1) = y(i)
                oob_output(j, 2) = y_oob(i)
            End If
        Next i
    End If
    Erase x1, y1
    
    'Identify tree that best fits training data
    pbest_tree = -1
    best_err = Exp(70)
    For i = 1 To pTrees.Count
        If i Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = "cRForest: Fit: Find best fitting tree" & iterate & "/" & pn_tree
        End If
        Set dT1 = pTrees(i)
        Call dT1.Predict(x, y1)
        tmp_x = Calc_Err(y1, y, y_categorical)
        If tmp_x < best_err Then
            best_err = tmp_x
            pbest_tree = i
        End If
    Next i
    
    Application.StatusBar = False
End Sub


Function Calc_Err(y_out As Variant, y_tgt As Variant, Optional y_categorical As Boolean = False)
Dim i As Long, n As Long
Dim tmp_x As Double
    tmp_x = 0
    n = UBound(y_out, 1)
    If y_categorical = False Then
        For i = 1 To n
            tmp_x = tmp_x + (y_out(i) - y_tgt(i)) ^ 2
        Next i
    Else
        For i = 1 To n
            If y_out(i) <> y_tgt(i) Then tmp_x = tmp_x + 1
        Next i
    End If
    Calc_Err = tmp_x / n
End Function


Private Sub Resampling(y As Variant, x As Variant, y1 As Variant, x1 As Variant, _
    Optional x_oob As Variant, Optional oob_list As Variant)
Dim i As Long, j As Long, n As Long, n_dimension As Long, ii As Long, n_oob As Long
Dim isPick() As Long
    Randomize
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim y1(1 To n)
    ReDim x1(1 To n, 1 To n_dimension)
    ReDim isPick(1 To n)
    For ii = 1 To n
        i = Int(Rnd() * n) + 1
        y1(ii) = y(i)
        For j = 1 To n_dimension
            x1(ii, j) = x(i, j)
        Next j
        isPick(i) = isPick(i) + 1
    Next ii
    'Save out-of-bag observations
    If IsMissing(x_oob) = False Then
        n_oob = 0
        ReDim oob_list(1 To n)
        For i = 1 To n
            If isPick(i) = 0 Then
                n_oob = n_oob + 1
                oob_list(n_oob) = i
            End If
        Next i
        ReDim Preserve oob_list(1 To n_oob)
        ReDim x_oob(1 To n_oob, 1 To n_dimension)
        ii = 0
        For i = 1 To n_oob
            ii = oob_list(i)
            For j = 1 To n_dimension
                x_oob(i, j) = x(ii, j)
            Next j
        Next i
    End If
End Sub
