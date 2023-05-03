VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================
'A single layer RNN in either GRU or LSTM setting
'The hidden states output is connected to a single layer
'feed forward network
'training data needs to be fed in a variant array
'containing separate 2D arrays
'=============================================

Private pUnitType As String
Private pstrMode As String
Private pOutputLayer As String
Private pn_input As Long, pn_hidden As Long
Private pn_output As Long
Private pn_row As Long
Private wz() As Double, uz() As Double, bz() As Double
Private wf() As Double, uf() As Double, bf() As Double
Private ws() As Double, us() As Double, bs() As Double
Private wo() As Double, uo() As Double, bo() As Double
Private vOut() As Double, bout() As Double

Private wz_tmp() As Double, uz_tmp() As Double, bz_tmp() As Double
Private wf_tmp() As Double, uf_tmp() As Double, bf_tmp() As Double
Private ws_tmp() As Double, us_tmp() As Double, bs_tmp() As Double
Private wo_tmp() As Double, uo_tmp() As Double, bo_tmp() As Double
Private Vout_tmp() As Double, bout_tmp() As Double

Private dwf() As Double, duf() As Double, dbf() As Double
Private dwz() As Double, duz() As Double, dbz() As Double
Private dws() As Double, dus() As Double, dbs() As Double
Private dwo() As Double, duo() As Double, dbo() As Double
Private dVout() As Double, dbout() As Double

Private pADAM_count As Long
Private dwf1() As Double, duf1() As Double, dbf1() As Double
Private dwz1() As Double, duz1() As Double, dbz1() As Double
Private dwo1() As Double, duo1() As Double, dbo1() As Double
Private dws1() As Double, dus1() As Double, dbs1() As Double
Private dVout1() As Double, dbout1() As Double
Private dwf2() As Double, duf2() As Double, dbf2() As Double
Private dwz2() As Double, duz2() As Double, dbz2() As Double
Private dwo2() As Double, duo2() As Double, dbo2() As Double
Private dws2() As Double, dus2() As Double, dbs2() As Double
Private dVout2() As Double, dbout2() As Double

Private xt() As Variant, zt() As Variant, ft() As Variant, st() As Variant
Private ot() As Variant, ct() As Variant, tanhct() As Variant, ht() As Variant

Property Get cellstate(Optional batchIdx As Long = -1) As Variant
Dim i As Long
    If batch_idx = -1 Then
        i = UBound(ct, 1)
        cellstate = ct(i)
    Else
        cellstate = ct(batchIdx)
    End If
End Property

Property Get hiddenstate(Optional batchIdx As Long = -1) As Variant
Dim i As Long
    If batch_idx = -1 Then
        i = UBound(ht, 1)
        hiddenstate = ht(i)
    Else
        hiddenstate = ht(batchIdx)
    End If
End Property


Property Get n_input() As Long
    n_input = pn_input
End Property

Property Get n_hidden() As Long
    n_hidden = pn_hidden
End Property

Property Get n_output() As Long
    n_output = pn_output
End Property

Property Get n_row() As Long
    n_row = pn_row
End Property

Property Get OutputLayer() As String
    OutputLayer = pOutputLayer
End Property

Property Get UnitType() As String
    UnitType = pUnitType
End Property

Property Get Wgts() As Variant
Dim v As Variant
    If pUnitType = "LSTM" Then
        ReDim v(1 To 14)
        v(1) = wf
        v(2) = uf
        v(3) = bf
        v(4) = wz
        v(5) = uz
        v(6) = bz
        v(7) = ws
        v(8) = us
        v(9) = bs
        v(10) = wo
        v(11) = uo
        v(12) = bo
        v(13) = vOut
        v(14) = bout
    ElseIf pUnitType = "GRU" Then
        ReDim v(1 To 11)
        v(1) = wf
        v(2) = uf
        v(3) = bf
        v(4) = wz
        v(5) = uz
        v(6) = bz
        v(7) = ws
        v(8) = us
        v(9) = bs
        v(10) = vOut
        v(11) = bout
    End If
    Wgts = v
End Property

Sub Clone(cNN As cRNN)
Dim v As Variant
    With cNN
        pUnitType = .UnitType
        pn_input = .n_input
        pn_hidden = .n_hidden
        pn_output = .n_output
        pOutputLayer = .OutputLayer
        v = .Wgts
    End With
    wf = v(1)
    uf = v(2)
    bf = v(3)
    wz = v(4)
    uz = v(5)
    bz = v(6)
    ws = v(7)
    us = v(8)
    bs = v(9)
    If pUnitType = "LSTM" Then
        wo = v(10)
        uo = v(11)
        bo = v(12)
        vOut = v(13)
        bout = v(14)
    ElseIf pUnitType = "GRU" Then
        v(10) = vOut
        v(11) = bout
    End If
End Sub


Sub ADAM_Clear()
    pADAM_count = 0
    
    Erase dwf1, duf1, dbf1
    Erase dwz1, duz1, dbz1
    Erase dws1, dus1, dbs1
    Erase dVout1, dbout1
    
    Erase dwf2, duf2, dbf2
    Erase dwz2, duz2, dbz2
    Erase dws2, dus2, dbs2
    Erase dVout2, dbout2
    
    If pUnitType = "LSTM" Then
        Erase dwo1, duo1, dbo1
        Erase dwo2, duo2, dbo2
    End If
End Sub

Private Sub ADAM_Init()
Dim i As Long, j As Long

    pADAM_count = 0
    
    ReDim dwf1(1 To pn_hidden, 1 To pn_input)
    ReDim duf1(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbf1(1 To pn_hidden)
    ReDim dwz1(1 To pn_hidden, 1 To pn_input)
    ReDim duz1(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbz1(1 To pn_hidden)
    ReDim dws1(1 To pn_hidden, 1 To pn_input)
    ReDim dus1(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbs1(1 To pn_hidden)
    ReDim dVout1(1 To pn_output, 1 To pn_hidden)
    ReDim dbout1(1 To pn_output)
    
    ReDim dwf2(1 To pn_hidden, 1 To pn_input)
    ReDim duf2(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbf2(1 To pn_hidden)
    ReDim dwz2(1 To pn_hidden, 1 To pn_input)
    ReDim duz2(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbz2(1 To pn_hidden)
    ReDim dws2(1 To pn_hidden, 1 To pn_input)
    ReDim dus2(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbs2(1 To pn_hidden)
    ReDim dVout2(1 To pn_output, 1 To pn_hidden)
    ReDim dbout2(1 To pn_output)
    
    If pUnitType = "LSTM" Then
        ReDim dwo1(1 To pn_hidden, 1 To pn_input)
        ReDim duo1(1 To pn_hidden, 1 To pn_hidden)
        ReDim dbo1(1 To pn_hidden)
        
        ReDim dwo2(1 To pn_hidden, 1 To pn_input)
        ReDim duo2(1 To pn_hidden, 1 To pn_hidden)
        ReDim dbo2(1 To pn_hidden)
    End If
End Sub


Private Sub ADAM_CalcMoment()
Dim i As Long, j As Long
    
    If pADAM_count = 0 Then Call ADAM_Init
    
    pADAM_count = pADAM_count + 1
    
    For i = 1 To pn_hidden
        dbf1(i) = 0.9 * dbf1(i) + 0.1 * dbf(i)
        dbz1(i) = 0.9 * dbz1(i) + 0.1 * dbz(i)
        dbs1(i) = 0.9 * dbs1(i) + 0.1 * dbs(i)
        dbf2(i) = 0.999 * dbf2(i) + 0.001 * dbf(i) ^ 2
        dbz2(i) = 0.999 * dbz2(i) + 0.001 * dbz(i) ^ 2
        dbs2(i) = 0.999 * dbs2(i) + 0.001 * dbs(i) ^ 2
        For j = 1 To pn_input
            dwf1(i, j) = 0.9 * dwf1(i, j) + 0.1 * dwf(i, j)
            dwz1(i, j) = 0.9 * dwz1(i, j) + 0.1 * dwz(i, j)
            dws1(i, j) = 0.9 * dws1(i, j) + 0.1 * dws(i, j)
            dwf2(i, j) = 0.999 * dwf2(i, j) + 0.001 * dwf(i, j) ^ 2
            dwz2(i, j) = 0.999 * dwz2(i, j) + 0.001 * dwz(i, j) ^ 2
            dws2(i, j) = 0.999 * dws2(i, j) + 0.001 * dws(i, j) ^ 2
        Next j
        For j = 1 To pn_hidden
            duf1(i, j) = 0.9 * duf1(i, j) + 0.1 * duf(i, j)
            duz1(i, j) = 0.9 * duz1(i, j) + 0.1 * duz(i, j)
            dus1(i, j) = 0.9 * dus1(i, j) + 0.1 * dus(i, j)
            duf2(i, j) = 0.999 * duf2(i, j) + 0.001 * duf(i, j) ^ 2
            duz2(i, j) = 0.999 * duz2(i, j) + 0.001 * duz(i, j) ^ 2
            dus2(i, j) = 0.999 * dus2(i, j) + 0.001 * dus(i, j) ^ 2
        Next j
    Next i
    
    If pUnitType = "LSTM" Then
        For i = 1 To pn_hidden
            dbo1(i) = 0.9 * dbo1(i) + 0.1 * dbo(i)
            dbo2(i) = 0.999 * dbo2(i) + 0.001 * dbo(i) ^ 2
            For j = 1 To pn_input
                dwo1(i, j) = 0.9 * dwo1(i, j) + 0.1 * dwo(i, j)
                dwo2(i, j) = 0.999 * dwo2(i, j) + 0.001 * dwo(i, j) ^ 2
            Next j
            For j = 1 To pn_hidden
                duo1(i, j) = 0.9 * duo1(i, j) + 0.1 * duo(i, j)
                duo2(i, j) = 0.999 * duo2(i, j) + 0.001 * duo(i, j) ^ 2
            Next j
        Next i
    End If
    
    For i = 1 To pn_output
        dbout1(i) = 0.9 * dbout1(i) + 0.1 * dbout(i)
        dbout2(i) = 0.999 * dbout2(i) + 0.001 * dbout(i) ^ 2
        For j = 1 To pn_hidden
            dVout1(i, j) = 0.9 * dVout1(i, j) + 0.1 * dVout(i, j)
            dVout2(i, j) = 0.999 * dVout2(i, j) + 0.001 * dVout(i, j) ^ 2
        Next j
    Next i
End Sub


'Initialize RNN network with random weights
'UniType        "LSTM" or "GRU"
'n_input        dimension of input data
'n_hidden       number of hidden units
'n_output       dimension of output data
'strOutputLayer activation functions of the feed forward layer connected to the output,
'               this layer will be skipped when set to ""
'strMode        "S2S", "S2V", or "V2S"
Sub Init(UnitType As String, n_input As Long, n_hidden As Long, n_output As Long, _
            Optional strOutputLayer As String = "SIGMOID", _
            Optional strMode As String = "S2S")
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double

    VBA.Randomize
    
    pUnitType = UCase(Trim(UnitType))
    pn_input = n_input
    pn_hidden = n_hidden
    pn_output = n_output
    pOutputLayer = UCase(strOutputLayer)
    pstrMode = UCase(strMode)
    
    ReDim wf(1 To n_hidden, 1 To n_input)
    ReDim uf(1 To n_hidden, 1 To n_hidden)
    ReDim bf(1 To n_hidden)
    ReDim wz(1 To n_hidden, 1 To n_input)
    ReDim uz(1 To n_hidden, 1 To n_hidden)
    ReDim bz(1 To n_hidden)
    ReDim ws(1 To n_hidden, 1 To n_input)
    ReDim us(1 To n_hidden, 1 To n_hidden)
    ReDim bs(1 To n_hidden)
    ReDim vOut(1 To n_output, 1 To n_hidden)
    ReDim bout(1 To n_output)
    
    tmp_x = Sqr(2 / n_hidden)
    For i = 1 To n_hidden
        bf(i) = 0
        bz(i) = 0
        bs(i) = 0
        For j = 1 To n_input
            wf(i, j) = (-0.5 + Rnd()) * tmp_x
            wz(i, j) = (-0.5 + Rnd()) * tmp_x
            ws(i, j) = (-0.5 + Rnd()) * tmp_x
        Next j
        For j = 1 To n_hidden
            uf(i, j) = (-0.5 + Rnd()) * tmp_x
            uz(i, j) = (-0.5 + Rnd()) * tmp_x
            us(i, j) = (-0.5 + Rnd()) * tmp_x
        Next j
    Next i
    
    
    If pUnitType = "LSTM" Then
        ReDim wo(1 To n_hidden, 1 To n_input)
        ReDim uo(1 To n_hidden, 1 To n_hidden)
        ReDim bo(1 To n_hidden)
        For i = 1 To n_hidden
            bo(i) = 0
            For j = 1 To n_input
                wo(i, j) = (-0.5 + Rnd()) * tmp_x
            Next j
            For j = 1 To n_hidden
                uo(i, j) = (-0.5 + Rnd()) * tmp_x
            Next j
        Next i
    End If
    
    For i = 1 To n_output
        bout(i) = 0
        For j = 1 To n_hidden
            vOut(i, j) = (-0.5 + Rnd()) * tmp_x
        Next j
    Next i
End Sub


Sub PrintNetwork(mysht As Worksheet, Optional start_row As Long = 1)
Dim i As Long, j As Long, m As Long
With mysht
    m = start_row
    .Cells(m, 1).Value = pUnitType
    .Cells(m + 1, 1).Value = pOutputLayer
    .Cells(m + 2, 1).Value = pstrMode
    .Cells(m + 3, 1).Value = pn_input
    .Cells(m + 4, 1).Value = pn_hidden
    .Cells(m + 5, 1).Value = pn_output
    m = m + 5
    .Cells(m + 1, 1).Resize(pn_hidden, pn_input).Value = wf
    m = m + pn_hidden
    .Cells(m + 1, 1).Resize(pn_hidden, pn_hidden).Value = uf
    m = m + pn_hidden
    .Cells(m + 1, 1).Resize(1, pn_hidden).Value = bf
    
    m = m + 1
    .Cells(m + 1, 1).Resize(pn_hidden, pn_input).Value = wz
    m = m + pn_hidden
    .Cells(m + 1, 1).Resize(pn_hidden, pn_hidden).Value = uz
    m = m + pn_hidden
    .Cells(m + 1, 1).Resize(1, pn_hidden).Value = bz

    m = m + 1
    .Cells(m + 1, 1).Resize(pn_hidden, pn_input).Value = ws
    m = m + pn_hidden
    .Cells(m + 1, 1).Resize(pn_hidden, pn_hidden).Value = us
    m = m + pn_hidden
    .Cells(m + 1, 1).Resize(1, pn_hidden).Value = bs
    
    If pUnitType = "LSTM" Then
        m = m + 1
        .Cells(m + 1, 1).Resize(pn_hidden, pn_input).Value = wo
        m = m + pn_hidden
        .Cells(m + 1, 1).Resize(pn_hidden, pn_hidden).Value = uo
        m = m + pn_hidden
        .Cells(m + 1, 1).Resize(1, pn_hidden).Value = bo
    End If
    
    m = m + 1
    .Cells(m + 1, 1).Resize(pn_output, pn_hidden).Value = vOut
    m = m + pn_output
    .Cells(m + 1, 1).Resize(1, pn_output).Value = bout
    pn_row = m + 1
End With
End Sub


Sub ReadNetwork(mysht As Worksheet, Optional start_row As Long = 1)
Dim i As Long, j As Long, m As Long
With mysht
    m = start_row
    pUnitType = .Cells(m, 1).Value
    pOutputLayer = .Cells(m + 1, 1).Value
    pstrMode = .Cells(m + 2, 1).Value
    pn_input = .Cells(m + 3, 1).Value
    pn_hidden = .Cells(m + 4, 1).Value
    pn_output = .Cells(m + 5, 1).Value
    m = m + 5
    ReDim wf(1 To pn_hidden, 1 To pn_input)
    ReDim uf(1 To pn_hidden, 1 To pn_hidden)
    ReDim bf(1 To pn_hidden)
    ReDim wz(1 To pn_hidden, 1 To pn_input)
    ReDim uz(1 To pn_hidden, 1 To pn_hidden)
    ReDim bz(1 To pn_hidden)
    If pUnitType = "LSTM" Then
        ReDim wo(1 To pn_hidden, 1 To pn_input)
        ReDim uo(1 To pn_hidden, 1 To pn_hidden)
        ReDim bo(1 To pn_hidden)
    End If
    ReDim ws(1 To pn_hidden, 1 To pn_input)
    ReDim us(1 To pn_hidden, 1 To pn_hidden)
    ReDim bs(1 To pn_hidden)
    ReDim vOut(1 To pn_output, 1 To pn_hidden)
    ReDim bout(1 To pn_output)

    For j = 1 To pn_input
        For i = 1 To pn_hidden
            wf(i, j) = .Cells(m + i, j).Value
        Next i
    Next j
    m = m + pn_hidden
    For j = 1 To pn_hidden
        For i = 1 To pn_hidden
            uf(i, j) = .Cells(m + i, j).Value
        Next i
    Next j
    m = m + pn_hidden
    For i = 1 To pn_hidden
        bf(i) = .Cells(m + 1, i).Value
    Next i

    m = m + 1
    For j = 1 To pn_input
        For i = 1 To pn_hidden
            wz(i, j) = .Cells(m + i, j).Value
        Next i
    Next j
    m = m + pn_hidden
    For j = 1 To pn_hidden
        For i = 1 To pn_hidden
            uz(i, j) = .Cells(m + i, j).Value
        Next i
    Next j
    m = m + pn_hidden
    For i = 1 To pn_hidden
        bz(i) = .Cells(m + 1, i).Value
    Next i
    
    m = m + 1
    For j = 1 To pn_input
        For i = 1 To pn_hidden
            ws(i, j) = .Cells(m + i, j).Value
        Next i
    Next j
    m = m + pn_hidden
    For j = 1 To pn_hidden
        For i = 1 To pn_hidden
            us(i, j) = .Cells(m + i, j).Value
        Next i
    Next j
    m = m + pn_hidden
    For i = 1 To pn_hidden
        bs(i) = .Cells(m + 1, i).Value
    Next i
    
    If pUnitType = "LSTM" Then
        m = m + 1
        For j = 1 To pn_input
            For i = 1 To pn_hidden
                wo(i, j) = .Cells(m + i, j).Value
            Next i
        Next j
        m = m + pn_hidden
        For j = 1 To pn_hidden
            For i = 1 To pn_hidden
                uo(i, j) = .Cells(m + i, j).Value
            Next i
        Next j
        m = m + pn_hidden
        For i = 1 To pn_hidden
            bo(i) = .Cells(m + 1, i).Value
        Next i
    End If
    
    m = m + 1
    For j = 1 To pn_hidden
        For i = 1 To pn_output
            vOut(i, j) = .Cells(m + i, j).Value
        Next i
    Next j
    m = m + pn_output
    For i = 1 To pn_output
        bout(i) = .Cells(m + 1, i).Value
    Next i
End With
End Sub


Private Sub clear_hist()
    Erase xt, ft, zt, st, ht
    If pUnitType = "LSTM" Then
        Erase ot, ct, tanhct
    End If
End Sub

Private Sub cache_current_wgts()
Dim i As Long, j As Long

    ReDim wf_tmp(1 To pn_hidden, 1 To pn_input)
    ReDim uf_tmp(1 To pn_hidden, 1 To pn_hidden)
    ReDim bf_tmp(1 To pn_hidden)
    ReDim wz_tmp(1 To pn_hidden, 1 To pn_input)
    ReDim uz_tmp(1 To pn_hidden, 1 To pn_hidden)
    ReDim bz_tmp(1 To pn_hidden)
    ReDim ws_tmp(1 To pn_hidden, 1 To pn_input)
    ReDim us_tmp(1 To pn_hidden, 1 To pn_hidden)
    ReDim bs_tmp(1 To pn_hidden)
    ReDim Vout_tmp(1 To pn_output, 1 To pn_hidden)
    ReDim bout_tmp(1 To pn_output)

    For i = 1 To pn_hidden
        bf_tmp(i) = bf(i)
        bz_tmp(i) = bz(i)
        bs_tmp(i) = bs(i)
        For j = 1 To pn_input
            wf_tmp(i, j) = wf(i, j)
            wz_tmp(i, j) = wz(i, j)
            ws_tmp(i, j) = ws(i, j)
        Next j
        For j = 1 To pn_hidden
            uf_tmp(i, j) = uf(i, j)
            uz_tmp(i, j) = uz(i, j)
            us_tmp(i, j) = us(i, j)
        Next j
    Next i
    
    If pUnitType = "LSTM" Then
        ReDim wo_tmp(1 To pn_hidden, 1 To pn_input)
        ReDim uo_tmp(1 To pn_hidden, 1 To pn_hidden)
        ReDim bo_tmp(1 To pn_hidden)
        For i = 1 To pn_hidden
            bo_tmp(i) = bo(i)
            For j = 1 To pn_input
                wo_tmp(i, j) = wo(i, j)
            Next j
            For j = 1 To pn_hidden
                uo_tmp(i, j) = uo(i, j)
            Next j
        Next i
    End If
    
    For i = 1 To pn_output
        bout_tmp(i) = bout(i)
        For j = 1 To pn_hidden
            Vout_tmp(i, j) = vOut(i, j)
        Next j
    Next i
End Sub


Private Sub clear_cache_wgts()
    Erase wf_tmp, uf_tmp, bf_tmp
    Erase wz_tmp, uz_tmp, bz_tmp
    Erase ws_tmp, us_tmp, bs_tmp
    If pUnitType = "LSTM" Then Erase wo_tmp, uo_tmp, bo_tmp
    Erase Vout_tmp, bout_tmp
End Sub

Private Sub clear_wgt_chg()
    Erase dwf, duf, dbf
    Erase dwz, duz, dbz
    Erase dws, dus, dbs
    If pUnitType = "LSTM" Then Erase dwo, duo, dbo
    Erase dVout, dbout
End Sub

Private Sub init_wgt_chg()
    ReDim dwf(1 To pn_hidden, 1 To pn_input)
    ReDim duf(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbf(1 To pn_hidden)
    ReDim dwz(1 To pn_hidden, 1 To pn_input)
    ReDim duz(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbz(1 To pn_hidden)
    ReDim dws(1 To pn_hidden, 1 To pn_input)
    ReDim dus(1 To pn_hidden, 1 To pn_hidden)
    ReDim dbs(1 To pn_hidden)
    If pUnitType = "LSTM" Then
        ReDim dwo(1 To pn_hidden, 1 To pn_input)
        ReDim duo(1 To pn_hidden, 1 To pn_hidden)
        ReDim dbo(1 To pn_hidden)
    End If
    ReDim dVout(1 To pn_output, 1 To pn_hidden)
    ReDim dbout(1 To pn_output)
End Sub

Private Sub restore_wgts_from_cache()
Dim i As Long, j As Long

    For i = 1 To pn_hidden
        bf(i) = bf_tmp(i)
        bz(i) = bz_tmp(i)
        bs(i) = bs_tmp(i)
        For j = 1 To pn_input
            wf(i, j) = wf_tmp(i, j)
            wz(i, j) = wz_tmp(i, j)
            ws(i, j) = ws_tmp(i, j)
        Next j
        For j = 1 To pn_hidden
            uf(i, j) = uf_tmp(i, j)
            uz(i, j) = uz_tmp(i, j)
            us(i, j) = us_tmp(i, j)
        Next j
    Next i
    
    If pUnitType = "LSTM" Then
        For i = 1 To pn_hidden
            bo(i) = bo_tmp(i)
            For j = 1 To pn_input
                wo(i, j) = wo_tmp(i, j)
            Next j
            For j = 1 To pn_hidden
                uo(i, j) = uo_tmp(i, j)
            Next j
        Next i
    End If
    
    For i = 1 To pn_output
        bout(i) = bout_tmp(i)
        For j = 1 To pn_hidden
            vOut(i, j) = Vout_tmp(i, j)
        Next j
    Next i
End Sub


Private Sub update_wgts(learn_rate As Double, Optional useSpeedUp As String = "")
Dim i As Long, j As Long
Dim adam_discount1 As Double, adam_discount2 As Double

    If UCase(useSpeedUp) = "ADAM" Then
        
        Call ADAM_CalcMoment
    
        adam_discount1 = 1 / (1 - 0.9 ^ pADAM_count)
        adam_discount2 = 1 / (1 - 0.999 ^ pADAM_count)
        For i = 1 To pn_hidden
            bf(i) = bf(i) - learn_rate * dbf1(i) * adam_discount1 / (Sqr(dbf2(i) * adam_discount2) + 0.00000001)
            bz(i) = bz(i) - learn_rate * dbz1(i) * adam_discount1 / (Sqr(dbz2(i) * adam_discount2) + 0.00000001)
            bs(i) = bs(i) - learn_rate * dbs1(i) * adam_discount1 / (Sqr(dbs2(i) * adam_discount2) + 0.00000001)
        Next i
        
        For j = 1 To pn_input
            For i = 1 To pn_hidden
                wf(i, j) = wf(i, j) - learn_rate * dwf1(i, j) * adam_discount1 / (Sqr(dwf2(i, j) * adam_discount2) + 0.00000001)
                wz(i, j) = wz(i, j) - learn_rate * dwz1(i, j) * adam_discount1 / (Sqr(dwz2(i, j) * adam_discount2) + 0.00000001)
                ws(i, j) = ws(i, j) - learn_rate * dws1(i, j) * adam_discount1 / (Sqr(dws2(i, j) * adam_discount2) + 0.00000001)
            Next i
        Next j
        
        For j = 1 To pn_hidden
            For i = 1 To pn_hidden
                uf(i, j) = uf(i, j) - learn_rate * duf1(i, j) * adam_discount1 / (Sqr(duf2(i, j) * adam_discount2) + 0.00000001)
                uz(i, j) = uz(i, j) - learn_rate * duz1(i, j) * adam_discount1 / (Sqr(duz2(i, j) * adam_discount2) + 0.00000001)
                us(i, j) = us(i, j) - learn_rate * dus1(i, j) * adam_discount1 / (Sqr(dus2(i, j) * adam_discount2) + 0.00000001)
            Next i
        Next j
        
        If pUnitType = "LSTM" Then
            For i = 1 To pn_hidden
                bo(i) = bo(i) - learn_rate * dbo1(i) * adam_discount1 / (Sqr(dbo2(i) * adam_discount2) + 0.00000001)
            Next i
            
            For j = 1 To pn_input
                For i = 1 To pn_hidden
                    wo(i, j) = wo(i, j) - learn_rate * dwo1(i, j) * adam_discount1 / (Sqr(dwo2(i, j) * adam_discount2) + 0.00000001)
                Next i
            Next j
            
            For j = 1 To pn_hidden
                For i = 1 To pn_hidden
                    uo(i, j) = uo(i, j) - learn_rate * duo1(i, j) * adam_discount1 / (Sqr(duo2(i, j) * adam_discount2) + 0.00000001)
                Next i
            Next j
        End If
        
        
        For i = 1 To pn_output
            bout(i) = bout(i) - learn_rate * dbout1(i) * adam_discount1 / (Sqr(dbout2(i) * adam_discount2) + 0.00000001)
        Next i
    
        For j = 1 To pn_hidden
            For i = 1 To pn_output
                vOut(i, j) = vOut(i, j) - learn_rate * dVout1(i, j) * adam_discount1 / (Sqr(dVout2(i, j) * adam_discount2) + 0.00000001)
            Next i
        Next j
        
    Else

        For i = 1 To pn_hidden
            bf(i) = bf(i) - learn_rate * dbf(i)
            bz(i) = bz(i) - learn_rate * dbz(i)
            bs(i) = bs(i) - learn_rate * dbs(i)
        Next i
        
        For j = 1 To pn_input
            For i = 1 To pn_hidden
                wf(i, j) = wf(i, j) - learn_rate * dwf(i, j)
                wz(i, j) = wz(i, j) - learn_rate * dwz(i, j)
                ws(i, j) = ws(i, j) - learn_rate * dws(i, j)
            Next i
        Next j
        
        For j = 1 To pn_hidden
            For i = 1 To pn_hidden
                uf(i, j) = uf(i, j) - learn_rate * duf(i, j)
                uz(i, j) = uz(i, j) - learn_rate * duz(i, j)
                us(i, j) = us(i, j) - learn_rate * dus(i, j)
            Next i
        Next j
        
        If pUnitType = "LSTM" Then
            For i = 1 To pn_hidden
                bo(i) = bo(i) - learn_rate * dbo(i)
            Next i
            
            For j = 1 To pn_input
                For i = 1 To pn_hidden
                    wo(i, j) = wo(i, j) - learn_rate * dwo(i, j)
                Next i
            Next j
            
            For j = 1 To pn_hidden
                For i = 1 To pn_hidden
                    uo(i, j) = uo(i, j) - learn_rate * duo(i, j)
                Next i
            Next j
        End If
        
        
        For i = 1 To pn_output
            bout(i) = bout(i) - learn_rate * dbout(i)
        Next i
    
        For j = 1 To pn_hidden
            For i = 1 To pn_output
                vOut(i, j) = vOut(i, j) - learn_rate * dVout(i, j)
            Next i
        Next j
    
    End If
End Sub


Private Function calc_lossBatch(vOut, vTgt, strCostType As String) As Double
Dim iterate As Long, i As Long, j As Long, k As Long, n_batch As Long
Dim tmp_x As Double

    n_batch = UBound(vOut, 1)
    tmp_x = 0
    For k = 1 To n_batch
        tmp_x = tmp_x + calc_loss(vOut(k), vTgt(k), strCostType)
    Next k
    calc_lossBatch = tmp_x / n_batch
    
End Function


Private Function calc_loss(y, y_tgt, strCostType As String)
Dim iterate As Long, i As Long, j As Long, k As Long, n_T As Long
Dim x_cost As Double, tmp_x As Double
Dim strtmp As String
    
    n_T = UBound(y, 2)
    If UBound(y_tgt, 1) <> pn_output Then
        Debug.Print "cRNN: calc_loss: input dimension not confromable"
        End
    End If
    strtmp = UCase(strCostType)

    If strtmp = "SSE" Then
    
        x_cost = 0
        For iterate = 1 To n_T
            For k = 1 To pn_output
                x_cost = x_cost + (y(k, iterate) - y_tgt(k, iterate)) ^ 2
            Next k
        Next iterate
        x_cost = x_cost / (2 * n_T)
        
    ElseIf strtmp = "CROSSENTROPY" Then
    
        x_cost = 0
        For iterate = 1 To n_T
            For k = 1 To pn_output
                x_cost = x_cost - y_tgt(k, iterate) * Log(y(k, iterate))
            Next k
        Next iterate
        x_cost = x_cost / (n_T)
        
    ElseIf strtmp = "LOGLOSS" Then
    
        x_cost = 0
        For iterate = 1 To n_T
            For k = 1 To pn_output
                x_cost = x_cost - y_tgt(k, iterate) * Log(y(k, iterate)) - (1 - y_tgt(k, iterate)) * Log(1 - y(k, iterate))
            Next k
        Next iterate
        x_cost = x_cost / (n_T)
    
    
    ElseIf strtmp = "MISCLASS" Then
    
        x_cost = 0
        For iterate = 1 To n_T
            i = 1
            tmp_x = y(1, iterate)
            For k = 2 To pn_output
                If y(k, iterate) > tmp_x Then
                    tmp_x = y(k, iterate)
                    i = k
                End If
            Next k
            If y_tgt(i, iterate) = 1 Then
                x_cost = x_cost + 1
                Exit For
            End If
        Next iterate
        x_cost = 1 - x_cost / (n_T)
    
    Else
    
        MsgBox "cRNN: calc_loss: " & strCostType & " is not a valid loss function."
    
    End If
    
    calc_loss = x_cost
    
End Function





'Generate future outputs given a time series of seed vector
'if RNN is trained in "S2S" mode, then output are recurrently generated
'by feeding the current output to next time step's input
'"S2V" mode is only applicable to univariate signal. The last length-n vector output
'is transposed to become the next n steps of input
'n      total length of output signals, including input signal
Function generate(y_seed() As Double, n As Long, Optional softmaxClean As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, iterate As Long
Dim n_seed As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_s As Double
Dim f() As Double, z() As Double, o() As Double, s() As Double
Dim c_prv() As Double, c() As Double, h() As Double, h_prv() As Double, y_tmp() As Double, y() As Double, y_last() As Double
    
    n_seed = UBound(y_seed, 2)
    
    If pstrMode = "S2S" Then
        
        y = FwdPass(y_seed, storeOutput:=True, softmaxClean:=softmaxClean)
        m = UBound(ht, 1)
        ReDim y_last(1 To pn_output, 1 To 1)
        For i = 1 To pn_output
            y_last(i, 1) = y(i, n_seed)
        Next i
        'use last output to recurrently generate next time step
        If pUnitType = "LSTM" Then
            y_tmp = FwdPass(y_last, storeOutput:=False, outputLen:=(n - n_seed), isRecurrent:=True, h0:=ht(m), c0:=ct(m), softmaxClean:=softmaxClean)
        ElseIf pUnitType = "GRU" Then
            y_tmp = FwdPass(y_last, storeOutput:=False, outputLen:=(n - n_seed), isRecurrent:=True, h0:=ht(m), softmaxClean:=softmaxClean)
        End If
        ReDim Preserve y(1 To pn_output, 1 To n)
        For i = n_seed + 1 To n
            For j = 1 To pn_output
                y(j, i) = y_tmp(j, i - n_seed)
            Next j
        Next i
        
    ElseIf pstrMode = "S2V" Then
    
        ReDim y(1 To 1, 1 To n)
        For i = 1 To n_seed
            y(1, i) = y_seed(1, i)
        Next i
        
        'Predict next pn_output time step from last vector output
        m = n_seed
        y_tmp = FwdPass(y_seed, storeOutput:=True)
        For i = 1 To pn_output
            m = m + 1
            If m <= n Then
                y(1, m) = y_tmp(i, 1)
            Else
                Exit For
            End If
        Next i
        
        Do While m < n
            k = UBound(ht, 1)
            h_prv = ht(k)
            If pUnitType = "LSTM" Then c_prv = ct(k)
            ReDim y_last(1 To 1, 1 To pn_output)
            For i = 1 To pn_output
                y_last(1, i) = y_tmp(i, 1)
            Next i
            If pUnitType = "LSTM" Then
                y_tmp = FwdPass(y_last, storeOutput:=True, isRecurrent:=False, h0:=h_prv, c0:=c_prv, softmaxClean:=softmaxClean)
            ElseIf pUnitType = "GRU" Then
                y_tmp = FwdPass(y_last, storeOutput:=True, isRecurrent:=False, h0:=h_prv, softmaxClean:=softmaxClean)
            End If
            For i = 1 To pn_output
                m = m + 1
                If m <= n Then
                    y(1, m) = y_tmp(i, 1)
                Else
                    Exit For
                End If
            Next i
        Loop

        
    End If
    
    Call clear_hist
    generate = y
End Function


Function FwdPassBatch(vIn() As Variant, Optional storeOutput As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_batch As Long
Dim x() As Double, y() As Double
Dim vOut() As Variant

    n_batch = UBound(vIn, 1)
    ReDim vOut(1 To n_batch)
    For k = 1 To n_batch
        x = vIn(k)
        y = FwdPass(x, storeOutput:=storeOutput)
        vOut(k) = y
    Next k
    
    FwdPassBatch = vOut
    
End Function


'Forward pass of a single time series of signal
'in "S2V" mode, hidden states are passed to output layer only at the final time step
'in "S2S" mode, output is generated at every time step
'isRecurrent    whether to use previous output as current input
'h0 and c0      hidden states and cell states at t=0 are zero unless otheriwse specified
'softmaxClean   to force softmax output to be 0 and 1
Function FwdPass(x As Variant, Optional storeOutput As Boolean = False, _
                Optional outputLen As Long = -1, _
                Optional isRecurrent As Boolean = False, Optional softmaxClean As Boolean = False, _
                Optional h0 As Variant, Optional c0 As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, n_T As Long, n_input As Long
Dim z() As Double, f() As Double, s() As Double, o() As Double, c() As Double, tanhc() As Double, h() As Double
Dim y() As Double, y_tmp() As Double
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_s As Double
Dim wfx As Variant, wzx As Variant, wsx As Variant, wox As Variant

    If outputLen > 0 Then
        n_T = outputLen
    Else
        n_T = UBound(x, 2)
    End If
    
    If UBound(x, 1) <> pn_input Then
        Debug.Print "cRNN: FwdPass: Input size not conformable"
        End
    End If
    If pstrMode = "S2V" Then
        ReDim y(1 To pn_output, 1 To 1)
    Else
        ReDim y(1 To pn_output, 1 To n_T)
    End If
    ReDim y_tmp(1 To pn_output)
    
    ReDim z(1 To pn_hidden, 1 To n_T)
    ReDim f(1 To pn_hidden, 1 To n_T)
    ReDim s(1 To pn_hidden, 1 To n_T)
    If pUnitType = "LSTM" Then
        ReDim o(1 To pn_hidden, 1 To n_T)
        ReDim c(1 To pn_hidden, 0 To n_T)
        ReDim tanhc(1 To pn_hidden, 1 To n_T)
    End If
    ReDim h(1 To pn_hidden, 0 To n_T)

    If Not IsMissing(h0) Then
        For i = 1 To pn_hidden
            h(i, 0) = h0(i, 1)
        Next i
    End If
    If Not IsMissing(c0) Then
        For i = 1 To pn_hidden
            c(i, 0) = c0(i, 1)
        Next i
    End If
    
    With Application.WorksheetFunction
        wfx = .MMult(wf, x)
        wzx = .MMult(wz, x)
        wsx = .MMult(ws, x)
        If pUnitType = "LSTM" Then
            wox = .MMult(wo, x)
        End If
    End With
    
    For iterate = 1 To n_T
        If pUnitType = "LSTM" Then
        
            For i = 1 To pn_hidden
                tmp_x = bf(i)
                tmp_y = bz(i)
                tmp_z = bo(i)
                tmp_s = bs(i)
                If Not (pstrMode = "V2S" And iterate > 1) Then
                    If isRecurrent And iterate > 1 Then
                        For j = 1 To pn_input
                            tmp_x = tmp_x + wf(i, j) * y(j, iterate - 1)
                            tmp_y = tmp_y + wz(i, j) * y(j, iterate - 1)
                            tmp_z = tmp_z + wo(i, j) * y(j, iterate - 1)
                            tmp_s = tmp_s + ws(i, j) * y(j, iterate - 1)
                        Next j
                    Else
                        tmp_x = tmp_x + wfx(i, iterate)
                        tmp_y = tmp_y + wzx(i, iterate)
                        tmp_z = tmp_z + wox(i, iterate)
                        tmp_s = tmp_s + wsx(i, iterate)
'                        For j = 1 To pn_input
'                            tmp_x = tmp_x + wf(i, j) * x(j, iterate)
'                            tmp_y = tmp_y + wz(i, j) * x(j, iterate)
'                            tmp_z = tmp_z + wo(i, j) * x(j, iterate)
'                            tmp_s = tmp_s + ws(i, j) * x(j, iterate)
'                        Next j
                    End If
                End If
                For j = 1 To pn_hidden
                    tmp_x = tmp_x + uf(i, j) * h(j, iterate - 1)
                    tmp_y = tmp_y + uz(i, j) * h(j, iterate - 1)
                    tmp_z = tmp_z + uo(i, j) * h(j, iterate - 1)
                    tmp_s = tmp_s + us(i, j) * h(j, iterate - 1)
                Next j
                f(i, iterate) = f_sigmoid(tmp_x)
                z(i, iterate) = f_sigmoid(tmp_y)
                o(i, iterate) = f_sigmoid(tmp_z)
                s(i, iterate) = f_tanh(tmp_s)
            Next i
        
            For i = 1 To pn_hidden
                c(i, iterate) = f(i, iterate) * c(i, iterate - 1) + z(i, iterate) * s(i, iterate)
                tanhc(i, iterate) = f_tanh(c(i, iterate))
                h(i, iterate) = o(i, iterate) * tanhc(i, iterate)
            Next i
            
        ElseIf pUnitType = "GRU" Then
        
            For i = 1 To pn_hidden
                tmp_x = bf(i)
                tmp_y = bz(i)
                tmp_s = bs(i)
                If Not (pstrMode = "V2S" And iterate > 1) Then
                    If isRecurrent And iterate > 1 Then
                        For j = 1 To pn_input
                            tmp_x = tmp_x + wf(i, j) * y(j, iterate - 1)
                            tmp_y = tmp_y + wz(i, j) * y(j, iterate - 1)
                            tmp_s = tmp_s + ws(i, j) * y(j, iterate - 1)
                        Next j
                    Else
                        tmp_x = tmp_x + wfx(i, iterate)
                        tmp_y = tmp_y + wzx(i, iterate)
                        tmp_s = tmp_s + wsx(i, iterate)
'                        For j = 1 To pn_input
'                            tmp_x = tmp_x + wf(i, j) * x(j, iterate)
'                            tmp_y = tmp_y + wz(i, j) * x(j, iterate)
'                            tmp_s = tmp_s + ws(i, j) * x(j, iterate)
'                        Next j
                    End If
                End If
                For j = 1 To pn_hidden
                    tmp_x = tmp_x + uf(i, j) * h(j, iterate - 1)
                    tmp_y = tmp_y + uz(i, j) * h(j, iterate - 1)
                Next j
                f(i, iterate) = f_sigmoid(tmp_x)
                z(i, iterate) = f_sigmoid(tmp_y)
                
                For j = 1 To pn_hidden
                    tmp_s = tmp_s + us(i, j) * h(j, iterate - 1) * f(i, iterate)
                Next j
                s(i, iterate) = f_tanh(tmp_s)
                
            Next i
        
            For i = 1 To pn_hidden
                h(i, iterate) = (1 - z(i, iterate)) * h(i, iterate - 1) + z(i, iterate) * s(i, iterate)
            Next i
            
        End If
        
        If pOutputLayer <> "" Then
            If pstrMode = "S2S" Or pstrMode = "V2S" Or (pstrMode = "S2V" And iterate = n_T) Then
                For i = 1 To pn_output
                    tmp_x = bout(i)
                    For j = 1 To pn_hidden
                        tmp_x = tmp_x + vOut(i, j) * h(j, iterate)
                    Next j
                    y_tmp(i) = tmp_x
                Next i
                y_tmp = OutputLayerFwd(y_tmp)
                If pOutputLayer = "SOFTMAX" And softmaxClean = True Then
                    tmp_x = y_tmp(1): k = 1
                    For i = 2 To pn_output
                        If y_tmp(i) > tmp_x Then
                            tmp_x = y_tmp(i): k = i
                        End If
                    Next i
                    ReDim y_tmp(1 To pn_output)
                    y_tmp(k) = 1
                End If
                If pstrMode = "S2S" Or pstrMode = "V2S" Then
                    For i = 1 To pn_output
                        y(i, iterate) = y_tmp(i)
                    Next i
                ElseIf (pstrMode = "S2V" And iterate = n_T) Then
                    For i = 1 To pn_output
                        y(i, 1) = y_tmp(i)
                    Next i
                End If
            End If
        End If
        
    Next iterate
    
    FwdPass = y

    'store intermediate outputs for backpropagation
    If storeOutput Then
        
        If ArrayIsEmpty(ht) Then
            m = 1
            ReDim xt(1 To 1)
            ReDim ft(1 To 1)
            ReDim zt(1 To 1)
            ReDim st(1 To 1)
            If pUnitType = "LSTM" Then
                ReDim ot(1 To 1)
                ReDim ct(1 To 1)
                ReDim tanhct(1 To 1)
            End If
            ReDim ht(1 To 1)
        Else
            m = UBound(ht, 1) + 1
            ReDim Preserve xt(1 To m)
            ReDim Preserve ft(1 To m)
            ReDim Preserve zt(1 To m)
            ReDim Preserve st(1 To m)
            If pUnitType = "LSTM" Then
                ReDim Preserve ot(1 To m)
                ReDim Preserve ct(1 To m)
                ReDim Preserve tanhct(1 To m)
            End If
            ReDim Preserve ht(1 To m)
        End If
        
        xt(m) = x
        ft(m) = f
        zt(m) = z
        st(m) = s
        If pUnitType = "LSTM" Then
            ot(m) = o
            ct(m) = c
            tanhct(m) = tanhc
        End If
        ht(m) = h
        
    End If
    
End Function

'Fitting input vIn() to generate vTgt()
'vIn() and vTgt() are variant arrays that contain multiple sequences
'each sequence itself is a double array
'sequences do not need to have the same lengths, but intput and output dimension should be consistent
Function Fit(vIn() As Variant, vTgt() As Variant, Optional learn_rate As Double = 0.001, _
            Optional n_epoch As Long = 200, Optional batch_size As Long = -1, _
            Optional strLossType As String = "SSE", _
            Optional err_tol As Double = 0.0001, _
            Optional err_tol_rel As Double = 0.0001, _
            Optional useSpeedUp As String = "", _
            Optional learnSchedule As String = "AGGRESSIVE", _
            Optional statusShown As Long = 5)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, i_epoch As Long, n_T As Long, n_batch As Long
Dim ii As Long, jj As Long, converge_count As Long
Dim tmp_x As Double, tmp_y As Double
Dim y() As Double, grad_out() As Double
Dim x_cost As Double, x_cost_prv As Double
Dim step_size As Double
Dim v_progress As Variant
Dim x() As Double, y_tgt() As Double
Dim vGradOut() As Variant, vOut() As Variant
Dim batchIdx() As Long
Dim vInSub() As Variant, vTgtSub() As Variant

    n_batch = UBound(vIn, 1)
    ReDim batchIdx(1 To n_batch)
    For i = 1 To n_batch
        batchIdx(i) = i
    Next i
    
    vOut = FwdPassBatch(vIn, storeOutput:=False)
    x_cost_prv = calc_lossBatch(vOut, vTgt, strLossType)
    
    converge_count = 0
    step_size = learn_rate
    ReDim v_progress(1 To 3, 1 To 1)
    
    For i_epoch = 1 To n_epoch
    
        DoEvents
        If (i_epoch - 1) Mod statusShown = 0 Then
            Application.StatusBar = "cRNN: Fit: Epoch " & i_epoch & "/" & n_epoch & "..."
        End If
        
        x_cost = 0
        Call cache_current_wgts
        
        If batch_size < 0 Or batch_size >= n_batch Then
            
            Call init_wgt_chg
            vOut = FwdPassBatch(vIn, storeOutput:=True)
            If pOutputLayer = "SOFTMAX" Then
                Call BackwardBatch(vOut, , vTgt)
            Else
                vGradOut = f_calcdEdyBatch(vOut, vTgt, strLossType)
                Call BackwardBatch(vOut, vGradOut)
            End If
            
            Call update_wgts(step_size, useSpeedUp:=useSpeedUp)
            Call clear_hist

        Else
        
            Call Shuffle(batchIdx)
            
            ii = 0: jj = 0
            Do While (jj + 1) <= n_batch

                'Extract a mini-batch
                ii = jj + 1
                jj = jj + batch_size
                If jj > n_batch Then jj = n_batch
                m = jj - ii + 1
                ReDim vInSub(1 To m)
                ReDim vTgtSub(1 To m)
                For i = 1 To m
                    vInSub(i) = vIn(batchIdx(ii + i - 1))
                    vTgtSub(i) = vTgt(batchIdx(ii + i - 1))
                Next i
                
                Call init_wgt_chg
                vOut = FwdPassBatch(vInSub, storeOutput:=True)
                If pOutputLayer = "SOFTMAX" Then
                    Call BackwardBatch(vOut, , vTgtSub)
                Else
                    vGradOut = f_calcdEdyBatch(vOut, vTgtSub, strLossType)
                    Call BackwardBatch(vOut, vGradOut)
                End If
                
                Call update_wgts(step_size, useSpeedUp:=useSpeedUp)
                Call clear_hist

            Loop

        End If

        vOut = FwdPassBatch(vIn, storeOutput:=False)
        x_cost = calc_lossBatch(vOut, vTgt, strLossType)

        ReDim Preserve v_progress(1 To 3, 1 To i_epoch)
        v_progress(1, i_epoch) = i_epoch
        v_progress(2, i_epoch) = x_cost
        v_progress(3, i_epoch) = step_size
        
        DoEvents
        If (i_epoch - 1) Mod statusShown = 0 Then
            Debug.Print "cRNN: Fit: epoch " & i_epoch & "; x_cost = " & Format(x_cost, "0.0000E+00") & "; step_size=" & Format(step_size, "0.0000E+00")
        End If
        
        If (x_cost <= x_cost_prv) And (Abs(x_cost_prv - x_cost) <= (err_tol_rel * Abs(x_cost_prv)) Or x_cost <= err_tol) Then
            converge_count = converge_count + 1
        Else
            converge_count = 0
        End If
        If converge_count >= 5 Then Exit For
        
        If UCase(learnSchedule) = "AGGRESSIVE" Then
        
            If (x_cost <= x_cost_prv) Then
                step_size = step_size * 1.05
                x_cost_prv = x_cost
            Else
                DoEvents
                Debug.Print "cRNN: Fit: cost increases, reducing learning rate to " & step_size * 0.1 & " (from " & step_size & "), discard current epoch"
                step_size = step_size * 0.1
                Call restore_wgts_from_cache
                If useSpeedUp = "ADAM" Then Call ADAM_Init
                If step_size < 0.00000001 Then Exit For
            End If
            
        ElseIf UCase(learnSchedule) = "DECAY" Then
        
            step_size = learn_rate * (1 - i_epoch / n_epoch)
            x_cost_prv = x_cost
            
        ElseIf UCase(learnSchedule) = "" Then
        
            step_size = step_size
            x_cost_prv = x_cost
            
        End If
        
    Next i_epoch
    
    Call clear_wgt_chg
    Call clear_cache_wgts
    If useSpeedUp = "ADAM" Then Call ADAM_Clear
    
    Application.StatusBar = False
    
    Fit = v_progress
    
End Function


Private Function f_LocalGrad(y As Variant, strType As String)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, iterate As Long
Dim y_grad() As Double
Dim tmp_x As Double
    
    If UCase(strType) = "SOFTMAX" Then
        Debug.Print "cRNN: f_LocalGrad: should not be necessary to calculate local grad for SOFTMAX."
        Exit Function
    End If
    
    n_T = UBound(y, 2)
    ReDim y_grad(1 To pn_output, 1 To n_T)
    
    Select Case UCase(strType)
    Case "LINEAR"
        For iterate = 1 To n_T
            For i = 1 To pn_output
                y_grad(i, iterate) = 1
            Next i
        Next iterate
    Case "SIGMOID"
        For iterate = 1 To n_T
            For i = 1 To pn_output
                y_grad(i, iterate) = y(i, iterate) * (1 - y(i, iterate))
            Next i
        Next iterate
    Case "TANH"
        For iterate = 1 To n_T
            For i = 1 To pn_output
                y_grad(i, iterate) = 1 - y(i, iterate) ^ 2
            Next i
        Next iterate
    End Select

    f_LocalGrad = y_grad
    
End Function


Private Function f_calcdEdyBatch(vOut, vTgt, strLossType)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_batch As Long
Dim dEdy() As Double
Dim vGradOut() As Variant

    n_batch = UBound(vOut, 1)
    ReDim vGradOut(1 To n_batch)
    For k = 1 To n_batch
        dEdy = f_calcdEdy(vOut(k), vTgt(k), strLossType)
        vGradOut(k) = dEdy
    Next k
    
    f_calcdEdyBatch = vGradOut
End Function


Private Function f_calcdEdy(y, y_tgt, strLossType)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, iterate As Long
Dim dEdy() As Double
    
    n_T = UBound(y, 2)
    ReDim dEdy(1 To pn_output, 1 To n_T)
    
    If UCase(pOutputLayer) <> "SOFTMAX" Then

        Select Case UCase(strLossType)
        Case "SSE"
            For iterate = 1 To n_T
                For i = 1 To pn_output
                    dEdy(i, iterate) = (y(i, iterate) - y_tgt(i, iterate))
                Next i
            Next iterate
        Case "CROSSENTROPY"
            For iterate = 1 To n_T
                For i = 1 To pn_output
                    dEdy(i, iterate) = -(y_tgt(i, iterate) / y(i, iterate))
                Next i
            Next iterate
        Case "LOGLOSS"
            For iterate = 1 To n_T
                For i = 1 To pn_output
                    dEdy(i, iterate) = -y_tgt(i, iterate) / y(i, iterate) - (1 - y_tgt(i, iterate)) / (1 - y(i, iterate))
                Next i
            Next iterate
        Case Else
            MsgBox "cRNN:Train_Step:" & strLossType & " is not valid."
        End Select
    End If
    
    f_calcdEdy = dEdy
    
End Function


Sub BackwardBatch(vOut() As Variant, Optional vGradOut As Variant, Optional vTgt As Variant, Optional isdEdx As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_batch As Long

    n_batch = UBound(vOut, 1)
    
    If Not IsMissing(vGradOut) Then
        For k = 1 To n_batch
            Call Backward(k, vOut(k), vGradOut(k), isdEdx:=isdEdx)
        Next k
    Else
        For k = 1 To n_batch
            Call Backward(k, vOut(k), , vTgt(k), isdEdx:=isdEdx)
        Next k
    End If

End Sub


Sub Backward(batchIdx As Long, y_out, Optional grad_out, Optional y_tgt As Variant, Optional isdEdx As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, ii As Long, jj As Long, n_T As Long
Dim n_len_in As Long, n_len_out As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_s As Double, tmp_c As Double
Dim dEdh() As Double
Dim dhdo() As Double, dhdf() As Double, dhdz() As Double, dhds() As Double
Dim xA() As Double, xA_prv() As Double
Dim grad_curr() As Double, grad_loc() As Double
Dim dhdh() As Double, grad_rs() As Double
Dim x() As Double, z() As Double, f() As Double, s() As Double
Dim o() As Double, c() As Double, tanhc() As Double
Dim h() As Double, y_tmp() As Double, y() As Double
    
    'extract batch from memory
    x = xt(batchIdx)
    z = zt(batchIdx)
    f = ft(batchIdx)
    s = st(batchIdx)
    If pUnitType = "LSTM" Then
        c = ct(batchIdx)
        o = ot(batchIdx)
        tanhc = tanhct(batchIdx)
    End If
    h = ht(batchIdx)
    
    n_T = UBound(h, 2)

    'Back Propagate Through Time
    ReDim dEdh(1 To pn_hidden)
    ReDim xA(1 To pn_hidden)
    ReDim xA_prv(1 To pn_hidden)
    
    n_len_in = UBound(x, 2)
    n_len_out = UBound(y_out, 2)
    ReDim grad_curr(1 To pn_output, 1 To n_len_out)
    If isdEdx Then
    
        For iterate = 1 To n_len_out
            For i = 1 To pn_output
                grad_curr(i, iterate) = grad_out(i, iterate)
            Next i
        Next iterate
        
    Else
  
        If pOutputLayer = "SOFTMAX" Then
        
            'For softmax directly calculate grad_nxt * grad_loc
            For iterate = 1 To n_len_out
                For i = 1 To pn_output
                    grad_curr(i, iterate) = y_out(i, iterate) - y_tgt(i, iterate)
                Next i
            Next iterate
            
        Else
        
            grad_loc = f_LocalGrad(y_out, pOutputLayer)
            For iterate = 1 To n_len_out
                For i = 1 To pn_output
                    grad_curr(i, iterate) = grad_out(i, iterate) * grad_loc(i, iterate)
                Next i
            Next iterate
            
        End If
        
    End If
    
    For iterate = n_T To 1 Step -1
        
        If (pstrMode = "S2S" Or pstrMode = "V2S" Or (pstrMode = "S2V" And iterate = n_T)) Then
            m = iterate - (n_T - n_len_out)
            For i = 1 To pn_output
                tmp_x = grad_curr(i, m)
                dbout(i) = dbout(i) + tmp_x
                For j = 1 To pn_hidden
                    dVout(i, j) = dVout(i, j) + tmp_x * h(j, iterate)
                Next j
            Next i
            
            For i = 1 To pn_hidden
                tmp_x = 0
                For k = 1 To pn_output
                    tmp_x = tmp_x + grad_curr(k, m) * vOut(k, i)
                Next k
                dEdh(i) = tmp_x
            Next i
 
        Else
        
            For i = 1 To pn_hidden
                dEdh(i) = 0
            Next i
            
        End If
        
        If pUnitType = "LSTM" Then
            
            ReDim dhdz(1 To pn_hidden)
            ReDim dhdf(1 To pn_hidden)
            ReDim dhds(1 To pn_hidden)
            ReDim dhdo(1 To pn_hidden)
            ReDim dhdh(1 To pn_hidden, 1 To pn_hidden)
            For i = 1 To pn_hidden
                tmp_x = o(i, iterate) * (1 - tanhc(i, iterate) ^ 2)
                dhdz(i) = tmp_x * s(i, iterate) * z(i, iterate) * (1 - z(i, iterate))
                dhdf(i) = tmp_x * c(i, iterate - 1) * f(i, iterate) * (1 - f(i, iterate))
                dhds(i) = tmp_x * z(i, iterate) * (1 - s(i, iterate) ^ 2)
                dhdo(i) = h(i, iterate) * (1 - o(i, iterate))
                For j = 1 To pn_hidden
                    dhdh(i, j) = dhdo(i) * uo(i, j) + dhds(i) * us(i, j) + dhdz(i) * uz(i, j) + dhdf(i) * uf(i, j)
                Next j
            Next i
    
            If iterate < n_T Then
                For k = 1 To pn_hidden
                    tmp_x = 0
                    For i = 1 To pn_hidden
                        tmp_x = tmp_x + xA_prv(i) * dhdh(i, k)
                    Next i
                    xA(k) = dEdh(k) + tmp_x
                Next k
            ElseIf iterate = n_T Then
                For k = 1 To pn_hidden
                    xA(k) = dEdh(k)
                Next k
            End If
            
            For i = 1 To pn_hidden
                
                tmp_x = xA(i) * dhdf(i)
                tmp_z = xA(i) * dhdz(i)
                tmp_y = xA(i) * dhdo(i)
                tmp_s = xA(i) * dhds(i)
                
                dbf(i) = dbf(i) + tmp_x
                dbz(i) = dbz(i) + tmp_z
                dbo(i) = dbo(i) + tmp_y
                dbs(i) = dbs(i) + tmp_s
                
                If Not (pstrMode = "V2S" And iterate > 1) Then
                    For j = 1 To pn_input
                        dwf(i, j) = dwf(i, j) + tmp_x * x(j, iterate)
                        dwz(i, j) = dwz(i, j) + tmp_z * x(j, iterate)
                        dwo(i, j) = dwo(i, j) + tmp_y * x(j, iterate)
                        dws(i, j) = dws(i, j) + tmp_s * x(j, iterate)
                    Next j
                End If
                
                For j = 1 To pn_hidden
                    duf(i, j) = duf(i, j) + tmp_x * h(j, iterate - 1)
                    duz(i, j) = duz(i, j) + tmp_z * h(j, iterate - 1)
                    duo(i, j) = duo(i, j) + tmp_y * h(j, iterate - 1)
                    dus(i, j) = dus(i, j) + tmp_s * h(j, iterate - 1)
                Next j
            
            Next i
        
        ElseIf pUnitType = "GRU" Then
        
            ReDim grad_rs(1 To pn_hidden)
            For i = 1 To pn_hidden
                grad_rs(i) = h(i, iterate - 1) * f(i, iterate) * (1 - f(i, iterate))
            Next i
            
            ReDim dhdz(1 To pn_hidden)
            ReDim dhds(1 To pn_hidden)
            ReDim dhdf(1 To pn_hidden, 1 To pn_hidden)
            ReDim dhdh(1 To pn_hidden, 1 To pn_hidden)
            For i = 1 To pn_hidden
                tmp_x = z(i, iterate) * (1 - s(i, iterate) ^ 2)
                dhds(i) = tmp_x
                dhdz(i) = (s(i, iterate) - h(i, iterate - 1)) * (z(i, iterate) * (1 - z(i, iterate)))
                For j = 1 To pn_hidden
                    dhdf(i, j) = tmp_x * us(i, j) * grad_rs(j)
                Next j
                
                dhdh(i, i) = 1 - z(i, iterate)
                For j = 1 To pn_hidden
                    tmp_y = us(i, j) * f(j, iterate)
                    For m = 1 To pn_hidden
                        tmp_y = tmp_y + us(i, m) * grad_rs(m) * uf(m, j)
                    Next m
                    dhdh(i, j) = dhdh(i, j) + dhdz(i) * uz(i, j) + dhds(i) * tmp_y
                Next j
            Next i
            
            If iterate < n_T Then
                For k = 1 To pn_hidden
                    tmp_x = 0
                    For i = 1 To pn_hidden
                        tmp_x = tmp_x + xA_prv(i) * dhdh(i, k)
                    Next i
                    xA(k) = dEdh(k) + tmp_x
                Next k
            ElseIf iterate = n_T Then
                For k = 1 To pn_hidden
                    xA(k) = dEdh(k)
                Next k
            End If
            
            For i = 1 To pn_hidden
                tmp_x = xA(i) * dhdz(i)
                tmp_y = xA(i) * dhds(i)
    
                dbz(i) = dbz(i) + tmp_x
                dbs(i) = dbs(i) + tmp_y
                For k = 1 To pn_hidden
                    dbf(i) = dbf(i) + xA(k) * dhdf(k, i)
                Next k
                
                If Not (pstrMode = "V2S" And iterate > 1) Then
                    For j = 1 To pn_input
                        tmp_z = x(j, iterate)
                        dwz(i, j) = dwz(i, j) + tmp_x * tmp_z
                        dws(i, j) = dws(i, j) + tmp_y * tmp_z
                        For k = 1 To pn_hidden
                            dwf(i, j) = dwf(i, j) + xA(k) * dhdf(k, i) * tmp_z
                        Next k
                    Next j
                End If
                
                For j = 1 To pn_hidden
                    tmp_z = h(j, iterate - 1)
                    duz(i, j) = duz(i, j) + tmp_x * tmp_z
                    dus(i, j) = dus(i, j) + tmp_y * tmp_z * f(j, iterate)
                    For k = 1 To pn_hidden
                        duf(i, j) = duf(i, j) + xA(k) * dhdf(k, i) * tmp_z
                    Next k
                Next j
            
            Next i

        End If
        
        For k = 1 To pn_hidden
            xA_prv(k) = xA(k)
        Next k
        
    Next iterate
    
End Sub



Private Function OutputLayerFwd(x() As Double)
Dim i As Long, n As Long
Dim y() As Double
Dim tmp_x As Double

    n = UBound(x, 1)
    ReDim y(1 To n)

    Select Case UCase(pOutputLayer)
    Case "SIGMOID"

        For i = 1 To n
            y(i) = f_sigmoid(x(i))
        Next i
        OutputLayerFwd = y
        
    Case "TANH"

        For i = 1 To n
            y(i) = f_tanh(x(i))
        Next i
        OutputLayerFwd = y
    
    Case "LINEAR"

        OutputLayerFwd = x
        
    Case "SOFTMAX"
        
        tmp_x = 0
        For i = 1 To n
            y(i) = Exp(x(i))
            tmp_x = tmp_x + y(i)
        Next i
        For i = 1 To n
            y(i) = y(i) / tmp_x
        Next i
        OutputLayerFwd = y
    
    Case Else
        
        MsgBox "cRNN: OutputlayerFwd: " & pOutputLayer & " is not a valid output layer."
    
    End Select
    
End Function


Private Function f_tanh(x) As Double
Dim x1 As Double, x2 As Double
    If x > 10 Then
        f_tanh = 1
    ElseIf x < (-10) Then
        f_tanh = -1
    Else
        x1 = Exp(x)
        x2 = Exp(-x)
        f_tanh = (x1 - x2) / (x1 + x2)
    End If
End Function

Private Function f_sigmoid(x) As Double
    If x > 20 Then
        f_sigmoid = 1
    ElseIf x < (-20) Then
        f_sigmoid = 0
    Else
        f_sigmoid = 1 / (1 + Exp(-x))
    End If
End Function

Private Function ArrayIsEmpty(x) As Boolean
Dim i As Long
    If Not IsArray(x) Then
        ArrayIsEmpty = True
    Else
        ArrayIsEmpty = False
        On Error Resume Next
        i = UBound(x, 1)
        If Err.Number <> 0 Then
            Err.Clear
            ArrayIsEmpty = True
        End If
    End If
End Function

'Return a subset of input array x()
Private Function SubsetIdx(x As Variant, Optional m As Long, Optional n As Long, Optional x_idx As Variant = Null)
Dim i As Long, j As Long, k As Long
Dim y As Variant, y_idx() As Long

    If IsArray(x_idx) Then
        k = UBound(x_idx, 1)
        If getDimension(x) = 1 Then
            ReDim y(1 To k)
            For i = 1 To k
                y(i) = x(x_idx(i))
            Next i
        ElseIf getDimension(x) = 2 Then
            m = UBound(x, 1)
            ReDim y(1 To m, 1 To k)
            For i = 1 To k
                For j = 1 To m
                    y(j, i) = x(j, x_idx(i))
                Next j
            Next i
        End If
        SubsetIdx = y
        Exit Function
    End If
    
    k = n
    If n > UBound(x, 1) Then
        k = UBound(x, 1)
    End If
    
    ReDim y_idx(1 To k - m + 1)
    For i = 1 To (k - m + 1)
        y_idx(i) = x(m + i - 1)
    Next i
    
    SubsetIdx = y_idx
    
End Function

Private Function getDimension(A As Variant) As Long
    Dim i As Long, j As Long
    i = 0
    On Error GoTo getDimension_Err:
    Do While True:
        i = i + 1
        j = UBound(A, i)
    Loop
getDimension_Err:
    getDimension = i - 1
End Function

'Shuffle an integer vector x()
Private Sub Shuffle(x() As Long)
Dim i As Long, j As Long, n As Long
Dim k As Long
Dim vtmp As Variant
    n = UBound(x)
    Randomize
    For i = n To 2 Step -1
        j = Int(Rnd() * i) + 1  'Random_Integer(1, i)
        vtmp = x(j)
        x(j) = x(i)
        x(i) = vtmp
    Next i
End Sub



