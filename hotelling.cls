VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Hotelling()
    '変数の定義
    Dim LastRow, LastColumn
    Dim Data, Data_T, S
    Dim ws0, ws1 As Worksheet
    'worksheet の指定
    Set ws0 = Worksheets("data")
    Set ws1 = Worksheets("result")
    'resultシートに色を付けるので初期化する
    ws1.Cells.ClearFormats
    
    'データの情報を取得する
    LastRow = ws0.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = ws0.Cells(1, Columns.Count).End(xlToLeft).Column
    Data = ws0.Range(ws0.Cells(2, 2), ws0.Cells(LastRow, LastColumn)).Value
    Data_all = ws0.Range(ws0.Cells(1, 1), ws0.Cells(LastRow, LastColumn)).Value
    N = LastRow - 1
    k = LastColumn - 1
    D = Data
    '共分散の計算
    ReDim h(k)
    For i = 1 To k
        S = 0
        For j = 1 To N
            S = S + Data(j, i)
        Next j
        h(i) = S / N
    Next i
    
    For i = 1 To k
        For j = 1 To N
           D(j, i) = (Data(j, i) - h(i)) / N
        Next j
    Next i
    
    D_T = WorksheetFunction.Transpose(D)
    S = WorksheetFunction.MMult(D_T, D)
    S_inv = WorksheetFunction.MInverse(S)
    'カイ二乗関連の計算と判定結果の記録
    C = WorksheetFunction.MMult(WorksheetFunction.MMult(D, S_inv), D_T)
    ReDim chi(N), test(N)
    chi_M = WorksheetFunction.ChiSq_Inv(0.05, k)
    For i = 1 To N
        chi(i) = C(i, i)
        If chi(i) < chi_M Then
            test(i) = "正常"
        Else
            test(i) = "異常"
        End If
        ws1.Cells(i + 1, 1) = test(i)
        
        If ws1.Cells(i + 1, 1) = "異常" Then
            ws1.Cells(i + 1, 1).Font.Color = RGB(255, 0, 0)
        End If
    Next i
    ws1.Cells(1, 1) = "test"
    
    'resultシートにもデータの写しを作成する
    For i = 1 To N + 1
        For j = 1 To k + 1
            ws1.Cells(i, j + 1) = Data_all(i, j)
        Next j
    Next i
    
    
End Sub
