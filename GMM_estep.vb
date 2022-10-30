'Estepまでのコード
Sub GMM()
    '変数の定義
    Dim LastRow, LastColumn
    Dim Data, Data_T
    Dim ws0, ws1 As Worksheet
    'worksheet の指定
    Set ws0 = Worksheets("data")
    Set ws1 = Worksheets("result")
    
    'データの情報を取得する
    LastRow = ws0.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = ws0.Cells(1, Columns.Count).End(xlToLeft).Column
    
    
    Dim N As Integer, M As Integer, k As Integer
    N = LastRow - 1
    M = LastColumn - 1
    ReDim Data(0 To N, 0 To M)
    Data = ws0.Range(ws0.Cells(2, 2), ws0.Cells(LastRow, LastColumn)).Value
    Data_T = WorksheetFunction.Transpose(Data)
    Data_all = ws0.Range(ws0.Cells(1, 1), ws0.Cells(LastRow, LastColumn)).Value
    D = Data

    
    
    'msgboxからKの値を受け取る
    k = InputBox(prompt:="クラスタの数を入力してください", Default:=2, Title:="K=?")
    
    'Kが整数か判定する
    If Not VarType(k) = 2 Then
        k = InputBox(prompt:="整数を入力してください", Default:=2, Title:="K=?")
        End If
    
    'パラメーターの初期化
    Dim pi()
    ReDim pi(1 To k)
    For i = 1 To k
        pi(i) = 1 / k
        Next i
        
    Dim mu()
    ReDim mu(1 To M, 1 To k)
    For i = 1 To M
        For j = 1 To k
            mu(i, j) = 1 / N
        Next j
    Next i
    
    Dim Sigma()
    ReDim Sigma(1 To M, 1 To M, 1 To k)
    For l = 1 To k
        For i = 1 To M
            For j = 1 To M
                If i = j Then
                    Sigma(i, j, l) = 1
                Else
                    Sigma(i, j, l) = 0
                End If
            Next j
        Next i
    Next l
    
    
    For i = 1 To k
        ws1.Cells(1, i).Value = "mu_" + CStr(i)
        ws1.Range(Cells(2, i), Cells(2 + M - 1, i)).Value = get_mu_k(mu, (i))
    Next i
    For i = 1 To k
        ws1.Cells(1, k + M * (i - 1) + 1).Value = "Sigma_" + CStr(i)
        ws1.Range(Cells(2, k + M * (i - 1) + 1), Cells(2 + M - 1, k + M * i)).Value = get_Sigma_k(Sigma, (i))
        Next i
               
    
    
    'パラメーターの初期化終わり
    
    
    'Estep

    Dim G As Double
    Dim gamma()
    ReDim gamma(1 To N, 1 To k)

    For i = 1 To N
        Dim x_n()
        ReDim x_n(1 To M)
        
        
        For j = 1 To M
            x_n(j) = Data(i, j)
        Next j
        For l = 1 To k

            mu_l = get_mu_k(mu, (l))

            Sigma_l = get_Sigma_k(Sigma, (l))
            
            Dim y
            ReDim y(1 To M, 1 To 1)
            For p = 1 To M
                y(p, 1) = x_n(p) - mu_l(p)
            Next p
            Sigma_inv = WorksheetFunction.MInverse(Sigma_l)
            
            shorder = (-1 / 2) * _
            WorksheetFunction.MMult(WorksheetFunction.MMult(WorksheetFunction.Transpose(y) _
            , Sigma_inv), y)(1)

            Mnormal = (2 * WorksheetFunction.pi) ^ (-M / 2) * WorksheetFunction.MDeterm(Sigma_l) ^ (-1 / 2) * Exp(shorder)
            
            
            G = G + pi(l) * Mnormal
        Next l
        For l = 1 To k
            mu_l = get_mu_k(mu, (l))
            Sigma_l = get_Sigma_k(Sigma, (l))
            Sigma_inv = WorksheetFunction.MInverse(Sigma_l)
            shorder = (-1 / 2) * WorksheetFunction.MMult(WorksheetFunction.MMult(WorksheetFunction.Transpose(y), Sigma_inv), y)(1)
            Mnormal = (2 * WorksheetFunction.pi) ^ (-M / 2) * WorksheetFunction.MDeterm(Sigma_l) ^ (-1 / 2) * Exp(shorder)
            
            gamma(i, l) = pi(l) * Mnormal / G
        Next l
    Next i
    
    'gamma_nk をワークシートに書いておく
        For i = 1 To k
            ws1.Cells(1, k + M * k + i).Value = "gamma_" + CStr(i)
            Dim gamma_k()
            ReDim gamma_k(1 To N, 1 To 1)
            For j = 1 To N
                gamma_k(j, 1) = gamma(j, i)
            Next j
            
            ws1.Range(Cells(2, k + M * k + i), Cells(2 + N - 1, k + M * k + i)).Value = gamma_k
        Next i
    

    Dim N_k()
    ReDim N_k(1 To k)
    For i = 1 To k
        N_k(i) = WorksheetFunction.Sum(ws1.Range(Cells(2, k + M * k + i), Cells(2 + N - 1, k + M * k + i)))
        Debug.Print N_k(i)
    Next i
End Sub

'μの一部を抜き出す関数
Function get_mu_k(mu(), k As Long)
    M = UBound(mu, 1)
    Dim mu_k
    ReDim mu_k(1 To M)
    
    For i = 1 To M
        mu_k(i) = mu(i, k)
    Next i
    get_mu_k = mu_k
End Function

'Σの一部を抜き出す関数
Function get_Sigma_k(Sigma(), k As Long)
    M = UBound(Sigma, 1)
    Dim Sigma_k
    ReDim Sigma_k(1 To M, 1 To M)
    
    For i = 1 To M
        For j = 1 To M
            Sigma_k(i, j) = Sigma(i, j, k)
        Next j
    Next i
    get_Sigma_k = Sigma_k
End Function

'データの一部を抜き出す関数
Function get_x_n(X(), k As Integer)
    M = UBound(X, 2)
    Dim x_n
    ReDim x_n(1 To M)
    For i = 1 To M
        x_n(i) = X(k, i)
    Next i
    get_x_n = x_n
        
End Function
