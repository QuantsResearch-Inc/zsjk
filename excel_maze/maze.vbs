
'迷路を探索するVBA
'迷路は黒いセルで作成する。迷路は外周を黒いセルで囲んでおく必要がある。
'開始地点にはS、終了地点にはFを記述しておく
'Solveで迷路を探索する
' 探索は、”O"という文字が現在の位置を表し、一度通った道はグレーに塗られる。
' 複数回同じ道を通るとだんだんとグレーが濃くなっていく
'Resetで迷路を初期状態に戻す。Solveを実行した後に実行する必要がある
Const MAX_ROW As Integer = 26 '迷路を作成する最大行、Z列まで
Const MAX_COL As Integer = 26 '迷路を作成する最大列、26行まで
Const MAX_LOOP As Integer = 1000  '自動探索を行う最大試行回数
Dim START_I, START_J, FINISH_I, FINISH_J As Integer 'スタート位置と終了位置、リセットのために記憶しておく

'セルに記述されている文字の位置検索する
' mark: 入力、"S" もしくは "F"
' row: 出力、markの文字が記述されているセルの行番号
' col: 出力、markの文字が記述されているセルの列番号
' 戻り値: Trueの時はmarkの文字が見つかった、Falseの時はmarkの文字が見つからなかった
Function SearchPoint(mark, row, col) As Boolean
    Dim i, j As Integer
    row = 0
    col = 0
    For i = 1 To MAX_ROW
        For j = 1 To MAX_COL
            If Cells(i, j).Value = mark Then
                row = i
                col = j
                SearchPoint = True
                Exit Function
            End If
        Next j
    Next i
    SearchPoint = False
End Function


' 迷路を探索する
Sub Solve()
    Dim start_row, start_col As Integer
    Dim finish_row, fInish_col As Integer
    Dim i As Integer
    Dim now_row, now_col As Integer
    Dim prev_row, prev_col As Integer
    Dim up_col, down_col, left_col, right_col As Long
    Dim prev_color As Long

    ' "S"の位置を検索する
    If Not SearchPoint("S", start_row, start_col) Then
        MsgBox "Sが見つかりませんでした"
        Exit Sub
    End If

    ' "F"の位置を検索する
    If Not SearchPoint("F", finish_row, fInish_col) Then
        MsgBox "Fが見つかりませんでした"
        Exit Sub
    End If

    START_I = start_row  ' リセット用に"S"の行番号を保存しておく
    START_J = start_col  ' リセット用に"S"の列番号を保存しておく
    FINISH_I = finish_row ' リセット用に"F"の行番号を保存しておく
    FINISH_J = fInish_col ' リセット用に"F"の列番号を保存しておく
    prev_row = start_row  ' 直前の”O"の行位置を保存しておく
    prev_col = start_col  ' 直前の”O"の列位置を保存しておく
    now_row = start_row   ' 現在の"O"の行位置
    now_col = start_col   ' 現在の”O"の列位置

    ' MAX_LOOPまで探索を行う、
    For i = 1 To MAX_LOOP
        prev_color = Cells(prev_row, prev_col).Interior.Color  ' 直前の"O"の位置のセルの色を取得する
        prev_color = prev_color - RGB(32, 32, 32)              ' 直前の"O"の位置のセルの色にグレーを塗っていく。引き算にしているのは複数回通った際に、前回よりも濃いグレーにするため。
        Cells(prev_row, prev_col).Interior.Color = prev_color  ' 直前の"O"のセルをグレーに塗る
        Cells(prev_row, prev_col).Value = ""                   ' 直前の"O"の文字を消す
        Cells(start_row, start_col).Value = "S"                ' スタート地点に"S"を上書きする。スタート位置に"O"があるときの対応のため
        Cells(now_row, now_col).Value = "O"                    ' 現在の位置に"O"を記述する
        If now_row = finish_row And now_col = fInish_col Then  ' 現在の位置が"F"と同じならば終了
            MsgBox "おめでとうございます。終了です。"
            Exit Sub
        End If
        prev_row = now_row        ' 現在のセルの行位置を前回の行位置に代入する
        prev_col = now_col        ' 現在のセルの行位置を前回の列位置に代入する

        up_col = Cells(now_row - 1, now_col).Interior.Color     ' 現在の位置の上のセルの色を取得する
        down_col = Cells(now_row + 1, now_col).Interior.Color   ' 現在の位置の下のセルの色を取得する
        left_col = Cells(now_row, now_col - 1).Interior.Color   ' 現在の位置の左のセルの色を取得する
        right_col = Cells(now_row, now_col + 1).Interior.Color  ' 現在の位置の右のセルの色を取得する

        ' 上下左右のセルの色を比較し、最も色の薄いセルに進む
        ' 色が薄いほど数値が大きい
        If up_col >= down_col And up_col >= left_col And up_col >= right_col Then '上のセルが一番薄い色か？
            now_row = now_row - 1
        ElseIf down_col >= up_col And down_col >= left_col And down_col >= right_col Then '下のセルが一番薄い色か？
            now_row = now_row + 1
        ElseIf left_col >= up_col And left_col >= down_col And left_col >= right_col Then '左のセルが一番薄い色か？
            now_col = now_col - 1
        ElseIf right_col >= up_col And right_col >= down_col And right_col >= left_col Then '右のセルが一番薄い色か？
            now_col = now_col + 1
        Else
            MsgBox "次に進む道が見つかりません"
            Exit Sub
        End If
        Application.wait [Now() + "0:00:00.2"]  ' 0.2秒まつ
    Next i
    MsgBox "ギブアップ。"
End Sub

' 迷路をリセットし、元の状態に戻す。通路を白に塗り直し、”O"の文字を削除する。
' ”S"と”F"をリセット用にほぞんしたSTART_I, START_J, FINISH_I, FINISH_Jで描き直す
Sub Reset()
    Dim i, j As Integer
    If START_I = 0 Or START_J = 0 Or FINISH_I = 0 Or FINISH_J = 0 Then
        MsgBox "Cannot reset"
        Exit Sub
    End If

    For i = 1 To MAX_COL
        For j = 1 To MAX_ROW
            ' もしセルが白でもなく、黒でもない場合には、白にする
            If Cells(i, j).Interior.Color <> RGB(255, 255, 255) And Cells(i, j).Interior.Color <> RGB(0, 0, 0) Then
                Cells(i, j).Interior.ColorIndex = 0
            End If
            Cells(i, j).Value = "" ' 文字を消す
        Next j
    Next i
    Cells(START_I, START_J).Value = "S"      ' スタート位置に"S"を記述する
    Cells(FINISH_I, FINISH_J).Value = "F"    ' 終了位置に"F"を記述する
End Sub
