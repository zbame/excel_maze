' 迷路を生成して、迷うこと無く解く

' 迷路を入れる2次元配列
' mazeMatriz(2,2) が迷路内での現在地の場合
' mazeMatrix(1,1): 左上の柱　mazeMatrix(2,1): 上の壁　mazeMatrix(3,1): 右上の柱
' mazeMatrix(1,2): 左の壁　　mazeMatrix(2,2): 現在地　mazeMatrix(3,2): 右の壁
' mazeMatrix(1,3): 左下の柱　mazeMatrix(2,3): 下の壁　mazeMatrix(3,3): 右下の柱
' 現在地を移動するときは２つ飛びで上下左右に移動する
' 柱と壁に '1' または '2' をセットする
' 壁の値は描画する時に参照する（解くときには使用されない）
' 柱の値は迷路を生成するときと迷路を解くときに参照する
Private mazeMatrix()

' 迷路のX軸方向のサイズ
Private sizeX As Integer

' 迷路をY軸方向のサイズ
Private sizeY As Integer

' 壁を作るX軸上の現在地
Private currentPoleX As Integer

' 壁を作るY軸上の現在地
Private currentPoleY As Integer

' 迷路を解くときのX軸上の現在地
Private currentMazeX As Integer

' 迷路を解くときのY軸上の現在地
Private currentMazeY As Integer

' 迷路を解くのときの進行方向
' 0:左
' 1:上
' 2:右
' 3:下
Private direction As Integer

' 迷路を描く行オフセット
Private Const xOffset As Integer = 1

' 迷路を描く列オフセット
Private Const yOffset As Integer = 3

' 迷路の初期化
' x: X軸方向の迷路の大きさ
' y: Y軸方向の迷路の大きさ
Private Sub initializeMaze(ByVal x As Integer, ByVal y As Integer)
    Dim ii, jj As Integer
    sizeX = x
    sizeY = y

    ' 迷路の大きさに併せて配列の大きさを初期化する
    ReDim mazeMatrix(1 To x * 2 + 2, 1 To y * 2 + 2)

    ' 配列の中の値を初期化する
    For ii = 1 To sizeX * 2 + 1
        For jj = 1 To sizeY * 2 + 1
            If ii = 1 Or jj = sizeY * 2 + 1 Then
                ' 迷路の左端と下端の柱と壁を '2' に設定する
                mazeMatrix(ii, jj) = 2
            Else
                If ii = sizeX * 2 + 1 Or jj = 1 Then
                    ' 迷路の右端と上の柱と壁を '1' に設定する
                    mazeMatrix(ii, jj) = 1
                Else
                    ' それ以外は 0 に設定する
                    mazeMatrix(ii, jj) = 0
                End If
            End If
        Next jj
    Next ii
    
    ' 入り口の壁を '0' に設定する
    mazeMatrix(1, 2) = 0
    
    ' 出口の壁を '0' に設定する
    mazeMatrix(sizeX * 2 + 1, sizeY * 2) = 0
End Sub

' 壁を伸ばす時の現在の柱を設定する
Private Sub setCurrentPole(ByVal x As Integer, ByVal y As Integer)
    currentPoleX = x
    currentPoleY = y
End Sub

' 壁を下に伸ばす
Private Sub moveWallBottom()

    ' 現在の柱の値を x に代入する
    x = mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 - 1)
    
    ' 壁を伸ばす
    mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2) = x
    
    ' 柱を立てる
    mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 + 1) = x
    
    ' 現在地を下に移動する
    currentPoleY = currentPoleY + 1
End Sub

' 壁を上に伸ばす
Private Sub moveWallTop()

    ' 現在の柱の値を x に代入する
    x = mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 - 1)

    ' 壁を伸ばす
    mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 - 2) = x

    ' 柱を立てる
    mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 - 3) = x

    ' 現在地を上に移動する
    currentPoleY = currentPoleY - 1
End Sub

' 壁を右へ伸ばす
Private Sub moveWallRight()

    ' 現在の柱の値を x に代入する
    x = mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 - 1)

    ' 壁を伸ばす
    mazeMatrix(currentPoleX * 2, currentPoleY * 2 - 1) = x

    ' 柱を立てる
    mazeMatrix(currentPoleX * 2 + 1, currentPoleY * 2 - 1) = x

    ' 現在地を右に移動する
    currentPoleX = currentPoleX + 1
End Sub
 
' 壁を左へ伸ばす
Private Sub moveWallLeft()

    ' 現在の柱の値を x に代入する
    x = mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 - 1)

    ' 壁を伸ばす
    mazeMatrix(currentPoleX * 2 - 2, currentPoleY * 2 - 1) = x

    ' 柱を立てる
    mazeMatrix(currentPoleX * 2 - 3, currentPoleY * 2 - 1) = x

    ' 現在地を左に移動する
    currentPoleX = currentPoleX - 1
End Sub

' 壁をランダムな方向へ伸ばす
' 伸ばすことができない場合 1 を返す
Private Function buildWall() As Integer
    buildWall = 0
    If currentPoleX > sizeX + 1 Or currentPoleY > sizeY + 1 Then buildWall = 2
    result = detectEnd()
    Select Case result
        Case 0
            ' このケースは使いません
            R = Int(Rnd() * 4)
            Select Case R
                Case 0
                    Call moveWallTop
                Case 1
                    Call moveWallBottom
                Case 2
                    Call moveWallLeft
                Case 3
                    Call moveWallRight
            End Select

        Case 1
            ' 上、下、右に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 3)
            Select Case R
                Case 0
                    Call moveWallTop
                Case 1
                    Call moveWallBottom
                Case 2
                    Call moveWallRight
            End Select
        
        Case 2
            ' 上、下、左に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 3)
            Select Case R
                Case 0
                    Call moveWallTop
                Case 1
                    Call moveWallBottom
                Case 2
                    Call moveWallLeft
            End Select

        Case 3
            ' 上、下に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 2)
            Select Case R
                Case 0
                    Call moveWallTop
                Case 1
                    Call moveWallBottom
            End Select
        
        Case 4
            ' 下、左、右に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 3)
            Select Case R
                Case 0
                    Call moveWallBottom
                Case 1
                    Call moveWallLeft
                Case 2
                    Call moveWallRight
            End Select
        
        Case 5
            ' 上、右に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 2)
            Select Case R
                Case 0
                    Call moveWallBottom
                Case 1
                    Call moveWallRight
            End Select

        Case 6
            ' 下、左に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 2)
            Select Case R
                Case 0
                    Call moveWallBottom
                Case 1
                    Call moveWallLeft
            End Select
        
        Case 7
            ' 下方向へ壁を伸ばす
            Call moveWallBottom

        Case 8
            ' 上、左、右に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 3)
            Select Case R
                Case 0
                    Call moveWallTop
                Case 1
                    Call moveWallLeft
                Case 2
                    Call moveWallRight
           End Select

        Case 9
            ' 上、右に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 2)
            Select Case R
                Case 0
                    Call moveWallTop
                Case 1
                    Call moveWallRight
            End Select

        Case 10
            ' 上、下、右に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 2)
            Select Case R
                Case 0
                    Call moveWallTop
                Case 1
                    Call moveWallLeft
            End Select

        Case 11
            ' 上へ壁を伸ばす
            Call moveWallTop
        
        Case 12
            ' 左、右に柱がない場合にそのどれかランダムな方向へ壁を伸ばす
            R = Int(Rnd() * 2)
            Select Case R
                Case 0
                    Call moveWallLeft
                Case 1
                    Call moveWallRight
            End Select

        Case 13
            ' 右へ壁を伸ばす
            Call moveWallRight
        
        Case 14
            ' 左へ壁を伸ばす
            Call moveWallLeft
        
        Case Else
            ' 柱に囲まれたら 1 を返す
            buildWall = 1
    End Select
    
    ' 40分1の確率で、壁を伸ばすのをやめる
    If Int(Rnd() * 40) = 0 Then buildWall = 1
End Function

' 現在の柱から見て左に柱があるか調べる
Private Function isLeftPole() As Boolean
    If currentPoleX = 1 Then
        isLeftPole = True
    Else
        If mazeMatrix(currentPoleX * 2 - 3, currentPoleY * 2 - 1) <> 0 Then
            isLeftPole = True
        Else
            isLeftPole = False
        End If
    End If
End Function

' 現在の柱から見て右に柱があるか調べる
Private Function isRightPole() As Boolean
    If currentPoleX = sizeX + 1 Then
        isRightPole = True
    Else
        If mazeMatrix(currentPoleX * 2 + 1, currentPoleY * 2 - 1) <> 0 Then
            isRightPole = True
        Else
            isRightPole = False
        End If
    End If
End Function

' 現在の柱から見て上に柱があるか調べる
Private Function isTopPole() As Boolean
    If currentPoleY = 1 Then
        isTopPole = True
    Else
        If mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 - 3) <> 0 Then
            isTopPole = True
        Else
            isTopPole = False
        End If
    End If
End Function

' 現在の柱から見て下に柱があるか調べる
Private Function isBottomPole() As Boolean
    If currentPoleY = sizeY + 1 Then
        isBottomPole = True
    Else
        If mazeMatrix(currentPoleX * 2 - 1, currentPoleY * 2 + 1) <> 0 Then
            isBottomPole = True
        Else
            isBottomPole = False
        End If
    End If
End Function

' 現在の柱から見て左右上下の柱の有無を調べる
' 1: 左
' 2: 右
' 4: 上
' 8: 下
' 上記の合計を戻り値とする
Private Function detectEnd() As Integer
    detectEnd = 0
    If isLeftPole() Then detectEnd = 1
    If isRightPole() Then detectEnd = detectEnd + 2
    If isTopPole() Then detectEnd = detectEnd + 4
    If isBottomPole() Then detectEnd = detectEnd + 8
End Function



'迷路を表示する
Private Sub viewMaze()
    Dim xx, yy As Integer
    Application.ScreenUpdating = False
    
    ' 全てのセルの塗りを消去
    Cells.Interior.Pattern = xlNone
    
    ' 全てのセルの罫線を消去
    Cells.Borders.LineStyle = xlNone
    
    ' 迷路を表示するセルの高さを 22px にする
    Range(Cells(yOffset, xOffset), Cells(sizeY + yOffset, sizeX + xOffset)).RowHeight = 22
    
    ' 迷路を表示するセルの幅を 22px にする
    Range(Cells(yOffset, xOffset), Cells(sizeY + yOffset, sizeX + xOffset)).ColumnWidth = 2.43
    
    ' 迷路を表示する
    For xx = 1 To sizeX + 1
         For yy = 1 To sizeY + 1
            If mazeMatrix(xx * 2 - 1, yy * 2) <> 0 Then
                Cells(yy + yOffset, xx + xOffset).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Cells(yy + yOffset, xx + xOffset).Borders(xlEdgeLeft).Weight = xlThick
            Else
                Cells(yy + yOffset, xx + xOffset).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
            End If
            If mazeMatrix(xx * 2, yy * 2 - 1) <> 0 Then
                Cells(yy + yOffset, xx + xOffset).Borders(xlEdgeTop).LineStyle = xlContinuous
                Cells(yy + yOffset, xx + xOffset).Borders(xlEdgeTop).Weight = xlThick
            Else
                Cells(yy + yOffset, xx + xOffset).Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            End If
        Next
    Next
    Application.ScreenUpdating = True
End Sub

' 表示上の迷路の現在地のセルの塗りを指定色にする
Private Sub setCellColor(ByVal color As Long)
    Cells(currentMazeY + yOffset, currentMazeX + xOffset).Interior.color = color
End Sub

' 迷路内での現在地を設定する
Private Sub setCurrentMaze(ByVal x As Integer, ByVal y As Integer)
    Application.ScreenUpdating = False
    currentMazeX = x
    currentMazeY = y
    
    ' 表示上の迷路の現在地のセルの塗りを赤にする
    Call setCellColor(vbRed)

    Application.ScreenUpdating = True
End Sub

' 迷路内での現在地を上に移動する
Private Sub moveMazeTop()
    Application.ScreenUpdating = False

    ' 表示上の迷路の現在地のセルの塗りを黄色にする
    Call setCellColor(vbYellow)
    
    ' 現在地を上に移動する
    currentMazeY = currentMazeY - 1
    
    ' 表示上の迷路の現在地のセルの塗りを赤にする
    Call setCellColor(vbRed)
    Application.ScreenUpdating = True
End Sub

' 迷路内での現在地を下に移動する
Private Sub moveMazeBottom()
    Application.ScreenUpdating = False

    ' 表示上の迷路の現在地のセルの塗りを黄色にする
    Call setCellColor(vbYellow)

    ' 現在地を上に移動する
    currentMazeY = currentMazeY + 1

    ' 表示上の迷路の現在地のセルの塗りを赤にする
    Call setCellColor(vbRed)
    Application.ScreenUpdating = True
End Sub

' 迷路内での現在地を左に移動する
Private Sub moveMazeLeft()
    Application.ScreenUpdating = False

    ' 表示上の迷路の現在地のセルの塗りを黄色にする
    Call setCellColor(vbYellow)

    ' 現在地を上に移動する
    currentMazeX = currentMazeX - 1

    ' 表示上の迷路の現在地のセルの塗りを赤にする
    Call setCellColor(vbRed)
    Application.ScreenUpdating = True
End Sub

' 迷路内での現在地を右に移動する
Private Sub moveMazeRight()
    Application.ScreenUpdating = False

    ' 表示上の迷路の現在地のセルの塗りを黄色にする
    Call setCellColor(vbYellow)

    ' 現在地を上に移動する
    currentMazeX = currentMazeX + 1

    ' 表示上の迷路の現在地のセルの塗りを赤にする
    Call setCellColor(vbRed)
    Application.ScreenUpdating = True
End Sub

' 迷路内での現在位置の四隅の柱に '2' の柱が何本あるか調べる
Private Function pole() As Integer
    Dim LT, RT, LB, RB As Integer
    LT = mazeMatrix(currentMazeX * 2 - 1, currentMazeY * 2 - 1)
    RT = mazeMatrix(currentMazeX * 2 + 1, currentMazeY * 2 - 1)
    LB = mazeMatrix(currentMazeX * 2 - 1, currentMazeY * 2 + 1)
    RB = mazeMatrix(currentMazeX * 2 + 1, currentMazeY * 2 + 1)
    pole = LT + RT + LB + RB - 4
End Function

' 最初の進行方向を決める
Private Function initDirection() As Integer
    Select Case pole()
        Case 2
            ' '2'の柱が2本の場合は最初の進行方向を下にする
            direction = 3
        Case 3
            ' '2'の柱が3本の場合は最初の進行方向を右にする
            direction = 2
    End Select
End Function

' 進行方向に進む
Private Sub moveMaze()
    Select Case direction
        Case 0
            Call moveMazeLeft
        Case 1
            Call moveMazeTop
        Case 2
            Call moveMazeRight
        Case 3
            Call moveMazeBottom
    End Select
End Sub

' 迷路を解く
' 進行方向へ進んで次の進行方向を決める
Private Function solveMaze() As Integer
    solveMaze = 0

    ' ゴールにたどり着いたら 1 を返す
    If currentMazeX = sizeX And currentMazeY = sizeY Then
        solveMaze = 1
        Exit Function
    Else
        ' 進行方向へ進む
        Call moveMaze

        ' 現在地の周りに'2'の柱が何本有るか調べる
        Select Case pole()
            Case 1
                ' '2'の柱が１本の場合、進行方向を時計回りに９０度回転する
                direction = (direction + 1) Mod 4
            Case 2
                ' '2'の柱が２本の場合、進行方向は変えない
                direction = direction
            Case 3
                ' '2'の柱が３本の場合、進行方向を反時計回りに９０度回転する
                direction = (direction + 3) Mod 4
            Case Else
                ' ここは使われない
                solveMaze = 1
        End Select
    End If
End Function

' 迷路を生成する
Private Sub createMaze(ByVal x As Integer, ByVal y As Integer)
    Dim xMax, yMax, xx, yy As Integer
    
    ' 迷路のX軸方向の大きさ
    xMax = x

    ' 迷路のY軸方向の大きさ
    yMax = y
    
    ' 迷路を初期化する
    Call initializeMaze(xMax, yMax)
    
    ' 迷路を表示する（外枠のみ）
    Call viewMaze
    DoEvents
    
    ' 迷路の下端から壁を伸ばす
    For xx = xMax + 1 To 1 Step -1
        Call setCurrentPole(xx, yMax + 1)
            While buildWall = 0
            Wend
    Next

    ' 迷路を表示する
    Call viewMaze
    DoEvents
    
    ' 迷路の右端から壁を伸ばす
    For yy = yMax + 1 To 1 Step -1
        Call setCurrentPole(xMax + 1, yy)
            While buildWall = 0
            Wend
    Next
    
    ' 迷路を表示する
    Call viewMaze
    DoEvents
    
    ' 迷路の上端から壁を伸ばす
    For xx = xMax + 1 To 1 Step -1
        Call setCurrentPole(xx, 1)
            While buildWall = 0
            Wend
    Next
    
    ' 迷路を表示する
    Call viewMaze
    DoEvents

    ' 迷路の左端から壁を伸ばす
    For yy = yMax + 1 To 1 Step -1
        Call setCurrentPole(1, yy)
            While buildWall = 0
            Wend
    Next
    
    ' 迷路を表示する
    Call viewMaze
    DoEvents

    ' 残りを壁で埋める
    For xx = xMax To 2 Step -1
        For yy = yMax To 2 Step -1
            Call setCurrentPole(xx, yy)
            While buildWall = 0
            Wend
        Next
    Next
    
    ' 迷路を表示する
    Call viewMaze
    DoEvents
    
    ' 迷路内での現在地をスタート地点に初期化する
    Call setCurrentMaze(1, 1)

    ' 進行方向を初期化する
    Call initDirection

End Sub

' ボタンクリックイベント
' 迷路を生成する
Sub Button1_Click()
    Call createMaze(40, 25)
End Sub

' ボタンクリックイベント
' 迷路を解く
Sub Button2_Click()
   ' ゴールにたどり着くまで進行方向を変えながら進む
   While solveMaze = 0
         DoEvents
   Wend
End Sub
