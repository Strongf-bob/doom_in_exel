Attribute VB_Name = "modExcelDoom"
Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const MAP_WIDTH As Long = 16
Private Const MAP_HEIGHT As Long = 16
Private Const VIEW_WIDTH As Long = 48
Private Const VIEW_HEIGHT As Long = 24
Private Const VIEW_ROW As Long = 2
Private Const VIEW_COL As Long = 2
Private Const MAP_ROW As Long = 2
Private Const MAP_COL As Long = 55
Private Const FOV As Double = PI / 3#
Private Const MAX_DEPTH As Double = 20#
Private Const MOVE_STEP As Double = 0.3
Private Const STRAFE_STEP As Double = 0.22
Private Const TURN_STEP As Double = 0.18
Private Const PLAYER_RADIUS As Double = 0.16
Private Const ENEMY_COUNT As Long = 4

Private Type EnemyState
    X As Double
    Y As Double
    Alive As Boolean
End Type

Private mMap(0 To MAP_HEIGHT - 1) As String
Private mEnemies(1 To ENEMY_COUNT) As EnemyState
Private mPlayerX As Double
Private mPlayerY As Double
Private mPlayerAngle As Double
Private mAmmo As Long
Private mHealth As Long
Private mKills As Long
Private mStarted As Boolean
Private mKeysBound As Boolean
Private mMuzzleFlash As Boolean
Private mWallDistances(1 To VIEW_WIDTH) As Double

Public Sub ExcelDoom_ConfigureSheet()
    SetupSheet GetGameSheet()
    ShowIdleScreen
End Sub

Public Sub ExcelDoom_StartGame()
    Dim ws As Worksheet

    Set ws = GetGameSheet()
    InitializeMap
    InitializeEnemies
    SetupSheet ws

    mPlayerX = 2.5
    mPlayerY = 2.5
    mPlayerAngle = 0#
    mAmmo = 18
    mHealth = 100
    mKills = 0
    mMuzzleFlash = False
    mStarted = True

    BindKeys
    RenderFrame
End Sub

Public Sub ExcelDoom_ResetGame()
    ExcelDoom_StartGame
End Sub

Public Sub ExcelDoom_StopGame()
    mStarted = False
    mMuzzleFlash = False
    UnbindKeys
    ShowIdleScreen
    RenderHud GetGameSheet(), "Остановлено. Запусти ExcelDoom_StartGame, чтобы вернуться в бой."
End Sub

Public Sub ExcelDoom_MoveForward()
    If Not EnsureRunning() Then Exit Sub
    AttemptMove Cos(mPlayerAngle) * MOVE_STEP, Sin(mPlayerAngle) * MOVE_STEP
    StepSimulation
End Sub

Public Sub ExcelDoom_MoveBackward()
    If Not EnsureRunning() Then Exit Sub
    AttemptMove -Cos(mPlayerAngle) * MOVE_STEP, -Sin(mPlayerAngle) * MOVE_STEP
    StepSimulation
End Sub

Public Sub ExcelDoom_TurnLeft()
    If Not EnsureRunning() Then Exit Sub
    mPlayerAngle = NormalizeAngle(mPlayerAngle - TURN_STEP)
    StepSimulation
End Sub

Public Sub ExcelDoom_TurnRight()
    If Not EnsureRunning() Then Exit Sub
    mPlayerAngle = NormalizeAngle(mPlayerAngle + TURN_STEP)
    StepSimulation
End Sub

Public Sub ExcelDoom_StrafeLeft()
    If Not EnsureRunning() Then Exit Sub
    AttemptMove Cos(mPlayerAngle - PI / 2#) * STRAFE_STEP, Sin(mPlayerAngle - PI / 2#) * STRAFE_STEP
    StepSimulation
End Sub

Public Sub ExcelDoom_StrafeRight()
    If Not EnsureRunning() Then Exit Sub
    AttemptMove Cos(mPlayerAngle + PI / 2#) * STRAFE_STEP, Sin(mPlayerAngle + PI / 2#) * STRAFE_STEP
    StepSimulation
End Sub

Public Sub ExcelDoom_Shoot()
    Dim hitIndex As Long

    If Not EnsureRunning() Then Exit Sub
    If mAmmo <= 0 Then
        RenderHud GetGameSheet(), "Патроны закончились. Нажми RESET."
        Exit Sub
    End If

    mAmmo = mAmmo - 1
    mMuzzleFlash = True
    hitIndex = FindShootTarget()

    If hitIndex > 0 Then
        mEnemies(hitIndex).Alive = False
        mKills = mKills + 1
    End If

    StepSimulation
    mMuzzleFlash = False
End Sub

Private Function EnsureRunning() As Boolean
    If mStarted Then
        EnsureRunning = True
    Else
        RenderHud GetGameSheet(), "Игра не запущена. Запусти ExcelDoom_StartGame."
    End If
End Function

Private Sub StepSimulation()
    ApplyEnemyPressure
    RenderFrame
End Sub

Private Sub AttemptMove(ByVal dx As Double, ByVal dy As Double)
    Dim nextX As Double
    Dim nextY As Double

    nextX = mPlayerX + dx
    nextY = mPlayerY + dy

    If CanOccupy(nextX, mPlayerY) Then
        mPlayerX = nextX
    End If

    If CanOccupy(mPlayerX, nextY) Then
        mPlayerY = nextY
    End If
End Sub

Private Function CanOccupy(ByVal x As Double, ByVal y As Double) As Boolean
    CanOccupy = Not IsWallAt(x - PLAYER_RADIUS, y - PLAYER_RADIUS) _
        And Not IsWallAt(x + PLAYER_RADIUS, y - PLAYER_RADIUS) _
        And Not IsWallAt(x - PLAYER_RADIUS, y + PLAYER_RADIUS) _
        And Not IsWallAt(x + PLAYER_RADIUS, y + PLAYER_RADIUS)
End Function

Private Sub ApplyEnemyPressure()
    Dim i As Long
    Dim aliveCount As Long

    aliveCount = 0

    For i = 1 To ENEMY_COUNT
        If mEnemies(i).Alive Then
            aliveCount = aliveCount + 1
            If Distance(mPlayerX, mPlayerY, mEnemies(i).X, mEnemies(i).Y) < 1.5 Then
                If HasLineOfSight(mEnemies(i).X, mEnemies(i).Y) Then
                    mHealth = mHealth - 6
                End If
            End If
        End If
    Next i

    If mHealth <= 0 Then
        mHealth = 0
        mStarted = False
        UnbindKeys
    ElseIf aliveCount = 0 Then
        mStarted = False
        UnbindKeys
    End If
End Sub

Private Function FindShootTarget() As Long
    Dim i As Long
    Dim bestDistance As Double
    Dim enemyDistance As Double
    Dim enemyAngle As Double

    bestDistance = 1E+30

    For i = 1 To ENEMY_COUNT
        If mEnemies(i).Alive Then
            enemyDistance = Distance(mPlayerX, mPlayerY, mEnemies(i).X, mEnemies(i).Y)
            enemyAngle = NormalizeRelativeAngle(Atan2(mEnemies(i).Y - mPlayerY, mEnemies(i).X - mPlayerX) - mPlayerAngle)

            If Abs(enemyAngle) <= 0.12 And enemyDistance < bestDistance Then
                If HasLineOfSight(mEnemies(i).X, mEnemies(i).Y) Then
                    bestDistance = enemyDistance
                    FindShootTarget = i
                End If
            End If
        End If
    Next i
End Function

Private Function HasLineOfSight(ByVal targetX As Double, ByVal targetY As Double) As Boolean
    Dim i As Long
    Dim steps As Long
    Dim sampleX As Double
    Dim sampleY As Double

    steps = CLng(Distance(mPlayerX, mPlayerY, targetX, targetY) / 0.05)
    If steps < 1 Then steps = 1

    For i = 1 To steps - 1
        sampleX = mPlayerX + (targetX - mPlayerX) * (i / steps)
        sampleY = mPlayerY + (targetY - mPlayerY) * (i / steps)
        If IsWallAt(sampleX, sampleY) Then Exit Function
    Next i

    HasLineOfSight = True
End Function

Private Sub RenderFrame()
    Dim ws As Worksheet
    Dim col As Long
    Dim row As Long
    Dim distanceToWall As Double
    Dim hitTile As String
    Dim hitSide As Long
    Dim wallHeight As Long
    Dim ceilingRow As Long
    Dim floorRow As Long
    Dim wallColorValue As Long
    Dim statusText As String

    Set ws = GetGameSheet()
    RenderBackground ws

    For col = 1 To VIEW_WIDTH
        CastRay col, distanceToWall, hitTile, hitSide
        mWallDistances(col) = distanceToWall
        wallHeight = CLng((VIEW_HEIGHT * 0.8) / MaxDouble(0.2, distanceToWall))
        If wallHeight > VIEW_HEIGHT Then wallHeight = VIEW_HEIGHT

        ceilingRow = (VIEW_HEIGHT - wallHeight) \ 2
        floorRow = ceilingRow + wallHeight
        wallColorValue = GetWallColor(hitTile, distanceToWall, hitSide)

        For row = ceilingRow + 1 To floorRow
            PaintCell ws, VIEW_ROW + row - 1, VIEW_COL + col - 1, wallColorValue, " "
        Next row
    Next col

    RenderEnemies ws
    RenderWeapon ws
    RenderCrosshair ws
    RenderMinimap ws

    If mStarted Then
        statusText = "WASD/стрелки: ходьба и поворот, Q/E: шаг вбок, SPACE: выстрел."
    ElseIf mHealth = 0 Then
        statusText = "GAME OVER. Нажми RESET."
    Else
        statusText = "VICTORY. Все демоны уничтожены. Нажми RESET."
    End If

    RenderHud ws, statusText
End Sub

Private Sub RenderBackground(ByVal ws As Worksheet)
    Dim row As Long
    Dim col As Long
    Dim colorValue As Long

    For row = 1 To VIEW_HEIGHT
        If row <= VIEW_HEIGHT \ 2 Then
            colorValue = RGB(55, 12, 10)
        Else
            colorValue = RGB(28, 24, 24)
        End If

        For col = 1 To VIEW_WIDTH
            PaintCell ws, VIEW_ROW + row - 1, VIEW_COL + col - 1, colorValue, " "
        Next col
    Next row
End Sub

Private Sub RenderEnemies(ByVal ws As Worksheet)
    Dim i As Long
    Dim relAngle As Double
    Dim enemyDistance As Double
    Dim screenX As Long
    Dim sizeCells As Long
    Dim drawX As Long
    Dim drawY As Long
    Dim topRow As Long
    Dim leftCol As Long
    Dim targetColumn As Long

    For i = 1 To ENEMY_COUNT
        If mEnemies(i).Alive Then
            enemyDistance = Distance(mPlayerX, mPlayerY, mEnemies(i).X, mEnemies(i).Y)
            relAngle = NormalizeRelativeAngle(Atan2(mEnemies(i).Y - mPlayerY, mEnemies(i).X - mPlayerX) - mPlayerAngle)

            If Abs(relAngle) <= (FOV / 2#) + 0.1 Then
                screenX = CLng((VIEW_WIDTH / 2#) + ((relAngle / (FOV / 2#)) * (VIEW_WIDTH / 2#)))
                If screenX < 1 Then screenX = 1
                If screenX > VIEW_WIDTH Then screenX = VIEW_WIDTH

                targetColumn = screenX
                If enemyDistance < mWallDistances(targetColumn) Then
                    sizeCells = CLng(8# / MaxDouble(0.6, enemyDistance))
                    If sizeCells < 1 Then sizeCells = 1
                    If sizeCells > 6 Then sizeCells = 6

                    topRow = VIEW_ROW + (VIEW_HEIGHT \ 2) - sizeCells
                    leftCol = VIEW_COL + screenX - (sizeCells \ 2)

                    For drawY = 0 To sizeCells
                        For drawX = 0 To sizeCells
                            If topRow + drawY >= VIEW_ROW And topRow + drawY < VIEW_ROW + VIEW_HEIGHT Then
                                If leftCol + drawX >= VIEW_COL And leftCol + drawX < VIEW_COL + VIEW_WIDTH Then
                                    PaintCell ws, topRow + drawY, leftCol + drawX, RGB(220, 88, 0), " "
                                End If
                            End If
                        Next drawX
                    Next drawY
                End If
            End If
        End If
    Next i
End Sub

Private Sub RenderWeapon(ByVal ws As Worksheet)
    Dim row As Long
    Dim col As Long
    Dim baseRow As Long
    Dim baseCol As Long
    Dim colorValue As Long

    baseRow = VIEW_ROW + VIEW_HEIGHT - 4
    baseCol = VIEW_COL + (VIEW_WIDTH \ 2) - 2

    For row = 0 To 2
        For col = 0 To 4
            colorValue = RGB(95, 95, 95)
            If mMuzzleFlash And row = 0 And col >= 1 And col <= 3 Then
                colorValue = RGB(255, 180, 0)
            End If
            PaintCell ws, baseRow + row, baseCol + col, colorValue, " "
        Next col
    Next row
End Sub

Private Sub RenderCrosshair(ByVal ws As Worksheet)
    Dim centerRow As Long
    Dim centerCol As Long

    centerRow = VIEW_ROW + (VIEW_HEIGHT \ 2) - 1
    centerCol = VIEW_COL + (VIEW_WIDTH \ 2) - 1

    PaintCell ws, centerRow, centerCol, RGB(255, 244, 214), "+"
End Sub

Private Sub RenderMinimap(ByVal ws As Worksheet)
    Dim mapX As Long
    Dim mapY As Long
    Dim cellColor As Long
    Dim playerCellX As Long
    Dim playerCellY As Long
    Dim enemyCellX As Long
    Dim enemyCellY As Long
    Dim i As Long

    ws.Cells(MAP_ROW - 1, MAP_COL).Value2 = "MAP"
    ws.Cells(MAP_ROW - 1, MAP_COL).Font.Bold = True
    ws.Cells(MAP_ROW - 1, MAP_COL).Font.Color = RGB(255, 214, 168)

    For mapY = 0 To MAP_HEIGHT - 1
        For mapX = 0 To MAP_WIDTH - 1
            Select Case TileAt(mapX, mapY)
                Case "#"
                    cellColor = RGB(90, 38, 32)
                Case "X"
                    cellColor = RGB(160, 36, 36)
                Case "D"
                    cellColor = RGB(185, 130, 25)
                Case Else
                    cellColor = RGB(18, 18, 18)
            End Select

            PaintCell ws, MAP_ROW + mapY, MAP_COL + mapX, cellColor, " "
        Next mapX
    Next mapY

    playerCellX = Int(mPlayerX)
    playerCellY = Int(mPlayerY)
    PaintCell ws, MAP_ROW + playerCellY, MAP_COL + playerCellX, RGB(0, 176, 255), "P"

    For i = 1 To ENEMY_COUNT
        If mEnemies(i).Alive Then
            enemyCellX = Int(mEnemies(i).X)
            enemyCellY = Int(mEnemies(i).Y)
            PaintCell ws, MAP_ROW + enemyCellY, MAP_COL + enemyCellX, RGB(255, 110, 0), "E"
        End If
    Next i
End Sub

Private Sub RenderHud(ByVal ws As Worksheet, ByVal statusText As String)
    ws.Cells(1, 2).Value2 = "EXCEL DOOM"
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 2).Font.Size = 16
    ws.Cells(1, 2).Font.Color = RGB(255, 214, 168)

    ws.Cells(VIEW_ROW + VIEW_HEIGHT + 1, VIEW_COL).Value2 = "HP " & Format$(mHealth, "000") & "   AMMO " & Format$(mAmmo, "00") & "   KILLS " & mKills & "/" & ENEMY_COUNT
    ws.Cells(VIEW_ROW + VIEW_HEIGHT + 1, VIEW_COL).Font.Color = RGB(255, 255, 255)
    ws.Cells(VIEW_ROW + VIEW_HEIGHT + 2, VIEW_COL).Value2 = statusText
    ws.Cells(VIEW_ROW + VIEW_HEIGHT + 2, VIEW_COL).Font.Color = RGB(255, 214, 168)

    ws.Cells(20, MAP_COL).Value2 = "Управление"
    ws.Cells(21, MAP_COL).Value2 = "W / ↑  вперёд"
    ws.Cells(22, MAP_COL).Value2 = "S / ↓  назад"
    ws.Cells(23, MAP_COL).Value2 = "A,D / ←,→  поворот"
    ws.Cells(24, MAP_COL).Value2 = "Q / E  шаг вбок"
    ws.Cells(25, MAP_COL).Value2 = "SPACE  выстрел"
    ws.Range(ws.Cells(20, MAP_COL), ws.Cells(25, MAP_COL + 6)).Font.Color = RGB(228, 228, 228)
End Sub

Private Sub ShowIdleScreen()
    Dim ws As Worksheet
    Dim row As Long

    Set ws = GetGameSheet()
    SetupSheet ws
    RenderBackground ws

    For row = 9 To 14
        PaintCell ws, row, VIEW_COL + 12, RGB(80, 20, 20), " "
        PaintCell ws, row, VIEW_COL + 35, RGB(80, 20, 20), " "
    Next row

    ws.Cells(12, 16).Value2 = "START"
    ws.Cells(13, 11).Value2 = "Запусти макрос ExcelDoom_StartGame"
    ws.Cells(14, 13).Value2 = "через Alt+F8"
    ws.Range(ws.Cells(12, 11), ws.Cells(14, 34)).Font.Color = RGB(255, 214, 168)
    ws.Range(ws.Cells(12, 11), ws.Cells(14, 34)).Font.Bold = True
End Sub

Private Sub SetupSheet(ByVal ws As Worksheet)
    Dim col As Long
    Dim row As Long

    ws.Cells.Clear
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    For col = 1 To VIEW_COL + VIEW_WIDTH + 20
        ws.Columns(col).ColumnWidth = 2.3
    Next col

    For row = 1 To VIEW_ROW + VIEW_HEIGHT + 4
        ws.Rows(row).RowHeight = 15
    Next row

    ws.Range(ws.Cells(1, 1), ws.Cells(40, 90)).Font.Name = "Consolas"
    ws.Range(ws.Cells(1, 1), ws.Cells(40, 90)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 1), ws.Cells(40, 90)).VerticalAlignment = xlCenter
End Sub

Private Sub PaintCell(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal targetCol As Long, ByVal colorValue As Long, ByVal cellText As String)
    With ws.Cells(targetRow, targetCol)
        .Interior.Color = colorValue
        .Value2 = cellText
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub CastRay(ByVal columnIndex As Long, ByRef outDistance As Double, ByRef outTile As String, ByRef outSide As Long)
    Dim rayAngle As Double
    Dim rayDirX As Double
    Dim rayDirY As Double
    Dim mapX As Long
    Dim mapY As Long
    Dim stepX As Long
    Dim stepY As Long
    Dim sideDistX As Double
    Dim sideDistY As Double
    Dim deltaDistX As Double
    Dim deltaDistY As Double

    rayAngle = mPlayerAngle - (FOV / 2#) + ((columnIndex - 1) / (VIEW_WIDTH - 1)) * FOV
    rayDirX = Cos(rayAngle)
    rayDirY = Sin(rayAngle)
    mapX = Int(mPlayerX)
    mapY = Int(mPlayerY)

    If Abs(rayDirX) < 0.00001 Then
        deltaDistX = 1E+30
    Else
        deltaDistX = Abs(1# / rayDirX)
    End If

    If Abs(rayDirY) < 0.00001 Then
        deltaDistY = 1E+30
    Else
        deltaDistY = Abs(1# / rayDirY)
    End If

    If rayDirX < 0# Then
        stepX = -1
        sideDistX = (mPlayerX - mapX) * deltaDistX
    Else
        stepX = 1
        sideDistX = (mapX + 1# - mPlayerX) * deltaDistX
    End If

    If rayDirY < 0# Then
        stepY = -1
        sideDistY = (mPlayerY - mapY) * deltaDistY
    Else
        stepY = 1
        sideDistY = (mapY + 1# - mPlayerY) * deltaDistY
    End If

    Do
        If sideDistX < sideDistY Then
            sideDistX = sideDistX + deltaDistX
            mapX = mapX + stepX
            outSide = 0
        Else
            sideDistY = sideDistY + deltaDistY
            mapY = mapY + stepY
            outSide = 1
        End If

        outTile = TileAt(mapX, mapY)
        If outTile <> "." Then Exit Do
    Loop

    If outSide = 0 Then
        outDistance = (mapX - mPlayerX + (1 - stepX) / 2#) / rayDirX
    Else
        outDistance = (mapY - mPlayerY + (1 - stepY) / 2#) / rayDirY
    End If

    If outDistance < 0.1 Then outDistance = 0.1
    If outDistance > MAX_DEPTH Then outDistance = MAX_DEPTH
End Sub

Private Function GetWallColor(ByVal tileValue As String, ByVal distanceToWall As Double, ByVal sideHit As Long) As Long
    Dim shade As Long

    shade = 220 - CLng(distanceToWall * 10)
    If shade < 40 Then shade = 40
    If sideHit = 1 Then shade = shade - 25
    If shade < 20 Then shade = 20

    Select Case tileValue
        Case "X"
            GetWallColor = RGB(shade, 38, 38)
        Case "D"
            GetWallColor = RGB(shade, 150, 24)
        Case Else
            GetWallColor = RGB(shade, 70, 24)
    End Select
End Function

Private Function MaxDouble(ByVal leftValue As Double, ByVal rightValue As Double) As Double
    If leftValue > rightValue Then
        MaxDouble = leftValue
    Else
        MaxDouble = rightValue
    End If
End Function

Private Function Distance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Distance = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
End Function

Private Function IsWallAt(ByVal x As Double, ByVal y As Double) As Boolean
    IsWallAt = TileAt(Int(x), Int(y)) <> "."
End Function

Private Function TileAt(ByVal tileX As Long, ByVal tileY As Long) As String
    If tileX < 0 Or tileX >= MAP_WIDTH Or tileY < 0 Or tileY >= MAP_HEIGHT Then
        TileAt = "#"
    Else
        TileAt = Mid$(mMap(tileY), tileX + 1, 1)
    End If
End Function

Private Sub InitializeMap()
    mMap(0) = "################"
    mMap(1) = "#.....#....#..X#"
    mMap(2) = "#.....#....#...#"
    mMap(3) = "#..#..#....#...#"
    mMap(4) = "#..#..####.#...#"
    mMap(5) = "#..#.........#.#"
    mMap(6) = "#..#######...#.#"
    mMap(7) = "#...........##.#"
    mMap(8) = "#.######......D#"
    mMap(9) = "#.#....#.......#"
    mMap(10) = "#.#....#.####..#"
    mMap(11) = "#.#....#....#..#"
    mMap(12) = "#...##......#..#"
    mMap(13) = "#.....#######..#"
    mMap(14) = "#..............#"
    mMap(15) = "################"
End Sub

Private Sub InitializeEnemies()
    mEnemies(1).X = 6.5
    mEnemies(1).Y = 2.5
    mEnemies(1).Alive = True

    mEnemies(2).X = 10.5
    mEnemies(2).Y = 7.5
    mEnemies(2).Alive = True

    mEnemies(3).X = 12.5
    mEnemies(3).Y = 11.5
    mEnemies(3).Alive = True

    mEnemies(4).X = 5.5
    mEnemies(4).Y = 13.5
    mEnemies(4).Alive = True
End Sub

Private Function GetGameSheet() As Worksheet
    On Error Resume Next
    Set GetGameSheet = ThisWorkbook.Worksheets("DOOM")
    On Error GoTo 0

    If GetGameSheet Is Nothing Then
        Set GetGameSheet = ThisWorkbook.Worksheets.Add
        GetGameSheet.Name = "DOOM"
    End If
End Function

Private Sub BindKeys()
    Application.OnKey "{UP}", "ExcelDoom_MoveForward"
    Application.OnKey "{DOWN}", "ExcelDoom_MoveBackward"
    Application.OnKey "{LEFT}", "ExcelDoom_TurnLeft"
    Application.OnKey "{RIGHT}", "ExcelDoom_TurnRight"
    Application.OnKey "w", "ExcelDoom_MoveForward"
    Application.OnKey "s", "ExcelDoom_MoveBackward"
    Application.OnKey "a", "ExcelDoom_TurnLeft"
    Application.OnKey "d", "ExcelDoom_TurnRight"
    Application.OnKey "q", "ExcelDoom_StrafeLeft"
    Application.OnKey "e", "ExcelDoom_StrafeRight"
    Application.OnKey " ", "ExcelDoom_Shoot"
    mKeysBound = True
End Sub

Private Sub UnbindKeys()
    If Not mKeysBound Then Exit Sub

    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "w"
    Application.OnKey "s"
    Application.OnKey "a"
    Application.OnKey "d"
    Application.OnKey "q"
    Application.OnKey "e"
    Application.OnKey " "
    mKeysBound = False
End Sub

Private Function NormalizeAngle(ByVal angleValue As Double) As Double
    Do While angleValue < 0#
        angleValue = angleValue + (2# * PI)
    Loop

    Do While angleValue >= 2# * PI
        angleValue = angleValue - (2# * PI)
    Loop

    NormalizeAngle = angleValue
End Function

Private Function NormalizeRelativeAngle(ByVal angleValue As Double) As Double
    Do While angleValue <= -PI
        angleValue = angleValue + (2# * PI)
    Loop

    Do While angleValue > PI
        angleValue = angleValue - (2# * PI)
    Loop

    NormalizeRelativeAngle = angleValue
End Function

Private Function Atan2(ByVal yValue As Double, ByVal xValue As Double) As Double
    If xValue = 0# Then
        If yValue > 0# Then
            Atan2 = PI / 2#
        ElseIf yValue < 0# Then
            Atan2 = -PI / 2#
        Else
            Atan2 = 0#
        End If
    ElseIf xValue > 0# Then
        Atan2 = Atn(yValue / xValue)
    ElseIf yValue >= 0# Then
        Atan2 = Atn(yValue / xValue) + PI
    Else
        Atan2 = Atn(yValue / xValue) - PI
    End If
End Function
