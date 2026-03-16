Attribute VB_Name = "modExcelDoom"
Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const MAP_WIDTH As Long = 18
Private Const MAP_HEIGHT As Long = 18
Private Const VIEW_WIDTH As Long = 120
Private Const VIEW_HEIGHT As Long = 40
Private Const FOV As Double = PI / 3#
Private Const MAX_DEPTH As Double = 24#
Private Const MOVE_STEP As Double = 0.32
Private Const STRAFE_STEP As Double = 0.24
Private Const TURN_STEP As Double = 0.17
Private Const PLAYER_RADIUS As Double = 0.16
Private Const ENEMY_COUNT As Long = 6
Private Const ENEMY_GRUNT As Long = 1
Private Const ENEMY_STALKER As Long = 2
Private Const ENEMY_BRUTE As Long = 3
Private Const UI_SHEET As String = "DOOM"
Private Const SHAPE_VIEWPORT As String = "doom_viewport"
Private Const SHAPE_TITLE As String = "doom_title"
Private Const SHAPE_HUD As String = "doom_hud"
Private Const SHAPE_MAP As String = "doom_map"

Private Type EnemyState
    X As Double
    Y As Double
    Kind As Long
    MaxHealth As Long
    Health As Long
    Speed As Double
    AttackRange As Double
    AttackDamage As Long
    CooldownMax As Long
    Cooldown As Long
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
Private mPaused As Boolean
Private mKeysBound As Boolean
Private mMuzzleFlashTicks As Long
Private mLastMessage As String
Private mTargetIndex As Long
Private mWallDistances(1 To VIEW_WIDTH) As Double

Public Sub ExcelDoom_ConfigureSheet()
    SetupSheet
    ShowIdleScreen
End Sub

Public Sub ExcelDoom_StartGame()
    InitializeMap
    InitializeEnemies
    SetupSheet

    mPlayerX = 2.5
    mPlayerY = 2.5
    mPlayerAngle = 0#
    mAmmo = 36
    mHealth = 100
    mKills = 0
    mStarted = True
    mPaused = False
    mMuzzleFlashTicks = 0
    mLastMessage = "Очисти уровень и не дай демонам подойти вплотную."

    BindKeys
    RenderFrame
End Sub

Public Sub ExcelDoom_ResetGame()
    ExcelDoom_StartGame
End Sub

Public Sub ExcelDoom_StopGame()
    mStarted = False
    mPaused = False
    mMuzzleFlashTicks = 0
    UnbindKeys
    SetupSheet
    ShowIdleScreen
    SetHudText "Остановлено. Нажми START или запусти ExcelDoom_StartGame."
End Sub

Public Sub ExcelDoom_TogglePause()
    If Not mStarted Then
        SetHudText "Игра не запущена. Нажми START."
        Exit Sub
    End If

    mPaused = Not mPaused
    If mPaused Then
        mLastMessage = "Пауза. P продолжает игру, R начинает заново."
    Else
        mLastMessage = "Пауза снята. Демоны снова активны."
    End If

    RenderFrame
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
        mLastMessage = "Патроны закончились. Нажми RESET."
        RenderFrame
        Exit Sub
    End If

    mAmmo = mAmmo - 1
    mMuzzleFlashTicks = 1
    hitIndex = mTargetIndex
    If hitIndex = 0 Then
        hitIndex = FindShootTarget()
    End If

    If hitIndex > 0 Then
        mEnemies(hitIndex).Health = mEnemies(hitIndex).Health - 1
        If mEnemies(hitIndex).Health <= 0 Then
            mEnemies(hitIndex).Alive = False
            mKills = mKills + 1
            mLastMessage = "Демон убит. Осталось: " & CStr(ActiveEnemyCount())
        Else
            mLastMessage = "Попадание. Добей его."
        End If
    Else
        mLastMessage = "Мимо. Держи врага на прицеле."
    End If

    StepSimulation
End Sub

Private Function EnsureRunning() As Boolean
    If mStarted And Not mPaused Then
        EnsureRunning = True
    ElseIf mPaused Then
        SetHudText "Пауза. Нажми P для продолжения или R для рестарта."
    Else
        SetHudText "Игра не запущена. Нажми START или запусти ExcelDoom_StartGame."
    End If
End Function

Private Sub StepSimulation()
    MoveEnemies
    ApplyEnemyAttacks
    If mMuzzleFlashTicks > 0 Then
        mMuzzleFlashTicks = mMuzzleFlashTicks - 1
    End If
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

Private Sub MoveEnemies()
    Dim i As Long
    Dim distanceToPlayer As Double
    Dim dirX As Double
    Dim dirY As Double
    Dim moveX As Double
    Dim moveY As Double
    Dim nextX As Double
    Dim nextY As Double

    For i = 1 To ENEMY_COUNT
        If mEnemies(i).Alive Then
            distanceToPlayer = Distance(mEnemies(i).X, mEnemies(i).Y, mPlayerX, mPlayerY)
            If distanceToPlayer > 1.2 Then
                dirX = (mPlayerX - mEnemies(i).X) / distanceToPlayer
                dirY = (mPlayerY - mEnemies(i).Y) / distanceToPlayer
                moveX = dirX * mEnemies(i).Speed
                moveY = dirY * mEnemies(i).Speed
                nextX = mEnemies(i).X + moveX
                nextY = mEnemies(i).Y + moveY

                If EnemyCanOccupy(i, nextX, mEnemies(i).Y) Then
                    mEnemies(i).X = nextX
                End If

                If EnemyCanOccupy(i, mEnemies(i).X, nextY) Then
                    mEnemies(i).Y = nextY
                End If
            End If

            If mEnemies(i).Cooldown > 0 Then
                mEnemies(i).Cooldown = mEnemies(i).Cooldown - 1
            End If
        End If
    Next i
End Sub

Private Function EnemyCanOccupy(ByVal enemyIndex As Long, ByVal x As Double, ByVal y As Double) As Boolean
    Dim i As Long

    If IsWallAt(x, y) Then Exit Function

    For i = 1 To ENEMY_COUNT
        If i <> enemyIndex Then
            If mEnemies(i).Alive Then
                If Distance(x, y, mEnemies(i).X, mEnemies(i).Y) < 0.45 Then Exit Function
            End If
        End If
    Next i

    EnemyCanOccupy = True
End Function

Private Sub ApplyEnemyAttacks()
    Dim i As Long
    Dim pressure As Long
    Dim distanceToPlayer As Double

    pressure = 0

    For i = 1 To ENEMY_COUNT
        If mEnemies(i).Alive Then
            distanceToPlayer = Distance(mEnemies(i).X, mEnemies(i).Y, mPlayerX, mPlayerY)
            If distanceToPlayer < mEnemies(i).AttackRange And HasLineOfSight(mEnemies(i).X, mEnemies(i).Y) Then
                If mEnemies(i).Cooldown = 0 Then
                    pressure = pressure + mEnemies(i).AttackDamage
                    mEnemies(i).Cooldown = mEnemies(i).CooldownMax
                End If
            End If
        End If
    Next i

    If pressure > 0 Then
        mHealth = mHealth - pressure
        If mHealth < 0 Then mHealth = 0
        mLastMessage = "Тебя обстреливают. Меняй позицию."
    End If

    If mHealth <= 0 Then
        mStarted = False
        UnbindKeys
        mLastMessage = "GAME OVER. Нажми RESET."
    ElseIf ActiveEnemyCount() = 0 Then
        mStarted = False
        UnbindKeys
        mLastMessage = "VICTORY. Все демоны уничтожены."
    End If
End Sub

Private Function FindShootTarget() As Long
    FindShootTarget = FindAimedEnemy(0.24)
End Function

Private Function FindAimedEnemy(ByVal aimWindow As Double) As Long
    Dim i As Long
    Dim enemyDistance As Double
    Dim enemyAngle As Double
    Dim bestDistance As Double

    bestDistance = 1E+30

    For i = 1 To ENEMY_COUNT
        If mEnemies(i).Alive Then
            enemyDistance = Distance(mPlayerX, mPlayerY, mEnemies(i).X, mEnemies(i).Y)
            enemyAngle = NormalizeRelativeAngle(Atan2(mEnemies(i).Y - mPlayerY, mEnemies(i).X - mPlayerX) - mPlayerAngle)

            If Abs(enemyAngle) <= aimWindow And enemyDistance < bestDistance Then
                If HasLineOfSight(mEnemies(i).X, mEnemies(i).Y) Then
                    bestDistance = enemyDistance
                    FindAimedEnemy = i
                End If
            End If
        End If
    Next i
End Function

Private Function HasLineOfSight(ByVal targetX As Double, ByVal targetY As Double) As Boolean
    Dim stepIndex As Long
    Dim steps As Long
    Dim sampleX As Double
    Dim sampleY As Double

    steps = CLng(Distance(mPlayerX, mPlayerY, targetX, targetY) / 0.05)
    If steps < 1 Then steps = 1

    For stepIndex = 1 To steps - 1
        sampleX = mPlayerX + (targetX - mPlayerX) * (stepIndex / steps)
        sampleY = mPlayerY + (targetY - mPlayerY) * (stepIndex / steps)
        If IsWallAt(sampleX, sampleY) Then Exit Function
    Next stepIndex

    HasLineOfSight = True
End Function

Private Sub RenderFrame()
    Dim previousScreenUpdating As Boolean
    Dim lines(1 To VIEW_HEIGHT) As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim distanceToWall As Double
    Dim hitTile As String
    Dim hitSide As Long
    Dim ceilingRow As Long
    Dim floorRow As Long
    Dim wallGlyph As String

    On Error GoTo RenderCleanup
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    mTargetIndex = FindAimedEnemy(0.24)

    For rowIndex = 1 To VIEW_HEIGHT
        lines(rowIndex) = String$(VIEW_WIDTH, " ")
    Next rowIndex

    For colIndex = 1 To VIEW_WIDTH
        CastRay colIndex, distanceToWall, hitTile, hitSide
        mWallDistances(colIndex) = distanceToWall

        ceilingRow = (VIEW_HEIGHT \ 2) - CLng((VIEW_HEIGHT * 0.8) / MaxDouble(0.25, distanceToWall))
        floorRow = VIEW_HEIGHT - ceilingRow

        If ceilingRow < 1 Then ceilingRow = 1
        If floorRow > VIEW_HEIGHT Then floorRow = VIEW_HEIGHT
        wallGlyph = GetWallGlyph(hitTile, distanceToWall, hitSide)

        For rowIndex = 1 To ceilingRow - 1
            Mid$(lines(rowIndex), colIndex, 1) = GetSkyGlyph(rowIndex)
        Next rowIndex

        For rowIndex = ceilingRow To floorRow
            Mid$(lines(rowIndex), colIndex, 1) = wallGlyph
        Next rowIndex

        For rowIndex = floorRow + 1 To VIEW_HEIGHT
            Mid$(lines(rowIndex), colIndex, 1) = GetFloorGlyph(rowIndex, distanceToWall)
        Next rowIndex
    Next colIndex

    RenderEnemiesInto lines
    RenderWeaponInto lines
    RenderCrosshairInto lines

    SetViewportText JoinLines(lines)
    SetMapText BuildMapText()
    SetHudText BuildHudText()
    FocusViewport

RenderCleanup:
    Application.ScreenUpdating = previousScreenUpdating
End Sub

Private Sub RenderEnemiesInto(ByRef lines() As String)
    Dim drawOrder(1 To ENEMY_COUNT) As Long
    Dim visibleCount As Long
    Dim i As Long
    Dim orderIndex As Long
    Dim chosenIndex As Long
    Dim farthestDistance As Double
    Dim currentDistance As Double
    Dim relAngle As Double
    Dim enemyDistance As Double
    Dim screenX As Long
    Dim spriteHeight As Long
    Dim spriteWidth As Long
    Dim topRow As Long
    Dim leftCol As Long
    Dim drawRow As Long
    Dim drawCol As Long
    Dim targetColumn As Long
    Dim glyph As String

    For orderIndex = 1 To ENEMY_COUNT
        farthestDistance = -1#
        chosenIndex = 0

        For i = 1 To ENEMY_COUNT
            If mEnemies(i).Alive Then
                If Not EnemyAlreadyQueued(drawOrder, visibleCount, i) Then
                    currentDistance = Distance(mEnemies(i).X, mEnemies(i).Y, mPlayerX, mPlayerY)
                    If currentDistance > farthestDistance Then
                        farthestDistance = currentDistance
                        chosenIndex = i
                    End If
                End If
            End If
        Next i

        If chosenIndex > 0 Then
            visibleCount = visibleCount + 1
            drawOrder(visibleCount) = chosenIndex
        End If
    Next orderIndex

    For orderIndex = 1 To visibleCount
        i = drawOrder(orderIndex)
        enemyDistance = Distance(mEnemies(i).X, mEnemies(i).Y, mPlayerX, mPlayerY)
        relAngle = NormalizeRelativeAngle(Atan2(mEnemies(i).Y - mPlayerY, mEnemies(i).X - mPlayerX) - mPlayerAngle)

        If Abs(relAngle) <= (FOV / 2#) + 0.12 Then
            screenX = CLng((VIEW_WIDTH / 2#) + ((relAngle / (FOV / 2#)) * (VIEW_WIDTH / 2#)))
            If screenX < 1 Then screenX = 1
            If screenX > VIEW_WIDTH Then screenX = VIEW_WIDTH

            targetColumn = screenX
            If enemyDistance < mWallDistances(targetColumn) Then
                spriteHeight = CLng((VIEW_HEIGHT * 0.9) / MaxDouble(0.5, enemyDistance))
                spriteWidth = spriteHeight \ 2
                If spriteHeight < 2 Then spriteHeight = 2
                If spriteHeight > 16 Then spriteHeight = 16
                If spriteWidth < 1 Then spriteWidth = 1
                If spriteWidth > 7 Then spriteWidth = 7

                topRow = (VIEW_HEIGHT \ 2) - (spriteHeight \ 2)
                leftCol = screenX - (spriteWidth \ 2)
                glyph = GetEnemyGlyph(i)

                For drawRow = 0 To spriteHeight - 1
                    For drawCol = 0 To spriteWidth - 1
                        If topRow + drawRow >= 1 And topRow + drawRow <= VIEW_HEIGHT Then
                            If leftCol + drawCol >= 1 And leftCol + drawCol <= VIEW_WIDTH Then
                                Mid$(lines(topRow + drawRow), leftCol + drawCol, 1) = glyph
                            End If
                        End If
                    Next drawCol
                Next drawRow
            End If
        End If
    Next orderIndex
End Sub

Private Function EnemyAlreadyQueued(ByRef drawOrder() As Long, ByVal queuedCount As Long, ByVal enemyIndex As Long) As Boolean
    Dim i As Long

    For i = 1 To queuedCount
        If drawOrder(i) = enemyIndex Then
            EnemyAlreadyQueued = True
            Exit Function
        End If
    Next i
End Function

Private Sub RenderWeaponInto(ByRef lines() As String)
    Dim baseRow As Long
    Dim baseCol As Long
    Dim flashGlyph As String

    baseRow = VIEW_HEIGHT - 4
    baseCol = (VIEW_WIDTH \ 2) - 4

    Mid$(lines(baseRow), baseCol + 2, 3) = "\|/"
    Mid$(lines(baseRow + 1), baseCol + 1, 5) = "[###]"
    Mid$(lines(baseRow + 2), baseCol + 1, 5) = "/###\"

    If mMuzzleFlashTicks > 0 Then
        flashGlyph = "*"
        Mid$(lines(baseRow - 1), baseCol + 3, 1) = flashGlyph
        Mid$(lines(baseRow), baseCol + 3, 1) = flashGlyph
    End If
End Sub

Private Sub RenderCrosshairInto(ByRef lines() As String)
    Dim centerRow As Long
    Dim centerCol As Long
    Dim crosshairGlyph As String

    centerRow = VIEW_HEIGHT \ 2
    centerCol = VIEW_WIDTH \ 2

    If mTargetIndex > 0 Then
        crosshairGlyph = "X"
    Else
        crosshairGlyph = "+"
    End If

    Mid$(lines(centerRow), centerCol, 1) = crosshairGlyph
End Sub

Private Function JoinLines(ByRef lines() As String) As String
    Dim rowIndex As Long

    For rowIndex = LBound(lines) To UBound(lines)
        If rowIndex = LBound(lines) Then
            JoinLines = lines(rowIndex)
        Else
            JoinLines = JoinLines & vbLf & lines(rowIndex)
        End If
    Next rowIndex
End Function

Private Function BuildMapText() As String
    Dim mapRow As Long
    Dim mapCol As Long
    Dim lineText As String
    Dim tileValue As String
    Dim enemyIndex As Long
    Dim playerMapX As Long
    Dim playerMapY As Long

    BuildMapText = "MAP  Facing " & GetFacingGlyph() & "  Enemies " & CStr(ActiveEnemyCount()) & vbLf
    playerMapX = Int(mPlayerX)
    playerMapY = Int(mPlayerY)

    For mapRow = 0 To MAP_HEIGHT - 1
        lineText = ""

        For mapCol = 0 To MAP_WIDTH - 1
            tileValue = TileAt(mapCol, mapRow)

            If mapCol = playerMapX And mapRow = playerMapY Then
                lineText = lineText & "P"
            ElseIf IsEnemyOnTile(mapCol, mapRow) Then
                lineText = lineText & "E"
            ElseIf tileValue = "." Then
                lineText = lineText & "."
            Else
                lineText = lineText & tileValue
            End If
        Next mapCol

        BuildMapText = BuildMapText & lineText
        If mapRow < MAP_HEIGHT - 1 Then BuildMapText = BuildMapText & vbLf
    Next mapRow

    BuildMapText = BuildMapText & vbLf & vbLf & "Controls" & vbLf
    BuildMapText = BuildMapText & "Up/Down or W/S Move" & vbLf
    BuildMapText = BuildMapText & "Left/Right or A/D Turn" & vbLf
    BuildMapText = BuildMapText & "Shift+Left/Right Strafe" & vbLf
    BuildMapText = BuildMapText & "Space Shoot" & vbLf
    BuildMapText = BuildMapText & "F8 Pause  F5 Reset" & vbLf & vbLf
    BuildMapText = BuildMapText & "g grunt  s stalker  B brute"
End Function

Private Function IsEnemyOnTile(ByVal mapX As Long, ByVal mapY As Long) As Boolean
    Dim enemyIndex As Long

    For enemyIndex = 1 To ENEMY_COUNT
        If mEnemies(enemyIndex).Alive Then
            If Int(mEnemies(enemyIndex).X) = mapX And Int(mEnemies(enemyIndex).Y) = mapY Then
                IsEnemyOnTile = True
                Exit Function
            End If
        End If
    Next enemyIndex
End Function

Private Function BuildHudText() As String
    BuildHudText = "HP " & Format$(mHealth, "000") & "   AMMO " & Format$(mAmmo, "00") & "   KILLS " & mKills & "/" & ENEMY_COUNT & "   " & BuildStateText() & vbLf
    BuildHudText = BuildHudText & BuildTargetText() & vbLf
    BuildHudText = BuildHudText & mLastMessage
End Function

Private Function GetSkyGlyph(ByVal rowIndex As Long) As String
    If rowIndex < VIEW_HEIGHT \ 6 Then
        GetSkyGlyph = "."
    Else
        GetSkyGlyph = " "
    End If
End Function

Private Function GetFloorGlyph(ByVal rowIndex As Long, ByVal wallDistance As Double) As String
    Dim blendValue As Double

    blendValue = (rowIndex - (VIEW_HEIGHT \ 2)) / (VIEW_HEIGHT \ 2)

    If blendValue < 0.18 Then
        GetFloorGlyph = "-"
    ElseIf blendValue < 0.35 Then
        GetFloorGlyph = "="
    ElseIf blendValue < 0.6 Then
        GetFloorGlyph = "+"
    ElseIf wallDistance < 4# Then
        GetFloorGlyph = "#"
    Else
        GetFloorGlyph = "."
    End If
End Function

Private Function GetWallGlyph(ByVal tileValue As String, ByVal distanceToWall As Double, ByVal sideHit As Long) As String
    Dim shadeIndex As Long
    Dim glyphs As String

    If tileValue = "X" Then
        glyphs = "@%#*+=-."
    ElseIf tileValue = "D" Then
        glyphs = "&8#*+=-."
    Else
        glyphs = "##*+=-.."
    End If

    shadeIndex = CLng(distanceToWall * 1.2) + 1 + sideHit
    If shadeIndex < 1 Then shadeIndex = 1
    If shadeIndex > Len(glyphs) Then shadeIndex = Len(glyphs)

    GetWallGlyph = Mid$(glyphs, shadeIndex, 1)
End Function

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

    If outDistance < 0.15 Then outDistance = 0.15
    If outDistance > MAX_DEPTH Then outDistance = MAX_DEPTH
End Sub

Private Function ActiveEnemyCount() As Long
    Dim enemyIndex As Long

    For enemyIndex = 1 To ENEMY_COUNT
        If mEnemies(enemyIndex).Alive Then
            ActiveEnemyCount = ActiveEnemyCount + 1
        End If
    Next enemyIndex
End Function

Private Function Distance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Distance = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
End Function

Private Function MaxDouble(ByVal leftValue As Double, ByVal rightValue As Double) As Double
    If leftValue > rightValue Then
        MaxDouble = leftValue
    Else
        MaxDouble = rightValue
    End If
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
    mMap(0) = "##################"
    mMap(1) = "#....#.......#..X#"
    mMap(2) = "#....#.......#...#"
    mMap(3) = "#....#..###..#...#"
    mMap(4) = "#....#..#....#...#"
    mMap(5) = "#....#..#....#...#"
    mMap(6) = "#....####....#...#"
    mMap(7) = "#...............##"
    mMap(8) = "#.######........D#"
    mMap(9) = "#.#....#....###..#"
    mMap(10) = "#.#....#....#....#"
    mMap(11) = "#.#....#.##.#....#"
    mMap(12) = "#.#....#....#....#"
    mMap(13) = "#...##......#....#"
    mMap(14) = "#.....#######....#"
    mMap(15) = "#................#"
    mMap(16) = "#................#"
    mMap(17) = "##################"
End Sub

Private Sub InitializeEnemies()
    SetEnemy 1, 6.5, 2.5, ENEMY_GRUNT
    SetEnemy 2, 10.5, 6.5, ENEMY_STALKER
    SetEnemy 3, 13.5, 9.5, ENEMY_GRUNT
    SetEnemy 4, 14.5, 13.5, ENEMY_STALKER
    SetEnemy 5, 8.5, 15.5, ENEMY_BRUTE
    SetEnemy 6, 4.5, 14.5, ENEMY_GRUNT
End Sub

Private Sub SetEnemy(ByVal enemyIndex As Long, ByVal posX As Double, ByVal posY As Double, ByVal enemyKind As Long)
    mEnemies(enemyIndex).X = posX
    mEnemies(enemyIndex).Y = posY
    mEnemies(enemyIndex).Kind = enemyKind
    mEnemies(enemyIndex).MaxHealth = GetEnemyMaxHealth(enemyKind)
    mEnemies(enemyIndex).Health = mEnemies(enemyIndex).MaxHealth
    mEnemies(enemyIndex).Speed = GetEnemySpeed(enemyKind)
    mEnemies(enemyIndex).AttackRange = GetEnemyAttackRange(enemyKind)
    mEnemies(enemyIndex).AttackDamage = GetEnemyAttackDamage(enemyKind)
    mEnemies(enemyIndex).CooldownMax = GetEnemyCooldown(enemyKind)
    mEnemies(enemyIndex).Cooldown = 0
    mEnemies(enemyIndex).Alive = True
End Sub

Private Sub SetupSheet()
    Dim ws As Worksheet

    Set ws = GetGameSheet()
    ws.Cells.Clear
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A:Z").ColumnWidth = 2.2
    ws.Rows("1:45").RowHeight = 15

    EnsureTextBox ws, SHAPE_TITLE, 16, 10, 860, 28, "EXCEL DOOM", 18, RGB(255, 214, 168), True, RGB(24, 16, 16)
    EnsureTextBox ws, SHAPE_VIEWPORT, 16, 44, 860, 540, "", 8.5, RGB(255, 236, 210), False, RGB(8, 8, 8)
    EnsureTextBox ws, SHAPE_HUD, 16, 592, 860, 72, "", 11, RGB(255, 214, 168), False, RGB(24, 16, 16)
    EnsureTextBox ws, SHAPE_MAP, 896, 44, 250, 596, "", 9, RGB(226, 226, 226), False, RGB(12, 12, 12)

    EnsureButton ws, "doom_start", 896, 10, 76, 24, "START", "ExcelDoom_StartGame"
    EnsureButton ws, "doom_reset", 980, 10, 76, 24, "RESET", "ExcelDoom_ResetGame"
    EnsureButton ws, "doom_pause", 1064, 10, 76, 24, "PAUSE", "ExcelDoom_TogglePause"
    EnsureButton ws, "doom_stop", 1148, 10, 76, 24, "STOP", "ExcelDoom_StopGame"
End Sub

Private Sub ShowIdleScreen()
    Dim idleText As String

    idleText = "   ______  _____  _____  __  __" & vbLf
    idleText = idleText & "  |  ____|/ ____|/ ____||  \/  |" & vbLf
    idleText = idleText & "  | |__  | |    | |     | \  / |" & vbLf
    idleText = idleText & "  |  __| | |    | |     | |\/| |" & vbLf
    idleText = idleText & "  | |____| |____| |____ | |  | |" & vbLf
    idleText = idleText & "  |______|\_____|\_____||_|  |_|" & vbLf & vbLf
    idleText = idleText & "  Faster ASCII renderer loaded." & vbLf
    idleText = idleText & "  Higher resolution: 120x40." & vbLf
    idleText = idleText & "  Real enemies chase and shoot back." & vbLf
    idleText = idleText & "  P pauses, R restarts, X crosshair means target lock." & vbLf & vbLf
    idleText = idleText & "  Press START or run ExcelDoom_StartGame."

    SetViewportText idleText
    SetMapText "MAP" & vbLf & vbLf & "Up/Down or W/S Move" & vbLf & "Left/Right or A/D Turn" & vbLf & "Shift+Left/Right Strafe" & vbLf & "Space Shoot" & vbLf & "F8 Pause  F5 Reset"
    SetHudText "Готово. Запускай игру."
    FocusViewport
End Sub

Private Function GetGameSheet() As Worksheet
    On Error Resume Next
    Set GetGameSheet = ThisWorkbook.Worksheets(UI_SHEET)
    On Error GoTo 0

    If GetGameSheet Is Nothing Then
        Set GetGameSheet = ThisWorkbook.Worksheets.Add
        GetGameSheet.Name = UI_SHEET
    End If
End Function

Private Sub EnsureTextBox(ByVal ws As Worksheet, ByVal shapeName As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal widthPos As Double, ByVal heightPos As Double, ByVal textValue As String, ByVal fontSize As Double, ByVal fontColor As Long, ByVal isBold As Boolean, ByVal fillColor As Long)
    Dim shp As Shape

    Set shp = GetOrCreateTextBox(ws, shapeName)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = widthPos
        .Height = heightPos
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = RGB(80, 42, 28)
        .Line.Weight = 1.25
        .TextFrame2.TextRange.Text = textValue
        .TextFrame2.TextRange.Font.Name = "Consolas"
        .TextFrame2.TextRange.Font.Size = fontSize
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fontColor
        .TextFrame2.TextRange.Font.Bold = isBold
        .TextFrame2.MarginLeft = 8
        .TextFrame2.MarginRight = 8
        .TextFrame2.MarginTop = 6
        .TextFrame2.MarginBottom = 6
        If shapeName = SHAPE_VIEWPORT Or shapeName = SHAPE_MAP Then
            .TextFrame2.WordWrap = msoFalse
        Else
            .TextFrame2.WordWrap = msoTrue
        End If
        .Placement = xlFreeFloating
    End With
End Sub

Private Sub EnsureButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal widthPos As Double, ByVal heightPos As Double, ByVal labelText As String, ByVal actionName As String)
    Dim shp As Shape

    Set shp = GetOrCreateShape(ws, shapeName, msoShapeRoundedRectangle)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = widthPos
        .Height = heightPos
        .Fill.ForeColor.RGB = RGB(88, 24, 16)
        .Line.ForeColor.RGB = RGB(255, 214, 168)
        .Line.Weight = 1.25
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Name = "Consolas"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 214, 168)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = actionName
        .Placement = xlFreeFloating
    End With
End Sub

Private Function GetOrCreateTextBox(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set GetOrCreateTextBox = ws.Shapes(shapeName)
    On Error GoTo 0

    If GetOrCreateTextBox Is Nothing Then
        Set GetOrCreateTextBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 20, 20)
        GetOrCreateTextBox.Name = shapeName
    End If
End Function

Private Function GetOrCreateShape(ByVal ws As Worksheet, ByVal shapeName As String, ByVal shapeType As MsoAutoShapeType) As Shape
    On Error Resume Next
    Set GetOrCreateShape = ws.Shapes(shapeName)
    On Error GoTo 0

    If GetOrCreateShape Is Nothing Then
        Set GetOrCreateShape = ws.Shapes.AddShape(shapeType, 0, 0, 20, 20)
        GetOrCreateShape.Name = shapeName
    End If
End Function

Private Sub SetViewportText(ByVal textValue As String)
    GetGameSheet.Shapes(SHAPE_VIEWPORT).TextFrame2.TextRange.Text = textValue
End Sub

Private Sub SetHudText(ByVal textValue As String)
    GetGameSheet.Shapes(SHAPE_HUD).TextFrame2.TextRange.Text = textValue
End Sub

Private Sub SetMapText(ByVal textValue As String)
    GetGameSheet.Shapes(SHAPE_MAP).TextFrame2.TextRange.Text = textValue
End Sub

Private Sub FocusViewport()
    On Error Resume Next
    GetGameSheet.Shapes(SHAPE_VIEWPORT).Select
    On Error GoTo 0
End Sub

Private Function BuildStateText() As String
    If Not mStarted Then
        BuildStateText = "STATE STOP"
    ElseIf mPaused Then
        BuildStateText = "STATE PAUSE"
    ElseIf mHealth <= 25 Then
        BuildStateText = "STATE CRITICAL"
    Else
        BuildStateText = "STATE COMBAT"
    End If
End Function

Private Function BuildTargetText() As String
    If mTargetIndex > 0 Then
        BuildTargetText = "TARGET LOCK " & GetEnemyName(mEnemies(mTargetIndex).Kind) & "  HP " & mEnemies(mTargetIndex).Health & "/" & mEnemies(mTargetIndex).MaxHealth & "  PRESS SPACE"
    ElseIf mPaused Then
        BuildTargetText = "PAUSED  Press P to resume or R to restart"
    Else
        BuildTargetText = "TARGET NONE  Keep enemy near center. X means target lock."
    End If
End Function

Private Function GetFacingGlyph() As String
    Dim facing As Double

    facing = NormalizeAngle(mPlayerAngle)

    If facing < PI / 4# Or facing >= (7# * PI / 4#) Then
        GetFacingGlyph = ">"
    ElseIf facing < (3# * PI / 4#) Then
        GetFacingGlyph = "v"
    ElseIf facing < (5# * PI / 4#) Then
        GetFacingGlyph = "<"
    Else
        GetFacingGlyph = "^"
    End If
End Function

Private Function GetEnemyGlyph(ByVal enemyIndex As Long) As String
    Select Case mEnemies(enemyIndex).Kind
        Case ENEMY_STALKER
            If mEnemies(enemyIndex).Health = 1 Then
                GetEnemyGlyph = "s"
            Else
                GetEnemyGlyph = "s"
            End If
        Case ENEMY_BRUTE
            If mEnemies(enemyIndex).Health <= 2 Then
                GetEnemyGlyph = "b"
            Else
                GetEnemyGlyph = "B"
            End If
        Case Else
            If mEnemies(enemyIndex).Health = 1 Then
                GetEnemyGlyph = "g"
            Else
                GetEnemyGlyph = "g"
            End If
    End Select

    If enemyIndex = mTargetIndex Then
        GetEnemyGlyph = UCase$(GetEnemyGlyph)
    End If
End Function

Private Function GetEnemyName(ByVal enemyKind As Long) As String
    Select Case enemyKind
        Case ENEMY_STALKER
            GetEnemyName = "STALKER"
        Case ENEMY_BRUTE
            GetEnemyName = "BRUTE"
        Case Else
            GetEnemyName = "GRUNT"
    End Select
End Function

Private Function GetEnemyMaxHealth(ByVal enemyKind As Long) As Long
    Select Case enemyKind
        Case ENEMY_STALKER
            GetEnemyMaxHealth = 2
        Case ENEMY_BRUTE
            GetEnemyMaxHealth = 4
        Case Else
            GetEnemyMaxHealth = 2
    End Select
End Function

Private Function GetEnemySpeed(ByVal enemyKind As Long) As Double
    Select Case enemyKind
        Case ENEMY_STALKER
            GetEnemySpeed = 0.16
        Case ENEMY_BRUTE
            GetEnemySpeed = 0.08
        Case Else
            GetEnemySpeed = 0.12
    End Select
End Function

Private Function GetEnemyAttackRange(ByVal enemyKind As Long) As Double
    Select Case enemyKind
        Case ENEMY_STALKER
            GetEnemyAttackRange = 4.8
        Case ENEMY_BRUTE
            GetEnemyAttackRange = 2.2
        Case Else
            GetEnemyAttackRange = 6#
    End Select
End Function

Private Function GetEnemyAttackDamage(ByVal enemyKind As Long) As Long
    Select Case enemyKind
        Case ENEMY_STALKER
            GetEnemyAttackDamage = 7
        Case ENEMY_BRUTE
            GetEnemyAttackDamage = 14
        Case Else
            GetEnemyAttackDamage = 6
    End Select
End Function

Private Function GetEnemyCooldown(ByVal enemyKind As Long) As Long
    Select Case enemyKind
        Case ENEMY_STALKER
            GetEnemyCooldown = 1
        Case ENEMY_BRUTE
            GetEnemyCooldown = 3
        Case Else
            GetEnemyCooldown = 2
    End Select
End Function

Private Sub BindKeys()
    Application.OnKey "{UP}", "ExcelDoom_MoveForward"
    Application.OnKey "{DOWN}", "ExcelDoom_MoveBackward"
    Application.OnKey "{LEFT}", "ExcelDoom_TurnLeft"
    Application.OnKey "{RIGHT}", "ExcelDoom_TurnRight"
    Application.OnKey "+{LEFT}", "ExcelDoom_StrafeLeft"
    Application.OnKey "+{RIGHT}", "ExcelDoom_StrafeRight"
    Application.OnKey "{F5}", "ExcelDoom_ResetGame"
    Application.OnKey "{F8}", "ExcelDoom_TogglePause"
    Application.OnKey "w", "ExcelDoom_MoveForward"
    Application.OnKey "s", "ExcelDoom_MoveBackward"
    Application.OnKey "a", "ExcelDoom_TurnLeft"
    Application.OnKey "d", "ExcelDoom_TurnRight"
    Application.OnKey "q", "ExcelDoom_StrafeLeft"
    Application.OnKey "e", "ExcelDoom_StrafeRight"
    Application.OnKey "p", "ExcelDoom_TogglePause"
    Application.OnKey "r", "ExcelDoom_ResetGame"
    Application.OnKey " ", "ExcelDoom_Shoot"
    mKeysBound = True
End Sub

Private Sub UnbindKeys()
    If Not mKeysBound Then Exit Sub

    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "+{LEFT}"
    Application.OnKey "+{RIGHT}"
    Application.OnKey "{F5}"
    Application.OnKey "{F8}"
    Application.OnKey "w"
    Application.OnKey "s"
    Application.OnKey "a"
    Application.OnKey "d"
    Application.OnKey "q"
    Application.OnKey "e"
    Application.OnKey "p"
    Application.OnKey "r"
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
