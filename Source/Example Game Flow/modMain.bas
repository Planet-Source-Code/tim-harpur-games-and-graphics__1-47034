Attribute VB_Name = "modMain"
Option Base 0

'declare public variables in this module, such as timeControl, gameState, etc.
'                         :
'                         :
'                         :
'                         :

Public timeControl As Long, gameState As Long, gameSubState As Long

'Timing routine in milliseconds
'this routine is actually from the DXDraw library but I only needed this one routine so instead
'of including the whole file I just copied the routine
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Function DelayTillTime(returnTime As Long, Optional maxCarryOver As Long = 0, Optional ByVal useRelativeTime As Boolean = False)
  Dim CarryOver As Long
  
  If useRelativeTime Then returnTime = timeGetTime() + returnTime
  
  Do
    DoEvents 'while idle let the system process other tasks - call at least once
    
    DelayTillTime = timeGetTime()
  Loop While DelayTillTime < returnTime
  
  CarryOver = DelayTillTime - returnTime
  If CarryOver > maxCarryOver Then CarryOver = maxCarryOver
  
  If CarryOver > 0 Then DelayTillTime = DelayTillTime - CarryOver
End Function

'this is the game start routine
Sub main()
  'show a splash screen here if you wish and possibly load up game settings
  '                         :
  '                         :
  '                         :
  
  'initialize gameState, timeControl and show the primary form for the program
  gameState = 0
  frmMain.Show
  timeControl = DelayTillTime(0, 0, True) + 16
  
  'this is the central timing and game state controller (shown in orange on the form)
  Do While gameState >= 0
    'time each pass through the game - I've choosen to use 16 milliseconds
    timeControl = DelayTillTime(timeControl, 16, False) + 16
    
    'call routine to check for certain user input such as pause/new game/load game/quit/etc.
    'this routine should know to not do anything until the game state is at or above 2
    checkInput
    
    'this is the central game state dispatcher
    Select Case gameState
      Case 0  'still waiting for the form to load - don't do anything - this case doesn't actually need
                    'to be here
                    
      Case 1  'form has loaded so we can complete initialization of display, input and sound/music
                    'this is done here inside the main game loop so that the user may change game
                    'settings without requiring a restart of the program
                    
                    frmMain.gameInitialize
      
      Case 2  'main game splash and options - I'm going to tap into this location to run the demo
                    goDemo
                    '                         :
                    '                         :
                    '                         :
      
      '     :
      '     :
      '     :
      
      'other game state cases such as settings, level initialization, main game loop, etc.
      
      '     :
      '     :
      '     :
      
      Case n
                    '                         :
                    '                         :
                    '                         :
    End Select
  Loop
  
  'do any cleanup required and/or show a close down screen
  DXInput.CleanUp_DXInput
  '                         :
  '                         :
  '                         :
  
  'signal that everything is done by setting gameState and unload all loaded forms, etc.
  gameState = -2
  
  Unload frmMain
End Sub

Private Sub checkInput()
  'input has not yet been initialized
  If gameState < 2 Then Exit Sub
  
  'check for user input... poll the keyboard, mouse or controller and test for specific keys/buttons
  'and make nescessary calls - the results of which may change the gameState
  DXInput.Poll_Keyboard
  
  'check if user pressed  the escape key
  If DXInput.dx_KeyboardState.Key(DIK_ESCAPE) Then
    'user has pressed Esc
    'verify intent and if yes - signal that the game needs to cleanup by setting the gameState
    If MsgBox("Do you really wish to quit?", vbYesNo, "Game Flow Demo Quit") = vbYes Then gameState = -1
  End If
  
  '                         :
  '                         :
  '                         :
  
End Sub

'this routine is in here to handle the demo
Private Sub goDemo()
  Static gameSubStateCounter As Long
  
  'this is a game substate controller
  If gameSubState = -1 Then
    enableMessage 0
    
    gameSubState = 0
  Else
    If gameSubStateCounter Mod 30 = 0 Then
      If frmMain.shpFlow(gameSubState).BorderColor = 0 Then
        frmMain.shpFlow(gameSubState).BorderColor = &HFFFFFF
      Else
        frmMain.shpFlow(gameSubState).BorderColor = 0
      End If
    End If
    
    gameSubStateCounter = gameSubStateCounter + 1
    
    If gameSubStateCounter >= 900 Then
      gameSubStateCounter = 0
      gameSubState = gameSubState + 1
      If gameSubState > 7 Then gameSubState = 0
      
      enableMessage gameSubState
    End If
  End If
End Sub

Private Sub enableMessage(ByVal mNumber As Long)
  Dim loop1 As Long
  
  For loop1 = 0 To 7
    If loop1 = mNumber Then
      frmMain.lblMessage(loop1).Visible = True
      frmMain.shpFlow(loop1).BorderColor = &HFFFFFF
    Else
      frmMain.lblMessage(loop1).Visible = False
      frmMain.shpFlow(loop1).BorderColor = 0
    End If
  Next loop1
End Sub
