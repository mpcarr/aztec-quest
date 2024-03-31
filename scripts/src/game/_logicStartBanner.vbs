'***********************************************************************************************************************
'*****     GAME LOGIC START                                                 	                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************


Dim canAddPlayers : canAddPlayers = True
Dim currentPlayer : currentPlayer = Null
Dim autoPlunge : autoPlunge = False
Dim ballSaver : ballSaver = False
Dim gameStarted : gameStarted = False
Dim pinEvents : Set pinEvents = CreateObject("Scripting.Dictionary")
Dim playerEvents : Set playerEvents = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")

