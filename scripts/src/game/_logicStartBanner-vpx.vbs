'***********************************************************************************************************************
'*****     GAME LOGIC START                                                 	                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************

Const AllLightsOnMode = False
Dim canAddPlayers : canAddPlayers = True
Dim currentPlayer : currentPlayer = Null
Dim PlungerDevice
Dim gameStarted : gameStarted = False
Dim pinEvents : Set pinEvents = CreateObject("Scripting.Dictionary")
Dim pinEventsOrder : Set pinEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerEvents : Set playerEvents = CreateObject("Scripting.Dictionary")
Dim playerEventsOrder : Set playerEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")

Dim bcpController
Dim useBCP : useBCP = True
Public Sub ConnectToBCPMediaController
    Set bcpController = (new VpxBcpController)(5050, Null)
End Sub