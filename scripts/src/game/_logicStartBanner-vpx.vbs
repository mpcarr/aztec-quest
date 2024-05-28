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

Dim bcpController : bcpController = Null
Dim useBCP : useBCP = False
Public Sub ConnectToBCPMediaController
    Set bcpController = (new VpxBcpController)(5050, "aztecquest-mc.exe")
End Sub