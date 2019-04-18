Dim msg, sapi

msg = InputBox("Type a text below to speak it!!!", "BOOST YOUR TEXT!!! xD")
Set sapi = CreateObject("SAPI.SPVoice")
sapi.speak(msg)