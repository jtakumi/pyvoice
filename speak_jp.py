import win32com.client

voice=win32com.client.Dispatch("SAPI.SpVoice")
cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
cat.SetID(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_oneCore\Voices",False)
v=[i for i in cat.EnumerateTokens() if i.GetAttribute('Name')=='Microsoft Ayumi']
if v:
    oldv=voice.Voice
    voice.Voice=v[0]
    voice.Speak('こんにちは、皆さん')
    voice.Voice=oldv