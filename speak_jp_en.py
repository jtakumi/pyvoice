#未完成
import win32com.client

voice=win32com.client.Dispatch("SAPI.SpVoice")
cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
cat.SetID(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_oneCore\Voices",False)
v=[i for i in cat.EnumerateTokens()]

with open('voice_jp_en.csv','w',encoding='utf-8-sig') as fc:
    for j in cat.EnumerateTokens():
            if(v[j].GetAttribute('Language')=='English'):
                voice.Voice=v[j]
                voice.speak('hello world')
                print(v[j].GetAttribute('Name'),file=fc)
            elif(v[j].Getattribute('Language')=='Japanese'):
                voice.Voice=v[j]
                voice.speak('こんにちは、世界')
                print(v[j].GetAttribute('Name'),file=fc)