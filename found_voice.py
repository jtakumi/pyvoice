import win32com.client

cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
cat.SetID(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_oneCore\Voices",False)

with open('voice_found.csv','w',encoding='utf-8-sig') as fc:
    for i in cat.EnumerateTokens():
            print(i.GetDescription(),file=fc)