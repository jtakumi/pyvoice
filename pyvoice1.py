import win32com.client as wincl


voice = wincl.Dispatch("SAPI.SpVoice")

#文字列をwindowsスピーカーが発音
voice.Speak("hello, I'm learning programming.")