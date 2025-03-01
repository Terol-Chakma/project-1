import speech_recognition as sr
import webbrowser
import datetime
import win32com.client
import pyjokes



speaker = win32com.client.Dispatch('SAPI.SpVoice')
while 1:
    s='Hello im orion, your personal desktop assistant, how can i help you? '
    print("Hello I'm ORION")
    speaker.speak(s)
    break


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        a='ok'
        try:
            print("Recognizing...")
            query_in = r.recognize_google(audio, language="en-US")
            print(f"User said: {query_in}")
            return query_in
        except Exception as e:
            return "sorry didnt get that!!!"

if __name__ == '__main__':

    while True:bv 
        print("Listening...")
        text = takeCommand()
        if text!='sorry didnt get that!!!':
            speaker.speak(f"i think you said{text}")
        else:
            speaker.speak(text)

        sites=[["YouTube","https://www.youtube.com"],["wikipedia","https://www.wikipedia.org/"]
            ,["google","https://www.google.com"],["open AI","https://openai.com/"]]
        for site in sites:
            if f"open {site[0]}".lower() in text.lower():
                speaker.speak(f"ok Opening {site[0]} sir...")
                webbrowser.open(site[1])
        greets={
        "Hi":"hi sir do you need anything?",
        "hello":"hello sir do you need anything?",
        "good morning":"good morning sir have a nice day",
        "good night":"good night sir have a good sleep",
        "your name":"I mentioned earlier,that my name is orion sir",
        "about yourself" :"i am glad that you asked about myself,my name is orion sir,i am a desktop assistant and Srinjay Dutta is my inventor,let me know if you need anything else sir",
         "your birthday" : "i don't have a birthday,   but yeah ,   srinjay dutta created me on 25th july ,     in 2024",
                }
        for keys in greets:
            if keys.lower() in text.lower():
                speaker.speak(greets.get(keys))


        apps = [["camera", "C:\\Users\\SRINJAY DUTTA\\Desktop\\Camera - Shortcut.lnk"], ["chrome", "C:\\Users\\SRINJAY DUTTA\\Desktop\\Srinjay - Chrome.lnk"]
            , ["instagram", "C:\\Users\\SRINJAY DUTTA\\Desktop\\Instagram - Shortcut.lnk"], ["whatsapp", "C:\\Users\\SRINJAY DUTTA\\Desktop\\WhatsApp - Shortcut.lnk"], ["notepad", "C:\\Users\\SRINJAY DUTTA\\Desktop\\Notepad - Shortcut.lnk"],["music","C:\\Users\\SRINJAY DUTTA\\Desktop\\Media Player - Shortcut.lnk"]]
        for app in apps:
            if f"open {app[0]}".lower() in text.lower():
                speaker.speak(f"ok Opening {app[0]} sir...")
                webbrowser.open(app[1])


        if "the time" in text:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.speak(f"currently the time is{strfTime}")
        if "joke" in text:
            joke = pyjokes.get_joke()
            speaker.speak("ok here is a joke for you sir")
            print(joke)
            speaker.speak(joke)
        if "orion stop".lower() in text.lower():
            speaker.speak("Bye,let me know if you need anything else sir...")
            exit()
