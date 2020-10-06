from playsound import playsound as p
import playsound
import speech_recognition as sr
import win32com.client as w
import cv2
import webbrowser
from tkinter import *
def speak(x):
    voice = w.Dispatch("SAPI.SpVoice")
    voice.Voice = voice.GetVoices().Item(2)
    voice.Speak(x)
class Gui:
    def __init__(self):
        self.root = Tk()
        self.root.geometry("200x200")
        self.root.title("Search Engine")
        self.text1 = Entry(self.root, width=200)
        self.text1.pack()
        self.button1 = Button(self.root, text="Go!", command=self.decision, bg="green", fg="white")
        self.button1.pack()
        self.root.mainloop()
        self.root1 = None
        self.text2 = None
        self.button2 = None

    def decision(self):
        if self.text1.get() == str.lower("search in google"):
            self.google()
        elif self.text1.get() == str.lower("search in bing"):
            self.bing()

    def google(self):
        self.root.destroy()
        self.root1 = Tk()
        self.root1.geometry("200x200")
        self.root1.title("Search in Google")
        self.text2 = Entry(self.root1, width=200)
        self.text2.pack()
        self.button2 = Button(self.root1, text="Search", command=self.searchgoogle, bg="green", fg="white")
        self.button2.pack()
        self.root1.mainloop()

    def bing(self):
        self.root.destroy()
        self.root1 = Tk()
        self.root1.geometry("200x200")
        self.root1.title("Search in Bing")
        self.text2 = Entry(self.root1, width=200)
        self.text2.pack()
        self.button2 = Button(self.root1, text="Search", command=self.searchbing, bg="green", fg="white")
        self.button2.pack()
        self.root1.mainloop()

    def searchgoogle(self):
        webbrowser.get().open("https://www.google.co.in/search?q=" + self.text2.get())

    def searchbing(self):
        webbrowser.get().open("https://www.bing.com/search?q=" + self.text2.get())
tos = input("Do you want to enter text commands or tell me commands(text/speech): ")
if tos == "text":
    while True:
        try:
            a = str.lower(input(">>> "))
            if a == "play song one" or a == "play song 1":
                speak("Playing Amazing grace")
                p("E:\Songs\Amazing grace.mp3")
            elif a == "play song two" or a == "play song 2":
                speak("Playing Beethoven virus")
                p("E:\Songs\Beethoven virus.mp3")
            elif a == "play song three" or a == "play song 3":
                speak("Playing Canon")
                p("E:\Songs\Canon.mp3")
            elif a == "play song four" or a == "play song 4":
                speak("Playing Flight of the bumblebee")
                p("E:\Songs\Flight of the bumblebee.mp3")
            elif a == "play song 5" or a == "play song five":
                speak("Playing Fur elise")
                p("E:\Songs\Fur elise.mp3")
            elif a == "play song 6" or a == "play song six":
                speak("Playing Greensleeves")
                p("E:\Songs\Greensleeves.mp3")
            elif a == "play song 7" or a == "play song seven":
                speak("Playing Heroic Polonaise")
                p("E:\Songs\Heroic Polonaise.mp3")
            elif a == "play song 8" or a == "play song eight":
                speak("Playing Little Prelude")
                p("E:\Songs\Little prelude.mp3")
            elif a == "play song 9" or a == "play song nine":
                speak("Playing Moonlight sonata 1st movement")
                p("E:\Songs\Moonlight sonata 1st movement.mp3")
            elif a == "play song 10" or a == "play song ten":
                speak("Playing Moonlight sonata 2nd movement")
                p("E:\Songs\Moonlight sonata 2nd movement.mp3")
            elif a == "play song 11" or a == "play song eleven":
                speak("Playing Moonlight sonata 3rd movement")
                p("E:\Songs\Moonlight sonata 3rd movement.mp3")
            elif a == "play song 12" or a == "play song twelve":
                speak("Playing Ode to Joy")
                p("E:\Songs\Ode to joy.mp3")
            elif a == "play song 13" or a == "play song thirteen":
                speak("Playing Petit chien")
                p("E:\Songs\Petit chien.mp3")
            elif a == "play song 14" or a == "play song fourteen":
                speak("Playing Symphony number 5")
                p("E:\Songs\Symphony no 5.mp3")
            elif a == "play song 15" or a == "play song fifteen":
                speak("Playing Turkish march")
                p("E:\Songs\Turkish march.mp3")
            elif a == "play song 16" or a == "play song sixteen":
                speak("Playing Waldstein")
                p("E:\Songs\Waldstein.mp3")
            elif a == "play song 17" or a == "play song seventeen":
                speak("Playing Winter wind")
                p("E:\Songs\Winter wind.mp3")
            elif a == "show video":
                speak("Showing video")
                vid = cv2.VideoCapture(0)
                speak("Press q to stop video")
                while True:
                    ret, frame = vid.read()
                    cv2.imshow('This is your video', frame)
                    if cv2.waitKey(1) & 0xFF == ord('q'):
                        vid.release()
                        cv2.destroyAllWindows()
                        speak("Exiting")
                        break
            elif a == "speak":
                speak("What do you want me to speak")
                b = input("... ")
                speak(b)
            elif a == "search in bing":
                speak("What do you want to search")
                b = input("... ")
                webbrowser.get().open("https://www.bing.com/search?q=" + b)
            elif a == "find location":
                speak("I will show you the location using google maps")
            elif a == "search in google":
                speak("What do you want to search")
                b = input("... ")
                webbrowser.get().open("https://www.google.co.in/search?q=" + b)
            elif a == "who are you":
                speak("Hello, I am a Virtual Bot, Please tell me a command from the following below")
                print("""
play song <number> : This will play a song
show video : This will show your video
speak : You can tell me what you want me to speak
search in Bing : This will search your search item in Bing
find location : This will tell your location (Under development)
search in Google : This will search your search item in Google
gui search engine : This will make you search in a gui version of this program
sva version : This will tell you the version of the program
exit : This will exit the program""")
            elif a == "how are you":
                speak("I am fine")
            elif a == "gui search engine":
                n = Gui()
            elif a == "sva version":
                print("Version:- " + "0.3v45.6")
            elif a == "exit":
                speak("Exiting the program")
                speak("Thank you for using my program")
                break
                quit()
            elif a == "go to main menu":
                speak("Sorry I have to exit the program because this command is still in development")
                break
                quit()
            else:
                print("Please tell me a valid command")
        except playsound.PlaysoundException:
            pass

elif tos == "speech":
    r = sr.Recognizer()
    r.energy_threshold = 10000
    while True:
        with sr.Microphone() as source:
            try:
                speak("Sorry this is still under development so please tell a command in american english")
                speak("Please tell me a command from the below commands: ")
                print("""
play song <number> : This will play a song
show video : This will show your video
speak : You can tell me what you want me to speak
find location : This will tell your location (Under development)
search in Google : This will search your search item in Google
gui runner : This will make you search in a gui version of this program
sva version : This will tell you the version of the program
exit : This will exit the program""")
                audio = r.listen(source)
                data = str.lower(r.recognize_google(audio))
                print("You spoke:- " + data)
                if "play song One" in data:
                    speak("Playing Amazing grace")
                    p("E:\Songs\Amazing grace.mp3")
                elif str.lower("play song Two") in data:
                    speak("Playing Beethoven virus")
                    p("E:\Songs\Beethoven virus.mp3")
                elif str.lower("play song Three") in data:
                    speak("Playing Canon")
                    p("E:\Songs\Canon.mp3")
                elif str.lower("play song Four") in data:
                    speak("Playing Flight of the bumblebee")
                    p("E:\Songs\Flight of the bumblebee.mp3")
                elif str.lower("play song Five") in data:
                    speak("Playing Fur elise")
                    p("E:\Songs\Fur elise.mp3")
                elif str.lower("play song Six") in data:
                    speak("Playing Greensleeves")
                    p("E:\Songs\Greensleeves.mp3")
                elif str.lower("play song seven") in data:
                    speak("Playing Heroic Polonaise")
                    p("E:\Songs\Heroic Polonaise.mp3")
                elif str.lower("play song eight") in data:
                    speak("Playing Little Prelude")
                    p("E:\Songs\Little prelude.mp3")
                elif str.lower("play song 9") in data:
                    speak("Playing Moonlight sonata 1st movement")
                    p("E:\Songs\Moonlight sonata 1st movement.mp3")
                elif str.lower("play song 10") in data:
                    speak("Playing Moonlight sonata 2nd movement")
                    p("E:\Songs\Moonlight sonata 2nd movement.mp3")
                elif str.lower("play song eleven") in data:
                    speak("Playing Moonlight sonata 3rd movement")
                    p("E:\Songs\Moonlight sonata 3rd movement.mp3")
                elif str.lower("play song 12") in data:
                    speak("Playing Ode to Joy")
                    p("E:\Songs\Ode to joy.mp3")
                elif str.lower("play song 13") in data:
                    speak("Playing Petit chien")
                    p("E:\Songs\Petit chien.mp3")
                elif str.lower("play song 14") in data:
                    speak("Playing Symphony number 5")
                    p("E:\Songs\Symphony no 5.mp3")
                elif str.lower("play song 15") in data:
                    speak("Playing Turkish march")
                    p("E:\Songs\Turkish march.mp3")
                elif str.lower("play song 16") in data:
                    speak("Playing Waldstein")
                    p("E:\Songs\Waldstein.mp3")
                elif str.lower("play song 17") in data:
                    speak("Playing Winter wind")
                    p("E:\Songs\Winter wind.mp3")
                elif str.lower("show video") in data:
                    speak("Showing video")
                    vid = cv2.VideoCapture(0)
                    speak("Press q to stop video")
                    while True:
                        ret, frame = vid.read()
                        cv2.imshow('This is your video', frame)
                        if cv2.waitKey(1) & 0xFF == ord('q'):
                            vid.release()
                            cv2.destroyAllWindows()
                            speak("Exiting")
                            break
                elif str.lower("speak") in data:
                    r1 = sr.Recognizer()
                    r1.energy_threshold = 10000
                    with sr.Microphone() as source1:
                        speak("What do you want me to speak")
                        audio1 = r1.listen(source1)
                        data1 = str.lower(r1.recognize_google(audio1))
                        print("You spoke:- " + data1)
                        speak(data1)
                elif str.lower("search in Bing") in data:
                    r1 = sr.Recognizer()
                    r1.energy_threshold = 10000
                    with sr.Microphone() as source1:
                        speak("What do you want to search for?: ")
                        audio1 = r1.listen(source1)
                        data1 = str.lower(r1.recognize_google(audio1))
                    print("You spoke:- " + data1)
                    webbrowser.get().open("https://www.bing.com/search?q=" + data1)
                elif str.lower("find location") in data:
                    speak("I will show you the location using google maps")
                elif str.lower("search in Google") in data:
                    r1 = sr.Recognizer()
                    r1.energy_threshold = 10000
                    with sr.Microphone() as source1:
                        speak("What do you want to search for?: ")
                        audio1 = r1.listen(source1)
                        data1 = r1.recognize_google(audio1)
                        print("You spoke:- " + data1)
                        webbrowser.get().open("https://www.google.co.in/search?q=" + data1)
                elif str.lower("who are you") in data:
                    speak("Hello, I am a Virtual Bot, Please tell me a command from the following below")
                    print("""
play song <number> : This will play a song
show video : This will show your video
speak : You can tell me what you want me to speak
search in Bing : This will search your search item in Bing
find location : This will tell your location (Under development)
search in Google : This will search your search item in Google
gui search engine : This will make you search in a gui version of this program
sva version : This will tell you the version of the program
exit : This will exit the program""")
                elif str.lower("how are you") in data:
                    speak("I am fine")
                elif str.lower("gui search engine") in data:
                    n = Gui()
                elif str.lower("SVA version") in data:
                    print("Version:- " + "0.3v45.6")
                elif str.lower("exit") in data:
                    speak("Exiting the program")
                    speak("Thank you for using my program")
                    break
                    quit()
                else:
                    print("Please tell me a valid command")
            except sr.UnknownValueError as e:
                print("Sorry could not understand please tell me again")
            except FileNotFoundError as e:
                raise FileNotFoundError
                print("I think you forgot to put the pen drive inside the USB port")
            except playsound.PlaysoundException as e:
                print("you did not put the pen drive inside the USB port")
else:
    print("Please enter a valid command (text/speech), exit the program and enter it again")
