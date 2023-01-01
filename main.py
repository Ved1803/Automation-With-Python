
# Importing some useful packages
from win32com.client import Dispatch
import speech_recognition as sr
import webbrowser
import wikipedia
import pywhatkit 
import requests
import datetime
import random
import json
import time
import smtplib
import os

# Basic structure of Jarvis
class jarvis:
    def __init__(self):
        # Getting the current hour
        hour = self.hour = int(datetime.datetime.now().hour)
        currTime = self.currTime = datetime.datetime.now().strftime('%H:%M')

    def Send_email(self, to, mssg):
        userName = "ved.tiwari982@gmail.com"
        passcode="Ved@18192"
        try:
            smtpObj = smtplib.SMTP('smtp.gmail.com', 587) #587 port running
            smtpObj.starttls()
            smtpObj.login(user=userName, password=passcode)
            smtpObj.sendmail(userName, to, mssg)
            print("Successfully sent email")

        except SMTPException:
            print("Error: unable to send email")

    def wish_me(self):
        '''This is a function to wish sommone'''
        if self.hour >= 0 and self.hour < 12:
            self.speak_jarvis(f'Good Morning sir! Its {self.currTime}')
        elif self.hour >= 12 and self.hour < 16:
            self.speak_jarvis(f'Good Afternoon sir! Its {self.currTime}')
        elif self.hour >= 16 and self.hour < 18:
            self.speak_jarvis(f'Good Evening sir! Its {self.currTime}')
        elif self.hour >= 18 and self.hour < 24:
            self.speak_jarvis(f'Good Night sir! Its {self.currTime}')

    def speak_jarvis(self, audio):
        '''This is function to convert text into audioData '''
        speak = Dispatch('SAPI.Spvoice')
        speak.speak(audio)

    # Converts audioData into text
    def Take_command(self):
        '''This is function take microphone input and returns a string '''
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening....")
            r.pause_threshold = 1
            audio = r.listen(source)

        try:
            print("Recognizing....")
            #querry = r.recognize_google(audio, language='en-IN')
            querry= r.recognize(audio)
            print(querry)
            print(f"User said : {querry} \n")
            return querry # returning string 

        except Exception as e:
            print(e)
            print("Unable to recognize! Please say that again.....")
            return 'None'

    def calling_weather(self):
        '''Fetching current weather condition of city'''
        self.speak_jarvis("weather ")
        city = "indore"
        api= "f913b388391d76922ad68e74ac1bc884"
        url = f'http://api.openweathermap.org/data/2.5/weather?q={city} &appid={api}' ##using api
        json_data = requests.get(url)
        json_data1 = json_data.text
        data = json.loads(json_data1)
       #  ori_data = data['weather'][0]['description']

        
        temp10 = data['main']['temp']
        temp = int(temp10 - 273)
        temp = str(temp)
        self.speak_jarvis(f'sir! The weather outside is {temp} degree celcius. It will be  outside')
        #self.speak_jarvis(f'sir! The weather outside is {temp} degree celcius. It will be {ori_data} outside')

    def Executing_task(self):
        '''Logic for executing the tasks based on querry'''
        querry = self.Take_command().lower()
        print(f"querry---------------{querry}")

        if 'wikipedia'  in querry:
            self.speak_jarvis('Searching Wikipedia.....')
            querry = querry.replace('wikipedia is', "")
            # Setting language hindi
            # wikipedia.set_lang('hi')
            result = wikipedia.summary(querry, sentences=2)
            self.speak_jarvis('According to Wikipedia')
            print(result)
            self.speak_jarvis(result)

        elif 'sms' in querry:
            querry = querry.replace('sms to', "")
            self.sms(querry.lower())
        elif 'message' in querry:
            pywhatkit.sendwhatmsg("+919424046609", "hello world", 20,7)

        elif 'shutdown' in querry:
            pywhatkit.shutdown(time=20)

        elif 'wait' in querry:
            pywhatkit.cancelShutdown()
        elif 'pagal' in querry:
            self.speak_jarvis('Sorry sir, but iam not mad. I did not hear you clearly')

        elif 'search' in querry:
            self.speak_jarvis('Searching your results...')
            querry = querry.replace('search', "")
            pywhatkit.search(querry)

        elif 'play' in querry:
            self.speak_jarvis('searching your results...')
            querry = querry.replace('play', "")
            pywhatkit.playonyt(querry)

        elif 'work' in querry or 'task' in querry:
            with open('task.txt', 'r') as re:
                task1 = re.read()
                if task1 == "":
                    self.speak_jarvis('Sir! you dont have any task')
                else:
                    self.speak_jarvis("Sir! today you have following task")
                    self.speak_jarvis(task1)
                    self.speak_jarvis("sir, thats all for now")

        elif 'open youtube' in querry:
            webbrowser.open('youtube.com')

        elif 'open instagram' in querry:
            webbrowser.open('instagram.com')

        elif 'open google' in querry:
            webbrowser.open('google.com')

        elif 'weather' in querry:
            self.calling_weather()

        elif 'time' in querry:
            strtime = datetime.datetime.now().strftime('%H:%M')
            self.speak_jarvis(f"Sir, it's {strtime}")

        elif 'open vs code' in querry:
            vs_code_path = "D:\\Microsoft VS Code\\Code.exe"
            os.startfile(vs_code_path)

        elif 'open chrome' in querry:
            chrome_path = "C:\\Users\\HP\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(chrome_path)

        elif 'notepad' in querry:
            notepad = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Accessories"
            os.startfile(notepad)

        elif 'command prompt' in querry:
            os.system('start cmd')

        elif 'email' in querry:
            #reciever= {'abhishek': 'rams31824@gmail.com', 'aditya': 'shyamsinghrtm457001@gmail.com'}
            #querry = querry.replace('email to', "").lower()
            #nameList= list(reciever.keys())
            #if querry in nameList:
             #       print('inside if ---->', querry)
              #      print(querry)
                
                   # try:
                        print('initialising pro')
                        self.speak_jarvis('Initialising your process sir!')
                        self.speak_jarvis('What is the subject')
                        subject = self.Take_command()
                        self.speak_jarvis('What should i send...')
                        message1 = self.Take_command()
                        message = f"""From: From Person <from@fromdomain.com>
                                    To: To Person <to@todomain.com>
                                    Subject: {subject}

                                    {message1}
                                    """
                        self.Send_email("rams31824@gmail.com", message)
                        self.speak_jarvis('Email has been sent! sir')
                   # except Exception as e:
                    #    print(e)
                     #   self.speak_jarvis("Sorry sir! iam not able to send this email du to some technical issue!")
                        # return 0

            #else:
                    #self.speak_jarvis(f'Sorry sir! unable to find person {querry}')

        # elif 'music' in querry:
        #     ''''using os library'''
        #         music_dir = 'E:\\my_usic'
        #         songs = os.listdir(music_dir)
        #         print(songs)
        #         os.startfile(os.path.join(music_dir, songs[random.randint(0, 12)]))

        elif 'sleep' in querry:
                self.speak_jarvis(f'Iam sleeping down sir!')
                exit()


if __name__ == "__main__":
    # Creating object or intance of class jarvis

    j1 = jarvis()
    
    j1.wish_me()
    j1.speak_jarvis('Iam Jarvis Sir. Please tell me how may i help you')
    while 1:
        j1.Take_command()
        j1.Executing_task()
