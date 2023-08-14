import os
import datetime
import speech_recognition as sr
import win32com.client
import webbrowser
import openai
import subprocess
from config import apikey

speaker =win32com.client.Dispatch("SAPI.Spvoice")
voices = speaker.GetVoices()
speaker.Voice = voices[1]

def get_date_and_day():
    now = datetime.datetime.now()
    date = now.strftime("%Y-%m-%d")
    day = now.strftime("%A")
    return date, day
    
chatStr = ""
def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Shaishta:{query}\nLuna:"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        say(response["choices"][0]["text"])
        chatStr+= f"{response['choices'][0]['text']}\n"
        return response["choices"][0]["text"]
    except Exception as e:
        say("Sorry I did not get you. PLease repeat.")

def ai(prompt):
  openai.api_key = apikey
  text = f"Openai response for Prompt:  {prompt}\n\n\t\t\t\t\t\t\t****\n\n"
  response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
      {
        "role": "user",
        "content": prompt
      },
      {
        "role": "assistant",
        "content": ""
      }
    ],
    temperature=1,
    max_tokens=500,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
  )
  try:
      print(response["choices"][0]["message"]["content"])
      text += response["choices"][0]["message"]["content"]
      if not os.path.exists("Openai"):
          os.mkdir("Openai")
      with open(f"Openai/{''.join(prompt.split('using')[1:]).strip()}.txt", "w") as f:
          f.write(text)
  except  Exception as e:
      say("Sorry there was an error generating your response. Please try again later.")

def say(text):
    speaker.Speak(f"{text}")
def takeCommand():
    r =sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        r.energy_threshold = 150
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Sorry I did not get you. PLease repeat."

if __name__=='__main__':
    speaker.Speak("Hello I am Luna. How can I assist you?")
    sites = [["youtube", "https://youtube.com"], ["google", "https://google.com"],
             ["whatsapp", "https://web.whatsapp.com"], ["wikipedia", "https://wikipedia.com/"],
             ["linkedin", "https://linkedin.com/"], ["instagram", "https://instagram.com"]]

    # todo: Modify the App path according to your system
    apps = [["edge", "<path>"],
            ["spotify", "<path"],
            ["brave", "<path>"],
            ["firefox", "<path>"],["chrome", "<path>"],["teams", "<path>"],["python", "<path>"]]
    while True:
        print("\nListening:")
        query=takeCommand()
        for site in sites:
            try:
                if f"Open {site[0]}".lower() in query.lower():
                    speaker.Speak(f"Opening {site[0]} right away.")
                    webbrowser.open(site[1])
            except Exception as e:
                say("Sorry there was an error opening this site.")

        for app in apps:
            try:
                if f"Open {app[0]}".lower() in query.lower():
                    speaker.Speak(f"Opening {app[0]} .")
                    subprocess.Popen([app[1]])
            except Exception as e:
                say("Sorry there was an error opening this app.")

        #todo: add a feature to play a specific song
        if "play music" in query:
                musicPath="C://Users//PRINCE//Music//Midnight.mpeg"
                speaker.Speak("Playing music..")
                os.startfile(musicPath)

        elif "the time" in query:
                strfTime= datetime.datetime.now().strftime("%H:%M:%S")
                speaker.Speak(f"The time is {strfTime}")

        elif "what's the date" in query.lower() or "tell me the date" in query.lower() or "what day is it" in query.lower() or "what is the day today" in query.lower():
            date, day = get_date_and_day()
            speaker.Speak(f"Today is {date} and it's a {day}.")


        elif "using ai".lower() in query.lower():
            print("Generating...\n")
            ai(prompt=query)


        elif "Luna stop".lower() in query.lower():
            speaker.Speak("Luna signing off. Goodbye")
            exit()

        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)
