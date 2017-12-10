import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.scatter import Scatter
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.core.window import Window
import win32com.client as wincl
import datetime

import wolframalpha
import wikipedia

Window.clearcolor = (0, 0, 1, 1)
Window.size = (500,650)
class MyApp(App):
    title = 'J.A.R.V.I.S'
    def build(self):
        self.usevoice = True
        self.now = datetime.datetime.now().strftime("%I:%M:%S")
        b = BoxLayout(orientation="vertical")
        self.q = TextInput(font_size=25,
                      size_hint_y=None,
                      height=75,
                      multiline=True)
        self.r = TextInput(font_size=20,
                  size_hint_y=None,
                  height=250,
                  multiline=True,
                  pos=(0,150),
                  text='Results: ')

        f = FloatLayout()
        s = Scatter()
        self.l = Label(text="What can I do for you? ",pos=(0,400))
        i = Label(text="Hello, My Name Is J.A.R.V.I.S ",pos=(0,450))
        bx = Button(text="SEND", size_hint_y=None, height=50, pos=(0,100))
        bx.bind(on_press=self.callback)
        t = Label(text=self.now,pos=(200,475))
        bx1 = Button(text="Volume OFF", pos_hint={'x': 0, 'center_y': .19}, size_hint=(None, None))
        bx1.bind(on_press=self.voicebool)
        bx2 = Button(text="Volume ON", pos_hint={'right': 1, 'center_y': .19}, size_hint=(None, None))
        bx2.bind(on_press=self.voicebool1)
        
        f.add_widget(s)
        f.add_widget(t)
        f.add_widget(bx1)
        f.add_widget(bx2)
        f.add_widget(i)
        f.add_widget(self.l)

        b.add_widget(f)
        b.add_widget(self.r)
        b.add_widget(self.q)
        b.add_widget(bx)
        self.speak = wincl.Dispatch("SAPI.SpVoice")
        self.rate = 1
        self.speak.Rate = self.rate
        if self.usevoice == True:
            self.speak.Speak("Hello, My Name Is Jarvis")
        return b
    
    def callback(self, obj):
        self.searchonline() 
    def searchonline(self):
        try:
            #wolframalpha
            client = wolframalpha.Client("YK9VR6-PKLETKARWT")
            res = client.query(self.q.text)
            self.ans = next(res.results).text
            print(self.ans)
            self.r.text = self.ans
            self.q.text = ''
            if self.usevoice == True:
                self.speak.Speak(self.ans)
        except:
            #wikipedia
            self.answer = wikipedia.summary(self.q.text, sentences=2)
            print(self.answer)
            self.r.text = self.answer
            self.q.text = ''
            if self.usevoice == True:
                self.speak.Speak(self.answer)
    def voicebool(self, obj):
        self.usevoice = False
        print(self.usevoice)
    def voicebool1(self, obj):
        self.usevoice = True
        print(self.usevoice)
MyApp().run()
