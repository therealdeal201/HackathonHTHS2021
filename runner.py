from __future__ import print_function
import os.path
import kivy
from kivy.app import App
from kivy.uix.button import Button
from kivy.graphics import Rectangle, Color
from kivy.uix.textinput import TextInput
from kivy.uix.label import Label
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.scatter import Scatter
from kivymd.app import MDApp
from kivymd.uix.screen import Screen
from kivymd.uix.textfield import MDTextField
import os.path
import speech_recognition as sr
import pyaudio
import numpy as np

import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import numpy as np
import csv


class DemoApp(MDApp):

    def build(self):
        screen = Screen()

        btn = Button(text ="Start Recording",
                   background_color=(20,1,1,1),
                   #font_size="20sp",
                   #color=(1,1,1,4),
                   size = (30,10),
                   size_hint =(.6, .4),
                   pos =(175,200))
        screen.add_widget(btn)
        btn.bind(on_press=self.audioScript)

        return screen

    def audioScript(self, *args):
        exec(open("SpeechRecog.py").read())
        print("hello")


d = DemoApp()
d.run()
