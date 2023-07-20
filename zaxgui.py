import typing
from PyQt5.QtCore import QObject
from PyQt5.QtWidgets import QWidget
from zaxui import  Ui_zaxui
from zaxui import  *
from PyQt5 import QtCore,QtWidgets,QtGui
from PyQt5.QtGui import QMovie
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.uic import loadUiType
import warnings
import speech_recognition as aa
import pyttsx3
import datetime
import time
import subprocess
import smtplib
import json
import requests
import pandas as pd
import numpy as np
import openpyxl
#import pprint
import json
import pymongo
import sys
from io import StringIO

client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["griet"]
machine=pyttsx3.init()
voices = machine.getProperty('voices')       #getting details of current voice
#engine.setProperty('voice', voices[0].id)  #changing index, changes voices. o for male
machine.setProperty('voice', voices[1].id) 
machine.setProperty('rate',170)  #changing index, changes voices. 1 for female
listener=aa.Recognizer()

def talk(text):
    # self.ui.outputLabel.setText(text)
    machine.say(text)
    machine.runAndWait()

class MainThread(QThread):

    def __init__(self):
        super(MainThread,self).__init__()

    
    def run(self):
        self.Task_Gui()

    def wish(self):
        hour=int(datetime.datetime.now().hour)
        if(hour>=0 and hour<=12):
            talk("Good Morning!")
            talk("Welcome to ZAX,your G R I E T virtual assistant . Please tell me how may I help you")
        elif hour>=12 and hour<18:
            talk("Good Afternoon!")
            talk("Welcome to ZAX,your G R I E T virtual assistant . Please tell me how may I help you")  
        else:
            talk("Good Evening!") 
            talk("Welcome to ZAX,your G R I E T virtual assistant . Please tell me how may I help you")

    def input_instruction(self):

        with aa.Microphone() as source:
                    listener.pause_threshold =1
                    listener.adjust_for_ambient_noise(source)
                    print("Listening....")
                    speech=listener.listen(source,phrase_time_limit = 5)

        try:
            print("Recognizing...")
            self.instruction=listener.recognize_google(speech,language='en-in')
                
        except:
            return "None"
        
        return self.instruction
    
    def Task_Gui(self):

        self.wish()
        while(True):
            instruction=self.input_instruction().lower()
            print(instruction)
            f=0
            instruction=instruction.replace("details of","")
            instruction=instruction.replace("dr","")   
            instruction=instruction.replace("babu","")
            instruction=instruction.replace("rao","")
            instruction=instruction.replace("raju","")
            instruction=instruction.replace("lakshmi","")
            instruction=instruction.replace("chandra","")

            if "contact " in instruction:        
                collection = db["newcontactdetails"]
                if "sri " in instruction:
                    results = collection.find({"name":'Dr.V Sri lakshmi'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)               
                    if(len(l_c)!=0):                       
                            for doc in documents:
                                contact = doc['contact']
                                print(contact)
                                talk(contact)     
                            f=f+1
                            f=f+1

                elif "krishna" in instruction:
                    x=''
                    if "bhargavi" in instruction:
                            x=x+'Dr.Krishna Bhargavi Y'
                    else:
                        x=x+'N Krishna Chythanya'
                    results = collection.find({"name":x})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                
                    if(len(l_c)!=0):
                            for doc in documents:
                                contact = doc['contact']
                                print(contact)
                                talk(contact)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "bhargavi" in instruction:
                    results = collection.find({"name":'Dr.S.Bhargavi Latha'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                contact = doc['contact']
                                print(contact)
                                talk(contact)
                        f=f+1
                        df=pd.DataFrame(l_c)
                
                        f=f+1

                else:
                    results = collection.find({"$text": {"$search":instruction}})
                    documents = [doc.copy() for doc in results]
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for x in documents:
                                contact = x['contact']
                                print(contact)
                                talk(contact)
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1

            elif "experience " in instruction:
                collection = db["experience"]
                if "sri " in instruction:
                    results = collection.find({"name":'Dr.V Sri lakshmi'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                            for doc in documents:
                                contact = doc['experience']
                                x=doc['link_source']
                                print(contact)
                                print(x)
                                talk(contact)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "krishna" in instruction:
                    x=''
                    if "bhargavi" in instruction:
                            x=x+'Dr.Krishna Bhargavi Y'
                    else:
                        x=x+'N Krishna Chythanya'
                    results = collection.find({"name":x})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                            for doc in documents:
                                contact = doc['experience']
                                x=doc['link_source']
                                print(contact)
                                print(x)
                                talk(contact)      
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "bhargavi" in instruction:
                    #results = collection.find({"$text": {"$search":instruction}})
                    results = collection.find({"name":'Dr.S.Bhargavi Latha'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                            contact = doc['experience']
                            x=doc['link_source']
                            print(contact)
                            print(x)
                            talk(contact)      
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1
                        
                else:
                    results = collection.find({"$text": {"$search":instruction}})
                    documents = [doc.copy() for doc in results]
                    
                    l_c=list(documents)
            
                    if(len(l_c)!=0):
                    
                        #print(l_c)
                        for x in documents:
                                contact = x['experience']
                                y = x['link_source']
                                print(contact)
                                print(y)
                                talk(contact)
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1

            elif "about" in instruction:            
                collection = db["newabout"]
                if "sri " in instruction:
                    results = collection.find({"name":'Dr.V Sri lakshmi'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                            for doc in documents:
                                about = doc['about']
                                print(about)
                                talk(about)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "krishna" in instruction:
                    x=''
                    if "bhargavi" in instruction:
                            x=x+'Dr.Krishna Bhargavi Y'
                    else:
                        x=x+'N Krishna Chythanya'
                    results = collection.find({"name":x})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)              
                    if(len(l_c)!=0):
                            for doc in documents:
                                    about = doc['about']
                                    print(about)
                                    talk(about)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            print(df)
                            f=f+1

                elif "bhargavi" in instruction:
                    results = collection.find({"name":'Dr.S.Bhargavi Latha'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                about = doc['about']
                                print(about)
                                talk(about)
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1

                else:
                    results = collection.find({"$text": {"$search":instruction}})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                about = doc['about']
                                print(about)
                                talk(about)
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1

            elif "designation"  in instruction:
                collection = db["newdesgn"]
                if "sri " in instruction:
                    results = collection.find({"name":'Dr.V Sri lakshmi'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)             
                    if(len(l_c)!=0):
                            for doc in documents:
                                designation = doc['designation']
                                print(designation)
                                talk(designation)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "Mallikarjuna" in instruction:
                    print("Professor")
                    talk("Professor")

                elif "krishna" in instruction:
                    x=''
                    if "bhargavi" in instruction:
                            x=x+'Dr.Krishna Bhargavi Y'
                    else:
                        x=x+'N Krishna Chythanya'
                    results = collection.find({"name":x})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                            for doc in documents:
                                    designation = doc['designation']
                                    print(designation)
                                    talk(designation)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "bhargavi" in instruction:
                    #results = collection.find({"$text": {"$search":instruction}})
                    results = collection.find({"name":'Dr.S.Bhargavi Latha'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                designation = doc['designation']
                                print(designation)
                                talk(designation)
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1

                else:
                    results = collection.find({"$text": {"$search":instruction}})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                designation = doc['designation']
                                print(designation)
                                talk(designation)
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1
            elif "research papers " in instruction:
                collection = db["newpublications"]
                if "sri " in instruction:
                    results = collection.find({"name":'Dr.V Sri lakshmi'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                            for doc in documents:
                                x= doc['publications']
                                y=doc['link_source']
                                print(x)
                                print(y)
                                talk(x)
                                print(y)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1
                
                elif "krishna" in instruction:
                    x=''
                    if "bhargavi" in instruction:
                            x=x+'Dr.Krishna Bhargavi Y'
                    else:
                        x=x+'N Krishna Chythanya'
                    results = collection.find({"name":x})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                            for doc in documents:
                                    x= doc['publications']
                                    y=doc['link_source']
                                    print(x)
                                    print(y)
                                    talk(x)   
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "bhargavi" in instruction:
                    #results = collection.find({"$text": {"$search":instruction}})
                    results = collection.find({"name":'Dr.S.Bhargavi Latha'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                x= doc['publications']
                                y=doc['link_source']
                                print(x)
                                print(y)
                                talk(x)            
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1

                else:
                    results = collection.find({"$text": {"$search":instruction}})
                    documents = [doc.copy() for doc in results]
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                x=doc['publications']
                                y=doc['link_source']
                                print(x)
                                print(y)
                                talk(x)
                        f=f+1
                        df=pd.DataFrame(l_c)
                        f=f+1

            elif "courses" or "subject" in instruction:
                collection = db["subject"]
                if "sri " in instruction:
                    results = collection.find({"name":'Dr.V Sri lakshmi'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                            for doc in documents:
                                Courses_Taught = doc['Courses Taught']
                                print(Courses_Taught)
                                talk(Courses_Taught)
                            f=f+1
                            df=pd.DataFrame(l_c)
                            f=f+1

                elif "krishna" in instruction:
                    x=''
                    if "chaitanya" in instruction:
                            x=x+'N Krishna Chythanya'      
                    else:
                        x=x+'Dr.Krishna Bhargavi Y'
                    results = collection.find({"name":x})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                            del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                                for doc in documents:
                                    Courses_Taught = doc['Courses Taught']
                                    print(Courses_Taught)
                                    talk(Courses_Taught)
                                f=f+1
                                df=pd.DataFrame(l_c)
                                f=f+1    

                elif "bhargavi" in instruction:
                    #results = collection.find({"$text": {"$search":instruction}})
                    results = collection.find({"name":'Dr.S.Bhargavi Latha'})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)
                    if(len(l_c)!=0):
                        for doc in documents:
                                Courses_Taught = doc['Courses Taught']
                                print(Courses_Taught)
                                talk(Courses_Taught)
                        f=f+1
                        df=pd.DataFrame(l_c)    
                else:
                    results = collection.find({"$text": {"$search":instruction}})
                    documents = [doc.copy() for doc in results]
                    for doc in documents:
                        del doc['_id']
                    l_c=list(documents)              
                    if(len(l_c)!=0):
                        for doc in documents:
                                Courses_Taught = doc['Courses Taught']
                                print(Courses_Taught)
                                talk(Courses_Taught)
                        f=f+1 
                        df=pd.DataFrame(l_c)  
                         
            if "exit" in instruction or "quit" in instruction:
                talk("thank you,have a nice day")
                exit()
            elif "none" in instruction:
                    print("...")
            elif(f==0):
                talk("data not found")
                talk("please repeat it correctly")
            else:
                    print("")            

startFunction= MainThread()

class Gui_Start(QMainWindow):
    def __init__(self):
        super().__init__()
        self.zax_ui = Ui_zaxui()
        self.zax_ui.setupUi(self)
        self.zax_ui.pushButton.clicked.connect(self.startFunc)
        self.zax_ui.pushButton_2.clicked.connect(self.close)

    def __del__(self):
        sys.stdout = sys.__stdout__

    def startFunc(self):
            self.zax_ui.jargif=QtGui.QMovie("material/Aqua.gif")
            self.zax_ui.label_3.setMovie(self.zax_ui.jargif)
            self.zax_ui.jargif.start()
            self.zax_ui.movie1=QtGui.QMovie("material\B.G_Template_1.gif")
            self.zax_ui.label_2.setMovie(self.zax_ui.movie1)
            self.zax_ui.movie1.start()            
            startFunction.start()

Gui_App=QApplication(sys.argv)
Gui_zax=Gui_Start()
Gui_zax.show()
exit(Gui_App.exec_())