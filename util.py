from __future__ import print_function
import os.path
import speech_recognition as sr
import pyaudio
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import numpy as np
import csv

def main():
    r = sr.Recognizer()
    print(sr.Microphone.list_microphone_names())
    strWords = ""

    with sr.Microphone(device_index=1) as source:

        audio = r.listen(source)
    try:
        strWords = r.recognize_google(audio)
        print("System Predicts:\n" + strWords)
    except Exception:
        print("Something Went Wrong")

    strlist = strWords.lower().split("splitter")

    i = 0
    finalstrlist = []

    for slide in strlist:
        words = slide.split(" ")
        while(i < len(words)):
            if words[i].lower() == "stopping":
                if i + 1 < len(words) and words[i+1].lower() == "stopping":
                    #stop stop -> one lowercase stop, used in case someone says stop
                    words.pop(i)
                else:
                    #stop -> (Period)
                    words.pop(i)
                    words.insert(i, "newline")
            elif words[i].lower() == "header":
                finalstrlist.append(" ".join(words).split("header")[0])
                words = words[i+1:]

            i = i + 1
        #print(words)
        i = 0
        finalstrlist.append(" ".join(words))

    print(finalstrlist)
    np.savetxt("SpeechList.csv", finalstrlist, delimiter =", ", fmt  ='% s')



    speechList = []

    with open("SpeechList.csv") as csvfile:
        print(csvfile)
        reader = csv.reader(csvfile, delimiter=',')
        for row in reader:
            speechList.append(row)

    for thing in speechList:
        print(thing)
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    # The ID of a sample presentation.
    PRESENTATION_ID = '1v97TDXl0gFzyLYFcIuQJdxGD1tsK1WrXWiNEed6MXN8'

    """Shows basic usage of the Slides API.
    Prints the number of slides and elments in a sample presentation.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('thisFileDoesNotExistThisIsJustFiller.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('slides', 'v1', credentials=creds)
    driveService = build('drive', 'v2', credentials=creds)

    """
    #code for creating new presentation from scratch
    body = {
        'title': "HI THERE"
    }

    presentation = service.presentations() \
        .create(body=body).execute()
    print('Created presentation with ID: {0}'.format(
        presentation.get('presentationId')))

    """

    # code for creating new presentation from existing
    body = {
        'name': 'Sample Presentation - Hackathon'
    }
    #Creat Google Drive Presentation Copy
    presentation = driveService.files().copy(
        fileId='1Hs58Wxe0dadwYQW7AGNDFMF1m5Frl6gNGXr49fkUh9g', body=body).execute()
    print('Created presentation with ID: {0}'.format(presentation.get('id')))

    #Access Presentation Copy from Google Slides API
    presentation = service.presentations().get(presentationId='1Hs58Wxe0dadwYQW7AGNDFMF1m5Frl6gNGXr49fkUh9g').execute()
    # Get info for title slide
    titleSlide =presentation['slides'][0]

    titleID = titleSlide['pageElements'][0]['objectId']
    subtitleID = titleSlide['pageElements'][1]['objectId']

    titleSlideMain = str(speechList[0][0].strip()).capitalize() #accounts for the fact that each element in speechList is a list with one value
    titleSlideSub = str(speechList[1][0].strip()).capitalize() #accounts for the fact that each element in speechList is a list with one value
    reqs = [
        #create main text
        {'createSlide': {'slideLayoutReference': {'predefinedLayout': 'MAIN_POINT'}}},
        {'insertText': {'objectId': titleID, 'text': titleSlideMain}},
        #format main text
        {'updateTextStyle': {
            'objectId': titleID,
                                        'style': {'fontSize': {'magnitude': 48, 'unit': 'PT'}},
                                        'textRange':{'type': 'ALL'},
                                        'fields': 'fontSize',
        }},
        #create sub text
        {'insertText': {'objectId': subtitleID, 'text': titleSlideSub}},
        # format sub text
        {'updateTextStyle': {
            'objectId': subtitleID,
            'style': {'fontSize': {'magnitude': 24, 'unit': 'PT'}},
            'textRange': {'type': 'ALL'},
            'fields': 'fontSize',
        }}
    ]

    response = service.presentations() \
        .batchUpdate(presentationId=presentation.get('presentationId'), body={'requests': reqs}).execute()
    #create_slide_response = response.get('replies')[0].get('createSlide')



    num = 0
    enum = 2
    try:
        for slidePhrase in speechList[2::2]: #Accesses every title by default
            print(slidePhrase)
            runWhileNum = 1

            requests = [
                {
                    'createSlide': {
                    }
                }
            ]

            # Execute the request.
            body = {
                'requests': requests
            }
            response = service.presentations() \
                .batchUpdate(presentationId=presentation.get('presentationId'), body=body).execute()
            create_slide_response = response.get('replies')[0].get('createSlide')
            print('Created slide with ID: {0}'.format(
                create_slide_response.get('objectId')))
            # [END slides_create_slide]


            #insert text
            # [START slides_create_textbox_with_text]
            # Create a new square textbox, using the supplied element ID.
            print(presentation)
            while(runWhileNum <= 2): #Runs two times and will change text box size/position based on whether it is a header or body box

                element_id = 'My_Text_New_Box_' + str(num)
                if(runWhileNum == 1):
                    scaleX = 1.915
                    scaleY = 0.125
                    translateX = 25
                    translateY = 35
                    textSize = 27
                    text = str(slidePhrase)

                else:
                    scaleX =1.915
                    scaleY =0.8
                    translateX = 25
                    translateY = 90
                    textSize = 18
                    print(enum)
                    text = str(speechList[enum+1])

                #Delete extra brackets/spaces and capitalize first letter of each text entry
                for i,  c in enumerate(text):
                    if not(c == '[' or c == "'" or c == '"' or c == ' '): #Don't mind this horrible code, it's very late at night!
                        break
                text = text[i:].capitalize()[:-2]
                temporaryNewLineList = [x.strip().capitalize() + "\n" for x in text.split("newline")]
                text = "".join(temporaryNewLineList)

                pt350 = {
                    'magnitude': 350,
                    'unit': 'PT'
                }
                requests = [
                    {
                        'createShape': {
                            'objectId': element_id,
                            'shapeType': 'TEXT_BOX',
                            'elementProperties': {
                                'pageObjectId': create_slide_response.get('objectId'),
                                'size': {
                                    'height': pt350,
                                    'width': pt350
                                },
                                'transform': {
                                    'scaleX': scaleX,
                                    'scaleY': scaleY,
                                    'translateX': translateX,
                                    'translateY': translateY,
                                    'unit': 'PT'
                                }
                            }
                        }
                    },

                    # Insert text into the box, using the supplied element ID.
                    {
                        'insertText': {
                            'objectId': element_id,
                            'insertionIndex': 0,
                            'text': text
                        }
                    },
                        {'updateTextStyle': {
                            'objectId': element_id,
                            'style': {'fontSize': {'magnitude': textSize, 'unit': 'PT'}},
                            'textRange': {'type': 'ALL'},
                            'fields': 'fontSize',
                        }
                        }

                ]


                # Execute the request.
                body = {
                    'requests': requests
                }
                response = service.presentations() \
                    .batchUpdate(presentationId=presentation.get('presentationId'), body=body).execute()
                create_shape_response = response.get('replies')[0].get('createShape')
                print('Created textbox with ID: {0}'.format(
                    create_shape_response.get('objectId')))





                num += 1  # Adds 1 to the numList to create a new textbox (can't be at end of loop because another one
                # must be created for body text)
                runWhileNum +=1
            enum +=2
                # [END slides_create_textbox_with_text]
    except Exception:
        pass


if __name__ == '__main__':
    main()
