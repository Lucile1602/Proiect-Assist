import sys
import speech_recognition as sr
import win32com.client as win32
import pyodbc
from datetime import date, datetime, timedelta
days_ahead=0
task_day=''
tmp_Description=''
date_index=0
recognized_date = None 

while True:
    # Create a recognizer object
    recognizer = sr.Recognizer()
    
    def set_date():
        global recognized_date,task_day,date_index
          # Use Google Speech Recognition to recognize the audio
          # Use the default microphone as the audio source
        with sr.Microphone() as source:
            print("Listening Day...")
            audio = recognizer.listen(source)

        try:
            text3 = recognizer.recognize_google(audio)
            print("You said:", text3)
            recognized_date = text3
                        
            today = date.today()
            current_weekday = today.weekday()
            
            match text3.lower():
                case 'monday':
                    #print("The date is Monday")
                    days_ahead = (0 - current_weekday) % 7  # 0 corresponds to Monday
                    date_index = 2
                case 'tuesday':
                    #print("The date is Tuesday")
                    days_ahead = (1 - current_weekday) % 7  # 1 corresponds to Tuesday (Monday is 0)
                    date_index = 3
                case 'wednesday':
                    #print("The date is Wednesday")
                    days_ahead = (2 - current_weekday) % 7  # 2 corresponds to wednesday (Monday is 0)
                    date_index = 4
                case 'thursday':
                    #print("The date is Thursday")
                    days_ahead = (3 - current_weekday) % 7  # 3 corresponds to Thursday (Monday is 0)
                    date_index = 5
                case 'friday':
                    #print("The date is Friday")
                    days_ahead = (4 - current_weekday) % 7  # 4 corresponds to Friday (Monday is 0)
                    date_index = 6
                case 'saturday':
                    #print("The date is Saturday")
                    days_ahead = (5 - current_weekday) % 7  # 5 corresponds to Saturday (Monday is 0)
                    date_index = 7
                case 'sunday':
                    #print("The date is Sunday")
                    days_ahead = (6 - current_weekday) % 7  # 6 corresponds to Sunday (Monday is 0)
                    date_index = 1                        
                case _:
                    print("The value is something else")             
            
            
            task_day= today + timedelta(days=days_ahead)
            #print(task_day)        
            
        except Exception as e:
            print("Error:", str(e))
            recognize_speech()
        # Return the recognized date (or None if no date was recognized)
        return recognized_date, date_index
    
    def recognize_speech():
        
        # Use the default microphone as the audio source
        with sr.Microphone() as source:
            print("Listening...")
            print("\nOptions:")
            print(" - Create")
            print(" - Report")
            print(" - Exit")
            print("")
            audio = recognizer.listen(source)

        try:
            # Use Google Speech Recognition to recognize the audio
            text = recognizer.recognize_google(audio)
            print("You said", text)
            
            if "report" in text.lower():
                              
                print("Your report has been opened!")

                # Create an instance of the Access application
                ac = win32.Dispatch("Access.Application")

                # Set Access to be visible (optional)
                ac.Visible = True

                # Open the Access database
                ac.OpenCurrentDatabase("C:\\Users\\lucil\\OneDrive\\Documents\\assist\\REPORT_nov2023_2.accdb")

                # Run the macro
                ac.DoCmd.RunMacro('Autoexec')


                    
            if "create" in text.lower():
                while True:
                        
                    # Use Google Speech Recognition to recognize the audio
                    # Use the default microphone as the audio source
                    with sr.Microphone() as source:
                        print("Listening Description...")
                        audio = recognizer.listen(source)

                    try:
                        text2 = recognizer.recognize_google(audio)
                        print("You said:", text2)
                          
                        # Get the description
                        tmp_Description=text2.lower()
                        
                        set_date()
                        
               
                        try:
                            
                            
                            # Create a connection to the local Access database
                            cnxn_str = (
                                r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
                                r'DBQ=C:\\Users\\lucil\\OneDrive\\Documents\\assist\\REPORT_nov2023_2.accdb;'                      )
                            cnxn = pyodbc.connect(cnxn_str)

                            # Create a cursor object to execute SQL statements
                            cursor = cnxn.cursor()
                            
                            # Modify your SQL query to insert records into the local database
                            insert_query = "INSERT INTO TW_Tasks (Tasks_Date, Tasks_Date_Index, Tasks_Description_short, Tasks_Description_long) VALUES ( ?, ?, ?, ?)"
                            values = (task_day, date_index, tmp_Description, 'Value3')

                            cursor.execute(insert_query, values)
                            cnxn.commit()
                           
                            print("Record inserted successfully into local database!")

                            recognize_speech()

                        except Exception as e:
                            print("Error:", str(e))
                            recognize_speech()
                        
                    finally:
                        # Close the cursor and the connection
                        cursor.close()
                        cnxn.close()
                        
            if "exit" in text.lower():
                sys.exit()
            
        except Exception as e:
            print("Error:", str(e))
            recognize_speech()

        except sr.UnknownValueError:
            print("Could not understand audio")
        except sr.RequestError as e:
            print("Could not request results from Google Speech Recognition service; {0}".format(e))

    # Call the function to start voice recognition
    recognize_speech()

