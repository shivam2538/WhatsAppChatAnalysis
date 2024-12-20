import regex
import pandas as pd
import numpy as np
import emoji
import plotly.express as px
from collections import Counter
import matplotlib.pyplot as plt # type: ignore
from os import path
from PIL import Image
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator

import re

#For Audio Sound
from win32com.client import Dispatch
def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

speak("Hello..., Welcome!...")

def startsWithDateAndTime(s):
    pattern = r'^([0-9]+)(\/)([0-9]+)(\/)([0-9]+), ([0-9]+):([0-9]+)[ ]?(AM|PM|am|pm)? -'
    result = re.match(pattern, s)
    if result:
        return True
    return False

def FindAuthor(s):
  s=s.split(":")
  if len(s)==2:
    return True
  else:
    return False

def getDataPoint(line):   
    splitLine = line.split(' - ') 
    dateTime = splitLine[0]
    date, time = dateTime.split(', ') 
    message = ' '.join(splitLine[1:])
    if FindAuthor(message): 
        splitMessage = message.split(': ') 
        author = splitMessage[0] 
        message = ' '.join(splitMessage[1:])
    else:
        author = None
    return date, time, author, message



import pandas as pd
import re

def startsWithDateAndTime(line):
    # Check if the line starts with the date and time pattern "dd/mm/yy, hh:mm am/pm"
    return bool(re.match(r'^\d{2}/\d{2}/\d{2}, \d{1,2}:\d{2}\s[ap]m - ', line))

def getDataPoint(line):
    # Split the line into date-time and message
    date_time, message = line.split(' - ', 1)
    
    # Split date and time
    date, time = date_time.split(', ')
    
    # Determine if the message contains an author or is a system message
    if ': ' in message:
        author, content = message.split(': ', 1)
        # Check if the author is likely a mobile number
        if not re.match(r'^\+?\d+$', author):
            # If the author is a name, use it
            return date, time, author, content
        else:
            # If the author is a mobile number, handle accordingly
            return date, time, author, content
    else:
        # If no author, handle as a system message or generic message
        return date, time, None, message

data = []  # List to keep track of data so it can be used by a Pandas dataframe
conversation = 'WhatsApp Chat.txt'  # Ensure this path is correct

with open(conversation, encoding="utf-8") as fp:
    fp.readline()  # Skipping the first line
    messageBuffer = []
    date, time, author = None, None, None
    while True:
        line = fp.readline()
        if not line:
            break
        line = line.strip()
        if startsWithDateAndTime(line):
            if len(messageBuffer) > 0:
                data.append([date, time, author, ' '.join(messageBuffer)])
            messageBuffer.clear()
            date, time, author, message = getDataPoint(line)
            messageBuffer.append(message)
        else:
            messageBuffer.append(line)

    # Append the last message
    if len(messageBuffer) > 0:
        data.append([date, time, author, ' '.join(messageBuffer)])

# Create a DataFrame
df = pd.DataFrame(data, columns=['Date', 'Time', 'Author', 'Message'])

# Reorder columns if needed and ensure the DataFrame has the correct format
df = df[['Date', 'Time', 'Author', 'Message']]

# Example: Print the first few rows of the DataFrame
print(df.head(30))

# Save to CSV if needed
df.to_csv('WhatsApp_Chat_Analysis.csv', index=False)

print(df.Author.unique())


import pandas as pd
import re
import numpy as np
import emoji  # Make sure to install the `emoji` library if not already installed
import regex  # Make sure to install the `regex` library if not already installed

def split_count(text):
    # Extract emojis from text
    emoji_list = []
    data = regex.findall(r'\X', text)
    for word in data:
        if any(emoji.is_emoji(char) for char in word):
            emoji_list.append(word)
    return emoji_list

# Assuming df is already defined from the previous steps

# Count media messages
media_messages = df[df['Message'] == '<Media omitted>'].shape[0]
print(f"Media messages: {media_messages}")

# Count emojis
df["emoji"] = df["Message"].apply(split_count)
emojis = sum(df['emoji'].apply(len))
print(f"Emojis: {emojis}")

# Count URLs
URLPATTERN = r'(https?://\S+)'
df['urlcount'] = df['Message'].apply(lambda x: re.findall(URLPATTERN, x)).str.len()
links = np.sum(df['urlcount'])
print(f"Links: {links}")

# Total number of messages
total_messages = df.shape[0]

# Print summary
print("Data Science Community")
print(f"Messages: {total_messages}")
print(f"Media: {media_messages}")
print(f"Emojis: {emojis}")
print(f"Links: {links}")


media_messages_df = df[df['Message'] == '<Media omitted>']
messages_df = df.drop(media_messages_df.index)
messages_df.info()
messages_df['Letter_Count'] = messages_df['Message'].apply(lambda s : len(s))
messages_df['Word_Count'] = messages_df['Message'].apply(lambda s : len(s.split(' ')))
messages_df["MessageCount"]=1

l = ["HOD Sir", '+91 96618 07601', 'Dhilipanrajkumar Sir',"Shanmugapriya Ma'am","Gnanakumar XSE KARE"]
for i in range(len(l)):
  # Filtering out messages of particular user
  req_df= messages_df[messages_df["Author"] == l[i]]
  # req_df will contain messages of only one particular user
  print(f'Stats of {l[i]} -')
  # shape will print number of rows which indirectly means the number of messages
  print('Messages Sent', req_df.shape[0])
  #Word_Count contains of total words in one message. Sum of all words/ Total Messages will yield words per message
  words_per_message = (np.sum(req_df['Word_Count']))/req_df.shape[0]
  print('Words per message', words_per_message)
  #media conists of media messages
  media = media_messages_df[media_messages_df['Author'] == l[i]].shape[0]
  print('Media Messages Sent', media)
  # emojis conists of total emojis
  emojis = sum(req_df['emoji'].str.len())
  print('Emojis Sent', emojis)
  #links consist of total links
  links = sum(req_df["urlcount"])   
  print('Links Sent', links)   
  print()
  
  