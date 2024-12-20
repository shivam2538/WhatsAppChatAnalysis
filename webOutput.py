import pandas as pd
import numpy as np
import regex
import emoji
import re

#For Audio Sound
from win32com.client import Dispatch
def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)


# Load the DataFrame from the CSV file or re-process the data as needed
df = pd.read_csv('WhatsApp_Chat_Analysis.csv')

# Ensure the 'Message' column contains only strings
df["Message"] = df["Message"].astype(str)

def split_count(text):
    if not isinstance(text, str):  # Ensure that 'text' is a string
        return []
    
    emoji_list = []
    data = regex.findall(r'\X', text)
    for word in data:
        if emoji.is_emoji(word):
            emoji_list.append(word)
    return emoji_list

# Calculate additional columns
df['Letter_Count'] = df['Message'].apply(len)
df['Word_Count'] = df['Message'].apply(lambda s: len(s.split(' ')))

# Calculate statistics
total_messages = df.shape[0]
media_messages = df[df['Message'] == '<Media omitted>'].shape[0]
df["emoji"] = df["Message"].apply(split_count)
emojis = sum(df['emoji'].apply(len))
URLPATTERN = r'(https?://\S+)'
df['urlcount'] = df['Message'].apply(lambda x: re.findall(URLPATTERN, x)).str.len()
links = np.sum(df['urlcount'])

# User stats (assuming 'l' contains the list of users)
l = [ "HOD Sir" , '+91 96618 07601', 'Dhilipanrajkumar Sir',"Shanmugapriya Ma'am","Gnanakumar XSE KARE"]
user_stats_html = ""
for user in l:
    req_df = df[df["Author"] == user]
    if req_df.shape[0] > 0:
        words_per_message = (np.sum(req_df['Word_Count'])) / req_df.shape[0]
        media = df[(df['Message'] == '<Media omitted>') & (df['Author'] == user)].shape[0]
        emojis = sum(req_df['emoji'].apply(len))
        links = sum(req_df["urlcount"])

        user_stats_html += f"""
        <h3>Stats of {user}</h3>
        <p><strong>Messages Sent:</strong> {req_df.shape[0]}</p>
        <p><strong>Words per Message:</strong> {words_per_message:.2f}</p>
        <p><strong>Media Messages Sent:</strong> {media}</p>
        <p><strong>Emojis Sent:</strong> {emojis}</p>
        <p><strong>Links Sent:</strong> {links}</p>
        """

# Create the HTML content
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WhatsApp Chat Analysis</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            color: #333;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }}
        th, td {{
            padding: 8px 12px;
            border: 1px solid #ddd;
            text-align: left;
        }}
        th {{
            background-color: #f4f4f4;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        .emoji {{
            font-size: 1.2em;
        }}
        .media {{
            color: #007bff;
        }}
    </style>
</head>
<body>
    <h1>WhatsApp Chat Analysis</h1>
    <h2>Chat Data</h2>
    {table}
    <h2>Summary</h2>
    <p><strong>Messages:</strong> {total_messages}</p>
    <p><strong>Media Messages:</strong> {media_messages}</p>
    <p><strong>Emojis:</strong> {emojis}</p>
    <p><strong>Links:</strong> {links}</p>
    {user_stats}
</body>
</html>
"""

# Convert DataFrame to HTML
table_html = df.to_html(classes='data', header=True, index=False, escape=False)

# Fill in the HTML template with the actual data
try:
    html_content = html_template.format(
        table=table_html,
        total_messages=total_messages,
        media_messages=media_messages,
        emojis=emojis,
        links=links,
        user_stats=user_stats_html
    )
except KeyError as e:
    print(f"KeyError: {e}. Check if all placeholders in the HTML template match with the format keys.")
    raise

# Write the HTML content to a file
with open("WhatsApp_Chat_Analysis.html", "w", encoding="utf-8") as f:
    f.write(html_content)
speak("File Generated for the Chat Analysis...")
print("HTML file has been generated successfully!")
