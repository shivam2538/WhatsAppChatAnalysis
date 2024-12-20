# from google.colab import files
# uploaded = files.upload()

import pdfkit

# Set the path to the wkhtmltopdf executable
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)




#For Audio Sound
from win32com.client import Dispatch
def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

import pandas as pd
data = pd.read_csv('data/ner_dataset.csv', encoding= 'unicode_escape')
data.head()
print(data)

# Data Preparation for Neural Networks
from itertools import chain
def get_dict_map(data, token_or_tag):
    tok2idx = {}
    idx2tok = {}
    
    if token_or_tag == 'token':
        vocab = list(set(data['Word'].to_list()))
    else:
        vocab = list(set(data['Tag'].to_list()))
    
    idx2tok = {idx:tok for  idx, tok in enumerate(vocab)}
    tok2idx = {tok:idx for  idx, tok in enumerate(vocab)}
    return tok2idx, idx2tok
token2idx, idx2token = get_dict_map(data, 'token')
tag2idx, idx2tag = get_dict_map(data, 'tag')

# Now I will transform the columns in the data to extract the sequential data for our neural network:

import pandas as pd

# Assuming token2idx and tag2idx are already defined
data['Word_idx'] = data['Word'].map(token2idx)
data['Tag_idx'] = data['Tag'].map(tag2idx)
data_fillna = data.fillna(method='ffill', axis=0)

# Groupby and collect columns
data_group = data_fillna.groupby(
    ['Sentence #'], as_index=False
)[['Word', 'POS', 'Tag', 'Word_idx', 'Tag_idx']].agg(lambda x: list(x))

# Alternatively, using a more explicit approach for readability
data_group = data_fillna.groupby('Sentence #', as_index=False).agg({
    'Word': lambda x: list(x),
    'POS': lambda x: list(x),
    'Tag': lambda x: list(x),
    'Word_idx': lambda x: list(x),
    'Tag_idx': lambda x: list(x)
})

print(data_group.head())

### +++++++++++++++++++++++++++++++++
from sklearn.model_selection import train_test_split
from keras.preprocessing.sequence import pad_sequences
from keras.utils import to_categorical

def get_pad_train_test_val(data_group, data):

    #get max token and tag length
    n_token = len(list(set(data['Word'].to_list())))
    n_tag = len(list(set(data['Tag'].to_list())))

    #Pad tokens (X var)    
    tokens = data_group['Word_idx'].tolist()
    maxlen = max([len(s) for s in tokens])
    pad_tokens = pad_sequences(tokens, maxlen=maxlen, dtype='int32', padding='post', value= n_token - 1)

    #Pad Tags (y var) and convert it into one hot encoding
    tags = data_group['Tag_idx'].tolist()
    pad_tags = pad_sequences(tags, maxlen=maxlen, dtype='int32', padding='post', value= tag2idx["O"])
    n_tags = len(tag2idx)
    pad_tags = [to_categorical(i, num_classes=n_tags) for i in pad_tags]
    
    #Split train, test and validation set
    tokens_, test_tokens, tags_, test_tags = train_test_split(pad_tokens, pad_tags, test_size=0.1, train_size=0.9, random_state=2020)
    train_tokens, val_tokens, train_tags, val_tags = train_test_split(tokens_,tags_,test_size = 0.25,train_size =0.75, random_state=2020)

    print(
        'train_tokens length:', len(train_tokens),
        '\ntrain_tokens length:', len(train_tokens),
        '\ntest_tokens length:', len(test_tokens),
        '\ntest_tags:', len(test_tags),
        '\nval_tokens:', len(val_tokens),
        '\nval_tags:', len(val_tags),
    )
    
    return train_tokens, val_tokens, test_tokens, train_tags, val_tags, test_tags

train_tokens, val_tokens, test_tokens, train_tags, val_tags, test_tags = get_pad_train_test_val(data_group, data)



##########Training Neural Network for Named Entity Recognition (NER)########
import numpy as np
import tensorflow
from tensorflow.keras import Sequential, Model, Input
from tensorflow.keras.layers import LSTM, Embedding, Dense, TimeDistributed, Dropout, Bidirectional
from tensorflow.keras.utils import plot_model
from numpy.random import seed
seed(1)
tensorflow.random.set_seed(2)

#################
input_dim = len(list(set(data['Word'].to_list())))+1
output_dim = 64
input_length = max([len(s) for s in data_group['Word_idx'].tolist()])
n_tags = len(tag2idx)


#######get_bilstm_lstm_model Function#########
from tensorflow.keras import Sequential
from tensorflow.keras.layers import Embedding, LSTM, Bidirectional, TimeDistributed, Dense
from tensorflow.keras.utils import plot_model
import numpy as np
import pandas as pd

# Define dimensions
input_dim = 10000  # Example dimension
output_dim = 128
input_length = 50
n_tags = 10

def get_bilstm_lstm_model(input_dim, output_dim, input_length, n_tags):
    model = Sequential()

    # Add Embedding layer
    model.add(Embedding(input_dim=input_dim, output_dim=output_dim, input_length=input_length))

    # Add bidirectional LSTM
    model.add(Bidirectional(LSTM(units=output_dim, return_sequences=True, dropout=0.2, recurrent_dropout=0.2), merge_mode='concat'))

    # Add LSTM
    model.add(LSTM(units=output_dim, return_sequences=True, dropout=0.5, recurrent_dropout=0.5))

    # Add TimeDistributed Layer
    model.add(TimeDistributed(Dense(n_tags, activation="relu")))

    # Compile model
    model.compile(loss='categorical_crossentropy', optimizer='adam', metrics=['accuracy'])

    # Explicitly build the model
    model.build(input_shape=(None, input_length))

    model.summary()
    
    return model

def train_model(X, y, model):
    loss = list()
    for i in range(25):
        # fit model for one epoch on this sequence
        hist = model.fit(X, y, batch_size=1000, verbose=1, epochs=1, validation_split=0.2)
        loss.append(hist.history['loss'][0])
    return loss

# Driver Code
results = pd.DataFrame()
model_bilstm_lstm = get_bilstm_lstm_model(input_dim, output_dim, input_length, n_tags)

# Plot model after building it
plot_model(model_bilstm_lstm, to_file='model_plot.png', show_shapes=True, show_layer_names=True)

# Example data for training
train_tokens = np.zeros((100, input_length))  # Replace with actual data
train_tags = np.zeros((100, input_length, n_tags))  # Replace with actual data

results['with_add_lstm'] = train_model(train_tokens, train_tags, model_bilstm_lstm)

##################### Original #######################

# import os
# import pandas as pd

# # Set the absolute path to the CSV file
# csv_path = r'C:\Users\Adarsha Kumar\Downloads\WhatsApp Chat Analysis\WhatsApp_Chat_Analysis.csv'

# # Check if the file exists
# if not os.path.exists(csv_path):
#     print(f"Error: The file {csv_path} does not exist.")
# else:
#     # Read the CSV file
#     df = pd.read_csv(csv_path, encoding='utf-8')
#     print("CSV file loaded successfully.")
#     print(df.head())

#     # Combine all messages into a single text
#     text = ' '.join(df['Message'].dropna().tolist())

#     import spacy
#     from spacy import displacy

#     # Load the SpaCy model
#     nlp = spacy.load('en_core_web_sm')

#     # Process the text to create a Doc object
#     doc = nlp(text)

#     # Render the named entities in the text
#     html = displacy.render(doc, style='ent')

#     # Save the HTML to a file
#     with open('entities.html', 'w', encoding='utf-8') as f:
#         f.write(html)

#     print("Named entity recognition visualization saved to 'entities.html'.")


############ Ends Here #############################

import os
import pandas as pd
import spacy
from spacy import displacy

# Set the absolute path to the CSV file
csv_path = r'C:\Users\Adarsha Kumar\Downloads\WhatsApp Chat Analysis\WhatsApp_Chat_Analysis.csv'

# Check if the file exists
if not os.path.exists(csv_path):
    print(f"Error: The file {csv_path} does not exist.")
else:
    # Read the CSV file
    df = pd.read_csv(csv_path, encoding='utf-8')
    print("CSV file loaded successfully.")
    print(df.head())

    # Create a table with date, time, author, and message
    df_table = df[['Date', 'Time', 'Author', 'Message']].dropna()
    
    # Group messages by author
    author_groups = df_table.groupby('Author')

    # Load the SpaCy model
    nlp = spacy.load('en_core_web_sm')

    # Create a directory to save the NER results for each author
    output_dir = 'NER_Results'
    os.makedirs(output_dir, exist_ok=True)

    # Process and save NER results for each author separately
    for author, group in author_groups:
        # Initialize HTML content for the author
        author_html = f'<h2 style="font-family: Arial, sans-serif; color: #333;">Named Entity Recognition for {author}</h2>'
        author_html += '''
        <table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
            <thead style="background-color: #f2f2f2;">
                <tr>
                    <th style="padding: 8px; text-align: left;">Date</th>
                    <th style="padding: 8px; text-align: left;">Time</th>
                    <th style="padding: 8px; text-align: left;">Message</th>
                    <th style="padding: 8px; text-align: left;">Entities</th>
                </tr>
            </thead>
            <tbody>
        '''
        
        # Process each message individually
        for index, row in group.iterrows():
            date = row['Date']
            time = row['Time']
            message = row['Message']
            
            # Process the message with SpaCy
            doc = nlp(message)
            
            # Render the named entities in the message
            entities_html = displacy.render(doc, style='ent', jupyter=False, options={'compact': True})
            
            # Add a row to the table with the message and its entities
            row_style = 'background-color: #f9f9f9;' if index % 2 == 0 else ''
            author_html += f'''
                <tr style="{row_style}">
                    <td style="padding: 8px; text-align: left; vertical-align: top;">{date}</td>
                    <td style="padding: 8px; text-align: left; vertical-align: top;">{time}</td>
                    <td style="padding: 8px; text-align: left; vertical-align: top;">{message}</td>
                    <td style="padding: 8px; text-align: left; vertical-align: top;">{entities_html}</td>
                </tr>
            '''
        
        author_html += '''
            </tbody>
        </table>
        '''
        
        # Save the HTML to a file for each author
        output_path = os.path.join(output_dir, f'{author}_entities.html')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(author_html)

        print(f"Named entity recognition table saved for {author} to '{output_path}'.")
    speak("File Generated...and categorize entities like people, organizations, locations, and dates within text, offering a detailed understanding of the content.")
    print("Named entity recognition processing completed for all authors.")




########## 3rd #########################
# import os
# import pandas as pd
# import spacy
# from spacy import displacy
# import pdfkit

# # Set the absolute path to the CSV file
# csv_path = r'C:\Users\Adarsha Kumar\Downloads\WhatsApp Chat Analysis\WhatsApp_Chat_Analysis.csv'

# # Set the path to the wkhtmltopdf executable
# path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
# config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

# # Check if the file exists
# if not os.path.exists(csv_path):
#     print(f"Error: The file {csv_path} does not exist.")
# else:
#     # Read the CSV file
#     df = pd.read_csv(csv_path, encoding='utf-8')
#     print("CSV file loaded successfully.")
#     print(df.head())

#     # Create a table with date, time, author, and message
#     df_table = df[['Date', 'Time', 'Author', 'Message']].dropna()
    
#     # Group messages by author
#     author_groups = df_table.groupby('Author')

#     # Load the SpaCy model
#     nlp = spacy.load('en_core_web_sm')

#     # Create a directory to save the NER results for each author
#     output_dir = 'NER_Results'
#     os.makedirs(output_dir, exist_ok=True)

#     # Process and save NER results for each author separately
#     for author, group in author_groups:
#         # Initialize HTML content for the author
#         author_html = f'<h2 style="font-family: Arial, sans-serif; color: #333;">Named Entity Recognition for {author}</h2>'
#         author_html += '''
#         <table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
#             <thead style="background-color: #f2f2f2;">
#                 <tr>
#                     <th style="padding: 8px; text-align: left;">Date</th>
#                     <th style="padding: 8px; text-align: left;">Time</th>
#                     <th style="padding: 8px; text-align: left;">Message</th>
#                     <th style="padding: 8px; text-align: left;">Entities</th>
#                 </tr>
#             </thead>
#             <tbody>
#         '''
        
#         csv_data = []

#         # Process each message individually
#         for index, row in group.iterrows():
#             date = row['Date']
#             time = row['Time']
#             message = row['Message']
            
#             # Process the message with SpaCy
#             doc = nlp(message)
            
#             # Render the named entities in the message
#             entities_html = displacy.render(doc, style='ent', jupyter=False, options={'compact': True})
#             entities_text = ', '.join([ent.text for ent in doc.ents])
            
#             # Add a row to the table with the message and its entities
#             row_style = 'background-color: #f9f9f9;' if index % 2 == 0 else ''
#             author_html += f'''
#                 <tr style="{row_style}">
#                     <td style="padding: 8px; text-align: centre; vertical-align: top;">{date}</td>
#                     <td style="padding: 8px; text-align: centre; vertical-align: top;">{time}</td>
#                     <td style="padding: 8px; text-align: centre; vertical-align: top;">{message}</td>
#                     <td style="padding: 8px; text-align: centre; vertical-align: top;">{entities_html}</td>
#                 </tr>
#             '''

#             # Prepare data for CSV
#             csv_data.append([date, time, message, entities_text])

#         author_html += '''
#             </tbody>
#         </table>
#         '''

#         # Save the HTML to a file for each author
#         output_html_path = os.path.join(output_dir, f'{author}_entities.html')
#         with open(output_html_path, 'w', encoding='utf-8') as f:
#             f.write(author_html)

#         print(f"Named entity recognition table saved for {author} to '{output_html_path}'.")

#         # Save the CSV file for each author
#         csv_output_path = os.path.join(output_dir, f'{author}_entities.csv')
#         csv_df = pd.DataFrame(csv_data, columns=['Date', 'Time', 'Message', 'Entities'])
#         csv_df.to_csv(csv_output_path, index=False, encoding='utf-8')
#         print(f"CSV file saved for {author} to '{csv_output_path}'.")

#         # Convert the HTML to PDF
#         pdf_output_path = os.path.join(output_dir, f'{author}_entities.pdf')
#         try:
#             pdfkit.from_file(output_html_path, pdf_output_path, configuration=config)
#             print(f"PDF file saved for {author} to '{pdf_output_path}'.")
#         except Exception as e:
#             print(f"Error converting HTML to PDF for {author}: {e}")

#     print("Named entity recognition processing completed for all authors.")
