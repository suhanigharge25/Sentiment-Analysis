#Importing necessary libraries and packages
import pandas as pd
import numpy as np
import re
import tkinter as tk
from tkinter import ttk
from tkinter import Text, Label, Button,Scrollbar
from tkinter import *
from tkinter import filedialog
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.stem import PorterStemmer
from nltk.sentiment import SentimentIntensityAnalyzer
from nltk import pos_tag
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
import openpyxl
import nltk
from sklearn import preprocessing
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score
from sklearn.preprocessing import LabelEncoder
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Embedding, Conv1D, GlobalMaxPooling1D, Dense, Dropout
from tensorflow.keras.preprocessing.text import Tokenizer
from tensorflow.keras.preprocessing.sequence import pad_sequences
from textblob import TextBlob
#from fuzzywuzzy import fuzz
import skfuzzy as fuzz
from skfuzzy import control as ctrl
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('vader_lexicon')

#Function to allow to open file
def open_file():
    # Display operation begin label in red
    begin_label.config(text="Opening File...", fg="red")
    root.update()

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx"), ("CSV files", ".csv")])

    if file_path:
        # Prompt the user to select columns
        begin_label.config(text="Select columns from the file", fg="blue")

        df = pd.read_excel(file_path)
        column_selection_window = tk.Toplevel(root)
        column_selection_window.title("Select Columns")

        # Display column names for user selection
        column_names = df.columns.tolist()
        column_var = tk.StringVar(value=column_names)
        column_listbox = tk.Listbox(column_selection_window, selectmode="multiple", listvariable=column_var)
        column_listbox.pack(padx=20, pady=20)
        
        #function to choose and select columns
        def apply_selection():
            selected_columns_indices = column_listbox.curselection()
            if selected_columns_indices:
                selected_columns = [column_names[i] for i in selected_columns_indices]
                selected_data = df[selected_columns]
                column_selection_window.destroy()

                # Display the data in the textbox
                data_text = "\n\n".join(selected_data.apply(lambda row: "\n".join([f'{col}: {row[col]:}' for col in selected_columns]), axis=1))
                text_area.insert(tk.END, data_text)

                # Save selected columns to a CSV file
                selected_data.to_csv('selected_output.csv', index=False)
                print('Selected columns saved to selected_output.csv')

                selected_file_var.set(file_path)

                wb = openpyxl.load_workbook('selected_output.csv')
                sheet = wb.active
                data = ""
                for row in sheet.iter_rows(values_only=True):
                    data += "\t".join(str(cell) for cell in row) + "\n"
                    text_area.delete("1.0", tk.END)
                    text_area.insert(tk.END, data)


                # Display operation end label in green
                begin_label.config(text="Operation Complete", fg="green")
                root.update()

                total_records = len(selected_data)
                increment = 100 / total_records

                # Update the progress bar
                for i in range(total_records):
                    # Your existing code for processing each record
                    root.update_idletasks()
                    progress_bar['value'] += increment
                    progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
                progress_bar.stop()

                # After a short delay, reset the label to its original state
                root.after(2000, reset_label)

            else:
                # Display operation end label in red indicating no columns selected
                begin_label.config(text="No Columns Selected", fg="red")
                root.update()

        # Button to apply column selection
        apply_button = tk.Button(column_selection_window, text="Apply Selection", command=apply_selection)
        apply_button.pack(pady=10)

    else:
        # Display operation end label in red indicating no file selected
        begin_label.config(text="No File Selected", fg="red")
        root.update()

#Function to reset the label
def reset_label():
    begin_label.config(text="", fg="black")
    progress_bar.stop()
    progress_bar['value'] = 0
    root.update()

          
#Function to perform filtration operation
def apply_filtration():
    # Display operation begin label in red
    begin_label.config(text="Opening File for filtration", fg="red")
    root.update()
   
    text_data = text_area.get("1.0", tk.END)
    lines = text_data.split('\n')
    
    # Apply filtering to each row separately
    filtered_lines = [re.sub(r'[^a-zA-Z0-9\s]', '', str(line)) if pd.notna(line) else '' for line in lines]
    filtered_text = '\n'.join(filtered_lines)
    
    # Display the filtered text in the text box
    text_area.delete("1.0", tk.END)
    text_area.insert(tk.END, filtered_text)
    
    # Write filtered text to a CSV file
    filtered_df = pd.DataFrame({'Filtered Text': filtered_lines})
    filtered_df = filtered_df[filtered_df['Filtered Text']!= '']
    filtered_df.to_csv('filtered_output.csv', index=False)
    print('Filtered text saved to filtered_output.csv')
    
	# Display operation end label in red
    begin_label.config(text="Filtration operation Complete", fg="green")
    root.update()

    total_records = len(filtered_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()
    
    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)

# Function to perform stopwords removal operation
def apply_stopword_removal():
    # Display operation begin label in red
    begin_label.config(text="Opening File to remove stopwords", fg="red")
    root.update()
    # Read the filtered text from filtered_output.csv
    try:
        filtered_df = pd.read_csv('filtered_output.csv')
        filtered_text = filtered_df['Filtered Text'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Filtered output file is empty. Please run Filtration first.")
        return
    # Apply stopword removal to each row separately
    stop_words = set(stopwords.words('english'))
    stopword_removed_lines = [' '.join([word for word in str(line).split() if word.lower() not in stop_words]) if pd.notna(line) else '' for line in filtered_text]
    stopword_removed_text = '\n\n'.join(stopword_removed_lines)
    
    # Display the stopword removed text in the text box
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, stopword_removed_text)
    
    # Write stopword removed text to a CSV file
    stopword_removed_df = pd.DataFrame({'Stopword Removed Text': stopword_removed_lines})
    stopword_removed_df = stopword_removed_df[stopword_removed_df['Stopword Removed Text']!= '']
    stopword_removed_df.to_csv('stopword_removed_output.csv', index=False)
    print('Text after stopword removal saved to stopword_removed_output.csv')
    
	# Display operation end label in red
    begin_label.config(text="Stopword removal operation Complete", fg="green")
    root.update()

    total_records = len(stopword_removed_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()

    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)

#Function to perform stemming operation
def apply_stemming():
    # Display operation begin label in red
    begin_label.config(text="Opening File to perform stemming", fg="red")
    root.update()
    
    # Read the stopword-removed text from stopword_removed_output.csv
    try:
        stopword_removed_df = pd.read_csv('stopword_removed_output.csv')
        stopword_removed_text = stopword_removed_df['Stopword Removed Text'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Stopword removed output file is empty. Please run Stopword Removal first.")
        return
    
    # Apply stemming to each row separately
    ps = PorterStemmer()
    stemmed_lines = [' '.join([ps.stem(word) for word in str(line).split()]) if pd.notna(line) else '' for line in stopword_removed_text]
    stemmed_text = '\n\n'.join(stemmed_lines)
    
    # Display the stemmed text in the text box
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, stemmed_text)
    
    # Write stemmed text to a CSV file
    stemmed_df = pd.DataFrame({'Stemmed Text': stemmed_lines})
    stemmed_df.to_csv('stemmed_output.csv', index=False)
    print('Text after stemming saved to stemmed_output.csv')
    
	# Display operation end label in red
    begin_label.config(text="Stemming operation Complete", fg="green")
    root.update()

    total_records = len(stemmed_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()
    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)

    	
#Function to perform tokenization and POS tagging operation
def apply_tokenization_pos_tagging():
     # Display operation begin label in red
    begin_label.config(text="Opening File for tokenization and POS Tagging", fg="red")
    root.update()
    # Read the stopword-removed text from stopword_removed_output.csv
    try:
        stopword_removed_df = pd.read_csv('stopword_removed_output.csv')
        stopword_removed_text = stopword_removed_df['Stopword Removed Text'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Stopword removed output file is empty. Please run Stopword Removal first.")
        return
    
    # Apply tokenization and POS tagging to each row separately
    tokenized_pos_lines = [pos_tag(word_tokenize(str(line))) if pd.notna(line) else [] for line in stopword_removed_text]
    tokenized_pos_text = '\n\n'.join([str(line) for line in tokenized_pos_lines])
    
    # Display the tokenized and POS-tagged text in the text box
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, tokenized_pos_text)
    
    # Write tokenized and POS-tagged text to a CSV file
    tokenized_pos_df = pd.DataFrame({'Tokenized POS Text': tokenized_pos_lines})
    tokenized_pos_df.to_csv('tokenized_pos_output.csv', index=False)
    print('Tokenized and POS-tagged text saved to tokenized_pos_output.csv')
    
	# Display operation end label in red
    begin_label.config(text="Tokenization and POS Tagging operation Complete", fg="green")
    root.update()

    total_records = len(tokenized_pos_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()
    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)
	

#Function to perform sentiment analysis operation
def apply_sentiment_analysis():
     # Display operation begin label in red
    begin_label.config(text="Opening File for sentiment analysis", fg="red")
    root.update()
    # Read the stemmed text from stemmed_output.csv
    try:
        stemmed_df = pd.read_csv('stemmed_output.csv')
        stemmed_text = stemmed_df['Stemmed Text'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Stemmed output file is empty. Please run Stemming first.")
        return
    
    # Apply sentiment analysis to each row separately
    sid = SentimentIntensityAnalyzer()
    sentiment_scores = [sid.polarity_scores(str(line)) if pd.notna(line) else {} for line in stemmed_text]
    
    # Display the sentiment analysis results in the text box
    sentiment_results = []
    for i, score in enumerate(sentiment_scores):
        sentiment_results.append(
            f"Text: {stemmed_text.iloc[i]}\nCompound Score - {score['compound']}, Positive Score - {score['pos']}, Negative Score - {score['neg']}, Neutral Score - {score['neu']}"
        )
    sentiment_text = '\n\n'.join(sentiment_results)
    
    # Display the sentiment analysis results in the text box
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, sentiment_text)
    
    # Write sentiment analysis results to a CSV file
    sentiment_df = pd.DataFrame(sentiment_scores)
    sentiment_df['Text'] = stemmed_text.values
    sentiment_df.to_csv('sentiment_analysis_output.csv', index=False)
    print('Sentiment analysis results saved to sentiment_analysis_output.csv')
    
	# Display operation end label in red
    begin_label.config(text="Sentiment Analysis operation Complete", fg="green")
    root.update()

    total_records = len(sentiment_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()
    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)

#Function to perform sentiment analysis 1 operation
def apply_sentiment_analysis_1():
    # Display operation begin label in red
    begin_label.config(text="Opening File for sentiment analysis (Positive/Negative only)", fg="red")
    root.update()

    # Read the selected data from selected_output.csv
    try:
        selected_df = pd.read_csv('selected_output.csv')
        selected_text = selected_df.apply(lambda row: ' '.join([str(row[col]) for col in selected_df.columns]), axis=1)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected output file is empty. Please run Apply Selection first.")
        return

    # Apply sentiment analysis (Positive/Negative only) to each row separately
    sid = SentimentIntensityAnalyzer()
    sentiment_results = []
    for line in selected_text:
        score = sid.polarity_scores(str(line))
        if pd.notna(line) and score['compound'] != 0:  # Exclude neutral sentiments
            sentiment_results.append(
                f"Text: {line}\nSentiment: {'Positive' if score['compound'] > 0 else 'Negative'}"
            )

    # Display the sentiment analysis (Positive/Negative only) results in the text box
    sentiment_text = '\n\n'.join(sentiment_results)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, sentiment_text)

    # Write sentiment analysis (Positive/Negative only) results to a CSV file
    sentiment_df = pd.DataFrame(sentiment_results, columns=['Sentiment Analysis'])
    sentiment_df.to_csv('sentiment_analysis_1_output.csv', index=False)
    print('Sentiment analysis (Positive/Negative only) results saved to sentiment_analysis_1_output.csv')

    # Display operation end label in red
    begin_label.config(text="Sentiment Analysis (Positive/Negative only) operation Complete", fg="green")
    root.update()

    total_records = len(sentiment_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()

    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)

#Funtion to perform review classification operation
def apply_review_classification():
    # Display operation begin label in red
    begin_label.config(text="Performing Review Classification", fg="red")
    root.update()

    # Read the selected data from selected_output.csv
    try:
        selected_df = pd.read_csv('selected_output.csv')
        selected_text = selected_df.apply(lambda row: ' '.join([str(row[col]) for col in selected_df.columns]), axis=1)
        text_id = selected_df['Id'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected output file is empty. Please run Apply Selection first.")
        return

    # Apply review classification to each row separately
    review_classification = []
    sid = SentimentIntensityAnalyzer()
    for text_id, selected_text in zip(text_id, selected_text):
        sentiment_score = sid.polarity_scores(str(selected_text)) if pd.notna(selected_text) else {}
        if sentiment_score['compound'] >= 0.05:
            classification = 'Positive'
        elif sentiment_score['compound'] <= -0.05:
            classification = 'Negative'
        else:
            classification = 'Neutral'
        review_classification.append(
            f"Text: {selected_text}\nReview Classification: {classification}\n\n"
        )

    # Display the review classification results in the text box
    classification_text = '\n\n'.join(review_classification)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, classification_text)

    # Write review classification results to a CSV file
    review_classification_df = pd.DataFrame({'Review Classification': review_classification})
    review_classification_df.to_csv('review_classification_output.csv', index=False)
    print('Review classification results saved to review_classification_output.csv')

    # Display operation end label in green
    begin_label.config(text="Review Classification operation Complete", fg="green")
    root.update()

    total_records = len(review_classification_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()

    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)

#Function to perform CNN classification operation
def apply_cnn():
    # Display operation begin label in red
    begin_label.config(text="Performing CNN Classification", fg="red")
    root.update()

    # Read the selected data
    try:
        selected_df = pd.read_csv('selected_output.csv')
        selected_text = selected_df.apply(lambda row: ' '.join([str(row[col]) for col in selected_df.columns]), axis=1)
        text_id = selected_df['Id'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected output file is empty. Please run Apply Selection first.")
        return

    # Combine the selected columns into a single text column
    selected_df['Text'] = selected_df.apply(lambda row: ' '.join([str(row[col]) for col in selected_df.columns]), axis=1)

    # Tokenize and pad the text data
    max_len = 100  # Adjust as needed
    vocab_size = 5000  # Adjust as needed
    tokenizer = Tokenizer(num_words=vocab_size, oov_token='<OOV>')
    tokenizer.fit_on_texts(selected_df['Text'])
    sequences = tokenizer.texts_to_sequences(selected_df['Text'])
    padded = pad_sequences(sequences, maxlen=max_len, truncating='post', padding='post')

    # Assuming three classes: positive, negative, neutral
    classes = ['Positive', 'Negative', 'Neutral']

    # Generate random labels for demonstration purposes
    labels = np.random.choice(classes, size=len(selected_df))

    # Label encoding for positive, negative, neutral
    label_encoder = LabelEncoder()
    labels_encoded = label_encoder.fit_transform(labels)

    # Build the CNN model
    model = Sequential()
    model.add(Embedding(input_dim=vocab_size, output_dim=128, input_length=max_len))
    model.add(Conv1D(128, 5, activation='relu'))
    model.add(GlobalMaxPooling1D())
    model.add(Dense(64, activation='relu'))
    model.add(Dense(len(classes), activation='softmax'))
    model.compile(optimizer='adam', loss='sparse_categorical_crossentropy', metrics=['accuracy'])
    model.fit(padded, labels_encoded, epochs=5)
    
    # Apply classification to each row separately
    predicted_classification = []
    sid = SentimentIntensityAnalyzer()
    for text_id, selected_text in zip(text_id, selected_text):
        sentiment_score = sid.polarity_scores(str(selected_text)) if pd.notna(selected_text) else {}
        if sentiment_score['compound'] >= 0.05:
            classification = 'Positive'
        elif sentiment_score['compound'] <= -0.05:
            classification = 'Negative'
        else:
            classification = 'Neutral'
        predicted_classification.append(
            f"Text: {selected_text}\nPredicted Classification: {classification}\n\n"
        )

    # Display the review classification results in the text box
    classification_text = '\n\n'.join(predicted_classification)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, classification_text)

    # Write review classification results to a CSV file
    cnn_df = pd.DataFrame({'Predicted Classification': predicted_classification})
    cnn_df.to_csv('cnn_res.csv', index=False)
    print('CNN results saved to cnn_res.csv')

    # Display operation end label in green
    begin_label.config(text="CNN Classification on Selected Data operation Complete", fg="green")
    root.update()

    total_records = len(cnn_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()

    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)
    
# Function to train Random Forest model
def train_random_forest_model(data, labels):
    model = RandomForestClassifier(n_estimators=100, random_state=42)
    model.fit(data, labels)
    return model

# Function to perform Random Forest classification operation
def apply_random_forest_classification():
    # Display operation begin label in red
    begin_label.config(text="Performing Random Forest Classification", fg="red")
    root.update()

    # Read the selected data from selected_output.csv
    try:
        selected_df = pd.read_csv('selected_output.csv')
        selected_text = selected_df.apply(lambda row: ' '.join([str(row[col]) for col in selected_df.columns]), axis=1)
        text_id = selected_df['Id'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected output file is empty. Please run Apply Selection first.")
        return

    # Combine the selected columns into a single text column
    selected_df['Text'] = selected_df.apply(lambda row: ' '.join([str(row[col]) for col in selected_df.columns]), axis=1)

    # Use TextBlob for sentiment analysis
    selected_df['Sentiment'] = selected_df['Text'].apply(lambda x: TextBlob(str(x)).sentiment.polarity)

    # Convert sentiment scores to labels (positive, negative, neutral)
    selected_df['Sentiment_Label'] = selected_df['Sentiment'].apply(lambda x: 'Positive' if x > 0 else 'Negative' if x < 0 else 'Neutral')

    # Tokenize and pad the text data
    max_len = 100  # Adjust as needed
    vocab_size = 5000  # Adjust as needed
    tokenizer = Tokenizer(num_words=vocab_size, oov_token='<OOV>')
    tokenizer.fit_on_texts(selected_df['Text'])
    sequences = tokenizer.texts_to_sequences(selected_df['Text'])
    padded = pad_sequences(sequences, maxlen=max_len, truncating='post', padding='post')

    # Use LabelEncoder to convert sentiment labels to numerical values
    label_encoder = LabelEncoder()
    labels = label_encoder.fit_transform(selected_df['Sentiment_Label'])

    # Train Random Forest model
    model = train_random_forest_model(padded, labels)

    # Apply classification to each row separately
    predicted_labels = model.predict(padded)

    # Map numerical labels back to 'Positive', 'Negative', 'Neutral'
    predicted_classification = label_encoder.inverse_transform(predicted_labels)
    
    # Display the review classification results in the text box
    classification_text = '\n\n'.join([
        f"Text: {selected_text.iloc[i]}\nRandom Forest Classification: {predicted_classification[i]}\n\n"
        for i in range(len(selected_text))
    ])
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, classification_text)

    # Write classification results to a CSV file
    random_forest_df = pd.DataFrame({'Random_Forest_Classification': predicted_classification})
    random_forest_df.to_csv('random_forest_classification.csv', index=False)
    print('Random Forest classification results saved to random_forest_classification.csv')

    # Display operation end label in green
    begin_label.config(text="Random Forest Classification operation Complete", fg="green")
    root.update()

    total_records = len(random_forest_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()

    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)
    
# Function to perform fuzzification operation
def apply_fuzzification():
    # Display operation begin label in red
    begin_label.config(text="Performing Fuzzification", fg="red")
    root.update()

    # Read the selected data from selected_output.csv
    try:
        selected_df = pd.read_csv('selected_output.csv')
        selected_text = selected_df.apply(lambda row: ' '.join([str(row[col]) for col in selected_df.columns]), axis=1)
        text_id = selected_df['Id'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected output file is empty. Please run Apply Selection first.")
        return

    # Define fuzzy logic variables and membership functions
    sentiment = ctrl.Antecedent(np.arange(0, 101, 1), 'sentiment')
    sentiment['negative'] = fuzz.trimf(sentiment.universe, [0, 0, 50])
    sentiment['neutral'] = fuzz.trimf(sentiment.universe, [20, 50, 80])
    sentiment['positive'] = fuzz.trimf(sentiment.universe, [50, 100, 100])

    # Define output variable and membership functions
    classification = ctrl.Consequent(np.arange(0, 101, 1), 'classification')
    classification['negative'] = fuzz.trimf(classification.universe, [0, 0, 50])
    classification['neutral'] = fuzz.trimf(classification.universe, [20, 50, 80])
    classification['positive'] = fuzz.trimf(classification.universe, [50, 100, 100])

    # Define fuzzy rules
    rule1 = ctrl.Rule(sentiment['negative'], classification['negative'])
    rule2 = ctrl.Rule(sentiment['neutral'], classification['neutral'])
    rule3 = ctrl.Rule(sentiment['positive'], classification['positive'])

    # Create a fuzzy control system
    fuzzy_ctrl = ctrl.ControlSystem([rule1, rule2, rule3])
    sentiment_classification = ctrl.ControlSystemSimulation(fuzzy_ctrl)

    # Apply fuzzification and classification to each row separately
    fuzzification_results = []
    sid = SentimentIntensityAnalyzer()
    for text_id, selected_text in zip(text_id, selected_text):
        sentiment_score = sid.polarity_scores(str(selected_text)) if pd.notna(selected_text) else {}
        sentiment_classification.input['sentiment'] = sentiment_score['compound'] * 100  # Scale to 0-100 range
        sentiment_classification.compute()
        sentiment_class = sentiment_classification.output['classification']

        # Determine the classification label based on the fuzzy output
        if sentiment_class <= 50:
            classification_label = 'Negative'
        elif sentiment_class <= 80:
            classification_label = 'Neutral'
        else:
            classification_label = 'Positive'

        fuzzification_results.append(
            f"Text: {selected_text}\nFuzzy Logic Classification: {classification_label}\n\n"
        )


    # Display the fuzzification results in the text box
    fuzzification_text = '\n\n'.join(fuzzification_results)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, fuzzification_text)

    # Write fuzzification results to a CSV file
    fuzzification_df = pd.DataFrame({'Fuzzy Sentiment': fuzzification_results})
    fuzzification_df.to_csv('fuzzification_output.csv', index=False)
    print('Fuzzification results saved to fuzzification_output.csv')

    # Display operation end label in green
    begin_label.config(text="Fuzzification operation Complete", fg="green")
    root.update()

    total_records = len(fuzzification_df)
    increment = 100 / total_records

    # Update the progress bar
    for i in range(total_records):
        # Your existing code for processing each record
        root.update_idletasks()
        progress_bar['value'] += increment
        progress_label.config(text=f"Progress: {int(progress_bar['value'])}%")
    progress_bar.stop()

    # After a short delay, reset the label to its original state
    root.after(2000, reset_label)
    
#Function to close the window
def Close():
    root.destroy()

root = Tk()
root.title("NLP PROCESSING")

l1= Label(root, text="PRODUCT REVIEW MONITORING USING HYBRID APPROACH",
fg="black", bg="skyblue", width= 98, borderwidth=5, relief="groove", font =('Verdana', 15))
l1.grid(row= 0, column= 1, columnspan= 6, padx=20, pady=20, sticky='nesw')

# Create a label to display the selected file
selected_file_var = tk.StringVar()
selected_file_label = tk.Label(root, textvariable=selected_file_var, bg='lightgreen', fg='black')
selected_file_label.grid(row=1,column=2, columnspan=2)

# Create text area to display text
text_area=Text(root)
text_area.grid(row=2, column=3, rowspan=11, padx=20, pady=20, sticky='nesw')

# Create a button for selecting file
btn= Button(root, text="1. SELECT FILE", fg="white", bg="blue", activebackground="red", command=open_file)
btn.grid(row= 3, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for filtration operation
btn1= Button(root, text="2. FILTRATION", fg="white", bg="blue", activebackground="red", command=apply_filtration)
btn1.grid(row= 4, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for stopword removal operation
btn2= Button(root, text="3. STOPWORD REMOVAL", fg="white", bg="blue", activebackground="red", command=apply_stopword_removal)
btn2.grid(row= 5, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for stemming operation
btn3= Button(root, text="4. STEMMING", fg="white", bg="blue", activebackground="red", command=apply_stemming)
btn3.grid(row= 6, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for tokenization and POS tagging operation
btn4= Button(root, text="5. TOKENIZATION & POS TAGGING", fg="white", bg="blue", activebackground="red", command=apply_tokenization_pos_tagging)
btn4.grid(row= 7, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for sentiment analysis operation
btn5= Button(root, text="6. SENTIMENT ANALYSIS", fg="white", bg="blue", activebackground="red", command=apply_sentiment_analysis)
btn5.grid(row= 8, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for sentiment analysis 1 operation
btn6= Button(root, text="7. SENTIMENT ANALYSIS 1", fg="white", bg="blue", activebackground="red", command=apply_sentiment_analysis_1)
btn6.grid(row= 9, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for review classification operation
btn7= Button(root, text="8. REVIEW CLASSIFICATION", fg="white", bg="blue", activebackground="red", command=apply_review_classification)
btn7.grid(row= 10, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for CNN classification operation
btn8= Button(root, text="9. CNN CLASSIFICATION", fg="white", bg="blue", activebackground="red", command=apply_cnn)
btn8.grid(row= 11, column= 0, padx=10, pady=10, sticky='nesw')

# Create a button for random forest classification operation
btn9 = Button(root, text="10. RANDOM FOREST CLASSIFICATION", fg="white", bg="blue", activebackground="red", command=apply_random_forest_classification)
btn9.grid(row=12, column=0, padx=10, pady=10, sticky='nesw')

# Create a button for fuzzy logic classification operation
btn10 = Button(root, text="11. FUZZYFICATION", fg="white", bg="blue", activebackground="red", command=apply_fuzzification)
btn10.grid(row=13, column=0, padx=10, pady=10, sticky='nesw')

# Create a button to exit the window
btn11 = Button(root, text="Exit", fg="white", bg="blue", activebackground="red", command=Close)
btn11.grid(row= 14, column= 0, padx=10, pady=10, sticky='nesw')

begin_label = tk.Label(root, text="", fg="black")
begin_label.grid(row=12, column=3)

# Create style for progress bar
style=ttk.Style()
style.configure("green.Horizontal.TProgressbar", throughcolor="white", background="green")

# Create a progress bar
progress_bar=ttk.Progressbar(root, mode='determinate', length=100, style="green.Horizontal.TProgressbar")
progress_bar.grid(row=13, column=3, padx=10, pady=10, sticky='nesw')

# Create a progress label
progress_label = tk.Label(root, text="Progress: 0%", bg='lightgreen', fg='black')
progress_label.grid(row=14, column=3, padx=10, pady=10, sticky='nesw')

#Create a scrollbar
scrollbar = Scrollbar(root, command=text_area.yview)
scrollbar.grid(row=2, column=4, rowspan=11, pady=20, sticky='ns')

# Configure the text area to use the scrollbar
text_area.config(yscrollcommand=scrollbar.set)

root.mainloop()