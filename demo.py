import pandas as pd
import re
import tkinter as tk
from tkinter import ttk
from tkinter import Scrollbar
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
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('vader_lexicon')
import tensorflow as tf
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Embedding, Conv1D, GlobalMaxPooling1D, Dense, Flatten
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score

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

        def apply_selection():
            selected_columns_indices = column_listbox.curselection()
            if selected_columns_indices:
                selected_columns = [column_names[i] for i in selected_columns_indices]
                selected_data = df[selected_columns]

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

                # Close the column selection window
                column_selection_window.destroy()

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

def reset_label():
    begin_label.config(text="", fg="black")
    progress_bar.stop()
    progress_bar['value'] = 0
    root.update()

          
#filtration
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

#stopwords removal
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

#stemming
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

    	
#tokenization
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
	

#sentiment analysis
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
#
def apply_sentiment_analysis_1():
    # Display operation begin label in red
    begin_label.config(text="Opening File for sentiment analysis (Pos/Neg Only)", fg="red")
    root.update()

    # Read the original data from the selected file
    try:
        df = pd.read_excel(selected_file_var.get())
        text_ids = df['Id'].astype(str)
        original_text = df['Text'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected file is empty. Please run '1. SELECT FILE' first.")
        return

    # Apply sentiment analysis to each row separately
    sid = SentimentIntensityAnalyzer()
    sentiment_results = []
    for text_id, text in zip(text_ids, original_text):
        sentiment_score = sid.polarity_scores(str(text)) if pd.notna(text) else {}
        # Display only positive and negative scores, exclude neutral
        if sentiment_score['compound'] > 0:
            sentiment_results.append(f"ID: {text_id}\nText: {text}\nPositive Score: {sentiment_score['pos']}")
        elif sentiment_score['compound'] < 0:
            sentiment_results.append(f"ID: {text_id}\nText: {text}\nNegative Score: {sentiment_score['neg']}")

    # Display the sentiment analysis results in the text box
    sentiment_text = '\n\n'.join(sentiment_results)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, sentiment_text)

    # Write sentiment analysis results to a CSV file (optional)
    # You can choose to save this information if needed

    # Display operation end label in red
    begin_label.config(text="Sentiment Analysis (Pos/Neg Only) operation Complete", fg="green")
    root.update()

    total_records = len(sentiment_results)
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
    
#review classification
def apply_review_classification():
    #
    # Display operation begin label in red
    begin_label.config(text="Opening File for review classification", fg="red")
    root.update()

    # Read the original data from the selected file
    try:
        df = pd.read_excel(selected_file_var.get())
        text_ids = df['Id'].astype(str)
        original_text = df['Text'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected file is empty. Please run '1. SELECT FILE' first.")
        return

    # Apply review classification to each row separately
    review_classification = []
    sid = SentimentIntensityAnalyzer()
    for text_id, text in zip(text_ids, original_text):
        sentiment_score = sid.polarity_scores(str(text)) if pd.notna(text) else {}
        if sentiment_score['compound'] >= 0.05:
            classification = 'Positive'
        elif sentiment_score['compound'] <= -0.05:
            classification = 'Negative'
        else:
            classification = 'Neutral'
        review_classification.append(
            f"ID: {text_id}\nText: {text}\nReview Classification: {classification}\n\n"
        )

    # Display the review classification results in the text box
    classification_text = '\n\n'.join(review_classification)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, classification_text)

    # Write review classification results to a CSV file
    classification_df = pd.DataFrame({'ID': text_ids, 'Review Classification': review_classification})
    classification_df.to_csv('review_classification_on_full_text_output.csv', index=False)
    print('Review classification results saved to review_classification output.csv')

    # Display operation end label in red
    begin_label.config(text="Review Classification operation Complete", fg="green")
    root.update()

    total_records = len(classification_df)
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
    
# Function to create and train a simple CNN
def train_cnn():
    begin_label.config(text="Training CNN...", fg="red")
    root.update()

    # Load your labeled dataset (Assuming you have a 'Label' column indicating categories)
    try:
        df = pd.read_excel(selected_file_var.get())
        text_data = df['Text'].astype(str)
        labels = df['Id'].astype(int)  # Assuming 'Label' column contains integer labels
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected file is empty. Please run '1. SELECT FILE' first.")
        return

    # Tokenize and pad sequences
    tokenizer = tf.keras.preprocessing.text.Tokenizer()
    tokenizer.fit_on_texts(text_data)
    sequences = tokenizer.texts_to_sequences(text_data)
    padded_sequences = tf.keras.preprocessing.sequence.pad_sequences(sequences)

    # Split the data into training and testing sets
    X_train, X_test, y_train, y_test = train_test_split(padded_sequences, labels, test_size=0.2, random_state=42)

    # Define the CNN model
    model = Sequential()
    model.add(Embedding(input_dim=len(tokenizer.word_index) + 1, output_dim=100, input_length=padded_sequences.shape[1]))
    model.add(Conv1D(128, 5, activation='relu'))
    model.add(GlobalMaxPooling1D())
    model.add(Dense(1, activation='sigmoid'))

    # Compile the model
    model.compile(optimizer='adam', loss='binary_crossentropy', metrics=['accuracy'])

    # Train the model
    model.fit(X_train, y_train, epochs=5, batch_size=32, validation_data=(X_test, y_test), verbose=2)

    begin_label.config(text="Training Complete", fg="green")
    root.update()
    
def test_cnn():
    begin_label.config(text="Testing CNN...", fg="red")
    root.update()

    # Load new data for testing
    try:
        new_data = pd.read_excel(selected_file_var.get())['Text'].astype(str)
    except pd.errors.EmptyDataError:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Selected file is empty. Please run '1. SELECT FILE' first.")
        return

    # Tokenize and pad sequences for new data
    new_sequences = tokenizer.texts_to_sequences(new_data)
    new_padded_sequences = tf.keras.preprocessing.sequence.pad_sequences(new_sequences)

    # Predict labels using the trained model
    predictions = model.predict(new_padded_sequences)

    # Display the results in the text box
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, f"Predictions:\n{predictions}")

    begin_label.config(text="Testing Complete", fg="green")
    root.update()
 
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

text_area=Text(root)
text_area.grid(row=2, column=3, rowspan=11, padx=20, pady=20, sticky='nesw')

btn= Button(root, text="1. SELECT FILE", fg="white", bg="blue", activebackground="red", command=open_file)
btn.grid(row= 3, column= 0, padx=10, pady=10, sticky='nesw')

btn1= Button(root, text="2. FILTRATION", fg="white", bg="blue", activebackground="red", command=apply_filtration)
btn1.grid(row= 4, column= 0, padx=10, pady=10, sticky='nesw')

btn2= Button(root, text="3. STOPWORD REMOVAL", fg="white", bg="blue", activebackground="red", command=apply_stopword_removal)
btn2.grid(row= 5, column= 0, padx=10, pady=10, sticky='nesw')

btn3= Button(root, text="4. STEMMING", fg="white", bg="blue", activebackground="red", command=apply_stemming)
btn3.grid(row= 6, column= 0, padx=10, pady=10, sticky='nesw')

btn4= Button(root, text="5. TOKENIZATION & POS TAGGING", fg="white", bg="blue", activebackground="red", command=apply_tokenization_pos_tagging)
btn4.grid(row= 7, column= 0, padx=10, pady=10, sticky='nesw')

btn5= Button(root, text="6. SENTIMENT ANALYSIS", fg="white", bg="blue", activebackground="red", command=apply_sentiment_analysis)
btn5.grid(row= 8, column= 0, padx=10, pady=10, sticky='nesw')

btn6= Button(root, text="7. SENTIMENT ANALYSIS 1", fg="white", bg="blue", activebackground="red", command=apply_sentiment_analysis_1)
btn6.grid(row= 9, column= 0, padx=10, pady=10, sticky='nesw')

btn7= Button(root, text="8. REVIEW CLASSIFICATION", fg="white", bg="blue", activebackground="red", command=apply_review_classification)
btn7.grid(row= 10, column= 0, padx=10, pady=10, sticky='nesw')

btn8= Button(root, text="9. TRAIN CNN", fg="white", bg="blue", activebackground="red", command=train_cnn)
btn8.grid(row= 11, column= 0, padx=10, pady=10, sticky='nesw')

btn9= Button(root, text="10. TEST CNN", fg="white", bg="blue", activebackground="red", command=test_cnn)
btn9.grid(row= 12, column= 0, padx=10, pady=10, sticky='nesw')

btn10= Button(root, text="Exit", fg="white", bg="blue", activebackground="red", command=Close)
btn10.grid(row= 13, column= 0, padx=10, pady=10, sticky='nesw')

begin_label = tk.Label(root, text="", fg="black")
begin_label.grid(row=12, column=3)

style=ttk.Style()
style.configure("green.Horizontal.TProgressbar", throughcolor="white", background="green")

progress_bar=ttk.Progressbar(root, mode='determinate', length=100, style="green.Horizontal.TProgressbar")
progress_bar.grid(row=13, column=3, padx=10, pady=10, sticky='nesw')

progress_label = tk.Label(root, text="Progress: 0%", bg='lightgreen', fg='black')
progress_label.grid(row=14, column=3, padx=10, pady=10, sticky='nesw')

scrollbar = Scrollbar(root, command=text_area.yview)
scrollbar.grid(row=2, column=4, rowspan=11, pady=20, sticky='ns')

# Configure the text area to use the scrollbar
text_area.config(yscrollcommand=scrollbar.set)

root.mainloop()