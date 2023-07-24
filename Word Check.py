#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# Importing the necessary packages. If you don't have them, you need to install them.

import docx
from collections import Counter
import re


def main():
    
# First, we define a minimum word length of 4 letters. You can change it if desired.

    while True:
        min_word_length = 4
        
# Now, we ask the user to enter the file path to be analyzed. If they don't want to analyze, they are requested to enter the word "exit" to close the program.

        try:
            file_path = input("Enter the file path (or type 'exit' to quit): ")

            if file_path.lower() == "exit":
                print("Exiting the program...")
                break

# We also added an error message in case the file or the path is not valid.

            if not file_path.endswith('.docx'):
                raise ValueError("Invalid file format. Please provide a valid .docx file.\n")

# Using the functions from the docx package, we read and split the document into paragraphs.

            doc = docx.Document(file_path)
            paragraphs = get_paragraphs(doc.paragraphs)

# After that, we create a list with the paragraphs, naming them to locate them in the document.
        
            for i in range(len(paragraphs) - 1):
                paragraphs_set_1 = paragraphs[i]
                paragraphs_set_2 = paragraphs[i + 1]

# Then, we perform the extraction of words, creating lists.

                words_list_1 = extract_words(paragraphs_set_1, min_length=min_word_length)
                words_list_2 = extract_words(paragraphs_set_2, min_length=min_word_length)

# Now, we compare the lists to identify the repeated words.

                repeated_words = set(words_list_1) & set(words_list_2)

# Finally, we show the user the repeated words in the paragraphs.

                if repeated_words:
                    print(f"\nRepeated words between paragraph {i+1} and the next paragraph {i+2}: {', '.join(repeated_words)}")
                else:
                    print(f"\nNo repeated words between paragraph {i+1} and paragraph {i+2}.")

# To avoid unnecessary interruption, we added some error messages.

        except FileNotFoundError:
            print("File not found. Please check the file path.\n")
        except ValueError as ve:
            print(ve)

            
if __name__ == "__main__":
    main()


# In[ ]:




