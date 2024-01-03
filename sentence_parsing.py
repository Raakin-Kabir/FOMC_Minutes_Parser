import nltk, random
from nltk.corpus import words
from nltk.corpus import wordnet 
from tika import parser
from num2words import num2words
import re
from datetime import datetime
import xlsxwriter
from nltk.tokenize.treebank import TreebankWordDetokenizer

# STEPS IN PREPROCESSING 
# (1) Remove any footnotes
# (2) Remove any "." that has some word not in NLTK corpus of words (e.g. middle names) or is only one letter (e.g. "a.m.")
# (3) Tokenize into sentences with NLTK sent_tokenize and make it all lowercase
# (4) Convert each number to its character equivalent (including time!)
# (5) Remove blank lines
# (6) Take random lines

# What needs to be done now is take only the relevant sections...
# This means starting from "Developments in Financial Markets and Open Market Operations"
# until reachiung "Voting for this action!"


# First, load in the raw text using tika parser
# .txt file is ued because PDF file is less parsable
raw = parser.from_file('fomc_minutes_2017_hand.txt')
content = raw['content']
titles = """Developments in Financial Markets and Open Market Operations, 
            Staff Review of the Economic Situation,
            Staff Review of the Financial Situation,
            Staff Economic Outlook,
            Committee Policy Action
            """

# Take only the relevant pieces of text
index = content.index("Developments in Financial Markets and Open Market Operations")
index_2 = content.index("Voting for this action:")
content = content[index:index_2]

# Remove any footnotes!
index = 0
while index < len(content):
    # if it's a footnote, remove it
    if (content[index].isnumeric() and content[index - 1] != ' ' and not content[index-1].isnumeric() and content[index-1] != ":"): 
        content = content[:index] + content[index+1:] # remove the footnote
    else:
        index += 1

content = content.replace("\n\n", ".\n\n")
content = content.replace("..", ".")
content = content.replace(":.", ":")
content = content.replace(",.", ",")
# Adding periods to section headers
content = content.replace("Developments in Financial Markets and Open Market Operations", "Developments in Financial Markets and Open Market Operations.")
content = content.replace("Staff Review of the Economic Situation", "Staff Review of the Economic Situation.")
content = content.replace("Staff Review of the Financial Situation", "Staff Review of the Financial Situation.")
content = content.replace("Staff Economic Outlook", "Staff Economic Outlook.")
content = content.replace("Committee Policy Action", "Committee Policy Action.")


# Remove any "." that has some word not in the NLTK corpus of words (e.g. middle names) or is only one letter (e.g. "a.m.")
index = 0
word_text = re.split(r'[\s]\s*', content)
vocab_1 = set(words.words())
vocab_2 = set(wordnet.words())
vocab = vocab_1.union(vocab_2)

while index < len(word_text):
    if ("." in word_text[index]):
        word_split = word_text[index].split(".")
        word = nltk.word_tokenize(word_split[0])
        if (len(word_split[0]) == 1 or (len(word) == 0) or (word[0].lower() not in vocab) or (not word[0].isnumeric())): # if right before the period is not a word or only one letter
            if (len(word) == 0 or (word[0].upper() != word[0] and word[0].lower() != word[0])):
                if (word[0] not in titles):
                    new_word = word_text[index].replace(".", "") # take away the "."s for the proper replacement
                    word_text[index] = new_word
    index += 1

content = " ".join(word_text)

# Add periods to sections of text that are divided by blank lines (if there isn't any other punctuation)
# So section titles end up being separate from paragraphs
# Other punctuations can get messed up with this, so fix them by removing periods that come right after


content = content.replace("\n\n", ".\n\n")
content = content.replace("..", ".")
content = content.replace(":.", ":")
content = content.replace(",.", ",")


# Tokenize the text into sentences
sent_text = nltk.sent_tokenize(content)



# lowercase all text
index = 0
while index < len(sent_text):
    sent_text[index] = sent_text[index].lower()
    index += 1

# Convert time stuff

# This method checks if something is in time format
def is_valid_time_format(input_str):
    try:
        datetime.strptime(input_str, '%I:%M')
        return True
    except ValueError:
        return False



index = 0
while index < len(sent_text):
    sentence = sent_text[index]
    words = sentence.split(' ')
    index_2 = 0
    for word in words:
        if is_valid_time_format(word): # if something is in valid time format, rewrite it
            dt_object = datetime.strptime(word, '%I:%M')
            hour_in_words = num2words(dt_object.strftime("%I"))
            minutes_in_words = num2words(dt_object.strftime("%M")) 
            if (minutes_in_words == "zero"): # NOTE, there is an extra space after the time (e.g. "one  pm")
                minutes_in_words = ""
            time_in_words = f"{hour_in_words} {minutes_in_words}"
            words[index_2] = time_in_words
        index_2 += 1
    sent_text[index] = ' '.join(words) # joining back the split elements of the sentence
    index += 1



# Convert all numbers to their alphabetical equivalent!
index = 0
while index < len(sent_text):
    sentence = sent_text[index]
    numbers = re.findall(r'\d+', sentence)
    for number in sorted(numbers, reverse=True): # using reverse so that digits inside of large numbers don't get changed first
        num_to_word = num2words(number)
        sentence = sentence.replace(number, num_to_word)
    sent_text[index] = sentence
    index += 1

# Replace newline characters with spaces
index = 0
while index < len(sent_text):
    sent_text[index] = sent_text[index].replace("\n", " ")
    index += 1





# Removing any bullet points
bullet_points = ["a", "b", "c", "d", "e", "i", "ii", "iii", "iv"]
index = 0
while index < len(sent_text):
    words = sent_text[index].split(" ")
    index_2 = 0
    for word in words:
        if word in bullet_points:
            words[index_2] = ""
        index_2 += 1
    words = " ".join(words)
    sent_text[index] = words
    index += 1


# Remove any blank lines
index = 0
while index < len(sent_text):
    if len(sent_text[index]) == 0:
        sent_text.pop(index)
    else:
        index += 1

sentences_num = len(sent_text)
numbers = list(range(0, sentences_num))
random_numbers = random.sample(numbers, 20)
# adding each sentence to an excel document 
workbook = xlsxwriter.Workbook("fomc_minute_2017.xlsx")
worksheet = workbook.add_worksheet("first_sheet")

worksheet.write(0, 0, "Sentence Number")
worksheet.write(0, 1, "Sentence")

row = 1
index = 0
for sentence in sent_text:
    if index in random_numbers:
        worksheet.write(row, 0, str(row))
        worksheet.write(row, 1, sentence)
        row += 1
    index += 1

workbook.close()
