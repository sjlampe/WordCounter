#! python 3
# WordCounter2.py -- A program that allows the user to enter a file name
# and count the frequency of the most common words or phrases of a defined
# length.

# This version in the College Files folder will always be the most recent.

import docx, os, pyperclip
os.chdir(r'C:\WordCounter')

# TODO: allow user to enter their own file directory, OR create a default
# directory via an installer (the latter would probably be more convenient).

# TODO: Set up a GUI

def getTextList(filename, phraseSize):
    # Gets a file and outputs a list with each token
    # word/phrase found in the file.
    
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    textString = '\n'.join(fullText)
    textList = textString.split()

    punctuations = '''!()[]{};:'"\,<>./?@#$%^&*_~'''
    for i in range(len(textList)):
        textList[i] = textList[i].lower()
    for i in range(len(textList)):
        noPunctString = ''
        for char in textList[i]:
            if char not in punctuations:
                noPunctString += char
        textList[i] = noPunctString

    if phraseSize == 1:
        print("Searching for single words...")
    else:
        print("Searching for " + str(phraseSize) + "-word phrases...")
        newList = [''] * len(textList)
        for i in range(len(textList) - (phraseSize)):
            for n in range(phraseSize):
                newList[i] += textList[(i + n)]
                newList[i] += ' '
        textList = newList
        
    return(textList)

def mostCommonWords(textlist,numWords):
    # Gets a textlist and finds the top numWords most
    # common words/phrases in it, prints results to
    # screen and copies results to clipboard.
    
    wordFreq = {}
    for i in range(len(textlist)):
        if textlist[i] in wordFreq.keys():
            wordFreq[textlist[i]] += 1
        else:
            wordFreq[textlist[i]] = 1 

    wordFreqKeys = list(wordFreq.keys())
    topWords = ['',]
    topCounts = [0,]
    for i in range(len(wordFreqKeys)):
        for x in range(len(topCounts)):
            if wordFreq[wordFreqKeys[i]] > topCounts[x]:
                topWords.insert(x,wordFreqKeys[i])
                topCounts.insert(x,wordFreq[wordFreqKeys[i]])
                break

    # TODO: Get rid of unnecessary words like a, the, and, etc.

    topWords = topWords[0:int(numWords)]
    topCounts = topCounts[0:int(numWords)]
    
    clipboardContents = ''
    for i in range(len(topWords)):
        clipboardContents += topWords[i]
        clipboardContents += ' '
        clipboardContents += str(topCounts[i])
        clipboardContents += '\n'
    pyperclip.copy(clipboardContents)

    print("WORD/PHRASE:".rjust(41) + "   FREQUENCY")
    for i in range(len(topCounts)):
        print(topWords[i].rjust(40) + ": " + str(topCounts[i]).rjust(5))

# TODO: Use Tkinter for GUI and main program loop instead

def main():
    # Organizes and handles prompts.

    print("What file would you like to use?:")
    print("(Enter the name exactly as it is listed, capitals and file extensions included. Example: Hello.docx)")
    file = input()

    while file == '': # TODO: have this loop check for a valid file name 
        # Catches a bug where if the user presses enter while a
        # previous file is processing, the program will not let
        # them enter a file name.
        print("Please enter a file name.")
        file = input()
    
    print('How many words would you like per phrase?')
    phraseSize = int(input())
    print('How many of the most common words/phrases would you like to look for?')
    numWords = input()

    textList = getTextList(file, phraseSize)
    mostCommonWords(textList, numWords)

    print("Results have been copied to clipboard.")
    print('')
    print("This program will repeat when you press enter.")
    x = input()

while True: # Main program loop.
    main()
