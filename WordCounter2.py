#! python 3
# WordCounter3.py -- A program that allows the user to enter a file name
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

    punctuations = '''!()[]{};:'"\,<>./?@#$%^&*_~–-'''
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

    connectorWords = ['the','of','and','in','to','a','by','it','that','is','for',
                      'as','was','which','this','its','their','be','on','at','has',
                      'but','from','they','are','into','were','have','an','or',
                      'we','all','any','but','our','not','with','i','me','he',
                      'her','him','his','hers','an','their','these','so','if','at',
                      'without','what','who','my','you','do','can','also','there',
                      'when','out','them','theirs','they','form','only','every','then',
                      'than','can','would','will','after','may','other','one','same',
                      'us','must','more','no','such','some','upon','though',
                      'although','nor','','had','been','up','most','way','thus',
                      'since','could','become','she','your','go','am','too','those']
    connectorPhrases2 = ['of the','in the','to the','by the','and the','for the',
                         'on the','with the', 'of all', 'at the','to be','from the'
                         ,'that the','of a','as a','it is','as the','of their',
                         'is the','all the','it was', 'into the','and in','in a','the most',
                         'of its','it has','and of',' the','  the','in its','of this',
                         'its own','will be','under the','the more','was the','but the',
                         'have been',] # may need to integrate this before the signle words are added, as the phrases are made
    connectorPhrases3 = []
    connectorPhrases4 = []

    for i in range(len(connectorPhrases2)):
        connectorWords.append(connectorPhrases2[i]+' ')

                        # TODO: add two or three length sets, integrate into code
                        # TODO: track length of phrases to use different sets

    newTopWords = []
    newTopCounts = []
    matchFound = False
    
    for i in range(len(topWords)):
        for j in range(len(connectorWords)):
            if topWords[i] == connectorWords[j]:
                matchFound = True
                break
        if not matchFound:
            newTopWords.append(topWords[i])
            newTopCounts.append(topCounts[i])
        matchFound = False

    topWords = newTopWords[0:int(numWords)]   #may introduce out of range bug for small values, double check this
    topCounts = newTopCounts[0:int(numWords)]
    
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

