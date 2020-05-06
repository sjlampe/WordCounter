#! python 3
# WordCounter2.py -- A program that allows the user to enter a file name
# and count the frequency of the most common words or phrases of a defined
# length.

# This version on the desktop is currently the most recent.

"""TODO: add ability to read PDFs
         add ability to read text documents properly
         make filter junk words automatic
         allow printing to output file
         make into executable"""

import docx, os, pyperclip, datetime, math
import tkinter as tk
from PIL import ImageTk, Image
from tkinter.filedialog import askopenfilename, asksaveasfilename

# os.chdir(r'C:\WordCounter') # TODO: may need to remove

def check_if_all_junk(phrase):
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
    
    words = phrase.split(' ')
    
    junk = False
    for word in words:
        if word in connectorWords:
            junk = True
        if word not in connectorWords:
            junk = False
            break
    if junk == True:
        return True
    else:
        return False
    

def filter_junk_words(top_words, top_counts, phraseSize):
    words = top_words
    counts = top_counts
    
    print(str(datetime.datetime.now()) + '--Words Ordered by Freq') #P2 done

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
                      'since','could','become','she','your','go','am','too','those','',' ']
    
    connectorPhrases2 = ['of the','in the','to the','by the','and the','for the',
                         'on the','with the', 'of all', 'at the','to be','from the',
                         'that the','of a','as a','it is','as the','of their',
                         'is the','all the','it was', 'into the','and in','in a','the most',
                         'of its','it has','and of',' the','  the','in its','of this',
                         'its own','will be','under the','the more','was the','but the',
                         'have been','the same','is not','we may','of our','which is','is a',
                         'of these', 'that it', 'such a', 'they are', 'we are', 'of any',
                         'which are', 'of them', 'in this', 'must be', 'of that', 'when we',''] # may need to integrate this before the signle words are added, as the phrases are made
    
    connectorPhrases3 = ['',' ']

                        # TODO: add two or three length sets, integrate into code
                        # TODO: track length of phrases to use different sets

    too_big = False
    phrase_size = int(entry_num_in_phrase.get())
    print(phrase_size)
    if phrase_size == 1:
        junk_words = connectorWords
    elif phrase_size == 2:
        junk_words = connectorPhrases2
    elif phrase_size > 2:
        junk_words = connectorPhrases3
    else:
        too_big = True

    newTopWords = []
    newTopCounts = []

    print(filter_var.get())
    
    if (filter_var.get()) == 1 and (too_big == False):
        for i in range(len(words)):
            if words[i] not in junk_words:
                newTopWords.append(words[i])
                newTopCounts.append(counts[i])        
    else:
        newTopWords = words
        newTopCounts = counts

    return newTopWords,newTopCounts

def update_progress_bar(progress,not_fin):
    global notifications
    notifications.pack_forget()
    if not_fin == 'yes':
        print('its not me')
        notifications = tk.Label(
                master=frame_notify,
                text=("Loading..." + str(progress) + "%"),
                fg='black',
                bg='#ffee80',
                width=progress
                )
        notifications.pack(side=tk.LEFT)
    else:
        notifications = tk.Label(
                master=frame_notify,
                text=("Done"),
                fg='black',
                bg='#ffee80',
                width=progress
                )
        notifications.pack(side=tk.LEFT)

def roundup(x):
    return int(math.ceil(x / 10.0)) * 10

def select_input(): #TODO: Get a text file opener that interprets files correctly, adjust to fit
    global inputFilename
    
    filepath = askopenfilename(
        filetypes=[("Word Documents","*.docx"),
                   ("Text Files", "*.txt"),
                   ("All Files", "*.*"),]
        )
    if not filepath:
        return
    else:
        entry_select_input.delete(0, tk.END)
        entry_select_input.insert(0, filepath)
        inputFilename = filepath # change to reference text box
        
def select_output():
    global outputFilename
    
    filepath = asksaveasfilename(
        defaultextension="txt",
        filetypes=[("Text Files", ".txt"), ("All Files", "*.*")],
        )
    if not filepath:
        return
    else:
        entry_select_output.delete(0,tk.END)
        entry_select_output.insert(0, filepath)
        outputFilename = filepath # change to reference text box

def save_to_log_file(outputFilename):
    #TODO: save to either a common word or text or excel document
    print("in progress")
    
def execute(): #TODO: adjust to fit
    phraseSize = int(entry_num_in_phrase.get()) # TODO: allow user to enter phraseSize
    numWords = int(entry_tot_num.get()) # TODO: allow user to enter numWords
    copyToClipboard = True #TODO: allow user to select clipboard y/n
    outputFilename = 'hi' #TODO: get rid of this

    if (inputFilename != '') and (phraseSize != '') and (numWords != ''):
        textList = getTextList(inputFilename, phraseSize) # change to reference text box
    else:
        print("enter input, phrasesize, and numwords")
        return
        # TODO: give failure log

    finalWords, finalCounts = mostCommonWords(textList, numWords)
    updateTextbox(finalWords,finalCounts)

    if clipboard_var.get() == 1:
        setClipboard(finalWords,finalCounts)
    
    # TODO: finish execute function -- implement saving to files
    # save_to_log_file(outputFilename)

def getTextList(filename, phraseSize):
    # Gets a file and outputs a list with each token
    # word/phrase found in the file.
    
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    textString = '\n'.join(fullText)
    textList = textString.split()

    punctuations = '''!()[]{};:'"“”’‘\,<>./?@#$%^&*_~–-—'''
    for i in range(len(textList)):
        textList[i] = textList[i].lower()
    for i in range(len(textList)):
        noPunctString = ''
        for char in textList[i]:
            if char not in punctuations:
                noPunctString += char
        textList[i] = noPunctString

    if phraseSize == 1:
        print("Searching for single words...") # TODO: adjust to main GUI
    else: #handles muliple word phrases
        print("Searching for " + str(phraseSize) + "-word phrases...") #TODO: main GUI
        newList = [''] * len(textList)
        for i in range(len(textList) - (phraseSize)):
            nextPhrase = ''
            for n in range(phraseSize):
                nextPhrase += textList[(i + n)]
                if n < phraseSize-1:
                    nextPhrase += ' '
            if (filter_var.get()) == 1: # filters junk when on
                if not check_if_all_junk(nextPhrase):
                    newList[i] = nextPhrase
            else:
                newList[i] = nextPhrase 
        textList = newList
        
    return(textList)

def mostCommonWords(textlist,numWords): #TODO: integrate progress detection
    # Gets a textlist and finds the top numWords most
    # common words/phrases in it, returns lists with counts and words.

    print(str(datetime.datetime.now()) + "--Word Search Start") #P1 done
    
    wordFreq = {}
    for i in range(len(textlist)):
        if textlist[i] in wordFreq.keys():
            wordFreq[textlist[i]] += 1
        else:
            wordFreq[textlist[i]] = 1

    print(str(datetime.datetime.now()) + '--Word Frequencies Obtained') #P1 done

    sorted_words = {k: v for k, v in sorted(wordFreq.items(), reverse=True, key=lambda item:item[1])}
    topWords = list(sorted_words.keys())
    topCounts = list(sorted_words.values())

    progress = 0 #get rid of

    phraseSize = int(entry_num_in_phrase.get())
    topWords,topCounts = filter_junk_words(topWords,topCounts,phraseSize)

    topWords = topWords[0:int(numWords)]   #may introduce out of range bug for small values, double check this
    topCounts = topCounts[0:int(numWords)]

    print(str(datetime.datetime.now()) + '--Junk Words Filtered') #P3 done
    update_progress_bar(progress,True)

    return topWords,topCounts

def updateTextbox(topWords,topCounts):
    outputString = (' WORD/PHRASE:'.rjust(20) + 'FREQUENCY')
    for i in range(len(topWords)):
        outputString += '\n'
        outputString += topWords[i].rjust(20)
        outputString += ':'
        outputString += str(topCounts[i])
    text_box.delete("1.0",tk.END)
    text_box.insert("1.0", outputString)

def setClipboard(topWords,topCounts):
    clipboardContents = ''
    for i in range(len(topWords)):
        clipboardContents += topWords[i]
        clipboardContents += ' '
        clipboardContents += str(topCounts[i])
        clipboardContents += '\n'
    pyperclip.copy(clipboardContents)

# Window formatting

window = tk.Tk()
window.title("Word Counter v. 3 - by Sam Lampe")

# window, top/middle/bottom/right structure
window.rowconfigure(0, minsize=150, weight=1)
window.rowconfigure(1, minsize=200, weight=1)
window.rowconfigure(2, minsize=30, weight=1)
window.columnconfigure(0, minsize=200, weight=1)

frame_picture = tk.Frame(window) # top
frame_interact = tk.Frame(window) # middle
frame_notify = tk.Frame(window) # bottom
frame_options = tk.Frame(window, relief=tk.RAISED, borderwidth=5) # right

frame_interact.columnconfigure(0, minsize=75, weight=0)
frame_interact.columnconfigure(1, minsize=325, weight=1)
frame_interact.columnconfigure(2, minsize=10, weight=0)
frame_interact.rowconfigure([0,1], minsize=25, weight=1)
frame_interact.rowconfigure(2, minsize=30, weight=1)
frame_interact.rowconfigure(4, minsize=50, weight=1)

# contents of top frame
img = ImageTk.PhotoImage(Image.open(r'C:\Users\dragonslayer\Desktop\EMU\wordSearchPicture2.png'))
image_label = tk.Label(frame_picture, image = img)
image_label.pack(side='bottom', fill='both', expand='yes')

# contents of middle frame
# buttons
btn_select_input = tk.Button(frame_interact, text="Select Input File", command=select_input)
btn_select_output = tk.Button(frame_interact, text="Select Output File", command=select_output)
btn_search = tk.Button(frame_interact, text="Search", command=execute)

btn_select_input.grid(row=0, column=0, padx=10, pady=5, sticky='ew')
btn_select_output.grid(row=1, column=0, padx=10, pady=5, sticky='ew')
btn_search.grid(row=4, column=0, padx=10, pady=5, sticky='new')

# filepaths/output
entry_select_input = tk.Entry(frame_interact, width=50)
entry_select_output = tk.Entry(frame_interact, width=50)
text_box = tk.Text(frame_interact, height=10, width=38)

entry_select_input.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
entry_select_output.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
text_box.grid(row=4, column=1, padx=5, pady=10, sticky='new')

# options
# word numbers
options_frame = tk.Frame(master=frame_interact)
options_frame.rowconfigure([0,1],minsize=15,weight=1)
options_frame.columnconfigure(0, minsize=20, weight=1)
options_frame.columnconfigure([1,2,3],minsize=30,weight=1)

text_num_in_phrase = tk.Label(options_frame,text="   No. Words in Phrase :")
entry_num_in_phrase = tk.Entry(options_frame)
text_tot_num = tk.Label(options_frame,text="   No. Top Words/Phrases :")
entry_tot_num = tk.Entry(options_frame)

text_num_in_phrase.grid(column=0, row=0, sticky='w', padx=0, pady=5)
entry_num_in_phrase.grid(column=1, row=0, sticky='w', pady=5)
text_tot_num.grid(column=0, row=1, sticky='w', padx=0, pady=5)
entry_tot_num.grid(column=1, row=1, sticky='w', pady=5)

entry_num_in_phrase.insert(0,"1")
entry_tot_num.insert(0,"100")

options_frame.grid(column=0, row=2, columnspan=2, sticky='ew')

# copy results to clipboard/filter junk words
clipboard_var = tk.IntVar()
clip_box = tk.Checkbutton(options_frame,
                          text="Copy to Clipboard",
                          variable=clipboard_var,
                          onvalue=1,
                          offvalue=0)
clip_box.grid(column=3, row=0, sticky='e')

filter_var = tk.IntVar()
filter_box = tk.Checkbutton(options_frame,
                            text="Filter Junk Words",
                            variable=filter_var,
                            onvalue=1,
                            offvalue=0)
filter_box.grid(column=3, row=1, sticky='e')

# contents of bottom frame
notifications = tk.Label(
    master=frame_notify,
    text="",
    fg='black',
    bg='#ffee80',
    )
notifications.pack(side=tk.LEFT)

# frame organization
frame_picture.grid(row=0, column=0, sticky='nsew')
frame_interact.grid(row=1, column=0, sticky='nsew')
frame_notify.grid(row=2, column=0, sticky='ew', padx=10)

window.mainloop()
