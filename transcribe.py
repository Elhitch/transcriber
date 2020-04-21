#!/usr/bin/python
import os
import sys
import argparse
import json
import urllib3
from html.parser import HTMLParser

VERBOSE_DEBUG = False
FILENAME = ""

###     TO DO   ###
# - Изходът дума да се маркира от кой речник е взета и дали е по британски или американски стандард
#
#
#
###      END    ###

# The URL which is used to look up words
baseURL = "https://www.oxfordlearnersdictionaries.com/definition/english/"

# List of fricatives, needed to form possessive transcriptions. Source: https://pronuncian.com/introduction-to-fricatives
# Source 2 (not used): https://english.stackexchange.com/questions/5913/what-is-the-pronunciation-of-the-possessive-words-that-already-end-in-s
FRICATIVE_SOUNDS = ['v', 'f', 'ð', 'θ', 'z', 's', 'ʒ', 'ʃ', 'h']

# List of voiced consonants. Source: https://www.thoughtco.com/voiced-and-voiceless-consonants-1212092
VOICED_CONSONANTS = ['b', 'd', 'g', 'j', 'l', 'm', 'n', 'ng', 'r', 'sz', 'th', 'v', 'w', 'y', 'z']

# List of voiceless consonants. Source: https://www.thoughtco.com/voiced-and-voiceless-consonants-1212092
# Voiceless consonats which could also be found in the list of fricative sounds are discarded from the following list.
# VOICELESS_CONSONANTS = ['ch', 'f', 'k', 'p', 's', 'sh', 't', 'th']
VOICELESS_CONSONANTS = ['ch', 'f', 'k', 'p', 'sh', 't', 'th']

# List of vowels. Source: https://simple.wikipedia.org/wiki/Vowel
# Y is part of this list because the checks performed in this script only check the last letter which, if 'y', is interpreted as a vowel according to the source cited above.
VOWELS = ['a', 'а', 'e', 'i', 'o', 'u', 'y']

class DictionaryParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.isBritish = False
        self.record = False
        self.found = dict()
        self.counter = 0;
        self.type = ""

        # Record type - remembers which of the following is being recorded:
        #   t - Type (noun / verb / adjective)
        #   w - Word
        self.recordType = ''

    def handle_starttag(self, tag, attrs):
        if attrs is not None and attrs != []:
            if tag == "span":
                if attrs[0][0] == "class":
                    if attrs[0][1] == "phon":
                        # If self.type doesn't exist, then the page relates to a name, for example: kiss - verb, noun, rock band
                        if self.isBritish and self.type != "":
                            self.record = True
                            self.recordType = 'w'
                            self.isBritish = False
                    if attrs[0][1] == "pos":
                        self.record = True
                        self.recordType = 't'
            if tag == "div":
                if attrs[0][0] == "class":
                    if attrs[0][1] == "phons_br":
                        self.isBritish = True

    def handle_endtag(self, tag):
        #print("Encountered an end tag :", tag)
        if self.record:
            self.record = False
            if self.recordType == 't':
                self.found[self.type] = []
            elif self.recordType == 'w':
                pass

    def handle_data(self, data):
        if self.record == True:
            if self.recordType == 't':
                self.type = data
            elif self.recordType == 'w':
                self.found[self.type].append(data.strip('/'))
                print (data.strip('/'))

def syllableCount(word):
    word = word.lower()
    count = 0
    if word[0] in VOWELS:
        count += 1
    for index in range(1, len(word)):
        if word[index] in VOWELS and word[index - 1] not in VOWELS:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count

def getPluralOrThidPerson(type, word, items):
    lastChar = word[-1]
    lastTwoChars = word[-2:]
    if type == "verb":
        pos = 2
    else:
        pos = 0
    if lastChar in VOICELESS_CONSONANTS or lastTwoChars in VOICELESS_CONSONANTS:
        items[type][pos] += "s"
    elif lastChar in FRICATIVE_SOUNDS:
        items[type][pos] += "iz"
    elif lastChar in VOWELS or lastChar in VOICED_CONSONANTS or lastTwoChars in VOICED_CONSONANTS:
        items[type][pos] += "z"
    return items

def cleanData(data):
    hasNouns = False
    for dictionary in data:
        for key in dictionary:
            # print (key)
            # if key != "noun" and "noun" in data and endsInS:
            pass

def getTranscription(wordToTranscribe, wordType=None):
    data = []
    word = wordToTranscribe
    try:
        http = urllib3.PoolManager()
        
        iterateURLs = True
        iteration_URL = 1
        iteration = 1

        hasNoun = False
        notInWord = False
        endsInS = False
        isPossessive = False
        isPluralOrThirdPerson = False
        isPlural = False
        endsInIng = False
        stopIterating = False
        syllables = syllableCount(word)

        if "n't" in word:
            word = word[:-3]
            notInWord = True

        elif word[-2:] == "'s" or word[-2:] == "’s":
            word = word[:-2]
            isPossessive = True

        elif word[-4:] == "sses":
            word = word[:-2]
            isPluralOrThirdPerson = True

        elif word[-1] == "s" and word[-4:] != "ness":
            endsInS = True
            isPluralOrThirdPerson = True

        # If the words ends in 'ing', mark it down. This should be mainly used for present continuous form of verbs,
        # but there might be some false positives which are checked later on - such as "bring".
        elif word[-3:] == "ing":
            endsInIng = True

        URL = baseURL + word.strip()
        print("URL:", URL)
        
        # If the word doesn't have any homonyms, only look through the content on this page.
        # Otherwise, keep iterating until the server returns "Not Found".
        while iterateURLs:
            print ("Word:", word)
            parser = DictionaryParser()
            openURL = http.request("GET", URL)
            if VERBOSE_DEBUG: print ("[*] [DEBUG] Opening URL:", URL)
            redirectedURL = openURL.geturl()

            if openURL.status == 200:
                parser.feed(openURL.data.decode())
                # if endsInS:
                #     iterateURLs = False
                #     if VERBOSE_DEBUG: print ("[*] [DEBUG] STOPPED ITERATING DUE TO 200 AND ENDS IN S")
            elif openURL.status == 404:
                iterateURLs = False
                if VERBOSE_DEBUG: print ("[*] [DEBUG] URL", URL, "returned HTTP 404.")
                # if endsInS and iteration > 2:
                #     endsInS = False
            else:
                print("An error occurred.")

            # If the word is the only thing displayed in the URL after the last slash, then that's the only match
            # so the script stops iterating.
            # print (redirectedURL.split("/")[-1], word)
            if redirectedURL.split("/")[-1] == word:
                if VERBOSE_DEBUG: print ("Redirected URL:", redirectedURL)
                # If a match wasn't found and the word ends in 'ing', it's a verb.
                # Try to get the root form of the verb and search based on it.
                # Source: https://web2.uvcs.uvic.ca/courses/elc/sample/beginner/gs/gs_53.htm
                if endsInIng:
                    baseWord = word[:-3]
                    #print(syllableCount(baseWord), baseWord[-3], baseWord[-2])
                    if baseWord[-1] == baseWord[-2]:
                        baseWord = baseWord[:-1]
                    elif baseWord[-2] in VOWELS and ((baseWord[-1] in VOICED_CONSONANTS) or (baseWord[-1] in VOICELESS_CONSONANTS)):
                        baseWord += 'e'
                    # elif baseWord[-1] in VOICED_CONSONANTS or baseWord[-1] in VOICELESS_CONSONANTS:
                        # baseWord += 'e'
                    URL = baseURL + baseWord
                    print(URL)
                    iterateURLs = True
                elif endsInS:
                    word = word[:-1]
                    URL = baseURL + word
                    print ("New URL: ", URL)
                    if ((syllables == 1 and iteration > 1) or iteration > 2):
                        iterateURLs = False
                        if VERBOSE_DEBUG: print ("[*] [DEBUG] Stopped iterating due to exceeded iteration limit.")
                    else:
                        iterateURLs = True
                else:
                    iterateURLs = False
                    if VERBOSE_DEBUG: print ("[*] [DEBUG] Stopped iterating due to no more possible matches.")
            else:
                if VERBOSE_DEBUG: print ("[*] [DEBUG] NO MATCH DETECTED. URL:", redirectedURL, "\r\nWord:", word)
                iteration_URL += 1
                URL = redirectedURL[0:-1] + str(iteration_URL)
                if VERBOSE_DEBUG:
                    print ("[*] [DEBUG] TRYING TO REDIRECT TO:", URL)
                    print ("[*] [DEBUG] ITERATION STATUS:", iterateURLs)
            
            if bool(parser.found) != False:
                items = parser.found
                if "verb" in items:
                    if notInWord:
                        counter = 0
                        for item in items:
                            items[counter] = items[counter] + "n̩t"
                            counter = counter + 1
                        data.append(items)
                    else:
                        data.append(items)
                # If the word ends in 's', append the corresponding form. Source: https://english.stackexchange.com/a/23528
                elif "noun" in items:
                    hasNoun = True
                    if isPossessive or isPluralOrThirdPerson:
                        endsInS = False
                        data.append(getPluralOrThidPerson("noun", word, items))
                    else:
                        data.append(items)
                else:
                    print (data)
                    if not endsInS and "noun" in data[0]:
                        print ("GOT IT")
                    data.append(items)
            
            iteration += 1
            parser.close()
    except Exception as e:
        print ("Something went wrong, check the following details for more information:")
        print ("Exception Type:", type(e))
        print ("Arguments", e.args)
        print ("Exception txt:", e)

    #print ("noun" in data)
    if (endsInS or isPluralOrThirdPerson) and data is not None and hasNoun:
        for dictionary in data:
            for key in list(dictionary):
                if key != "noun":
                    print ("Deleting:", dictionary[key])
                    del dictionary[key]
    print (data)
    return data

def getComplexTranscription(wordsCombination):
    words = wordsCombination.split('-')
    endResult = dict()
    tempArray = ""
    for word in words:
        data = getTranscription(word)
        for dictionary in data:
            for key in dictionary:
                for element in dictionary[key]:
                    if key != "verb":
                        tempArray += element
                        tempArray += " "
    return tempArray.rstrip()

def is_valid_file(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
        sys.exit()
    return arg

argParser = argparse.ArgumentParser(description="Get transcriptions for words from the Oxford Learner's Dictionaries.", allow_abbrev=False)
argParser.add_argument("-v", "--verbose", help="Produce additional DEBUG output.", action="store_true")
argParser.add_argument("-f", "--file", help="Specify the name of the file where the input values are stored.", action="store",
                        type=lambda x: is_valid_file(argParser, x), nargs=1)
argParser.add_argument("-o", "--output", help="Save output to a separate file.", action="store", type=argparse.FileType('w'), nargs=1)

args = argParser.parse_args()
VERBOSE_DEBUG = args.verbose
if args.file is not None:
    FILENAME = args.file[0]
else:
    FILENAME = "transcribe.txt"

with open(FILENAME, 'r') as file:
    for word in file:
        # Ignore commented words
        if word[0] != '#':
            word = word.strip()
            if word != '':
                complex = False
                if '-' in word:
                    result = getComplexTranscription(word)
                    print (word, "(complex)")
                    print ("\t" + result)
                else:
                    if "n't" not in word and "'ll" not in word and "'s" not in word:
                        word = word.replace("'", '-').lower()
                    else:
                        word = word.lower()
                    result = getTranscription(word)
                    for transcriptions in result:
                        for wordType in transcriptions:
                            print(word + ' (' + wordType + ')')
                            if wordType == "verb":
                                if word[-3:] == "ing":
                                    print ("\t" + transcriptions["verb"][5])
                                elif word[-1] == "s":
                                    # 'sses' at the end of the verb indicates he / she / it form.
                                    # if word[-4:] == "sses":
                                    #     print ("\t" + transcriptions["verb"][2])
                                    if word[-2:] == "ss":
                                        print ("\t" + transcriptions["verb"][1])
                                    else:
                                        print ("\t" + transcriptions["verb"][2])
                                else:
                                    print ("\t" + transcriptions["verb"][1])
                            else:
                                for transcription in transcriptions[wordType]:
                                    print("\t" + transcription)