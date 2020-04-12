#!/usr/bin/python
import sys
import json
import urllib3
from html.parser import HTMLParser

# The URL which is used to look up words
baseURL = "https://www.oxfordlearnersdictionaries.com/definition/english/"

# List of fricatives, needed to form possessive transcriptions. Source: https://pronuncian.com/introduction-to-fricatives
# Source 2 (not used): https://english.stackexchange.com/questions/5913/what-is-the-pronunciation-of-the-possessive-words-that-already-end-in-s
FRICATIVE_SOUNDS = ['v', 'f', 'ð', 'θ', 'z', 's', 'ʒ', 'ʃ', 'h']

# List of voiced consonants. Source: https://www.thoughtco.com/voiced-and-voiceless-consonants-1212092
VOICED_CONSONANTS = ['b', 'd', 'g', 'j', 'l', 'm', 'n', 'ng', 'r', 'sz', 'th', 'v', 'w', 'y', 'z']

# List of voiceless consonants. Source: https://www.thoughtco.com/voiced-and-voiceless-consonants-1212092
VOICELESS_CONSONANTS = ['ch', 'f', 'k', 'p', 's', 'sh', 't', 'th']

# List of vowels. Source: https://simple.wikipedia.org/wiki/Vowel
# Y is part of this list because the checks performed in this script only check the last letter which, if 'y', is interpreted as a vowel according to the source cited above.
VOWELS = ['a', 'e', 'i', 'o', 'u', 'y']

class MyHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.isBritish = False
        self.record = False
        self.getPos = False
        self.isVerb = False
        self.found = []

    def handle_starttag(self, tag, attrs):
        if attrs is not None and attrs != []:
            if tag == "span":
                if attrs[0][0] == "class":
                    if attrs[0][1] == "phon":
                        if self.isBritish:
                            self.record = True
                            self.isBritish = False
                    elif attrs[0][1] == "pos":
                        self.getPos = True
            if tag == "div":
                if attrs[0][0] == "class":
                    if attrs[0][1] == "phons_br":
                        self.isBritish = True

    def handle_endtag(self, tag):
        #print("Encountered an end tag :", tag)
        self.record = False
        self.getPos = False

    def handle_data(self, data):
        if self.getPos == True and data == "verb":
            self.isVerb = True
        if self.record == True:
            self.found.append(data.strip('/'))

def getTranscription(wordToTranscribe):
    data = []
    word = wordToTranscribe
    try:
        http = urllib3.PoolManager()
        
        iterateURLs = True
        iteration = 1

        notInWord = False
        endsInS = False
        isPossessive = False

        if "n't" in word:
            word = word[:-3]
            notInWord = True

        elif word[-2:] == "'s" or word[-2:] == "’s":
            word = word[:-2]
            isPossessive = True

        elif word[-1] == "s":
            word = word[:-1]
            endsInS = True
        print("Word:", word)

        URL = baseURL + word.strip()
        print("URL:", URL)

        # If the word doesn't have any homonyms, only look through the content on this page.
        # Otherwise, keep iterating until the server returns "Not Found".
        while iterateURLs:
            parser = MyHTMLParser()

            openURL = http.request("GET", URL)
            redirectedURL = openURL.geturl()

            if openURL.status == 200:
                # print(URL, "200")
                parser.feed(openURL.data.decode())
            elif openURL.status == 404:
                # print(URL, "404")
                iterateURLs = False
            else:
                print("An error occurred.")

            if redirectedURL.split("/")[-1] == word:
                iterateURLs = False
            else:
                iteration += 1
                URL = redirectedURL[0:-1] + str(iteration)
            
            if parser.found != []:
                items = parser.found
                if notInWord:
                    counter = 0
                    for item in items:
                        items[counter] = items[counter] + "n̩t"
                        counter = counter + 1
                # If the word ends in 's', append the corresponding form. Source: https://english.stackexchange.com/a/23528
                if isPossessive:
                    counter = 0
                    for item in items:
                        lastChar = word[-1]
                        lastTwoChars = word[-2:]
                        
                        if lastChar in VOICELESS_CONSONANTS or lastTwoChars in VOICELESS_CONSONANTS:
                            items[counter] = items[counter] + "s"
                        elif lastChar in VOWELS or lastChar in VOICED_CONSONANTS or lastTwoChars in VOICED_CONSONANTS:
                            items[counter] = items[counter] + "z"
                        elif items[counter][-1] in FRICATIVE_SOUNDS:
                            items[counter] = items[counter] + "iz"
                        counter = counter + 1
                data.append(items)
            
            parser.close()
    except urllib3.URLError:
        print("Couldn't find this word!")

    return data

def getComplexTranscription(wordsCombination):
    words = wordsCombination.split('-')
    endResult = ""
    for word in words:
        data = getTranscription(word)
        endResult += data[0][0]
        endResult += ' '
    return endResult[0:-1]

with open("transcribe.txt", 'r') as file:
    for word in file:
        if word[0] != '#':
            word = word.strip()
            if word != '':
                complex = False
                if '-' in word:
                    result = getComplexTranscription(word)
                    print("\t" + result)
                else:
                    if "n't" not in word and "'ll" not in word and "'s" not in word:
                        word = word.replace("'", '-').lower()
                    result = getTranscription(word)
                    for transcriptions in result:
                        for transcription in transcriptions:
                            print("\t" + transcription)