#!/usr/bin/python
import sys
import json
import urllib3
from html.parser import HTMLParser

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
VOICELESS_CONSONANTS = ['ch', 'f', 'k', 'p', 's', 'sh', 't', 'th']

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
                        if self.isBritish:
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

def getTranscription(wordToTranscribe, wordType=None):
    data = []
    word = wordToTranscribe
    try:
        http = urllib3.PoolManager()
        
        iterateURLs = True
        iteration = 1

        notInWord = False
        endsInS = False
        isPossessive = False
        endsInIng = False

        if "n't" in word:
            word = word[:-3]
            notInWord = True

        elif word[-2:] == "'s" or word[-2:] == "’s":
            word = word[:-2]
            isPossessive = True

        elif word[-1] == "s":
            word = word[:-1]
            endsInS = True

        # If the words ends in 'ing', mark it down. This should be mainly used for present continuous form of verbs,
        # but there might be some false positives which are checked later on - such as "bring".
        elif word[-3:] == "ing":
            endsInIng = True
        
        # print("Word:", word)

        URL = baseURL + word.strip()
        print("URL:", URL)

        # If the word doesn't have any homonyms, only look through the content on this page.
        # Otherwise, keep iterating until the server returns "Not Found".
        while iterateURLs:
            parser = DictionaryParser()

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

            # If the word is the only thing displayed in the URL after the last slash, then that's the only match
            # so the script stops iterating.
            if redirectedURL.split("/")[-1] == word:
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
                else:
                    iterateURLs = False
            else:
                # If a match wasn't found and the word doesn't end in 'ing', then there are multiple forms available in the
                # dictionary. They look like "/word_1", "/word_2", etc. Iterate through them accordingly.
                iteration += 1
                URL = redirectedURL[0:-1] + str(iteration)
            
            if bool(parser.found) != False:
                items = parser.found
                if "verb" in items:
                    if notInWord:
                        counter = 0
                        for item in items:
                            items[counter] = items[counter] + "n̩t"
                            counter = counter + 1
                        data.append(items)
                    elif endsInIng and "verb" in items:
                        # data.append(items["verb"][5])
                        data.append(items)
                # If the word ends in 's', append the corresponding form. Source: https://english.stackexchange.com/a/23528
                elif "noun" in items:
                    if isPossessive:
                        counter = 0
                        for item in items:
                            lastChar = word[-1]
                            lastTwoChars = word[-2:]
                            
                            if lastChar in VOICELESS_CONSONANTS or lastTwoChars in VOICELESS_CONSONANTS:
                                items["noun"][0] += "s"
                            elif lastChar in VOWELS or lastChar in VOICED_CONSONANTS or lastTwoChars in VOICED_CONSONANTS:
                                items["noun"][0] += "z"
                            elif items[counter][-1] in FRICATIVE_SOUNDS:
                                items["noun"][0] += "iz"
                            counter = counter + 1
                    data.append(items)
                else:
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
        # Ignore commented words
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
                        # print ("\t" + transcriptions)
                        for wordType in transcriptions:
                            print(word + ' (' + wordType + ')')
                            if wordType == "verb":
                                if word[-3:] == "ing":
                                    print ("\t" + transcriptions[wordType][5])
                            else:
                                for transcription in transcriptions[wordType]:
                                    print("\t" + transcription)