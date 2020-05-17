#!/usr/bin/python
import os
import sys
import argparse
import urllib3
from html.parser import HTMLParser
from openpyxl import load_workbook

VERBOSE_DEBUG = False
PLAINTEXT = False
FILENAME = ""

###     TO DO   ###
# - Mark output and note whether it follows British or American English transcription
# - Add a list to the script containing all irregular verbs for easy lookup: https://myefe.com/english-irregular-verbs
# - Reduce iterations when checking for past tense of regular verbs - currently, there is 1 additional iteration
# - The verb "bring" incorrectly removes "ing" and queries for "br"
# - Past tenses of words don't seem to have word types listed which leads to the script crashing with TypeError: 'NoneType' object is not iterable.
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

# A list of irregular verbs that couldn't be found at the Oxford Learner's Dictionaries.
# Source: https://www.englishpage.com/irregularverbs/irregularverbs.html#
# Also some present forms of verbs are listed below due to them being represented in a different way in the URL of OLD, e.g. sow1.
# NB! "inbreed" has been omitted
# NB! wound has two forms (both not directly found using this script) - noun and verb. The wordform has thus not been included.
# NB! Some verbs, such as "withdraw", have alternative spellings. Only one of these has been taken into consideration
IRREGULAR_VERBS = {
    "arisen": "əˈrɪzn",
    "beaten": "ˈbiːtn",
    "bid": "bɪd",
    "bidden": "ˈbɪdn",
    "bled": "bled",
    "daydreamt": "ˈdeɪdremt",
    "disproven": "ˌdɪsˈpruːvn",
    "do": "duː",
    "dreamt": "dremt",
    "dwelled": "dwelt",
    "eaten": "ˈiːtn",
    "forewent": "fɔːˈwent",
    "foregone": "fɔːˈɡɒn",
    "foresaw": "fɔːˈsɔː",
    "foreseen": "fɔːˈsiːn",
    "forgiven": "fəˈɡɪvn",
    "forsook": "forsook",
    "forsaken": "fəˈseɪkən",
    "frostbit": "ˈfrɒstbɪt",
    "frostbitten": "ˈfrɒstbɪtn",
    "hewn": "hjuːn",
    "hid": "hɪd",
    "kept": "kept",
    "leant": "lent",
    "leapt": "lept",
    "learnt": "lɜːnt",
    "lie": "laɪ",
    "lay": "leɪ",
    "mown": "məʊn",
    "partaken": "pɑːˈteɪkən",
    "risen": "ˈrɪzn",
    "sawn": "sɔːn",
    "slid": "slɪd",
    "sow": "səʊ",
    "sowed": "səʊd",
    "sown": "səʊn",
    "spelt": "spelt",
    "spilt": "spɪlt",
    "strewn": "struːn",
    "striven": "ˈstrɪvn",
    "sunburnt": "ˈsʌnbɜːnt",
    "taken": "ˈteɪkən",
    "typewrite": "taɪpraɪt",
    "typewrote": "taɪprəʊt",
    "typewritten": "taɪprɪtn",
    # "unbent": "ˌʌnˈbent",
    # "unbind": "ˌʌnbaɪnd",
    # "unbound": "ˌʌnbaʊnd",
    # "unclothe": "ˌʌnkləʊð",
    # "unclothed": "ˌʌnkləʊðd",
    # "unclad": "ˌʌnklæd",
    "waylaid": "weɪˈleɪd",
    "withdraw": "wɪðˈdrɔː",
    "withdrew": "wɪðˈdruː",
    "withdrawn": "wɪðˈdrɔːn",
    "withheld": "wɪðˈheld",
    "withstood": "wɪðˈstʊd"
}

# A table of prefixes' transcriptions obtained by manual observation of the OLD.
PREFIX_TRANSCRIPTIONS = {
    "mis": "ˌmɪsˈ",
    "over": "ˌəʊvəˈ",
    "out": "ˌaʊtˈ",
    "pre": "ˌpriːˈ",
    "re": "ˌriːˈ",
    "un": "ˌʌn",
    "under": "ˌʌndəˈ"
}

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
                print("DATA:", data)
                self.type = data
            elif self.recordType == 'w':
                self.found[self.type].append(data.strip('/'))
                # print (data.strip('/'))

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

def getTranscription(wordToTranscribe, wordType=None):
    data = []
    word = wordToTranscribe
    prefix_re = False
    prefix_un = False
    prefix_out = False
    prefix_mis = False
    prefix_pre = False
    prefix_over = False
    prefix_under = False
    prefix = ""
    prefixExists = False
    prefixIterations = 1
    
    if word[:5] == "under":
        prefix_under = True
        prefixExists = True
        prefix = "under"

    elif word[:2] == "re":
        prefix_re = True
        prefixExists = True
        prefix = "re"

    elif word[:2] == "un":
        prefix_un = True
        prefixExists = True
        prefix = "un"
        
    elif word[:3] == "out":
        prefix_out = True
        prefixExists = True
        prefix = "out"

    elif word[:3] == "mis":
        prefix_mis = True
        prefixExists = True
        prefix = "mis"

    elif word[:3] == "pre":
        prefix_pre = True
        prefixExists = True
        prefix = "pre"

    elif word[:4] == "over":
        prefix_over = True
        prefixExists = True
        prefix = "over"    

    if prefixExists:
        prefixIterations = 2

    for prefixIteration in range(prefixIterations):
        # If this is the second iteration, this means that the original word has a prefix and a match for the prefixed word
        # hasn't been found - therefore, remove the prefix and try adding it manually later.
        # This is needed in some cases where it might seem like there is a prefix but there actually isn't one,
        # for example "outside" 
        if prefixIteration == 1:
            # word = prefix + word
            if prefix_re or prefix_un:
                word = word[2:]

            elif prefix_out or prefix_mis or prefix_pre:
                word = word[3:]

            elif prefix_over:
                word = word[4:]

            elif prefix_under:
                word = word[5:]

        # print ("ITERATION WORD:", word, "ITERATION:", prefixIteration)
        
        # First, check if the word is present in the list of irregular verbs and, if so, return it.
        for irregular_verb in IRREGULAR_VERBS:
            if irregular_verb == word:
                transcription = dict()
                transcription["verb"] = list()
                # Because the part of the script that takes out the results looks at the second item in the list, there should be a placeholder.
                transcription["verb"].append(None)
                
                if prefixIteration == 0:
                    transcription["verb"].append(IRREGULAR_VERBS[irregular_verb])
                else:
                    transcription["verb"].append(PREFIX_TRANSCRIPTIONS[prefix] + IRREGULAR_VERBS[irregular_verb])
                
                transcription["verb"].append(IRREGULAR_VERBS[irregular_verb])
                data.append(transcription)
                return data

        # Otherwise, start querying the Oxford Learner's Dictionaries.
        iterateURLs = True
        iteration_URL = 1
        iteration = 1

        hasNoun = False
        notInWord = False
        endsInS = False
        isPossessive = False
        isPluralOrThirdPerson = False
        endsInIng = False
        truncatedIng = False
        stopIterating = False
        syllables = syllableCount(word)
        isPastTense = False      
        doneExchange = False

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
            if word[-3:] == "ies":
                word = word[:-3] + 'y'
        
        elif word[-2:] == "ed":
            word = word [:-1]
            isPastTense = True

        # If the words ends in 'ing', mark it down. This should be mainly used for present continuous form of verbs,
        # but there might be some false positives which are checked later on - such as "bring".
        elif word[-3:] == "ing":
            endsInIng = True

        try:
            http = urllib3.PoolManager()
            URL = baseURL + word.strip()
            if VERBOSE_DEBUG: print ("[*] [DEBUG] URL:", URL)
            
            # If the word doesn't have any homonyms, only look through the content on this page.
            # Otherwise, keep iterating until the server returns "Not Found".
            while iterateURLs:
                if VERBOSE_DEBUG: 
                    print ("---- NEW ITERATION -", iteration, "----")
                    print ("[*] [DEBUG] Word:", word)
                parser = DictionaryParser()
                openURL = http.request("GET", URL)
                if VERBOSE_DEBUG: print ("[*] [DEBUG] Opening URL:", URL)
                redirectedURL = openURL.geturl()

                if openURL.status == 200:
                    parser.feed(openURL.data.decode())
                elif openURL.status == 404:
                    iterateURLs = False
                    if VERBOSE_DEBUG: print ("[*] [DEBUG] URL", URL, "returned HTTP 404.")
                else:
                    print("An error occurred.")

                # If the word is the only thing displayed in the URL after the last slash, then that's the only match
                # so the script stops iterating.
                redirectedWord = redirectedURL.split("/")[-1]
                if redirectedWord == word:
                    if VERBOSE_DEBUG: print ("[*] [DEBUG] Redirected URL:", redirectedURL)
                    # If a match wasn't found and the word ends in 'ing', it's a verb.
                    # Try to get the root form of the verb and search based on it.
                    # Source: https://web2.uvcs.uvic.ca/courses/elc/sample/beginner/gs/gs_53.htm
                    if endsInIng and not truncatedIng:
                        baseWord = word[:-3]
                        #print(syllableCount(baseWord), baseWord[-3], baseWord[-2])
                        if baseWord[-1] == baseWord[-2] and syllables == 1:
                            baseWord = baseWord[:-1]
                        elif baseWord[-2] in VOWELS and ((baseWord[-1] in VOICED_CONSONANTS) or (baseWord[-1] in VOICELESS_CONSONANTS)):
                            baseWord += 'e'
                        URL = baseURL + baseWord
                        iterateURLs = True
                        truncatedIng = True
                        # print (word)
                    elif endsInS:
                        word = word[:-1]
                        URL = baseURL + word
                        print ("New URL: ", URL)
                        if ((syllables == 1 and iteration > 1) or iteration > 2):
                            iterateURLs = False
                            if VERBOSE_DEBUG: print ("[*] [DEBUG] Stopped iterating due to exceeded iteration limit.")
                        else:
                            iterateURLs = True
                    elif isPastTense:
                        word = word[:-1]
                        URL = baseURL + word
                        print ("New URL: ", URL)
                        if iteration > 0 and not doneExchange:
                            if not doneExchange and word[-1] == 'i':
                                word = word[:-1] + 'y'
                                URL = baseURL + word
                                doneExchange = True
                                iterateURLs = True
                            elif iteration == 1:
                                iterateURLs = True
                            else:
                                iterateURLs = False
                                if VERBOSE_DEBUG: print ("[*] [DEBUG] Stopped iterating due to exceeded iteration limit.")
                            print (word, iteration)
                        else:
                            iterateURLs = True
                    else:
                        iterateURLs = False
                        if VERBOSE_DEBUG: print ("[*] [DEBUG] Stopped iterating due to no more possible matches.")
                else:
                    if endsInIng and truncatedIng:
                        if redirectedWord[-1] == 'e':
                            URL = baseURL + redirectedWord[:-1]
                            iterateURLs = True
                        elif redirectedURL[-1].isdigit() and openURL.status != 404:
                            iterateURLs = True
                            URL = redirectedURL[:-1] + str(int(redirectedURL[-1]) + 1)
                        else:
                            iterateURLs = False
                            if VERBOSE_DEBUG:
                                print ("[*] [DEBUG] Stopping iterations - there don't seem to be anymore matches.")
                    else:
                        if VERBOSE_DEBUG: print ("[*] [DEBUG] NO MATCH DETECTED. URL:", redirectedURL, "\r\nWord:", word)
                        iteration_URL += 1
                        URL = redirectedURL[:-1] + str(iteration_URL)
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
                        data.append(items)
                if VERBOSE_DEBUG: print ("---- END ITERATION -", iteration, "----")
                iteration += 1
                parser.close()
        except Exception as e:
            print ("Something went wrong, check the following details for more information:")
            print ("Exception Type:", type(e))
            print ("Arguments", e.args)
            print ("Exception txt:", e)

        if VERBOSE_DEBUG: print ("[*] [DEBUG]", data)
        #print ("noun" in data)
        if (endsInS or isPluralOrThirdPerson) and data is not None and hasNoun:
            for dictionary in data:
                for key in list(dictionary):
                    if key != "noun" and key != "verb":
                        print ("Deleting:", dictionary[key])
                        del dictionary[key]
        
        # The line below caused words ending in "ing" to produce unexpected results. For example:
        # Input - targeting; output - "targeting (noun) ˈtɑːɡɪt" and "targeting (verb) ˈtɑːɡɪtɪŋ"
        # The changes are not yet confirmed to work as expected.

        # elif isPastTense or (endsInIng and not truncatedIng):
        elif isPastTense or endsInIng:
            for dictionary in data:
                for key in list(dictionary):
                    if key != "verb":
                        del dictionary[key]

        if prefixIteration == 1:
            for dictionary in data:
                for key in list(dictionary):
                    arrLocation = 0
                    for transcription in dictionary[key]:
                        # print (dictionary, type(dictionary))
                        # print (key, type(key))
                        # print (transcription, type(transcription))
                        #dictionary[key] = PREFIX_TRANSCRIPTIONS[prefix] + dictionary[key]
                        dictionary[key][arrLocation] = PREFIX_TRANSCRIPTIONS[prefix] + transcription
                        print (dictionary[key][arrLocation])
                        arrLocation += 1
                    

        if (len(data) != 0):
            return data

    # print (data)
    # return data

# If this is a complex word, e.g. inter-change, split the words, get transcriptions for each and merge the results
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

# Check if the file provided via argument is a valid file.
def is_valid_file(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
        sys.exit()
    return arg

argParser = argparse.ArgumentParser(description="Get transcriptions for words from the Oxford Learner's Dictionaries.", allow_abbrev=False)
argParser.add_argument("-v", "--verbose", help="Produce additional DEBUG output.", action="store_true")
argParser.add_argument("-t", "--plaintext", help="If the file that has to be transcribed is not in Excel, use this flag.", action="store_true")
# argParser.add_argument("-f", "--file", help="Specify the name of the file where the input values are stored.", action="store",
#                         type=lambda x: is_valid_file(argParser, x), nargs=1)
argParser.add_argument('-f', "--file", help='Path to a .xls(x) or .txt file containing words to be transcribed. If .txt file, use -t flag', type=lambda x: is_valid_file(argParser, x), nargs=1, metavar='[File to get word forms from]')
argParser.add_argument("-o", "--output", help="Save output to a separate file.", action="store", type=argparse.FileType('w'), nargs=1)
# argParser.add_argument('Path', help='Path to a .xls(x) containing words to be transcribed. If .txt file, use -t flag', type=lambda x: is_valid_file(argParser, x), nargs=1, metavar='[File to get word forms from]')

args = argParser.parse_args()

VERBOSE_DEBUG = args.verbose
PLAINTEXT = args.plaintext

if args.file is not None:
    FILENAME = args.file[0]
else:
    FILENAME = "transcribe.xlsx"

if FILENAME[-3:] == "txt" and not PLAINTEXT:
    print ("It looks like you're pointing to a .txt file - please use -t.")
    sys.exit()
elif FILENAME[-3:] == "xls" or FILENAME[-4:] == "xlsx" and PLAINTEXT:
    print ("Looks like you're pointing to a .xls(x) file, please omit -t.")
    sys.exit()

if PLAINTEXT:
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
                                        if syllableCount(word) == 1:
                                            print ("\t" + transcriptions["verb"][0])
                                        else:
                                            if transcriptions["verb"][-1][-2:] == "ɪŋ":
                                                print ("\t" + transcriptions["verb"][-1])
                                            elif transcriptions["verb"][-2][-2:] == "ɪŋ":
                                                print ("\t" + transcriptions["verb"][-2])
                                            else:
                                                for transcription in transcriptions["verb"]:
                                                    if transcription[-2:] == "ɪŋ":
                                                        print ("\t" + transcription)
                                    elif word[-1] == "s":
                                        if word[-2:] == "ss":
                                            print ("\t" + transcriptions["verb"][1])
                                        else:
                                            print ("\t" + transcriptions["verb"][2])
                                    elif word[-2:] == "ed":
                                        print ("\t" + transcriptions["verb"][3])
                                    else:
                                        print ("\t" + transcriptions["verb"][1])
                                else:
                                    for transcription in transcriptions[wordType]:
                                        print("\t" + transcription)
else:
    workbook = load_workbook(filename=FILENAME)
    sheet = workbook.active
    i = 1
    rPos = "A" + str(1)
    wPos = "B" + str(1)
    while sheet[rPos].value != None:
        word = sheet[rPos].value
        word = word.strip()
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
                    transcribed = ""
                    print(word + ' (' + wordType + ')')
                    if wordType == "verb":
                        if word[-3:] == "ing":
                            if syllableCount(word) == 1:
                                transcribed = transcriptions["verb"][0]
                            else:
                                if transcriptions["verb"][-1][-2:] == "ɪŋ":
                                    transcribed = transcriptions["verb"][-1]
                                elif transcriptions["verb"][-2][-2:] == "ɪŋ":
                                    transcribed = transcriptions["verb"][-2]
                                else:
                                    for transcription in transcriptions["verb"]:
                                        if transcription[-2:] == "ɪŋ":
                                            transcribed = transcription
                            
                        elif word[-1] == "s":
                            if word[-2:] == "ss":
                                transcribed = transcriptions["verb"][1]
                            else:
                                transcribed = transcriptions["verb"][2]
                        elif word[-2:] == "ed":
                            transcribed = transcriptions["verb"][3]
                        else:
                            transcribed = transcriptions["verb"][1]
                    else:
                        for transcription in transcriptions[wordType]:
                            if transcription != "":
                                transcribed = "\r\n" + transcription
                    # print ("Before writing:", transcribed)
                    sheet[wPos] = transcribed
                    wPos = str(chr(ord(wPos[0]) + 1)) + str(i)
        i += 1
        rPos = "A" + str(i)
        wPos = "B" + str(i)

    workbook.save(filename=FILENAME)
    workbook.close()