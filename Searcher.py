# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
from nltk import data
from nltk import pos_tag
from nltk.corpus import wordnet
from nltk.stem import WordNetLemmatizer


class Searcher:
    def __init__(self, pre_load=None):
        data.path.append('./Resources/nltk_data')
        self.url = "http://dict.youdao.com/w/%s/"
        if pre_load is None:
            self.pre_load = {}
        else:
            self.pre_load = pre_load

    def search_base(self, word):
        respond = requests.request(method="GET", url=self.url % word)
        bs = BeautifulSoup(respond.text, "html.parser")
        meanings, pronounce = None, ""
        try:
            meanings = bs.find(id="phrsListTab").select('.trans-container')[0].select('ul')[0].select('li')
        except:
            pass
        try:
            pronounce = bs.find(id="phrsListTab").select('.pronounce')[0].select('.phonetic')[0].string
        except:
            pass
        return pronounce, meanings

    def search(self, word):
        try:
            pronounce, meanings = self.pre_load[word]
            return pronounce, meanings, 1
        except:
            pass
        pronounce, meanings = self.search_base(word)
        state = 1
        if meanings is None:
            state = -1
        elif pronounce == "":
            state = 0
        if state == 1:
            self.pre_load[word] = [pronounce, meanings]
        return pronounce, meanings, state

    def lemmas(self, word):
        tag = pos_tag([word])[0]
        wnl = WordNetLemmatizer()
        wordnet_pos = self.get_wordnet_pos(tag[1]) or wordnet.NOUN
        return wnl.lemmatize(tag[0], pos=wordnet_pos)

    def get_wordnet_pos(self, tag):
        if tag.startswith('J'):
            return wordnet.ADJ
        elif tag.startswith('V'):
            return wordnet.VERB
        elif tag.startswith('N'):
            return wordnet.NOUN
        elif tag.startswith('R'):
            return wordnet.ADV
        else:
            return None
