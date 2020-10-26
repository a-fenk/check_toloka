import re

import nltk
from nltk import download
from nltk.corpus import stopwords

download('stopwords', quiet=True)
download('punkt', quiet=True)

stop_words = stopwords.words('russian')

def stemmer(corpus):
    stem = nltk.SnowballStemmer("russian").stem
    stems = []
    for word in corpus:
        stems.append(stem(word))
    return stems


def tokenize(corpus):
    corpus = re.sub(r'[^\w\s]|_', ' ', corpus)  # замена скобок, пунктуации и "_" на " "

    tokens = [word for sent in nltk.sent_tokenize(corpus) for word in nltk.word_tokenize(sent)]
    valuable_words = []

    for token in tokens:
        token = token.lower().strip()
        if token.isalnum() and not token.isdigit() and token not in stop_words and stemmer([token])[0] not in stop_words:
            valuable_words.append(token)

    return valuable_words
