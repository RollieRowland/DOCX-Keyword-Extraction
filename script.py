import docx
import glob
import os
import sys
import ntpath
import pandas as pd
from rake_nltk import Rake


dataDirectory = os.getcwd() + '/Data'

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

if __name__ == '__main__':
    filenames = []
    totalDescription = ''
    fileIndex = 0

    for filename in glob.glob(os.path.join(dataDirectory, '*.docx')):
        print(filename)
        filenames.append(filename)

    for filename in filenames:
        print("Reading file " + str(fileIndex) + " of " + str(len(filenames)))
        fileIndex += 1

        totalDescription += getText(filename)

    r = Rake()

    print("Extracting keywords...")

    r.extract_keywords_from_text(totalDescription)

    print("Getting ranked phrases")

    keywords = r.get_ranked_phrases_with_scores()

    df = pd.DataFrame(columns = ['rank', 'keyword_set'])

    for pair in keywords:
        num = (len(df) + 1)
        df.loc[num] = pair

    dirtitle =  'KeywordExtraction.csv'

    df.to_csv(dirtitle, encoding='utf-8')

    print("File created for query data extraction")
