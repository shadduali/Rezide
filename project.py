import pandas as pd
import numpy as np
import re
import os
import glob
import nltk
import docx2txt
import PyPDF2
import pickle

from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from sklearn.ensemble import RandomForestClassifier
from sklearn.feature_extraction.text import CountVectorizer

nltk.download('stopwords')
nltk.download('wordnet')


def train():
    path_folder = input("Enter path of folder to read files:")
    path_folder = path_folder.replace("/", "//")
    class_type = int(input("Enter type of file(Resume-1, Not Resume-0):"))
    if os.path.exists(path_folder):
        allfilesDF, flist = read_files(path_folder)
        cleanDF = clean_data(allfilesDF)
        cleanDF['class'] = class_type
        mergedDF = merge_file(cleanDF)
        train_model(mergedDF)
    else:
        print("Directory not exists.")


def pred():
    path_folder = input("Enter path of folder to read files:")
    path_folder = path_folder.replace("/", "//")
    if os.path.exists(path_folder):
        allfilesDF, flist = read_files(path_folder)
        cleanDF = clean_data(allfilesDF)
        x_test = word_vector_pred(cleanDF)
        clf = pickle.load(open("RFCmodel.pickle", "rb"))
        y_pred = clf.predict(x_test)
        prob = clf.predict_proba(x_test)
        prob32 = np.float32(prob)
        i = 0
        for (ftype, pr) in zip(y_pred, prob32):
            if ftype == 1:
                print(str(i + 1) + "." + flist[i] + ':')
                print('Resume', end='\t')
                print(pr[1] * 100)
            else:
                print(str(i + 1) + "." + flist[i] + ':')
                print('Not Resume', end='\t')
                print(pr[0] * 100)
            i += 1

        cleanDF['class'] = y_pred
        fback = input('Do you want to give feedback(y/n)?')
        if (fback == 'y'):
            auto_train(cleanDF)
    else:
        print("Directory/File not exists.")


def auto_train(cleanDF):
    ch = input('Were all files correctly predicted(y/n)?')
    if ch == 'n':
        wr_pr = input('Files which were wrongly predicted(serial no from above):')
        wr_pr = [int(wr_pr) - 1 for wr_pr in (re.split(',| |\n', wr_pr))]
        for i in wr_pr:
            cleanDF.iloc[i, 1] = 1 - cleanDF.iloc[i, 1]
    mergedDF = merge_file(cleanDF)
    train_model(mergedDF)


def read_files(pathorfile):
    if (type(pathorfile) == str):
        os.chdir(pathorfile)
        filelist = glob.glob("*")
    else:
        filelist = pathorfile

    # reading files(.docx, .txt, .xlsx, .ppt, .pdf)
    dfList = []
    filename = []
    for file in filelist:
        try:
            extension = os.path.splitext(file)[1]
            if extension == '.docx':
                filename.append(file)
                print(file)
                text = docx2txt.process(file)
                text = text.replace('\n', ' ')
                text = re.sub('\s+', ' ', text).strip()
                dfList.append(text)
            elif extension == '.pdf':
                filename.append(file)
                print(file)
                pdfFileObj = open(file, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                pageObj = pdfReader.getPage(0)
                text = pageObj.extractText()
                text = text.replace('\n', ' ')
                text = re.sub('\s+', ' ', text).strip()
                dfList.append(text)
                pdfFileObj.close()
            elif extension == '.txt':
                filename.append(file)
                print(file)
                text = open(file, 'r').read()
                text = text.replace('\n', ' ')
                text = re.sub('\s+', ' ', text).strip()
                dfList.append(text)
            elif extension == '.xlsx':
                filename.append(file)
                print(file)
                xlsdf = pd.read_excel(file)
                xlsdf['new'] = xlsdf.astype(str).values.sum(axis=1)
                xls = ' '.join(xlsdf['new'].tolist())
                xls = xls.replace('\n', ' ')
                xls = re.sub('\s+', ' ', xls).strip()
                dfList.append(xls)
            elif extension == '.ppt':
                filename.append(file)
                print(file)
                prs = Presentation(file)
                text_runs = []
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if not shape.has_text_frame:
                            continue
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                text_runs.append(run.text)
                text_runs = ' '.join(text_runs)
                text_runs = text_runs.replace('\n', ' ')
                text_runs = re.sub('\s+', ' ', text_runs).strip()
                dfList.append(text_runs)
                prs.close()
            else:
                print("WRONG EXTENSION:")
                print(file)
                continue
        except:
            pass
    indir = 'C:/Users/Faiz Ali/Rezide/imp' #change with the path where all files and ML model is stored
    os.chdir(indir)
    allfilesDF = pd.DataFrame(dfList)
    allfilesDF.columns = ['data']
    allfilesDF.drop_duplicates(keep="first", inplace=True)
    allfilesDF.replace('', np.nan, inplace=True)
    allfilesDF.dropna(inplace=True)
    allfilesDF.reset_index(drop=True, inplace=True)
    return allfilesDF, filename


def clean_data(allfilesDF):
    corpus = []
    lm = WordNetLemmatizer()
    for i in range(0, allfilesDF.shape[0]):
        review = re.sub('[^a-zA-Z]', ' ', allfilesDF['data'][i])
        review = review.lower()
        review = review.split()
        review = [lm.lemmatize(word) for word in review if not word in set(stopwords.words('english'))]
        review = ' '.join(review)
        corpus.append(review)

    # saving cleaned file as cleanfinalfile
    cleanDF = pd.DataFrame(corpus)
    cleanDF.columns = ['data']
    cleanDF['data'].replace('', np.nan, inplace=True)
    cleanDF.dropna(inplace=True)
    cleanDF.reset_index(drop=True, inplace=True)
    return cleanDF


def merge_file(cleanDF):
    prev_file=input("Is there any previous file(y/n)?")
    if prev_file=='y' or prev_file=='Y':
        if os.path.exists('mergedlem.csv'):
            file1 = pd.read_csv('mergedlem.csv'), encoding="ISO-8859-1")
            mergedDF = pd.concat([file1, cleanDF])
            mergedDF.to_csv("mergedlem.csv", index=None)
            return mergedDF
        else:
            print("File not exists")
            return cleanDF;
    else:
        return cleanDF;


def train_model(DF):
    listvec = DF['data'].tolist()
    cv = CountVectorizer(max_features=1500)
    x_train = cv.fit_transform(listvec).toarray()
    y_train = DF.iloc[:, 1].values
    #saving vocabulary
    vocab = cv.vocabulary_
    pickle.dump(vocab, open("vocab.pickle", "wb"))
    #training model
    clf = RandomForestClassifier(n_estimators=10, criterion='entropy', random_state=0)
    clf.fit(x_train, y_train)
    #saving model
    pickle.dump(vocab, open("RFCmodel.pickle", "wb"))


def word_vector_pred(DF):
    listvec = DF['data'].tolist()
    vocab = pickle.load(open("vocab.pickle", "rb"))
    cv = CountVectorizer(vocabulary=vocab)
    x = cv.transform(listvec).toarray()
    return x


def main():
    ch = 'y'
    while ch != 'n':
        n = int(input("Choose option:\n1.Train Data\n2.Predict\n"))
        if n == 1:
            train()
        elif n == 2:
            pred()
        else:
            print("Wrong option")
        ch = input("Continue?(y/n):")


if __name__ == "__main__":
    main()
