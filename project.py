import pandas as pd
import numpy as np
import re
import os
import glob
import textract
import re
import nltk
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
nltk.download('stopwords')
from sklearn.naive_bayes import MultinomialNB
import win32com.client as win32
import docx2txt
import PyPDF2
import os.path
from nltk.stem import WordNetLemmatizer
nltk.download('wordnet')
from sklearn.ensemble import RandomForestClassifier
#import timeit
from sklearn.feature_extraction.text import CountVectorizer
import pickle

def train():
    #print("In train()")
    path_folder=input("Enter path of folder to read files(use slash '/'):")
    class_type=int(input("Enter type of file(Resume-1, Not Resume-0):"))
    if os.path.exists(path_folder):
        allfilesDF,flist=read_files(path_folder)
        cleanDF=clean_data(allfilesDF)
        mergedDF=merge_file(cleanDF,class_type)
        x_train=word_vector_train(mergedDF)
        y_train=mergedDF.iloc[:, 1].values
        #print(y_train)
        save_array(x_train,y_train,1)
    else:
        print("Directory not exists.")

def pred():
    #print("In pred()")
    path_folder=input("Enter path of folder to read files(use slash '/'):")
    if os.path.exists(path_folder):
        #flist=os.listdir(path_folder)
        allfilesDF,flist=read_files(path_folder)
        cleanDF=clean_data(allfilesDF)
        x_test=word_vector_pred(cleanDF)
        save_array(x_test,0,0)
        #path=input("Enter path of prev x_train,y_train(use slash '/'):")
        #if os.path.exists(path_folder):
        x_train,y_train=load()
        clf=RandomForestClassifier(n_estimators=10,criterion='entropy',random_state=0)
        clf.fit(x_train,y_train)
        #clf=MultinomialNB()
        #clf.fit(x_train,y_train)
        y_pred=clf.predict(x_test)
        prob=clf.predict_proba(x_test)
        prob32=np.float32(prob)
        outL=[]
        i=0
        for (ftype,pr) in zip(y_pred,prob32):
            if ftype==1:
                print(str(i+1)+"."+flist[i]+':')
                print('Resume',end='\t')
                print(pr[1]*100)
                #outL.append('Resume         '+str(pr[1]*100)) 
            else:
                print(str(i+1)+"."+flist[i]+':')
                print('Not Resume',end='\t')
                print(pr[0]*100)
                #outL.append('Not Resume     '+str(pr[0]*100)) 
            i+=1
        #output=pd.DataFrame(outL)
        #output.to_csv('output.csv',header=None, index=None, mode='a')
        
        #else:
         #   print("File not exists.")
    else:
        print("Directory/File not exists.")
        
def read_files(path_folder):
    #print("In read_files()")
    indir=path_folder
    os.chdir(indir)
    filelist=glob.glob("*")

    #reading files(.doc, .docx, .csv, .pdf)
    dfList=[]
    filename=[]
    #word = win32.Dispatch("Word.Application")
    for file in filelist:
        try :
            filename.append(file)
            extension = os.path.splitext(file)[1]
            '''if extension=='.csv':
                print(file)
                #read csv file
                df=pd.read_csv(file,header=None)
                df.columns=['col']
                df.dropna(inplace=True)
                text=df['col'].astype(str).str.cat(sep=' ')
                text = text.replace('\n', ' ')
                text=re.sub('\s+', ' ', text).strip()
                dfList.append(text)'''
            if extension=='.doc' :
                print(file)
                text=textract.process(file).decode("utf-8") 
                #file_remove=name[0]+'.doc'
                #os.remove(file_remove)
                text = text.replace('\n', ' ')
                text=' '.join(text.split())
                dfList.append(text)
                word.Quit()
            elif extension=='.docx' :
                print(file)
                text = docx2txt.process(file)
                text = text.replace('\n', ' ')
                text=re.sub('\s+', ' ', text).strip()
                dfList.append(text)
            elif extension=='.pdf' :
                print(file)
                pdfFileObj = open(file, 'rb') 
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                pageObj = pdfReader.getPage(0) 
                text=pageObj.extractText()
                text = text.replace('\n', ' ')
                text=re.sub('\s+', ' ', text).strip()
                dfList.append(text)
                pdfFileObj.close()
            elif extension=='.txt' :
                print(file)
                text = open(file,'r').read()
                text = text.replace('\n', ' ')
                text=re.sub('\s+', ' ', text).strip()
                dfList.append(text)
            elif extension=='.xlsx' :
                print(file)
                xlsdf=pd.read_excel(file)
                xlsdf['new'] = xlsdf.astype(str).values.sum(axis=1)
                #xls=[]
                xls=' '.join(xlsdf['new'].tolist())
                xls=xls.replace('\n',' ')
                xls=re.sub('\s+', ' ', xls).strip()
                #print(xls)
                dfList.append(xls)
            elif extension=='.ppt' :
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
                text_runs=' '.join(text_runs)
                text_runs=text_runs.replace('\n',' ')
                text_runs=re.sub('\s+',' ',text_runs).strip()
                dfList.append(text_runs)
                prs.close()
            else :
                print("WRONG EXTENSION:")
                print(file)
                continue
        except:
            pass
            #print(file)
    #word.Quit()
    #main_path=input("Enter path to all model data(use '/'):")
    #indir=main_path
    indir='C:/Users/shadd/Rezide/imp'
    os.chdir(indir)
    allfilesDF=pd.DataFrame(dfList)
    allfilesDF.columns=['data']
    allfilesDF.drop_duplicates(keep="first",inplace=True)
    allfilesDF.replace('', np.nan, inplace=True)
    allfilesDF.dropna(inplace=True)
    allfilesDF.reset_index(drop=True,inplace=True)
    allfilesDF.to_csv('allfiles.csv',index=None)
    #print(allfilesDF)
    return allfilesDF,filename

def clean_data(allfilesDF):
    #print("In clean_data()")
    corpus = []
    lm = WordNetLemmatizer()
    for i in range(0, allfilesDF.shape[0]):
        review = re.sub('[^a-zA-Z]', ' ', allfilesDF['data'][i])
        review = review.lower()
        review = review.split()
        review = [lm.lemmatize(word) for word in review if not word in set(stopwords.words('english'))]
        review = ' '.join(review)
        corpus.append(review)

    #saving cleaned file as cleanfinalfile
    cleanDF=pd.DataFrame(corpus)
    cleanDF.columns=['data']
    cleanDF['data'].replace('', np.nan, inplace=True)
    cleanDF.dropna(inplace=True)
    cleanDF.reset_index(drop=True,inplace=True)
    #cleanDF['class']=class_type
    cleanDF.to_csv("cleanfile.csv",index=None)
    #print(cleanDF)
    return cleanDF

def merge_file(cleanDF,class_type):
    #print("In merge_file()")
    path_prevfile=input("Enter path of prev file(use '/'):")
    cleanDF['class']=class_type
    if os.path.exists(path_prevfile):
        file1 = pd.read_csv(path_prevfile,encoding="ISO-8859-1")
        mergedDF = pd.concat([file1,cleanDF])
        mergedDF.to_csv("mergedlem.csv",index=None)
    else:
        print("File not exists")
    return mergedDF
    #else:
     #   print("File not exists.")

def word_vector_train(DF):
    #print("In wordvector_train()")
    listvec=[]
    listvec=DF['data'].tolist()
    pickle_in = open("vocab.pickle","rb")
    vocab = pickle.load(pickle_in)
    cv = CountVectorizer(max_features = 1500,vocabulary=vocab)
    x= cv.fit_transform(listvec).toarray()
    vocab=cv.vocabulary_
    pickle_out = open("vocab.pickle","wb")
    pickle.dump(vocab, pickle_out)
    pickle_out.close()
    #print(x)
    return x


def word_vector_pred(DF):
    #print("In wordvector_pred()")
    listvec=[]
    listvec=DF['data'].tolist()
    pickle_in = open("vocab.pickle","rb")
    vocab = pickle.load(pickle_in)
    cv = CountVectorizer(vocabulary=vocab)
    x= cv.transform(listvec).toarray()
    #print(x)
    return x

def save_array(x,y,train_or_pred):
    #print("In save_array()")
    #changing cwd
    #indir="C:\\Users\\shadd\\OneDrive\\Desktop\\New folder\\allfiles\\final"
    #os.chdir(indir)
    #saving array 
    if train_or_pred:
        np.save('x_train', x)
        np.save('y_train', y)
    else:
        np.save('x_test', x)
        
def load():
    #print("In load()")
    x=np.load(os.path.join("x_train.npy"))
    y=np.load(os.path.join("y_train.npy")) 
    return x,y

def main():
    ch='y'
    while ch!='n':
        n=int(input("Choose option:\n1.Train Data\n2.Predict\n"))
        if n==1:
            train()
        elif n==2:
            pred()
        else:
            print("Wrong option")
        ch=input("Continue?(y/n):")
        
if __name__== "__main__":
    main()