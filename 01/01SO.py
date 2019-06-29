import pandas as pd
df = pd.read_excel (r'D:\vineet\PythonProjects\Stack overflow\Book1.xlsx')
print (df)
mypath = r'D:\vineet\PythonProjects\Stack overflow\doc files'
from os import listdir,rename
from os.path import isfile, join
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
#print(onlyfiles)
currentfilename=onlyfiles[0].split(".")[0]
for i in range(3):
    #print(currentfilename,df['ref'][i])
    if str(currentfilename)==str(df['Reference'][i]):
        corrosponding_email=df['EmailAddress'][i]
        #print(mypath+"\\"+corrosponding_email)
        rename(mypath+"\\"+str(currentfilename)+".docx",mypath+"\\"+corrosponding_email+".docx")