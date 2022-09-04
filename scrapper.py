## Coded by JUSTZETA ##


from bs4 import BeautifulSoup as bs #used for scrapping the website
import xlsxwriter #used to create the xlsx file
import requests #used to access the site
import urllib.request #used this for checking internet connectivity
from tqdm import tqdm #used this for progess bar

import time
import pyfiglet
from rich import print as prnt

#Function to check internet Connectivity
def connect(host='http://google.com'):
    try:
        urllib.request.urlopen(host)
        return False
    except:
        return True

#Function to get the HTML Document
def getHTMLdocument(url):
  response=requests.get(url)
  return response.text

#Fuction to scrape the data from the HTML Document
def scrape(srchTxt,pgCount):
  searchText=srchTxt
  pagecount=pgCount
  pages=0
  url="https://pubmed.ncbi.nlm.nih.gov/?term="+searchText+"&page="+str(pages)
  title=['Title']
  author=['Author(s)']
  cite=['Citation']
  abstractUrls=[]
  abstract=['Abstract']

  pbar=tqdm(total=pgCount+1,desc="Scrapping the site...", colour="blue")
  while True:
    html_doc=getHTMLdocument(url)
    soupObject=bs(html_doc,'html.parser')
    
    #Extracting Title
    titleContents=soupObject.findAll('a',class_='docsum-title',href=True)
    for titleContent in titleContents:
        title.append(titleContent.text)
    
    #EXtracting Authors    
    authorContents=soupObject.findAll('span',class_='docsum-authors full-authors')
    for authorContent in authorContents:
        author.append(authorContent.text)
    
   #EXtracting Cite    
    citeContents=soupObject.findAll('span',class_='docsum-journal-citation full-journal-citation')
    for citeContent in citeContents:
        cite.append(citeContent.text)
        
    #Getting Abstract
    for link in soupObject.findAll('a',class_='docsum-title'):
        abstractUrls.append(link.get('href'))
    
    for items in abstractUrls:
        abstract.append(getAbstract(items))
      
    pages=pages+1
    url="https://pubmed.ncbi.nlm.nih.gov/?term="+searchText+"&page="+str(pages)
    
    global dataZip
    def dataZip():
        pipe=[title,author,abstract,cite]
        return pipe
    pbar.update(1)
    if(pages==int(pagecount)+1):
        break
  pbar.close()
  return 0;

#get Abstract from each page
def getAbstract(siteUrl):
    
    siteUrl="https://pubmed.ncbi.nlm.nih.gov/"+siteUrl
    
    html_doc=getHTMLdocument(siteUrl)
    soupObject=bs(html_doc,'html.parser')
    abstract=[]
    
    #Extracting Title
    abstractContents=soupObject.find('div',class_='abstract-content selected')
    #for absContent in abstractContents:
    #    abstract.append(absContent.text)
    #abstract.append(abstractContents[0].text)
    if(abstractContents==None):
        return abstractContents
    else:
        return abstractContents.text

##Main
#printing the app name
appName=pyfiglet.figlet_format('PUB MED Scrapper',font='big')
prnt(f'[green]{appName}[/green]')

#Checking for internet connenctivity
if(connect()):
      print("NO NETWORK CONNECTION!!!")
      time.sleep(6)
      exit()

choice=True
valid_txt="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMOPQRSTUVWXYZ '-"

while(choice):
  #validating search text
  while True:
    st=input("Enter the SEARCH Text: ")
    if all(char in valid_txt for char in st):
        break
    print("Invalid Search Text!, Please Try Again!")
  #validating page count
  while True:
    try:
      pg=int(input("Enter How Many Pages to PARSE: "))
      break
    except ValueError:
      print("Please Enter A NUMBER Value!")
      

  #call the scrapping function
  scrape(st,pg)
  #create the xlsx file
  fileName="Excel_files/"+st+"_dataEXTRACT.xlsx"
  workbook = xlsxwriter.Workbook(fileName)
  worksheet = workbook.add_worksheet()
    
  data = dataZip()
  data=zip(data[0],data[1],data[2],data[3])
  row=0
  column=0
  #write to the xlsx file
  for v,x,y,z in data:
        worksheet.write(row,column,v)
        worksheet.write(row,column+1,x)
        worksheet.write(row,column+2,y)
        worksheet.write(row,column+3,z)
        row+=1

  # Close the workbook before sending the data.
  workbook.close()
  print("DONE.\nExcel File Created in the Excel_files folder.")
  ch=input("EXIT the PROGRAM : (Y/N)")
  if(ch.lower()=="y"):
      choice=False

#Peace
