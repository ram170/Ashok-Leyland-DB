import docx
from PIL import Image
from docx import Document
from pytesseract import image_to_string
import os
def read():
    document=Document()
    x=os.listdir('/01/Image processing/Images/')
    s='/01/Image processing/Images/'+str(x[0])
    img=Image.open(s)
    text=image_to_string(img)
    #print(text)
    document.save('C:\\01\\Image processing\\Chasisno\\'+text)
    os.remove(s)
    return text
x=os.listdir('/01/Image processing/Backup/')
document= Document('/01/Image processing/Backup/bup.docx')
for para in document.paragraphs:
    temp=para.text
    #print(temp)
l=list(temp.split(" "))
#print(l)    
while(1):
   print("What do you want to do?")
   print("1. Including New vehicle")
   print("2. Map")  
   print("3. Remove")
   do=int(input())
   if(do==1):
        chasisno=read()
        print("Enter position")
        get=int(input())
        l[get-1]=chasisno
        text2=str()
        document2=Document()
        for i in range(0,len(l)):
          if(i!=len(l)-1):
            text2=text2+str(l[i])+" "
          else:
            text2=text2+str(l[i])
        para1=document2.add_paragraph(text2)
        #print(text2)      
        document2.save('C:\\01\\Image processing\\Backup\\bup.docx')
   if(do==2):
       for i in range(0,len(l)):
           print(i+1,l[i])
   if(do==3):
       print("The position to remove?")
       r=int(input())
       print("The chasis no at the given position is",l[r-1])
       print("Remove? press Y or N")
       if(input()=='Y'):
           l[r-1]=0
       text2=str()    
       document2=Document()
       for i in range(0,len(l)):
          if(i!=len(l)-1):
            text2=text2+str(l[i])+" "
          else:
            text2=text2+str(l[i])
       para1=document2.add_paragraph(text2)
       #print(text2)      
       document2.save('C:\\01\\Image processing\\Backup\\bup.docx')

