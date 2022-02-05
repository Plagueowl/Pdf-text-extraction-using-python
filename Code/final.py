# Import libraries
from PIL import Image
import pytesseract
import sys
from pdf2image import convert_from_path
import os
import docx
import shutil
pytesseract.pytesseract.tesseract_cmd=r'C:\Users\HP\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
def filter(c):
    codepoint = ord(c)
    # conditions ordered by presumed frequency
    return (
        0x20 <= codepoint <= 0xD7FF or
        codepoint in (0x9, 0xA, 0xD) or
        0xE000 <= codepoint <= 0xFFFD or
        0x10000 <= codepoint <= 0x10FFFF
        )


def OCR(filename):
    # Set filename to recognize text from
    # Again, these files will be:
    # page_1.jpg
    # page_2.jpg
    # ....
    # page_n.jpg
    text2=(pytesseract.image_to_string(Image.open(filename)))
    return text2

def image_extractor(num,path,text):
    pages=convert_from_path(path,500,output_folder='Temporary',first_page=num,last_page=num+1)
    filename = "page_"+str(num+1)+".jpg"
    # Save the image of the page in system
    pages[num].save(filename, 'JPEG')
    return filename


    

def text_processor(text):
    #filters text and replaces suitable formats
    text = text.replace('-\n', '')
    filtered_text=''.join(c for c in text if filter(c))
    return filtered_text


def main():
    flag1=True
    while flag1==True:
        try:
            path=input("Enter path of file: ")
            page_count=int(input("Enter number of pages you want to convert: "))
            flag1=False
        except Exception as f:
            print("Error: ",f)
        #Generating pages
    text=''
    for i in range(0,page_count):
        # Declaring filename for each page of PDF as JPG
        # For each page, filename will be:
        # PDF page 1 -> page_1.jpg
        # PDF page 2 -> page_2.jpg
        # PDF page 3 -> page_3.jpg
        # ....
        # PDF page n -> page_n.jpg
        print("Generating page ",str(i+1),'...\n')
        filename1=image_extractor(i,path,text)
        print("Extracting from page ",str(i+1))
        #Running OCR
        text+=OCR(filename1)
        #Deleting clutter
        os.remove(filename1)
        folder='D:\\Temporary'
        for filename2 in os.listdir(folder):
            file_path = os.path.join(folder, filename2)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    pathn=folder+'\\'+filename2
                    os.remove(pathn)
                    break
                elif os.path.isdir(file_path):
                    os.remove(pathn)
                    break
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))
    folder='D:\\Temporary'
    for filename2 in os.listdir(folder):
        file_path = os.path.join(folder, filename2)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                pathn=folder+'\\'+filename2
                os.remove(pathn)
            elif os.path.isdir(file_path):
                os.remove(pathn)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))

        
        
        
    # The recognized text is stored in variable text
    # Any string processing may be applied on text
    # Here, basic formatting has been done:
    # In many PDFs, at line ending, if a word can't
    # be written fully, a hyphen is added.
    # The rest of the word is written in the next line
    # Eg: This is a sample text this word here  Vish-
    # wakarma is half on first line, remaining on next.
    # To remove this, we replace every '-\n' to ''.
    text=text_processor(text)
    #finally, save extracted text
    saver(text)


def saver(text):
    flag=True
    while flag==True:
        try:
            choice=int(input("Store information in:\n1)txt file\n2)word file\n3)print here\n"))
            if choice==1:
                path_out=input("Enter output path: ")
                file1=open(path_out,'w')
                file1.write(text)
                file1.close
                print("Document generated at : ",path_out)
                break
            elif choice==2:
                mydoc = docx.Document()
                mydoc.add_paragraph(text)
                path_out=input("Enter output path: ")
                mydoc.save(path_out)
                print("Document generated at : ",path_out)
                break
            elif choice==3:
                print(text)
                break
            else:
                print("Invalid Input")
        except Exception as e:
            print("Error: ",e)
if __name__=="__main__":
    main()


        

