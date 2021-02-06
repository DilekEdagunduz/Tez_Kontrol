
from tkinter import filedialog
from tkinter import *
import tkinter as tk
import os
import docx
from Data import bolum,bolum2,bolum3,word_acma

root = tk.Tk()
root.title('Word Kontrol')
root.geometry("300x200")
root.eval('tk::PlaceWindow . center')
root.resizable(False, False)




word_acma.Clear_Console()




def browsefunc():


    root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Word Dosyasını Seçin",filetypes = (("Word dosyaları",".docx"),("Tüm Dosyalar",".*")))
    label1.config(text='{}'.format(os.path.basename(root.filename)))
    if os.path.exists(root.filename):
        b1.config(text="Yeniden Seç")
        word_acma.Clear_Console()
        #print (root.filename)
        dosyayol=root.filename
        #print(dosyayol)
        bolum3.cifttirnakfonk(dosyayol)
        bolum2.Paragrafmidegilmi(dosyayol)
        bolum.Iceriyomu(dosyayol)

        #os.startfile('WordRapor.docx')
        filename='WordRapor.docx'
        word_acma.Open_file(filename)



b1=tk.Button(root,text="Dosya Seç",font=40,command=browsefunc)
spaceLabel = tk.Label(root, text= "                     ")
label1 = tk.Label(root, text= "Lütfen bir word dosyası seçin.")
spaceLabel.pack()
label1.pack()
b1.pack()



root.mainloop()
