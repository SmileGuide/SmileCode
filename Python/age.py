import easygui
import os

#Age
fl= os.open("l.age","r+")
fn=fl.readline()
a = codemao.Image(fn).face_recognize("age")
wt=fl.write(a)
cl=fl.close()
fe=os.rename("l.age","a.age")
easygui.msgbox(a)



