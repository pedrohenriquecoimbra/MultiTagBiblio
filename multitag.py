"""USAGE :
In Word, annotate before references

"""
############################ IMPORTS #############################

import os
from tkinter import *

############################ FUNCTIONS ###########################


class Mainw:
    def __init__(self, win, p):
        
        self.pth = p
        self.TAG_PLAN = self.open_plan
        self.disp_list = Variable(value = [self.TAG_PLAN[k][4]*"_" + self.TAG_PLAN[k][0] for k in range(len(self.TAG_PLAN))])
        self.listbox = Listbox(
            win,
            listvariable=self.disp_list,
            height = 20,
            width = 50,
            selectmode = EXTENDED)


        self.listbox.place(x=600, y=40)
        self.listbox.bind('<<ListboxSelect>>', self.tag_select)

        self.sep_field = Text(
            win,
            height=1,
            width=25)
        self.sep_field.place(x=40, y=40)

        self.extract_field = Text(
            win,
            height=1,
            width=25)
        self.extract_field.place(x=275, y=40)

        self.input_but = Button(
            win,
            text = 'Input',
            height=1,
            width=10,
            command = self.open_blocs)
        self.input_but.place(x=500, y=35)

        self.merge_but = Button(
            win,
            text = 'Merge',
            height=1,
            width=10,
            command = self.bloc_merge)
        self.merge_but.place(x=300, y=600)
        
        self.next_but = Button(
            win,
            text = 'Next',
            height=1,
            width=10,
            command = self.bloc_next)
        self.merge_but.place(x=300, y=700)

        self.previous_label = Text(
            win,
            height=8,
            width=60)
        self.previous_label.place(x=40, y=200)

        self.current_label = Text(
            win,
            height=8,
            width=60)
        self.current_label.place(x=40, y=400)
    
    def open_plan(self):
        tpr = open(self.pth + '/tag_plan.txt', 'r').read().split("\n")
        TAG_PLAN = []
        for k in range(len(tpr)):
            ID = k
            order = tpr[k].count("_")
            TAG_PLAN += [[tpr[k][order:], ID, 0, 1, order]]
        nb = 0
        for k in range(1, len(TAG_PLAN)):
            if TAG_PLAN[k][4] > TAG_PLAN[k-1][4]:
                TAG_PLAN[k][2] = TAG_PLAN[k-1][1]
            elif TAG_PLAN[k][4] == TAG_PLAN[k-1][4]:
                TAG_PLAN[k][3] = TAG_PLAN[k-1][3] + 1
                TAG_PLAN[k][2] = TAG_PLAN[k-1][2]
            elif TAG_PLAN[k][4] < TAG_PLAN[k-1][4]:
                nb = 0
                for i in range(1, k+1):
                    if TAG_PLAN[:k][-i][4] == TAG_PLAN[k][4]:
                        TAG_PLAN[k][2] = TAG_PLAN[:k][-i][2]
                        TAG_PLAN[k][3] = TAG_PLAN[:k][-i][3] + 1
                        break
        return TAG_PLAN
    
    def open_blocs(self):
        sep = self.sep_field.get("1.0","end-1c")
        article = self.extract_field.get("1.0","end-1c")
        self.new = 0
        fr = self.load_file(self.pth + '/blocs.txt')
        Saved_Blocs = self.list_import(fr)[:-1]
        #print(Saved_Blocs)
        fr.close()
        
        
        if sep not in article or len(sep) == 0:
            self.previous_label.delete('1.0', "end-1c")
            self.previous_label.insert("1.0","Veuillez reessayer, separateur introuvable")
            return None
        
        
        if self.new == 1:
            saved_id_max = 0
        else:
            saved_id_max = max([int(k[1]) for k in Saved_Blocs])
        
        self.BLOCS = []
        bloc_s = 0
        self.new = 1
        
        for k in range(article.count(sep)):
            bloc_e = article.index(sep)
            if self.new == 1:
                ID = saved_id_max + 1
                self.new = 0
            else:
                ID = max(saved_id_max, max([k[1] for k in self.BLOCS])) + 1
            self.BLOCS += [[article[bloc_s:bloc_e].replace('\n', '').replace('\t', ''), ID, sep]]
            article = article[bloc_e + len(sep):]
        self.BLOCS += [[article[bloc_s:bloc_e].replace('\n', '').replace('\t', ''), ID+1, sep]]
        
        self.mrg_count = 0
        self.previous_label.delete('1.0', "end-1c")
        self.previous_label.insert("1.0",self.BLOCS[self.mrg_count][0])
        self.current_label.delete('1.0', "end-1c")
        self.current_label.insert("1.0",self.BLOCS[self.mrg_count+1][0])
        
        
        
    
    ### Widget Methods ###
    
    def tag_select(self, event):
        init = self.listbox.curselection()
        for i in init:
            current = self.TAG_PLAN[i]
            while current[4] != 0:
                parent = current[2]
                for k in range(len(self.TAG_PLAN)):
                    if self.TAG_PLAN[k][1] == parent:
                        current = self.TAG_PLAN[k]
                        self.listbox.select_set(self.TAG_PLAN.index(self.TAG_PLAN[k]))
    
    def bloc_merge(self):
        
        self.BLOCS[self.mrg_count][0] = self.BLOCS[self.mrg_count][0] + ' ' + self.BLOCS[self.mrg_count + 1][0]
        self.BLOCS.pop(self.mrg_count + 1)
        
        self.previous_label.delete('1.0', "end-1c")
        self.previous_label.insert("1.0",self.BLOCS[self.mrg_count][0])
        self.current_label.delete('1.0', "end-1c")
        self.current_label.insert("1.0",self.BLOCS[self.mrg_count+1][0])
        
        if self.mrg_count == len(self.BLOCS) - 1:
            self.previous_label.delete('1.0', "end-1c")
            self.previous_label.delete('1.0', "end-1c")
        
            fw = open(path + '/blocs.txt', 'a+')
            self.list_save(self.BLOCS, fw)
            fw.close()
        
    def bloc_next(self):
        
        self.mrg_count += 1
        
        self.previous_label.delete('1.0', "end-1c")
        self.previous_label.insert("1.0",self.BLOCS[self.mrg_count][0])
        self.current_label.delete('1.0', "end-1c")
        self.current_label.insert("1.0",self.BLOCS[self.mrg_count+1][0])
        
        if self.mrg_count == len(self.BLOCS) - 1:
            self.previous_label.delete('1.0', "end-1c")
            self.previous_label.delete('1.0', "end-1c")
        
            fw = open(path + '/blocs.txt', 'a+')
            self.list_save(self.BLOCS, fw)
            fw.close()
    
    ### General Methods ###
    
    def load_file(self, p):
        if not os.path.exists(p):
            self.new = 1
            F = open(p, 'a+')
            F.close()
        return open(p, 'r')

    def list_import(self, F):
        SB = F.read()
        SB = SB.split("\n")
        for k in range(len(SB)):
            SB[k] = SB[k].split(",;")
        return SB

    def list_save(self, X, F):
        for k in X:
            for j in k:
                F.write(str(j))
                F.write(",;")
            F.write("\n")
                        
    

####################### SCRIPT ##############################

path = "C:/Users/tigerault/PycharmProjects/Multitag/Storage"

window=Tk()
win=Mainw(window, path)
window.title('Add to plan')
window.geometry("1400x720")
window.mainloop()


### TAGS CORSS TABLE SECTION
# id, bloc numbers (0 if classe1)
tr = load_file(path + '/tags.txt')

TAGS = [[], [], []]

###




### TAG_PLAN MANAGEMENT SECTION
tpr = open(path + '/tag_plan.txt', 'r').read().split("\n")
TAG_PLAN = []
for k in range(len(tpr)):
    ID = k
    order = tpr[k].count("_")
    TAG_PLAN += [[tpr[k][order:], ID, 0, 1, order]]
nb = 0
for k in range(1, len(TAG_PLAN)):
    if TAG_PLAN[k][4] > TAG_PLAN[k-1][4]:
        TAG_PLAN[k][2] = TAG_PLAN[k-1][1]
    elif TAG_PLAN[k][4] == TAG_PLAN[k-1][4]:
        TAG_PLAN[k][3] = TAG_PLAN[k-1][3] + 1
        TAG_PLAN[k][2] = TAG_PLAN[k-1][2]
    elif TAG_PLAN[k][4] < TAG_PLAN[k-1][4]:
        nb = 0
        for i in range(1, k+1):
            if TAG_PLAN[:k][-i][4] == TAG_PLAN[k][4]:
                TAG_PLAN[k][2] = TAG_PLAN[:k][-i][2]
                TAG_PLAN[k][3] = TAG_PLAN[:k][-i][3] + 1
                break

###


#disp_list = Variable(value = [TAG_PLAN[k][4]*"_" + TAG_PLAN[k][0] for k in range(len(TAG_PLAN))])




# Now 1 how to multitag a block

# Add a Block save variable to merge informations after review
