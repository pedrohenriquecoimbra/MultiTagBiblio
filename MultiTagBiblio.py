# Imports
import os
import sys
import subprocess
cwd = os.getcwd()
p = cwd + "\\Storage"
# Check whether the specified path exists or not
dependencies = ['tk', 'tkhtmlview', 'nltk', 'sentence_transformers', 'matplotlib', 'scipy', 'sklearn']
if not os.path.exists(p):
    setup = input("First time use, install dependancies? (y/n) :")
    if setup == 'y':
        for k in dependencies:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', k])
        import nltk
        nltk.download('punkt')
        nltk.download("stopwords")
        nltk.download('wordnet')

from tkinter import *
from tkinter.filedialog import askdirectory
import pickle
import sqlite3
import json
from tkhtmlview import HTMLLabel
import nltk
from sentence_transformers import SentenceTransformer, util
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage
from sklearn.cluster import AgglomerativeClustering


# Functions


class Biblio:

    def __init__(self, window):
        cwd = os.getcwd()
        self.p = cwd + "\\Storage"
        self.blocs = self.import_dict(self.p, 'blocs')
        self.tag_list = self.build_tag_list(self.blocs)
        self.plan = self.import_dict(self.p, 'plan')
        self.zotero = self.import_dict(self.p, 'Zotero_data')
        self.tagging = 0
        self.window = window

        self.var = IntVar()

        self.merge_var = IntVar()

        self.save_var = IntVar()

        self.merge_tick = Checkbutton(
            window,
            text='Merge ?',
            variable=self.merge_var)
        self.merge_tick.place(x=1300, y=500)

        self.input_but = Button(
            window,
            text='Import Zot',
            height=1,
            width=10,
            command=self.add_to_blocs)
        self.input_but.place(x=380, y=460)

        self.info_but = Button(
            window,
            text='Info',
            height=1,
            width=7,
            command=self.send_key)
        self.info_but.place(x=460, y=460)

        self.del_art_but = Button(
            window,
            text='Remove',
            height=1,
            width=7,
            command=self.delete_article)
        self.del_art_but.place(x=430, y=510)

        self.tag_but = Button(
            window,
            text='Tagging',
            height=1,
            width=10,
            command=self.tag_blocs)
        self.tag_but.place(x=1100, y=500)

        self.up_but = Button(
            window,
            text='^',
            height=1,
            width=3,
            command=self.move_up_plan)
        self.up_but.place(x=70, y=493)

        self.down_but = Button(
            window,
            text='v',
            height=1,
            width=3,
            command=self.move_down_plan)
        self.down_but.place(x=70, y=520)

        self.left_but = Button(
            window,
            text='<',
            height=1,
            width=3,
            command=self.move_left_plan)
        self.left_but.place(x=40, y=505)

        self.right_but = Button(
            window,
            text='>',
            height=1,
            width=3,
            command=self.move_right_plan)
        self.right_but.place(x=102, y=505)

        self.add_tag_but = Button(
            window,
            text='Add',
            height=1,
            width=7,
            command=self.add_plan)
        self.add_tag_but.place(x=150, y=500)

        self.take_note_but = Button(
            window,
            text='Take notes',
            height=1,
            width=10,
            command=self.edit_notes_from_plan)
        self.take_note_but.place(x=1300, y=650)

        self.del_tag_but = Button(
            window,
            text='Delete',
            height=1,
            width=7,
            command=self.delete_plan)
        self.del_tag_but.place(x=220, y=500)

        self.edit_tag_but = Button(
            window,
            text='Edit',
            height=1,
            width=7,
            command=self.edit_plan)
        self.edit_tag_but.place(x=290, y=500)

        self.tag_next_but = Button(
            window,
            text='Next',
            height=1,
            width=10,
            command=lambda: self.var.set(1))
        self.tag_next_but.place(x=1200, y=500)

        self.search_but = Button(
            window,
            text='Search',
            height=1,
            width=7,
            command=self.blocs_filter_search)
        self.search_but.place(x=800, y=500)

        self.topics_but = Button(
            window,
            text='Topics',
            height=1,
            width=7,
            command=self.blocs_main_subjects)
        self.topics_but.place(x=860, y=500)

        self.save_note_but = Button(
            window,
            text='Save',
            height=1,
            width=10,
            command=lambda: self.save_var.set(1))
        self.save_note_but.place(x=1300, y=700)

        self.export_but = Button(
            window,
            text='Export all',
            height=1,
            width=10)
        self.export_but.place(x=1400, y=750)

        self.plan_listbox = Listbox(
            window,
            height=30,
            width=50,
            selectmode=EXTENDED)
        self.plan_listbox.place(x=50, y=10)
        new = self.build_plan()
        for k in range(len(new)):
            self.plan_listbox.insert(k, new[k])
        self.plan_listbox.bind('<<ListboxSelect>>', self.blocs_filter_plan)

        self.source_listbox = Listbox(
            window,
            height=27,
            width=20,
            selectmode=EXTENDED)
        self.source_listbox.place(x=400, y=10)
        for k in unique(self.blocs["source"]):
            self.source_listbox.insert(END, k[0])
        self.source_listbox.bind('<<ListboxSelect>>', self.blocs_filter_sources)

        self.blocs_listbox = Listbox(
            window,
            height=30,
            width=60,
            selectmode=EXTENDED)
        self.blocs_listbox.place(x=550, y=10)
        self.blocs_listbox.bind('<<ListboxSelect>>', self.read_blocs)

        self.search_text = Text(
            window,
            height=1,
            width=30)
        self.search_text.place(x=550, y=500)

        self.shell_label = Label(window, text="Shell :")
        self.shell_label.place(x=1000, y=10)

        self.shell_text = Text(
            window,
            height=28,
            width=60)
        self.shell_text.place(x=1000, y=40)

        self.notes_text = Text(
            window,
            height=17,
            width=150)
        self.notes_text.place(x=50, y=550)

    # Dictionary management

    def import_dict(self, path, name):
        with open(path + '\\' + name + '.pkl', 'rb') as f:
            return pickle.load(f)

    def save_dict(self, path, name, d):
        with open(path + '\\' + name + '.pkl', 'wb') as f:
            pickle.dump(d, f)

    # Tag plan management

    def build_tag_list(self, blocs):
        tl = []
        for i in blocs["tag"]:
            for j in i:
                if len(j) > 0:
                    if j[0] not in tl:
                        tl += [j]
        return tl

    def build_plan(self):
        title = [['I. ', 'II. ', 'III. ', 'IV. ', 'V. ', 'VI. ', 'VII. ', 'VIII. ', 'IX. ', 'X. '],
                    ['A. ', 'B. ', 'C. ', 'D. ', 'E. ', 'F. ', 'G. ', 'H. ', 'I. ', 'J. '],
                    ['1. ', '2. ', '3. ', '4. ', '5. ', '6. ', '7. ', '8. ', '9. ', '10. '],
                    ['a. ', 'b. ', 'c. ', 'd. ', 'e. ', 'f. ', 'g. ', 'h. ', 'i. ', 'j. '],
                    ['i. ', 'ii. ', 'iii. ', 'iv. ', 'v. ', 'vi. ', 'vii. ', 'viii. ', 'ix. ', 'x. '],
                    ['-> ', '-> ', '-> ', '-> ', '-> ', '-> ', '-> ', '-> ', '-> ', '-> ']]
        ct = [0, 0, 0, 0, 0, 0]
        built_plan = ['' for k in range(len(self.plan["position"]))]
        for p in range(len(self.plan["position"])):
            k = self.plan["position"].index(p)
            for l in self.tag_list:
                if self.plan["ID"][k] == l[1]:
                    order = self.plan["order"][k]
                    built_plan[self.plan["position"][k]] = '___' * order + title[order][ct[order]] + l[0]
                    ct[order] += 1
                    if p < len(self.plan["position"])-1:
                        diff = order - self.plan["order"][self.plan["position"].index(p+1)]
                        if diff > 0:
                            for j in range(diff):
                                ct[order - j] = 0
                    break
        return built_plan

    def move_left_plan(self):
        pos = self.plan_listbox.curselection()
        for j in pos:
            for k in range(len(self.plan["position"])):
                if j == self.plan["position"][k] and self.plan["order"][k] != 0:
                    self.plan["order"][k] -= 1

        self.plan_listbox.delete(0, END)
        new = self.build_plan()
        for k in range(len(new)):
            self.plan_listbox.insert(k, new[k])

        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

    def move_right_plan(self):
        pos = self.plan_listbox.curselection()
        for j in pos:
            for k in range(len(self.plan["position"])):
                if j == self.plan["position"][k] and self.plan["order"][k] != 5:
                    self.plan["order"][k] += 1

        self.plan_listbox.delete(0, END)
        new = self.build_plan()
        for k in range(len(new)):
            self.plan_listbox.insert(k, new[k])

        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

    def move_up_plan(self):
        pos = self.plan_listbox.curselection()[0]
        for k in range(len(self.plan["position"])):
            if pos == self.plan["position"][k]:
                current = k
            elif pos - 1 == self.plan["position"][k]:
                target = k
        tp = self.plan["position"][target]
        self.plan["position"][target] = self.plan["position"][current]
        self.plan["position"][current] = tp

        self.plan_listbox.delete(0, END)
        new = self.build_plan()
        for k in range(len(new)):
            self.plan_listbox.insert(k, new[k])

        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

    def move_down_plan(self):
        pos = self.plan_listbox.curselection()[0]
        for k in range(len(self.plan["position"])):
            if pos == self.plan["position"][k]:
                current = k
            elif pos + 1 == self.plan["position"][k]:
                target = k
        tp = self.plan["position"][target]
        self.plan["position"][target] = self.plan["position"][current]
        self.plan["position"][current] = tp

        self.plan_listbox.delete(0, END)
        new = self.build_plan()
        for k in range(len(new)):
            self.plan_listbox.insert(k, new[k])

        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

    def add_plan(self):
        # to be called on edit button press
        pos = self.plan_listbox.curselection()

        # Get plan tag name
        self.shell_text.delete("1.0", "end-1c")
        self.shell_label.configure(text='New category? :')
        self.tag_next_but.wait_variable(self.var)
        new_tag = self.shell_text.get("1.0", "end-1c")
        if new_tag not in [j[0] for j in self.tag_list]:
            if len(self.plan["ID"]) == 0:
                new_ID = 1
            else:
                new_ID = max(self.plan["ID"]) + 1
            self.plan["ID"] += [new_ID]
            # if no position has been selected
            if len(pos) == 0:
                position = 0
            else:
                position = pos[0] + 1
            if position in self.plan["position"]:
                for k in range(len(self.plan["position"])):
                    if position <= self.plan["position"][k]:
                        self.plan["position"][k] += 1
            self.plan["position"] += [position]
            # Get tag order
            self.shell_text.delete("1.0", "end-1c")
            self.shell_label.configure(text='order :')
            self.tag_next_but.wait_variable(self.var)
            self.plan["order"] += [int(self.shell_text.get("1.0", "end-1c"))]
            self.plan["note"] += ['']
            self.blocs["tag"][0] += [[new_tag, new_ID]]

            self.shell_text.delete("1.0", "end-1c")
            self.shell_label.configure(text='Shell :')

            self.save_dict(self.p, 'blocs', self.blocs)
            self.blocs = self.import_dict(self.p, 'blocs')
            self.tag_list = self.build_tag_list(self.blocs)
            self.save_dict(self.p, 'plan', self.plan)
            self.plan = self.import_dict(self.p, 'plan')

            self.plan_listbox.delete(0, END)
            new = self.build_plan()
            for k in range(len(new)):
                self.plan_listbox.insert(k, new[k])

    def delete_plan(self):
        pos = self.plan_listbox.curselection()[0]
        deleted = 0
        for i in range(len(self.plan["position"])):
            k = i - deleted
            if pos < self.plan["position"][k]:
                self.plan["position"][k] -= 1
            elif pos == self.plan["position"][k]:
                deleted = 1
                deleted_id = self.plan["ID"][k]
                del self.plan["position"][k]
                del self.plan["ID"][k]
                del self.plan["order"][k]
                del self.plan["note"][k]

        for k in range(len(self.blocs["tag"])):
            deleted = 0
            for i in range(len(self.blocs["tag"][k])):
                j = i - deleted
                if self.blocs["tag"][k][j][1] == deleted_id:
                    deleted += 1
                    del self.blocs["tag"][k][j]

        self.save_dict(self.p, 'blocs', self.blocs)
        self.blocs = self.import_dict(self.p, 'blocs')
        self.tag_list = self.build_tag_list(self.blocs)
        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

        self.plan_listbox.delete(0, END)
        new = self.build_plan()
        for k in range(len(new)):
            self.plan_listbox.insert(k, new[k])

    def edit_plan(self):
        pos = self.plan_listbox.curselection()[0]
        # Ask user input
        self.shell_text.delete("1.0", "end-1c")
        self.shell_label.configure(text='Change name to :')
        self.tag_next_but.wait_variable(self.var)
        new_name = self.shell_text.get("1.0", "end-1c")
        self.shell_text.delete("1.0", "end-1c")
        self.shell_label.configure(text='Shell :')

        for k in range(len(self.plan["position"])):
            if pos == self.plan["position"][k]:
                id = self.plan["ID"][k]

        for k in range(len(self.blocs["tag"])):
            for j in range(len(self.blocs["tag"][k])):
                if self.blocs["tag"][k][j][1] == id:
                    self.blocs["tag"][k][j][0] = new_name

        self.tag_list = self.build_tag_list(self.blocs)
        self.plan_listbox.delete(0, END)
        new = self.build_plan()
        for k in range(len(new)):
            self.plan_listbox.insert(k, new[k])

        self.save_dict(self.p, 'blocs', self.blocs)
        self.blocs = self.import_dict(self.p, 'blocs')
        self.tag_list = self.build_tag_list(self.blocs)
        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

    def edit_notes_from_plan(self):
        # Get corresponding notes
        pos = self.plan["position"].index(self.plan_listbox.curselection()[0])
        self.notes_text.delete('1.0', "end-1c")
        self.notes_text.insert("1.0", self.plan["note"][pos])
        # Wait for save button press
        self.save_note_but.wait_variable(self.save_var)
        note = self.notes_text.get("1.0", "end-1c")
        self.plan["note"][pos] = note

        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

    # Blocs management

    def add_to_blocs(self):
        print(self.zotero)
        sources, highlights, notes = self.zotero_import()
        for k in range(len(highlights)):
            if highlights[k] not in self.blocs['text']:
                self.blocs['text'] += [highlights[k]]
                self.blocs['source'] += [sources[k]]
                self.blocs['note'] += [notes[k]]
                self.blocs['tag'] += [[]]

        self.save_dict(self.p, 'blocs', self.blocs)
        self.blocs = self.import_dict(self.p, 'blocs')

        self.source_listbox.delete(0, END)
        for k in unique(self.blocs["source"]):
            self.source_listbox.insert(END, k[0])

    def delete_article(self):
        pos = self.source_listbox.curselection()[0]
        article = self.source_listbox.get(pos)
        deleted = 0
        for j in range(len(self.blocs["source"])):
            k = j - deleted
            if self.blocs["source"][k][0] == article:
                del self.blocs["text"][k]
                del self.blocs["source"][k]
                del self.blocs["tag"][k]
                deleted += 1

        self.save_dict(self.p, 'blocs', self.blocs)
        self.blocs = self.import_dict(self.p, 'blocs')

        self.source_listbox.delete(0, END)
        for k in unique(self.blocs["source"]):
            self.source_listbox.insert(END, k[0])

    def tag_blocs(self):
        self.tagging = 1
        self.tag_list = self.build_tag_list(self.blocs)
        selected = []
        source = self.source_listbox.curselection()
        if len(source) == 0:
            current = self.blocs_listbox.curselection()
            for k in current:
                selected += [self.blocs_listbox.get(k)]

            for k in range(len(current)):
                # Show text in shell
                self.shell_text.delete('1.0', "end-1c")
                self.blocs_listbox.itemconfig(current[k], bg='green')
                self.plan_listbox.selection_clear(0, END)
                add_tag = self.blocs["text"].index(selected[k])
                self.shell_text.insert(END, self.blocs["text"][add_tag])
                # Show existing tags
                # Reset colors
                for p in self.plan["position"]:
                    self.plan_listbox.itemconfig(p, bg='white')
                # Color existing tags
                for j in range(len(self.blocs["text"])):
                    if self.blocs["text"][j] == self.blocs["text"][add_tag]:
                        for i in self.blocs["tag"][j]:
                            for m in range(len(self.plan["ID"])):
                                if i[1] == self.plan["ID"][m]:
                                    self.plan_listbox.select_set(self.plan["position"][m])
                                    # switch to select set
                # Wait for button press
                self.tag_next_but.wait_variable(self.var)
                self.blocs_listbox.itemconfig(current[k], bg='white')
                add = []
                focus = self.plan_listbox.curselection()
                # if a tag selection has been made
                if len(focus) > 0:
                    for m in focus:
                        tag = [m]
                        # To associate bloc to selected tag but also his parents.
                        for o in range(self.plan["order"][self.plan["position"].index(tag[0])]):
                            tag += [self.get_parent(tag[-1])]
                        for i in tag:
                            id = self.plan["ID"][self.plan["position"].index(i)]
                            for j in self.tag_list:
                                if j[1] == id:
                                    add += [j]

                    self.blocs["tag"][add_tag] = add

        else:
            sep = self.source_listbox.get(source)
            for k in range(len(self.blocs["text"])):
                if self.blocs["source"][k][0] == sep:
                    selected += [self.blocs["text"][k]]

            deleted = 0
            for l in range(len(selected)):
                k = l - deleted
                # Color current focus block
                self.blocs_listbox.itemconfig(l, bg='green')
                # Show text in shell
                self.shell_text.delete('1.0', "end-1c")

                if self.merge_var.get() == 0:
                    add_tag = self.blocs["text"].index(selected[k])
                if l < len(selected) - 1:
                    next = self.blocs["text"].index(selected[k + 1])
                    self.shell_text.insert(END, self.blocs["text"][add_tag] + '\n\nNext : \n\n' + self.blocs["text"][next])
                else:
                    self.shell_text.insert(END, self.blocs["text"][add_tag])
                self.merge_var.set(0)

                # Show existing tags
                # Reset colors
                for k in self.plan["position"]:
                    self.plan_listbox.itemconfig(k, bg='white')
                # Color existing tags
                for j in range(len(self.blocs["text"])):
                    if self.blocs["text"][j] == self.blocs["text"][add_tag]:
                        for i in self.blocs["tag"][j]:
                            for m in range(len(self.plan["ID"])):
                                if i[1] == self.plan["ID"][m]:
                                    self.plan_listbox.itemconfig(self.plan["position"][m], bg='green')
                # Wait for button press
                self.tag_next_but.wait_variable(self.var)
                if self.merge_var.get() == 1:
                    deleted += 1
                    # Merge text
                    self.blocs["text"][add_tag] += self.blocs["text"][next]
                    # Merge existing tags
                    self.blocs["tag"][add_tag] += self.blocs["tag"][next]

                    del selected[k + 1]
                    del self.blocs["text"][next]
                    del self.blocs["source"][next]
                    del self.blocs["tag"][next]

                    self.save_dict(self.p, 'blocs', self.blocs)
                    self.blocs = self.import_dict(self.p, 'blocs')

                add = []
                focus = self.plan_listbox.curselection()
                # if a tag selection has been made
                if len(focus) > 0 :
                    for m in focus:
                        tag = [m]
                        # To associate bloc to selected tag but also his parents.
                        for k in range(self.plan["order"][self.plan["position"].index(tag[0])]):
                            tag += [self.get_parent(tag[-1])]
                        for i in tag:
                            id = self.plan["ID"][self.plan["position"].index(i)]
                            for j in self.tag_list:
                                if j[1] == id:
                                    add += [j]

                    self.blocs["tag"][add_tag] += add
                    unique_tags = []
                    for n in self.blocs["tag"][add_tag]:
                        if n not in unique_tags:
                            unique_tags += [n]
                    self.blocs["tag"][add_tag] = unique_tags

        self.save_dict(self.p, 'blocs', self.blocs)
        self.blocs = self.import_dict(self.p, 'blocs')
        self.tag_list = self.build_tag_list(self.blocs)
        self.save_dict(self.p, 'plan', self.plan)
        self.plan = self.import_dict(self.p, 'plan')

        self.tagging = 0

    def get_parent(self, position):
        parent = []
        for k in range(len(self.plan["position"])):
            if self.plan["position"][k] < position and self.plan["order"][k] == self.plan["order"][self.plan["position"].index(position)] - 1:
                parent += [self.plan["position"][k]]
        return max(parent)

    # Blocs visualization filters

    def access_from_plan(self):
        # to be called as soon as pressed
        # get selected from list
        position_selected = 0
        blocs_selected = []
        for k in range(len(self.blocs["tag"])):
            for l in range(len(self.blocs["tag"][k])):
                if self.plan["ID"][position_selected] == self.blocs["tag"][k][l]:
                    blocs_selected += [self.blocs["text"][k]]

    def blocs_filter_plan(self, event):
        if self.tagging == 0:
            cursor = self.plan_listbox.curselection()
            for i in cursor:
                pos = self.plan["position"].index(i)
                selected = []
                sources = []
                # Skip the firt 'default' source
                for k in range(1, len(self.blocs["text"])):
                    for l in range(len(self.blocs["tag"][k])):
                        if len(self.blocs["tag"][k][l]) > 0:
                            if self.blocs["tag"][k][l][1] == self.plan["ID"][pos]:
                                selected += [self.blocs["text"][k]]
                                sources += [self.blocs["source"][k]]
                sources = unique(sources)

                self.blocs_listbox.delete(0, END)
                for k in selected:
                    self.blocs_listbox.insert(END, k)
                for k in sources:
                    index = self.source_listbox.get(0, "end").index(k[0])
                    self.source_listbox.itemconfig(index, bg='green')

    def blocs_filter_sources(self, event):
        pos = self.source_listbox.curselection()[0]
        selected = []
        for k in range(len(self.blocs["text"])):
            if self.blocs["source"][k][0] == self.source_listbox.get(pos):
                selected += [self.blocs["text"][k]]

        self.blocs_listbox.delete(0, END)
        for k in selected:
            self.blocs_listbox.insert(END, k)

    def blocs_filter_search(self):
        self.blocs_listbox.delete(0, END)
        request = self.search_text.get("1.0", "end-1c")
        lemma = nltk.stem.WordNetLemmatizer().lemmatize(request)
        for k in self.blocs["text"]:
            if lemma in k:
                self.blocs_listbox.insert(END, k)

    def blocs_main_subjects(self):
        tensors = []
        # load pre-trained model
        model = SentenceTransformer('all-MiniLM-L6-v2')
        for k in self.blocs["text"]:
            if k != 'default':
                selected = model.encode(k)
                compare = model.encode(self.blocs["text"])
                cos_sim = util.cos_sim(selected, compare)
                tensors += [cos_sim.tolist()[0]]

        # Classifying vectorial products for each bloc
        linkage_data = linkage(tensors, method='ward', metric='euclidean')
        dendrogram(linkage_data)

        self.shell_text.delete('1.0', "end-1c")
        self.shell_label.configure(text='Number of clusters? :')
        plt.show()
        self.tag_next_but.wait_variable(self.var)
        nb_clusters = int(self.shell_text.get("1.0", "end-1c"))

        hierarchical_cluster = AgglomerativeClustering(n_clusters=nb_clusters, affinity='euclidean', linkage='ward')
        labels = hierarchical_cluster.fit_predict(tensors)
        groups = ['' for k in unique(labels)]

        for k in range(len(self.blocs["text"])-1):
            groups[labels[k]] += self.blocs["text"][k+1]

        self.shell_text.delete('1.0', "end-1c")
        for k in range(len(groups)):
            # Frequent keywords and word combinations in blocs
            words = nltk.tokenize.word_tokenize(groups[k])
            stop_words = set(nltk.corpus.stopwords.words("english"))
            stop_words.update(',', '.', '(', ')', '[', ']', 'Figure')
            filtered_words = []
            # Filter words and set to their lemma
            lemmatizer = nltk.stem.WordNetLemmatizer()
            for word in words:
                if word.casefold() not in stop_words:
                    filtered_words.append(lemmatizer.lemmatize(word))
            self.shell_text.insert(END, "Group " + str(k) + " :\n")
            for k in nltk.FreqDist(filtered_words).most_common(2):
                self.shell_text.insert(END, str(k)+"\n")

            # Study bi-collocation information
            bigram_measures = nltk.collocations.BigramAssocMeasures()
            finder = nltk.collocations.BigramCollocationFinder.from_words(filtered_words)
            for k in finder.nbest(bigram_measures.likelihood_ratio, 2):
                self.shell_text.insert(END, str(k) + "\n")

    def read_blocs(self, event):
        if self.tagging == 0:
            # Reset widgets
            self.shell_text.delete('1.0', "end-1c")
            for k in [unique(self.blocs["source"]).index(l) for l in self.blocs["source"]]:
                self.source_listbox.itemconfig(k, bg="white")
            for k in self.plan["position"]:
                self.plan_listbox.itemconfig(k, bg='white')

            for i in self.blocs_listbox.curselection():
                # Display in shell
                self.shell_text.insert(END, self.blocs_listbox.get(i))

                # Show corresponding tags and sources
                for j in range(len(self.blocs["text"])):
                    if self.blocs["text"][j] == self.blocs_listbox.get(i):
                        source = self.blocs["source"][j]
                        self.source_listbox.itemconfig(unique(self.blocs["source"]).index(source), bg='green')
                        note = self.blocs["note"][j]
                        if len(note) > 0:
                            self.shell_text.insert(END, '\n\nNOTE : ' + note)

                        for k in self.blocs["tag"][j]:
                            for m in range(len(self.plan["ID"])):
                                if k[1] == self.plan["ID"][m]:
                                    self.plan_listbox.itemconfig(self.plan["position"][m], bg='green')

                self.shell_text.insert(END, '\n----\n')

    def send_key(self):
        pos = self.source_listbox.curselection()[0]
        for k in self.blocs["source"]:
            if k[0] == self.source_listbox.get(pos):
                key = k[1]
                break
        ArticleInfo(key=key, zotero=self.zotero)

    # Exchanges with Zotero
    def zotero_import(self):
        # standard database connection
        conn = sqlite3.connect(self.zotero['path'] + '/zotero.sqlite')
        cur = conn.cursor()
        # Then disconnect?
        # access to database table "syncCache" for data retrieval
        sCacheData = cur.execute("SELECT data FROM syncCache").fetchall()
        names = """SELECT sql FROM sqlite_master
        WHERE tbl_name = 'collectionItems' AND type = 'table'"""
        print(cur.execute(names).fetchall())
        sql_query = """SELECT name FROM sqlite_master WHERE type='table';"""
        print(cur.execute(sql_query).fetchall())
        sql = """SELECT * FROM collectionItems"""
        print(cur.execute(sql).fetchall())
        #print(sCacheData)

        for k in sCacheData:
            document = json.loads(k[0])
            if 'name' in document['data'].keys():
                if self.zotero['target_collection'] == document['data']['name']:
                    parent = document['data']['key']
                    break

        items = []
        all_annotations = []
        for k in sCacheData:
            document = json.loads(k[0])
            if 'collections' in document['data'].keys():
                if parent in document['data']['collections'] and document['data']['itemType'] != 'attachment':
                    items += [document]
            elif 'annotationType' in document['data'].keys():
                if document['data']['annotationType'] == 'highlight' or document['data']['annotationType'] == 'note':
                    all_annotations += [document]

        annotations = [[] for k in items]
        for k in sCacheData:
            link = json.loads(k[0])
            for m in range(len(items)):
                for n in range(len(all_annotations)):
                    if 'parentItem' in link['data'].keys() and 'parentItem' in all_annotations[n]['data'].keys():
                        if link['data']['parentItem'] == items[m]['data']['key'] and link['data']['key'] == \
                                all_annotations[n]['data']['parentItem']:
                            annotations[m] += [all_annotations[n]['data']]

        sources, highlights, notes = [], [], []
        for k in range(len(annotations)):
            annotations[k] = sorted(annotations[k], key=lambda d: d['annotationSortIndex'])
            if 'creatorSummary' in items[k]['meta'].keys() and 'parsedDate' in items[k]['meta'].keys():
                parent_summary = items[k]['meta']['creatorSummary'] + ', ' + items[k]['meta']['parsedDate'][:4]
            elif 'parsedDate' in items[k]['meta'].keys():
                parent_summary = '?, ' + items[k]['meta']['parsedDate'][:4]
            elif'creatorSummary' in items[k]['meta'].keys():
                parent_summary = items[k]['meta']['creatorSummary'] + ', ?'
            else:
                parent_summary = 'Unknown'
            parent_id = items[k]['key']
            skip = 0
            for m in range(len(annotations[k])):
                j = m + skip
                if j < len(annotations[k]):
                    sources += [[parent_summary, parent_id]]
                    if 'annotationText' in annotations[k][j].keys():
                        highlights += [annotations[k][j]['annotationText']]
                        notes += ['']
                    elif 'annotationComment' in annotations[k][j].keys():
                        if j == len(annotations[k]) - 1:
                            highlights += ['Just a note']
                            notes += [annotations[k][j]['annotationComment']]
                        else:
                            highlights += [annotations[k][j + 1]['annotationText']]
                            notes += [annotations[k][j]['annotationComment']]
                            skip += 1

        return sources, highlights, notes


class ArticleInfo:
    def __init__(self, key, zotero):
        win2 = Tk()
        win2.title('Article metadata')
        win2.geometry("600x400")
        my_label = HTMLLabel(win2, height=20, width=70, html=self.get_meta(key, zotero))
        my_label.place(x = 10, y = 10)
        win2.mainloop()

    def get_meta(self, key, zotero):
        # standard database connection
        conn = sqlite3.connect(zotero['path'] + '/zotero.sqlite')
        cur = conn.cursor()
        # Then disconnect?
        # access to database table "syncCache" for data retrieval
        sCacheData = cur.execute("SELECT data FROM syncCache").fetchall()

        for k in sCacheData:
            document = json.loads(k[0])
            if document['key'] == key:
                parent = document
        return (parent['data']['title'] + '<br>-<br>'
                + ''.join([k['lastName'] + ', ' for k in parent['data']['creators']]) + '<br>-<br>'
                + parent['data']['date'][:4] + '<br>-<br>'
                + parent['data']['abstractNote'])


def init_dict():
    cwd = os.getcwd()
    p = cwd + "\\Storage"
    # Check whether the specified path exists or not
    if not os.path.exists(p):
        os.makedirs(p)
        d = dict(text=['default'], source=[['default', '']], note=['default'], tag=[[]])
        with open(p + "\\blocs.pkl", 'wb') as f:
            pickle.dump(d, f)

        d = dict(position=[], ID=[], order=[], note=[])
        with open(p + "\\plan.pkl", 'wb') as f:
            pickle.dump(d, f)

        print('No Zotero folder found')
        path = askdirectory(title='Select Zotero folder location')
        target_collection = input("Enter the parent Zotero collection you are working with : ")
        d = dict(path=path, target_collection=target_collection)
        with open(p + "\\Zotero_data.pkl", 'wb') as f:
            pickle.dump(d, f)


def unique(X):
    only_one = []
    for k in X:
        if k not in only_one:
            only_one += [k]
    return only_one


# Script

init_dict()
win = Tk()
win.title('Multi Tag Biblio')
win.state('zoomed')
mt = Biblio(win)
win.mainloop()
