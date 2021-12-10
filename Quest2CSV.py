# Copyright 2021 Sven Franz <sven.franz@jade-hs.de>
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

# 2021-07-09 Version 1.0 Beta 1
# 2021-10-27 Version 1.0 Beta 2
#   bugfix in keys2lower
#   remove survey-leaf
# 2021-11-29 Version 1.0 Beta 3
#   convert \r\n to ' '
#   NoneType to empty String
# 2021-12-01 Version 1.0 Beta 4
#   fill empty header
#   store in latin-1 (for windows)
# 2021-12-09 Version 1.0 Beta 5
#   removed: store in latin-1 (for windows)
#   fields added: date, time and DateTime
#   filename corrected
# 2021-12-10 Version 1.0 Beta 5
#   filewrite optimized

import sys, os
import configparser
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import xmltodict
import csv
import collections

configFileName = "settings.ini"
usingGui = False

def getConfig():
    config = configparser.RawConfigParser()
    if not os.path.isfile(configFileName):
        config.add_section('default')
        config.set('default', 'QuestionnaireFile', '')
        config.set('default', 'QuestionnaireResultFolder', '')
        config.set('default', 'CSVOutputFile', 'output.csv')
        with open(configFileName, 'w') as configfile:
            config.write(configfile)    
    config.read(os.path.join(os.getcwd(), configFileName))
    return config    
        
def onInput(event):
    if event.widget == entrys[0]:
        config.set('default', 'QuestionnaireFile', event.widget.get())
    elif event.widget == entrys[1]:
        config.set('default', 'QuestionnaireResultFolder', event.widget.get())
    elif event.widget == entrys[2]:
        config.set('default', 'CSVOutputFile', event.widget.get())
    
def handle_click(event):
    if event.widget == buttons[0]:
        value = filedialog.askopenfilename(filetypes=[(labels[0].cget("text"), "*.xml")])
        if value:
            entrys[0].delete(0, tk.END)
            entrys[0].insert(0, value)
            config.set('default', 'QuestionnaireFile', value)
    elif event.widget == buttons[1]:
        value = filedialog.askdirectory()
        if value:
            entrys[1].delete(0, tk.END)
            entrys[1].insert(0, value)
            config.set('default', 'QuestionnaireResultFolder', value)
    elif event.widget == buttons[2]:
        value = filedialog.askopenfilename(filetypes=[(labels[2].cget("text"), "*.csv")])
        if value:
            entrys[1].delete(0, tk.END)
            entrys[1].insert(0, value)
            config.set('default', 'CSVOutputFile', value)
    elif event.widget == buttons[3]:
        executeQuest2CSV(entrys[0].get(), entrys[1].get(), entrys[2].get(), False)
        messagebox.showinfo("Info", "Done")

        
def keys2lower(iterable):
    if type(iterable) is dict:
        keys = list(iterable.keys())
        for key in keys:
            iterable[key.lower()] = iterable.pop(key)
            if type(iterable[key.lower()]) is dict or type(iterable[key.lower()]) is list:
                iterable[key.lower()] = keys2lower(iterable[key.lower()])
    elif type(iterable) is list:
        for item in iterable:
            item = keys2lower(item)
    return iterable
        
def getValue(xmlData, searchItem, value = False):
    returnValue = False
    if type(xmlData) is dict:
        for data in xmlData:
            if data.lower() == searchItem.lower():
                if value != False and xmlData[searchItem].lower().replace("_", "") == value.lower().replace("_", ""):
                    return xmlData
                elif value == False:
                    return xmlData[searchItem]
            else:
                returnValue = getValue(xmlData[data], searchItem, value)
                if returnValue:
                    return returnValue
    elif type(xmlData) is list:
        for data in xmlData:
            returnValue = getValue(data, searchItem, value)
            if returnValue:
                return returnValue
    
def executeQuest2CSV(QuestionnaireFile, QuestionnaireResult, CSVOutputFile = "output.csv", Append = False):
    if Append == False and os.path.isfile(CSVOutputFile):
        os.remove(CSVOutputFile)
    error = False
    if not os.path.isfile(QuestionnaireFile):
        error = "Questionnaire File '" + QuestionnaireFile + "' does not exist!" 
    elif not os.path.isfile(QuestionnaireResult) and not os.path.isdir(QuestionnaireResult):
        error = "Questionnaire Results '" + QuestionnaireResult + "' does not exist!" 
    if error:
        if usingGui:
            messagebox.showerror("Error", error)
        else:
            raise ValueError(error)
    else:
        if os.path.isdir(QuestionnaireResult):
            for (dirpath, dirnames, filenames) in os.walk(QuestionnaireResult):
                for filename in filenames:
                    if filename.lower().endswith(".xml"):
                        executeQuest2CSV(QuestionnaireFile, os.path.join(dirpath, filename), CSVOutputFile, True)
        else:
            print("Executing: ", QuestionnaireResult)
            with open(QuestionnaireResult) as fd:
                result = keys2lower(xmltodict.parse(fd.read(), dict_constructor=dict))
            if "mobiquest" in result.keys() and "record" in result["mobiquest"].keys():
                if result["mobiquest"]["record"]["@survey_uri"].lower() == os.path.basename(QuestionnaireFile).lower():
                    with open(QuestionnaireFile) as fd:
                        questionnaire = keys2lower(xmltodict.parse(fd.read(), dict_constructor=dict))
                    if "mobiquest" in questionnaire.keys() and "survey" in questionnaire["mobiquest"].keys():
                        questionnaire["mobiquest"] = questionnaire["mobiquest"].pop("survey")  
                    if "mobiquest" in questionnaire.keys():
                        dataRow = {}
                        dataRow["File"] = os.path.basename(QuestionnaireResult)
                        dataRow["Subject"] = ""
                        dataRow["Date"] = QuestionnaireResult.split('_')[1]
                        dataRow["Time"] = QuestionnaireResult.split('_')[2].split('.')[0]
                        dataRow["DateTime"] = dataRow["Date"] + "_" + dataRow["Time"]
                        temppath = os.path.split(QuestionnaireResult)[0].split(os.sep)[-1]
                        if temppath.endswith("_Quest"):
                            dataRow["Subject"] = temppath.split("_Quest")[0]
                        dataRow["Motivation"] = getValue(result["mobiquest"], "@motivation")
                        for question in questionnaire["mobiquest"]["question"]:
                            if "label" in question.keys() and "text" in question["label"].keys() and "@id" in question.keys():
                                id = str(question["@id"]).replace("_", "")
                                if "@hidden" in question.keys() and question["@hidden"] == "true":
                                    dataRow[question['label']['text']] = getValue(result["mobiquest"], "@" + question["label"]["text"].replace(" ", "_").replace("(", "").replace(")", "").lower())
                                elif "@type" in question.keys() and question["@type"] == "checkbox" and "option" in question.keys():
                                    currAnswer = getValue(result["mobiquest"], "@question_id", question["@id"])
                                    dataRow[id] = {"text" : question["label"]["text"].replace("\r", "").replace("\n", " "), "value": 0, "values": {}}
                                    for option in question["option"]:
                                        option["@id"] = option["@id"].replace("_", "")
                                        dataRow[id]["values"][option["@id"]] = {"text": str(option["text"].replace("\r", "").replace("\n", " ")), "value" : 0}
                                        if "@option_ids" in currAnswer.keys() and option["@id"] in currAnswer["@option_ids"].split(";"):
                                            dataRow[id]["values"][option["@id"]]["value"] = 1
                                            dataRow[id]["value"] = 1;
                                else:
                                    dataRow[id] = {"text": question["label"]["text"].replace("\r", "").replace("\n", " "), "values": {}}
                                    res = getValue(result["mobiquest"], "@question_id", id)
                                    if "@option_ids" in res.keys():
                                        dataRow[id]["values"][res["@option_ids"]] = ""
                                        res_label = getValue(question, "@id", res["@option_ids"])
                                        if type(res_label) is dict and "text" in res_label.keys():
                                            dataRow[id]["values"][res["@option_ids"]] = str(res_label["text"] or '').replace("\r", "").replace("\n", " ")
                                        elif "@type" in question.keys() and (question["@type"] == "sliderFree" or question["@type"] == "text"):
                                            dataRow[id]["values"][res["@option_ids"]] = res["@option_ids"]
                        newFile = False
                        if not os.path.isfile(CSVOutputFile):
                            newFile = True
                        myEncoding = "utf-8"
                        #if os.name == "nt":
                        #    myEncoding = "latin-1"
                        with open(CSVOutputFile, "a", encoding=myEncoding) as file_object:
                            if newFile:
                                for item in dataRow:
                                    file_object.write(item + ";")
                                    if type(dataRow[item]) is dict and "values" in dataRow[item].keys():
                                        if len(dataRow[item]["values"]) <= 1:
                                            file_object.write(item + "_text;")
                                        else:
                                            for option in dataRow[item]["values"]:
                                                file_object.write(option + ";")
                                file_object.write("\n")
                                for item in dataRow:
                                    if type(dataRow[item]) is dict and "values" in dataRow[item].keys():
                                        if len(dataRow[item]["values"]) <= 1:
                                            file_object.write(";" + dataRow[item]["text"] + ";")
                                        else:
                                            file_object.write(dataRow[item]["text"] + ";")
                                            for option in dataRow[item]["values"]:
                                                file_object.write(dataRow[item]["values"][option]["text"] + ";")
                                    else:
                                        file_object.write(";")
                                file_object.write("\n")
                                file_object.flush()

                            for item in dataRow:
                                if type(dataRow[item]) is dict and "values" in dataRow[item].keys():
                                    if len(dataRow[item]["values"]) == 0:
                                        file_object.write(";;")
                                    elif len(dataRow[item]["values"]) == 1:
                                        file_object.write(list(dataRow[item]["values"].keys())[0] + ";" + dataRow[item]["values"][list(dataRow[item]["values"].keys())[0]] + ";")
                                    else:
                                        file_object.write(str(dataRow[item]["value"]) + ";")
                                        for option in dataRow[item]["values"]:
                                            file_object.write(str(dataRow[item]["values"][option]["value"]) + ";")
                                else:
                                    file_object.write(dataRow[item] + ";")
                            file_object.write("\n")
                            file_object.flush()
                    else:
                        raise ValueError("Questionnaire File '" + QuestionnaireFile + "' not valid!" )
                else:
                    raise ValueError("Questionnaire File '" + QuestionnaireFile + "' does not match to given survey_uri '" + result["mobiquest"]["record"]["@survey_uri"] + "'!" )
            else:
                print("\tFile '" + QuestionnaireResult + "' not valid! Ignored!")
                
if __name__ == "__main__":
    #print ('Number of arguments:', len(sys.argv), 'arguments.')
    #print ('Argument List:', str(sys.argv))
    config = getConfig();
    if len(sys.argv) == 1:
        usingGui = True
        labels = list()
        buttons = list()
        entrys = list()
        checkbuttons = list();
        labelTexts = ['Questionnaire File', 'Questionnaire Result Folder', 'CSV Output File']
        window = tk.Tk()
        for idx in range(len(labelTexts)):
            frame = list()
            for cell in range(3):
                frame.append(tk.Frame(master = window))
                frame[cell].grid(row = idx, column = cell)
            labels.append(tk.Label(text = labelTexts[idx], master = frame[0]))
            entrys.append(tk.Entry(width = 80, master = frame[1]))
            entrys[-1].insert(-1, config.get('default', config.options('default')[idx]))
            entrys[-1].bind("<KeyRelease>", onInput)
            buttons.append(tk.Button(text = "...", master = frame[2]))
            buttons[-1].bind("<Button-1>", handle_click);
            labels[-1].pack(fill=tk.X)
            entrys[-1].pack(fill=tk.X)
            buttons[-1].pack(fill=tk.X)
        frame = list();
        frame.append(tk.Frame(master = window))
        frame[0].grid(row = idx + 2, columnspan = 2, sticky='nesw')
        buttons.append(tk.Button(text = "Execute", master = frame[0]))
        buttons[-1].bind("<Button-1>", handle_click);
        buttons[-1].pack(expand = True)
        window.mainloop()
        with open(configFileName, 'w') as configfile:
            config.write(configfile)    
    elif len(sys.argv) == 3:
        executeQuest2CSV(sys.argv[1], sys.argv[2])
    elif len(sys.argv) == 4:
        executeQuest2CSV(sys.argv[1], sys.argv[2], sys.argv[3])
    else:
        raise ValueError("Wrong usage of parameters")

