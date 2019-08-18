import xlsxwriter as xl
import openpyxl as oxl
import os.path
from os import path
import glob
import sys

#List of strings to identify the code composition
def dict_Data(arg):
    switcher = {
        #Program Control Data
        0: 'su>', #SubRoutine
        1: 'fo>', #For loops
        2: 'wh>', #While loop
        3: 'if>', #if
        4: 'el>', #else
        5: 'ca>', #case
        6: 'br>', #break
        #Variable Data
        7: 'va>', #Variable
        8: 'in>', #Integer
        9: 'fl>', #Float
        10: 'ch>', # String
        11: 'ar>', # Array (Not string)
        12: 'co>', # Comment
    }
    return switcher.get(arg, "nothing")

#Return instances of each word defined in word, in a line
def word_find(line,words):
    linetocheck = line.strip().split()
    #Making a set of words to be removed from the list to be counted
    toremove= set()
    for i,x in enumerate(linetocheck, start=0):
        if x not in words:
            toremove.add(x)
    #Removing the unmatch words from the list
    instances = [x for x in linetocheck if x not in toremove]
    return list(instances)

#Returns the lines which instance of words can be found
def word_line(file,words):
    with open(file) as f:
        for i,x in enumerate(f, start=1):
            common = word_find(x,words)
            if common:
                print(i, "".join(common))

#Returns a dictionnary containing the number of instances of words found
def word_count(file, words):
    #make a dict with words we're looking for
    var_tags = ['in>', 'fl>', 'ar>', 'ch>']
    counts = dict.fromkeys(words,0)
    with open(file) as f:
        for i,x in enumerate(f, start=0):
            #Check if line contains words being searched
            common = word_find(x,words)
            if common:
                #Increment number of instances for word found
                for word in common:
                    #possibly add check to increment var if in fl or ar
                    counts[word] += 1
                    if word in var_tags:
                        counts['va>'] += 1

    return counts

#Take in a list of variable names and return the average length
def list_avg(variables):
    #To see the list of vars uncomment this
    print(variables)
    if variables:
        return sum(len(word) for word in variables) / len(variables)
    return 0

#makes a list of whatever is tagged in va to return avg length
    #find va> then add var name to list
def var_avg(file):
    var_tags = ['in>', 'fl>', 'ar>', 'ch>', 'va>']
    var_names = []
    with open(file) as f:
        for i,x in enumerate(f, start=0):
            common = set(word_find(x,var_tags))
            if common:
                line = x.strip().split()
                #for every index containing va> add the following index to the list
                for x in range(0,len(line)):
                    if line[x] in var_tags:
                        var_names.append(line[x+1])
                        x+=1
    return list_avg(var_names)

#Make a function to format data into excel sheet
def xl_format(outputName):
    #Check if file exists
    if path.exists(outputName):
        return -1
    #Format should be as follows
    #worksheet1
    #col 0 subject, 1 su, 2 for, 3 while 4 if 5 else 6 case 7 break
    #worksheet2
    #col 0 subject 1 #var 2 #int 3 #float 4 #string 5 #array 6 avg varname 7 comment

    workbook = xl.Workbook(outputName)
    PCworksheet = workbook.add_worksheet()
    PCworksheet.write('A1', 'Subject')
    PCworksheet.write('B1', 'Subroutine')
    PCworksheet.write('C1', 'For')
    PCworksheet.write('D1', 'While')
    PCworksheet.write('E1', 'If')
    PCworksheet.write('F1', 'Else')
    PCworksheet.write('G1', 'Case')
    PCworksheet.write('H1', 'GoTo')

    Vworksheet= workbook.add_worksheet()
    Vworksheet.write('A1', 'Subject')
    Vworksheet.write('B1', '# Vars')
    Vworksheet.write('C1', '# Ints')
    Vworksheet.write('D1', '# Floats')
    Vworksheet.write('E1', '# Chars')
    Vworksheet.write('F1', '# Arrays')
    Vworksheet.write('G1', 'Ave Len')
    Vworksheet.write('H1', 'Comments')
    workbook.close()

    wb = oxl.load_workbook(outputName)
    ws = wb['Sheet1']
    ws.title = 'Program Control'
    ws = wb['Sheet2']
    ws.title = 'Variables'
    wb.save(outputName)
    return 0

#Fill the excel sheet with occurences found
def xl_fill(outfile, subject, occurences, var_avg):
    #take care of program control variable
    #Modify the xcel sheet to contain the respective instance
    wb = oxl.load_workbook(outfile)
    ws = wb['Program Control']
    ws.cell(subject+1, 1).value = subject

    x = 0
    for x in range(0,7):
        #Modify row {sujbect}  with war data 0 through 7
        value = occurences.get(dict_Data(x))
        ws.cell(subject+1, x+2).value = value

    ws = wb['Variables']
    ws.cell(subject+1, 1).value = subject
    for x in range(7, 12):
        value = occurences.get(dict_Data(x))
        ws.cell(subject+1, x + 2 - 7).value = value
    ws.cell(subject+1, x + 3 -7).value = var_avg
    wb.save(outfile)
    return 0;


if __name__ == '__main__':
    outputFile = sys.argv[2]
    xl_format(outputFile)
    inputDir = sys.argv[1] + "/*.txt"
    print(inputDir, outputFile)

    words = []
    for x in range(13):
        words.append(dict_Data(x))
    files = glob.glob(inputDir)
    print(files)
    #Split the string to remove
    for x in range(13):
        words.append(dict_Data(x))

    for file in files:
        split_file = file.split('/')# take the / from path
        subject = int(split_file[1][:-4])# Taking out the .txt
        print(split_file)
        print(subject)
        xl_fill(outputFile, subject, word_count(file, words),var_avg(file))
    #word_line(file, words)
    #print(word_count(file, words))
