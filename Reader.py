import xlsxwriter as xl



def word_find(line,words):
    #Count the number of instance in each word
    linetowork = line.strip().split()

    #Making a set of words to be removed from the list to be counted
    toremove= set()
    for i,x in enumerate(linetowork, start=0 ):
        if x not in words:
            toremove.add(x)
    newTrial = [x for x in linetowork if x not in toremove]
    #print(newTrial)
    return list(newTrial)

def word_line(file,words):
    with open(file) as f:
        for i,x in enumerate(f, start=1):
            #print(x)
            common = word_find(x,words)
            if common:
                print(i, "".join(common))

#count specific words
def word_count(file, words):
    counts = dict.fromkeys(words,0)
    with open(file) as f:
        for i,x in enumerate(f, start=0):
            common = word_find(x,words)
            if common:
                for word in common:
                    counts[word] += 1
    return counts

#Take in a list of variable names and return the average length
def list_avg(variables):
    print(variables)
    if variables:
        return sum(len(word) for word in variables) / len(variables)
    return 0

#makes a list of whatever is tagged in va to return avg length
    #find va> then add var name to list
def var_avg(file):
    var_names = []
    with open(file) as f:
        for i,x in enumerate(f, start=0):
            common = set(word_find(x,'va>'))
            if common:
                line = x.strip().split()
                #for every index containing va> add the following index to the list
                for x in range(0,len(line)):
                    if line[x] == 'va>':
                        var_names.append(line[x+1])
                        x+=1
    return list_avg(var_names)

#Make a function to format data into excel sheet
def xl_format(wordcount, row):
    #Format should be as follows
    #worksheet1
    #col 0 subject 1 su 2 for 3 while 4 if 5 else 6 case 7 break
    #worksheet2
    #col 0 subject 1 #var 2 #int 3 #float 4 #string 5 #array 6 avg varname 7 comment
    return 0


if __name__ == '__main__':
    file = 'requirements.txt'
    words = ['su>', 'fo>','wh>','if>','el>','ca>', 'br>',
    'va>','in>','fl>','ch>','ar>','co']
    print('average is ', var_avg(file))
    #words2 = ['va>','in>','fl>','ch>','ar>','co']
    word_line(file, words)
    print( word_count(file, words))
