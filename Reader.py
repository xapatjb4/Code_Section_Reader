def word_find(line,words):
    #Count the number of instance in each word
    linetowork = line.strip().split()
    #print(linetowork, "original2")
    toremove= set()
    for x,i in enumerate(linetowork, start=0 ):
        if linetowork[x] not in words:
            toremove.add(linetowork[x])
    newTrial = [x for x in linetowork if x not in toremove]
    #print(newTrial)
    return list(newTrial)

def main(file,words):
    with open(file) as f:
        for i,x in enumerate(f, start=1):
            #print(x)
            common = word_find(x,words)
            if common:
                print(i, "".join(common))

#count specific words
def word_count(file, words):
    counts = dict()
    with open(file) as f:
        for i,x in enumerate(f, start=0):
            common = word_find(x,words)
            if common:
                for word in common:
                    if word in counts:
                        counts[word] += 1
                    else:
                        counts[word] = 1
    return counts

if __name__ == '__main__':
    file = 'requirements.txt'
    words = ['<tagname1>', '<SR>']
    main(file, words)
    print( word_count(file, words))
