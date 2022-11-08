def writetonote(ttpath,notedata):
    #ttpath=r"C:\Users\2040664\anuraj\EDI\es.txt"
    file1 = open(ttpath, "w")
    file1.write(str(int(notedata)))
    file1.close()
    return ttpath
