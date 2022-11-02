def textFileData(nppath):
    npfile = open(nppath,"r+")
    dtas=npfile.readlines()[0]
    return dtas
