import shutil
import os

#source_dir=r"\\acdev01\3M_CAC\ERP_Quality_Review\QualityReview"
#destination_dir=r"\\acdev01\3M_CAC\ERP_Quality_Review\FileManipulation"
#foldername="QualityCheck_Week"
def fileprocess(source_dir,destination_dir,foldername,lastfilerangefile):
    file2 = open(lastfilerangefile, "r+")
    rng=file2.read()
    file2.close()
    delete=FileManipulation().deleteFolder(destination_dir)
    copy=FileManipulation().copyFile(source_dir,destination_dir,foldername+"_"+str(int(rng)+1))
    return "File prepared successfully"

class FileManipulation:

    def copyFile(self,source_dir,destination_dir,foldername ):
        path=os.path.join(destination_dir, foldername)
        os.mkdir(path)
        for file in os.listdir(source_dir):
            shutil.copy2(os.path.join(source_dir, file),path)
        return source_dir

    def deleteFolder(self,destination_dir):
        for dir in os.listdir(destination_dir):
            shutil.rmtree(os.path.join(destination_dir,dir))
#fileprocess()
