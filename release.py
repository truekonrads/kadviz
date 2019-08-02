import datetime,os
#from zipfile import ZipFile
import zipfile
import subprocess
from glob import glob
basedir=os.path.dirname(__file__)
subprocess.call([os.path.join(basedir,"setup.py"), "py2exe"],shell=True)
name=os.path.join(basedir,"releases","kadviz_" + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")+".zip")
print "*** writing to zipfile %s ***" % name
with zipfile.ZipFile(os.path.join(basedir,name),"w",zipfile.ZIP_DEFLATED) as myzip:
    assert isinstance(myzip,zipfile.ZipFile)    
    
    print "\n*** adding binaries ***"    
    files=glob(os.path.join(basedir,"dist","*"))    
    for f in files:
        arcname="kadviz\\"+os.path.basename(f)
        print "adding " + arcname
        myzip.write(f,arcname)

    print "\n*** adding source ***"
    files=glob(os.path.join(basedir,"*.py"))
    files.extend(glob(os.path.join(basedir,"*.wp*")))
    for f in files:
        arcname="source\\"+os.path.basename(f)
        print "adding " + arcname
        myzip.write(f,arcname)
        
    print "\n*** adding docs ***"
    docs=glob(os.path.join(basedir,"*.doc*"))
    for f in docs:
        print "adding " + f
        myzip.write(os.path.basename(f))
        
    
        

