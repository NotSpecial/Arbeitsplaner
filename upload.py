import os
import pysftp


#from upload import upload
def upload_file(filename):
    sftp = pysftp.Connection('hg.n.ethz.ch',
            username='adietmue', password='Hkjg868Hnd8+ethz')
    sftp.cd('downloads/calendar')
    sftp.put(filename)

#Variables
cal_dir = os.getcwd() + "\\cal\\"

#Scan all calendar files
for f in os.listdir(cal_dir):
    print("uploading file %s..." % str(f))
    upload_file(os.path.join(cal_dir, f))
    print("Done")
