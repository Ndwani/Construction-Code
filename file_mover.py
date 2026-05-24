import os ,shutil
destinationFolder= "C:\\python\\Files Sorter\\decoy"

os.chdir(destinationFolder)
for i in range (30):
    try:
        shutil.move(fr"C:\\python\\Files Sorter\\{i}.12.2025.txt",destinationFolder)
    except:
        continue


print('done')

                             
