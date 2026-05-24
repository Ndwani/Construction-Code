import re, file_validator,os, shutil
primePath =r"C:\Users\mwale.k\Pictures\Samsung Flow" # insert relevant path \ source folder basically
os.chdir(primePath)
destinationFolder =r"C:\Users\mwale.k\Pictures\Phase II Photos\2025\December"
sourceFileList = file_validator.fileIden(primePath) # returns a  list of all the files in source folder
folderNameList = []
month = "12" #inserts the month of interest (end date to be ajusted accordingly)
for i in range(31):
    folderName = f"{i+1}.{month}.2025" # returns a list of the name of the destination folders ( dated folders for each snap)
    folderNameList.append(folderName)
for j in folderNameList:  # allows us to extract the date from the each list e,g  1, 2, 3 .....12 , 13 14 etc
    if len(j) <= 9:
        date = j[:1]
    else:
        date = j[:2]
    currentFolderName = j
    
    if len(j) <= 9:
        
        Pattern = re.compile(fr"^(.*)2025{month}0{date}(.*).jpg") # generates pattern to match desired file date with current folder for dates with leading zeroes 01, 02, 03 etc
    else:
         Pattern = re.compile(fr"^(.*)2025{month}{date}(.*).jpg") # generates pattern to match desired file date with current folder for date with leading 1 e.g 15, 16, 17
        
    currentFileStore = list()
    for k in  sourceFileList:
        try:                                        # to guard against Nontype matches
            mo = Pattern.search (f"{k}")
            result = mo.group()
            currentFileStore.append(result)  # stores files that match with current foler name "J"
            #print(currentFileStore)
        except:
            continue
        for l in currentFileStore:
            shutil.copy(fr"{primePath}\{l}",fr"{destinationFolder}\{currentFolderName}")
   # for m in currentFileStore:
    #    sourceFileList.remove(m)
    del currentFileStore
    print("Loading...\n")
#print(sourceFileList)
print("done")
