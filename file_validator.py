

import os
#primePath =r"C:/python/Files Sorter"
#os.chdir(primePath)

"""
spam = os.listdir("C:/python/Files Sorter") #list returned
print(spam)
print(len(spam))
pam=[]
for i in spam:
    if os.path.isfile(fr"C:/python/Files Sorter/{i}"):
       
        pam.append(i)#validate file
    else:
        continue



        
print("***********/n")

print(pam)
print(len(pam))"""


def fileIden(path): # path entered by user
    import os
    
    os.chdir(path) # links to source folder

    Container_1 = list() # stores the filtered result (files only and no folders)
    Container_2 = os.listdir(path) #list returned

    for file in Container_2:
        if os.path.isfile(fr"{path}/{file}"): # if true file is appended to list
            Container_1.append(file) 
        else:
            continue # skips folders
    return Container_1 # returns filtered list

#print(fileIden(r"C:/python/Files Sorter"))

#rimePath = r"C:/python/Files Sorter"
#os.chdir(primePath)

#result = file_validator.fileIden(primePath)

#G = re.compile(r"(\d)?\d.12.2025.txt$")
#N = []
#for i in result:
    # p =G.findall(f"{i}")

 #   try:

  #      mo = G.search(f"{i}")
   #     p = mo.group()
    #    N.append(p)
    #except:
     #   continue

#print("done")

#print(N)
     
