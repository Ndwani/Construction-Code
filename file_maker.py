import file_validator




import os
primePath =r"C:/python/Files Sorter"
os.chdir(primePath)

print(file_validator.fileIden(primePath))



"""import os

os.chdir(r"C:/python/Files Sorter")


for i in range(31):

    fileMaker = open(fr"C:/python/Files Sorter/{i}.12.2025.txt","w")


fileMaker.close()
print("done")
"""
