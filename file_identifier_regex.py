import re
f = "15"
G = re.compile(fr"(.*)202511{f}(.*).jpg")
mo = G.search("20251115_082813.jpg")
p = mo.group()

print(p)


"""
1. #G = re.compile(r"^IMG (/d)?/d.jpg$") iteration 1
2. #(r"^IMG_202510(/d)?/d.jpg$") iteration 2
3. #(r"(.*)202510(\d)?\d(.*)") iteration 3 successful
4. #(r"(^IMG(.*)202510(\d)?\d(.*).jpg)") iteration 4 more refined and particular
5. #r"(^IMG(.*)202510(\d)?\d_(.*).jpg)") iteration 5 lower case dash
6. fr"^IMG(.*)202511{No}(.*).jpg)") this iteration assists in matching folder date with file name.

"""


"""

Notes

1. (.*) useful in assisting in matching "other" objects
2. (\d)?\d_ pattern acknowledges  dates like 1st = 1 and 15th = 15 on dates
3. Follow file name pattern to make regex making easier
4. An iteration and append approach will be utilized to
    a. capture file names of desired date name on file
    b. append once more to os.path to facitate copy to desired folder.


next issue how to identify folder with regex

import os
primePath =r"C:/python/Files Sorter"
os.chdir(primePath)
""""""
print(file_validator.fileIden(primePath))
G = re.compile(r"(\d)?\d.12.2025.txt")
mo = G.search("17.12.2025.txt")
p = mo.group()

"""
