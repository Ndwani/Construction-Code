import re

G = re.compile(r"(\d)?\d.11.2025")
mo = G.search("5.11.2025")
p = mo.group()

print(p)
