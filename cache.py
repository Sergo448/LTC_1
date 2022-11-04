import os

path = ".\\TSN-2001"
directory = os.listdir(path=path)

w = 0
for file in directory:
    if file.endswith((".xlsx", 'xls')):
        w = w + 1
    else:
        continue
print(w)

