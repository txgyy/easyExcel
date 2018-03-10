import os
path = r"X:\Users\yukino\Desktop\B"
xlnames = [os.path.join(path,filename)  for filename in os.listdir(path) if filename.split('.')[-1]=='xls' or filename.split('.')[-1]=='xlsx']
print(xlnames)