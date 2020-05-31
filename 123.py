import os

def bianLi(rootDir):
    for root,dirs,files in os.walk(rootDir):
        for file in files:
            paths=os.path.join(root,file)
            print(os.path.splitext(file)[0])
            
        for dir in dirs:
            bianLi(dir)


rootDir=r'E:\Codes\auto-excel\test'
bianLi(rootDir)