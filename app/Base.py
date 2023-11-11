import time

def Print(values: object,sep: str  = " ",end: str = "\n"):
    now = time.strftime("%Y-%m-%d %I:%M:%S %p>",time.localtime())
    print(now + str(values),end=end,sep=sep)
    with open("log\\print.log","w") as logFile:
        print(now + str(values),end=end,sep=sep,file=logFile,flush=True)

def Input():
    pass