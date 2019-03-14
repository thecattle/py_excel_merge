# encoding: utf-8
import os
import time

if __name__ == '__main__':
    path=os.path.abspath('.')
    dirName = path+"/aaa"
    print(dirName)
    pathNew=os.path.join(path, 'a')
    print(pathNew)

    p1 = "明细"
    p2 = "明细表"
    if p1 in p2:
        print(p2)
    words = [
        "A",
        "B",
        "C",
    ]
    for i in words:
        time.sleep(1)
        print(i)


    raw_input('Please press enter key to exit ...')


