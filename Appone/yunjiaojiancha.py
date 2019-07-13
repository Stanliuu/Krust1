import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.colors as colors
import math
import xlrd
import pylab
import pypinyin
import itchat
import time, sys
from itchat.content import TEXT
from xlrd import open_workbook
from matplotlib.font_manager import FontProperties
wb1 = open_workbook('D:\Krust\Appone\yunjiao.xlsx')
sta={'2押':0}
wb3 = open('D:\Krust\Appone\geci.txt')
wb = open_workbook('D:\Krust\Appone\yanse.xlsx')
color=[]
colornumber=20
def Yunjiao(a):
        global sta
        global color
        global colornumber
        global wb3
        global wb1
        global wb
        Tempt=pypinyin.pinyin(a, style=pypinyin.NORMAL)
        tempt=str(Tempt[0][0])
        ans=''
        if tempt=='a' or tempt=='o' or tempt=='e' or tempt=='ai' or tempt=='ei' or tempt=='ao' or tempt=='ou' or tempt=='ei' or tempt=='an' or tempt=='en' or tempt=='ang' or tempt=='eng' :
            ans=tempt[:]
        elif tempt=='yu' or tempt=='qu' or tempt=='ju' or tempt=='xu':
            ans='uu'
        elif tempt=='yue' or tempt=='que' or tempt=='jue' or tempt=='xue':
            ans='ue'
        elif tempt=='yun' or tempt=='qun' or tempt=='jun' or tempt=='xun':
            ans='uun'
        elif tempt=='ye':
            ans='ie'
        elif tempt[0]=='z'or tempt[0]=='c' or tempt[0]=='s':
            if len(tempt)>1 and tempt[1]=='h':
                ans=tempt[2:len(tempt)]
        if ans=='':
            ans=tempt[1:len(tempt)]
        for i in range(36):
            for ss in wb1.sheets():
                 if ans==ss.cell(i,0).value:
                       return int(ss.cell(i,1).value)-1
        return int(0)
def Multimortgage():
        global sta
        global color
        global colornumber
        global wb3
        global wb1
        global wb
        while True:
            line=wb3.readline()
            if line=='\n':
                continue
            if not line:
                break
            #if line[len(line)-1]=='\n':
                #line=line[0:len(line)-1]
            for i in range(len(line)):
                for s in wb.sheets():
                    if line[i]=='\n' :
                        color.append(-2)
                    elif line[i] > '\u9fa5' or '\u4e00' > line[i] or line[i]==' ':
                        color.append(-1)
                    else:
                        color.append(Yunjiao(line[i]))
        #print(color)
        color.reverse()
        for i in range(len(color)):
            if color[i]<0:
                continue
            maxx=-1
            no=0
            for j in range(i+1,len(color)):
                if color[j]==-2:
                    no=no+1
                    if no>1:
                        break
                    continue
                if color[j]==-1:
                    continue
                if color[i]==color[j]:
                    tempt=0
                    while True:
                        if i+tempt>=len(color) or j+tempt>=len(color) or color[i+tempt]!=color[j+tempt] or i+tempt>=j or color[i+tempt]==-2 or color[j+tempt]==-2:
                            break
                        tempt+=1
                    if tempt >=2:
                        #print(i,j,tempt)
                        no=no-1
                        for p in range(tempt):
                            if color[i]>18:
                                color[j+p]=color[i]
                            else:
                                color[j+p]=int(colornumber)
                    if maxx<tempt:
                        maxx=tempt
            if maxx>1:
                name=str(maxx)+'押'
                if name in sta.keys():
                    sta[name]+=1
                else:
                    sta[name]=1
                if color[i]<=18:
                    for u in range(maxx):
                         color[u+i]=int(colornumber)
                    colornumber+=1
        color.reverse()
def yunjiao():
    global sta
    global color
    global colornumber
    global wb3
    global wb1
    global wb
    fig = plt.figure()
    ax = fig.add_subplot(111)
    startposx=0.0
    startposy=0.87
    filenum=0
    font1=FontProperties(fname='D:\Krust\Appone\SimHei.ttf',size=15)
    words=0
    Multimortgage()
    #print(color)
    wb3.close()
    wb3 = open('D:\Krust\Appone\geci.txt')
    while True:
        line=wb3.readline()
        if not line:
            break
        if line=='\n':
            continue
        for i in range(len(line)):
            if line[i]=='\n':
                 words+=1
                 continue
            elif line[i] > '\u9fa5' or '\u4e00' > line[i]:
                pylab.text(startposx+0.008,startposy+0.009, line[i], fontsize=10 ,fontproperties=font1, fontweight='bold', color='black', horizontalalignment='center')
                startposx+=0.02
                if startposx>1.0:
                    startposx=0
                    startposy-=0.10
                    if startposy<0.05:
                        pylab.savefig("heart3-1"+str(filenum)+".png")
                        filenum+=1
                        fig = plt.figure()
                        ax = fig.add_subplot(111)
                        startposx=0.0
                        startposy=0.87
                words+=1
                continue
            for s in wb.sheets():
                c=s.cell(int(color[words]),0).value
                words+=1
            pos = (startposx, startposy)
            ax.add_patch(patches.Rectangle(pos, 0.05, 0.07, color=c))
            ax.annotate('', xy=pos)
            pylab.text(startposx+0.025,startposy+0.009, line[i], fontsize=10 ,fontproperties=font1, fontweight='bold', color='black', horizontalalignment='center')
            startposx+=0.05
            if startposx>1.0:
                startposx=0
                startposy-=0.10
                if startposy<0.05:
                    pylab.axis('off')
                    pylab.savefig("heart3-1"+str(filenum)+".png")
                    filenum+=1
                    fig = plt.figure()
                    ax = fig.add_subplot(111)
                    startposx=0.0
                    startposy=0.87
        startposx=0
        startposy-=0.09
        if startposy<0.05:
            pylab.axis('off')
            pylab.savefig("colored_lyrics"+str(filenum)+".png")
            filenum+=1
            fig = plt.figure()
            ax = fig.add_subplot(111)
            startposx=0.0
            startposy=0.87
    pylab.axis('off')
    pylab.savefig("colored_lyrics"+str(filenum)+".png")
    filenum+=1
    fig = plt.figure()
    ax = fig.add_subplot(111)
    startpos=0.7
    for keys in sta:
        pylab.text(0.5,startpos, keys+':'+str(sta[keys]), fontsize=10 ,fontproperties=font1, fontweight='bold', color='black', horizontalalignment='center')
        startpos-=0.2
    pylab.axis('off')
    pylab.savefig("colored_lyrics"+str(filenum)+".png")
    return int(filenum)
    #plt.show()
