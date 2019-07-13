from django.shortcuts import render
from django.http import HttpResponse
from django.template import Template, Context
from .forms import AddForm
from .yunjiaojiancha import yunjiao
import zipfile
import os

def index(request):
    return HttpResponse(u"欢迎来到我的主页!")
def page(request):
    return render(request,'index.html')

def download(request):  
    file=open(r'C:\Users\lenovo\Desktop\1234.txt','rb')  
    response =HttpResponse(file)  
    #response['Content-Type']='text/plain'  
    response['Content-Disposition']='attachment;filename="1234.txt"'  
    return response
def tools1(request):
    if request.method == 'POST':# 当提交表单时
        form = AddForm(request.POST) # form 包含提交的数据
        if form.is_valid():# 如果提交的数据合法
            a = form.cleaned_data['a']
            f=open("D:\Krust\Appone\geci.txt",'w+')
            f.write(a)
            f.close()
            num=yunjiao()
            zfile=zipfile.ZipFile("rhyme.zip","w")
            for i in range(num):
                filename="colored_lyrics"+str(i)+".png"
                zfile.write(filename)
                if os.path.exists(filename): 
                    os.remove(filename)
            file=open(r'rhyme.zip','rb')
            response =HttpResponse(file)
            response['Content-Disposition']='attachment;filename="rhyme.zip"' 
            return response
            

# Create your views here.
