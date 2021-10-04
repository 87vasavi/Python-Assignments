import csv
import openpyxl
import os
class subject:
    
    def __init__(self,subno,subname,ltp,crd):
        self.subno=subno
        self.subname=subname
        self.ltp=ltp
        self.crd=crd

class sem:
    
    def __init__(self):
        self.spi=0.0
        self.courses=[]
        self.credits=0.0

class student:
    
    def __init__(self):
        self.roll=''
        self.name=''
        self.department=''
        self.cpi=[0 for i in range(8)]
        self.sem_data=list(None for i in range(8))
        self.total_credits=[0 for i in range(8)]
        
class course:
    
    def __init__(self,subj,grade,tp):
        self.subj=subj
        self.grade=grade
        self.tp=tp

file="names-roll.csv"
rn=dict()
with open(file, 'r') as csvfile:
    csvreader = (csv.reader(csvfile))
    next(csvreader)
    for x in (csvreader):
        k=student()
        k.roll=x[0]
        k.name=x[1]
        rn[x[0]]=k
        k.department=x[0][4:6]
    
file="subjects_master.csv"
sub=dict()
with open(file, 'r') as csvfile:
    csvreader = (csv.reader(csvfile))
    next(csvreader)
    for x in csvreader:
        k=subject(x[0],x[1],x[2],int(x[3]))
        sub[x[0]]=k
    
file="grades.csv"
with open(file, 'r') as csvfile:
    csvreader = (csv.reader(csvfile))
    next(csvreader)
    for x in csvreader:
        if not rn[x[0]].sem_data[int(x[1])-1] :
            rn[x[0]].sem_data[int(x[1])-1]=sem()
        sd=rn[x[0]].sem_data[int(x[1])-1]
        sd.credits=sd.credits+int(x[3])
        sd_courses=sd.courses
        k=course(sub[x[2]].subno,x[4],x[5])
        sd_courses.append(k)
        #print(sd_courses)

for x in rn:
    file=x+'.xlsx'
    if os.path.exists(file):
        os.remove(file)
    
def caltotcredits(s):
    sd=s.sem_data
    s.total_credits[0]=sd[0].credits
    for i in range(1,8,1):
        if sd[i]!=None:
            s.total_credits[i]=s.total_credits[i-1]+sd[i].credits
        else:
            s.total_credits[i]=s.total_credits[i-1]
            
def calspicpi(s):
    sd=s.sem_data
    data={'AA':10,'AB':9,'BB':8,'BC':7,'CC':6,'CD':5,'DD':4,'F':0,'F*':0,'I':0}
    for i in range(8):
        if sd[i]!=None:
            courses=sd[i].courses
            val=0
            for subs in courses:
                try:
                    sg=subs.grade
                    sg=sg.replace(" ", "")
                    val=data[sg[0:2]]*sub[subs.subj].crd +val
                except:
                    print(s.roll,i,subs.subj,subs.grade,len(subs.grade))
            sd[i].spi=val
            if i==0:
                s.cpi[0]=val
            else:
                s.cpi[i]=s.cpi[i-1]+sd[i].spi  
            #print(subs.subj,subs.grade,val,sd[i].credits)
                
#calspicpi(rn['0401CS02'])


for x in rn:
    caltotcredits(rn[x])
    calspicpi(rn[x])
    wb=openpyxl.Workbook()
    b=wb.active
    b.title="Overall"
    l=['Roll No','Name of Student','Discipline','Semester No','Semester wise Credit Taken','SPI','Total Credits Taken','CPI']
    for i in range(1,len(l)+1):
        v='A'+str(i)
        c=b[v]
        c.value=l[i-1]
    c=b['B1']
    c.value=x
    c=b['B2']
    c.value=rn[x].name
    c=b['B3']
    c.value=rn[x].department
    d=['B','C','D','E','F','G','H','I']
    sd=rn[x].sem_data
    for i in range(8):
        v=d[i]+'4'
        c=b[v]
        c.value=i+1
        v=d[i]+'5'
        c=b[v]
        if sd[i]!=None:
            c.value=sd[i].credits
        v=d[i]+'6'
        c=b[v]
        if sd[i]!=None:
            c.value=round((sd[i].spi)/sd[i].credits,2)
        v=d[i]+'7'
        c=b[v]
        c.value=rn[x].total_credits[i]
        v=d[i]+'8'
        c=b[v]
        c.value=round((rn[x].cpi[i])/rn[x].total_credits[i],2)
        
    for i in range(8):
        t='Sem'+str(i+1)
        wb.create_sheet(index = i+1 , title = t)
        bb=wb[t]  #wb.get_sheet_by_name(t)
        l=['SI No','Subject No.','Subject Name','L-T-P','Credit','Subject Type','Grade']
        d1=['A','B','C','D','E','F','G']
        for j in range(len(l)):
            v=d1[j]+'1'
            c=bb[v]
            c.value=l[j]
        sd=rn[x].sem_data[i]
        courses=[]
        if sd!=None:
            courses=sd.courses
        
        for k in range(len(courses)):
            sub_d=sub[courses[k].subj]
            CC={'A':k+1,'B':sub_d.subno,'C':sub_d.subname,'D':sub_d.ltp,'E':sub_d.crd,'F':courses[k].tp,'G':courses[k].grade}
            for j in range(len(d1)):
                v=d1[j]+str(k+2)
                c=bb[v]
                c.value=CC[d1[j]]
                
    wb.save(x+'.xlsx')