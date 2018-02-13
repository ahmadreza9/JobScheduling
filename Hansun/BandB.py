import xlrd
import pyodbc
import copy
import  numpy as np
from math import exp
import sys

trial = 5000

class CNC:
    def __init__(self):
        self.size = []
        self.cnc_num = 0
        self.cnc_shape = 0

        self.reservedTime = 0
        self.numOfJobs = 0
        self.joblist = []

    def __repr__(self):
        return repr(self.size,self.cnc_num,self.cnc_shape,self.reservedTime,self.numOfJobs,self.joblist)
    def setSize(self,size):
        self.size = size
    def setCNCnumber(self,num):
        self.cnc_num = num
    def setShape(self,shape):
        self.cnc_shape = shape
    def setReservedTime(self,time):
        self.reservedTime = time
    def increaseNumOfJobs(self):
        self.numOfJobs +=1

    def getCNCnumber(self):
        return self.cnc_num
    def getSizeFrom(self):
        return self.size[0]
    def getSizeTo(self):
        if (self.size[1] == 'MAX'):
            return 1000
        else:
            return self.size[1]
    def getReservedTime(self):
        return self.reservedTime
    def getShape(self):
        return self.cnc_shape
    def getNumOfJobs(self):
        return self.numOfJobs

    def jobAssign(self,job):
        self.joblist.append(job)
    def printJobList(self):
        print("-------------------------------")
        print("CNC ",end="")
        print(self.getCNCnumber(),end=": ")
        print(self.getNumOfJobs(),end=" ")
        print("Jobs")
        for job in self.joblist:
            print("Workno:", end="")
            print(job.getWorkNo())
            print("WorkDate:", end="")
            print(job.getWorkDate(), end=" WorkEnd:")
            print(job.getDeliveryDate())
            print(job.getProcessCd(), end=" ")
            print(job.getSpec(), end=" ")
            print(job.getRequiredTime())
        print(self.getReservedTime())
        print("-------------------------------")

    def printCode(self):
        print("CNC번호:",end="")
        print(self.cnc_num, end = " ")
        print(self.cnc_shape,end = "JAW ")
        print(self.getSizeFrom(),end = " ~ ")
        print(self.getSizeTo())



class Job:
    def __init__(self):
        self.workno=0
        self.workdate=0
        self.deliverydate=0
        self.goodcd=0
        self.orderqty =0
        self.gubun = 0
        self.spec=0
        self.processcd=0
        self.CT = 0
        self.required_time=0
        self.isFinish = False

    def __repr__(self):
        return repr((self.workno,self.workdate,self.deliverydate,self.goodcd,self.orderqty,self.gubun,self.spec,self.processcd,self.CT,self.required_time,self.isFinish))

    def __iter__(self):
        return self


    def setWorkNo(self,workno):
        self.workno = workno
    def getWorkNo(self):
        return self.workno

    def setWorkDate(self,workdate):
        self.workdate = workdate
    def getWorkDate(self):
        return self.workdate

    def setDeliveryDate(self,deliverydate):
        self.deliverydate = deliverydate
    def getDeliveryDate(self):
        return self.deliverydate

    def setGoodCd(self,goodcd):
        self.goodcd = goodcd
    def getGoodCd(self):
        return self.goodcd

    def setOrderQty(self,orderqty):
        self.orderqty = orderqty
        if (self.CT != 0):
            self.required_time = int(float(self.orderqty) * float(self.CT))
    def getOrderQty(self):
        return self.orderqty

    def setGubun(self,gubun):
        self.gubun = gubun
    def getGubun(self):
        return self.gubun

    def setSpec(self,spec):
        self.spec = spec
    def getSpec(self):
        return self.spec

    def setProcessCd(self,processcd):
        self.processcd = processcd
    def getProcessCd(self):
        return self.processcd

    def setCT(self,ct):
        self.CT = ct
        if (self.orderqty != 0):
            self.required_time = int(float(self.orderqty) * float(self.CT))
    def getCT(self):
        return self.ct

    def getRequiredTime(self):
        return self.required_time

    def jobFinish(self):
        self.isFinish = True
    def jobCheck(self):
        return self.isFinish

    def printJob(self):
        print("Workno:",end="")
        print(self.getWorkNo())
        print("WorkDate:",end="")
        print(self.getWorkDate(),end=" WorkEnd:")
        print(self.getDeliveryDate())
        print(self.getProcessCd(),end=" ")
        print(self.getSpec(),end=" ")
        print(self.getRequiredTime())

def makeSumTo1(a):
    sum =0
    for i in range(len(a)):
        sum += a[i]
    a[(np.random.choice(len(a),1))[0]] += (1-sum)
    return a

def makeProbability(a):
    temp = copy.deepcopy(a)
    sum_ = 0
    mean = float(sum(a,0.0)+1)/float(len(a))
    for i in range(len(temp)):
        temp[i] = float(1 / 1 + exp( (-6.0*temp[i]) / mean) )
        sum_ += temp[i]
    li = []
    for i in range(len(temp)):
        li.append(float(temp[i]/float(sum_)))
    return li
"""
def makeProbability(a):
    temp = copy.deepcopy(a)
    sum = 0
    for i in range(len(temp)):
        temp[i] = float(1/temp[i])
        sum += temp[i]
    li = []
    for i in range(len(temp)):
        li.append(float(temp[i]/float(sum)))
    return li
"""
workbook = xlrd.open_workbook("hansun2.xlsx")
worksheet = workbook.sheet_by_index(0)
rows = worksheet.nrows
cols = worksheet.ncols
numOf2JAW = 0
numOf3JAW = 0

cncs= []
empty_cncs = []
for i in range(rows):
    try:
        E = worksheet.cell_value(i,4)[0]
        if(E=='H'):

            size = worksheet.cell_value(i,4)
            size = size.replace('H','')
            temp =CNC()
            temp.setSize(size.split(' ~ '))
            temp.setCNCnumber(int(worksheet.cell_value(i,1)))
            temp.setShape(worksheet.cell_value(i,2)[0])
            if(int(worksheet.cell_value(i,2)[0])==2):
                numOf2JAW+=1
            else:
                numOf3JAW+=1
            cncs.append(temp)
        else:
            pass
    except:
        #print(i)
        pass

        #cncs = sorted(cncs,key =lambda c:int(c.cnc_shape),reverse=True) # for P3,JAW2
        #for i in range(40):
        #cnc[i].printCode()

interval_time = 3600 #다음작업까지 걸리는시간

conn = pyodbc.connect(driver ='{SQL Server}', host = '221.161.62.124,2433', database = '',user = 'Han_Eng_Back',pwd = 'HseAdmin1991')
cursor = conn.cursor()
cursor.execute(
    """
   select j.workno, j.processcd, j.Cycletime, en.Workdate, en.deliverydate, en. orderqty,en.goodcd, REPLACE(REPLACE(REPLACE(REPLACE(g.Spec,'HEX.',''),'HEX',''),'-IP',''),'-DIN','') as Spec, Cast(j.Cycletime as float)*Cast(en.orderqty as float) as required_time
from TWorkreport_Han_Eng en 
	inner join TGood g on en.Raw_Materialcd = g.GoodCd
	and en.PmsYn = 'N'
	and en.ContractYn = '1'
	and g.Class2 not in ('060002', '060006')
	and g.Class3 in ('061038', '061039')
	inner join(
	select c.workno, c.processcd, AVG(c.Cycletime) as Cycletime
		from TWorkreport_Han_Eng e, TWorkReport_CNC c
		where c.workno = e.workno and e.Workdate between '20171201' and '20171230' and (processcd ='P1' or processcd = 'P2' or processcd = 'P3')
		group by c.workno, c.processcd
	) j on en.workno = j.workno
	order by workno
    """
)
jobs = []
rows = cursor.fetchall()
for row in rows:
    temp = Job()
    temp.setWorkNo(row.workno)
    temp.setProcessCd(row.processcd)
    temp.setCT(row.Cycletime)
    temp.setWorkDate(row.Workdate)
    temp.setDeliveryDate(row.deliverydate)
    temp.setOrderQty(row.orderqty)
    temp.setGoodCd(row.goodcd)
    try:
        temp.setSpec(float(row.Spec))
    except:
        continue
    jobs.append(temp)

jobs = sorted(jobs,key = lambda j:int(j.deliverydate))
deliverydate_sorted_jobs = []
deliverydate_required_time = []
req_time = 0
total_required_time = 0
same_deliverydate_jobs = []
date = jobs[0].deliverydate
for job in jobs:
    if(job.deliverydate != date):
        deliverydate_sorted_jobs.append(same_deliverydate_jobs)
        deliverydate_required_time.append(req_time)
        req_time = 0
        same_deliverydate_jobs = []
        date = job.deliverydate
        same_deliverydate_jobs.append(job)
        req_time += job.getRequiredTime()
    else:
        same_deliverydate_jobs.append(job)
        req_time += job.getRequiredTime()
deliverydate_sorted_jobs.append(same_deliverydate_jobs)
deliverydate_required_time.append(req_time)
for t in range(len(deliverydate_required_time)):
    total_required_time += deliverydate_required_time[t]

sorted_jobs = [] # deliverydate -> required_time sorted
for job in deliverydate_sorted_jobs:
    sorted_jobs.append(sorted(job,key = lambda j:j.required_time,reverse=True))

empty_cncs = copy.deepcopy(cncs)
deadline_max_time_list = []
for jobs in sorted_jobs:
    date_idx = 0
    for job in jobs:
        # job.printJob()
        for cnc in cncs:
            if(float(cnc.getSizeFrom())<= float(job.getSpec()) <= float(cnc.getSizeTo())):
                if(job.getProcessCd() != 'P3' and job.getRequiredTime() <= 30000 and float(job.getSpec())<50):
                    if(cnc.getNumOfJobs()!= 0):
                        cnc.setReservedTime(cnc.getReservedTime()+job.getRequiredTime()+ interval_time)
                    else:
                        cnc.setReservedTime(cnc.getReservedTime() + job.getRequiredTime())
                    cnc.jobAssign(job)
                    cnc.increaseNumOfJobs()
                    #print(job.getWorkNo(), end=" ")
                    #print(job.getSpec())
                    #print(cnc.getCNCnumber())
                    cncs = sorted(cncs,key=lambda j:int(j.reservedTime))

                    break
                else:
                    if(cnc.getShape()=='2JAW' and job.getRequiredTime() <= 30000 and float(job.getSpec())<50):
                        if (cnc.getNumOfJobs() != 0):
                            cnc.setReservedTime(cnc.getReservedTime() + job.getRequiredTime() + interval_time)
                        else:
                            cnc.setReservedTime(cnc.getReservedTime() + job.getRequiredTime())
                        cnc.jobAssign(job)
                        cnc.increaseNumOfJobs()
                        cncs = sorted(cncs, key=lambda j: int(j.reservedTime))
                        #print(job.getWorkNo,end=" ")
                        #print(job.getSpec())
                        #print(cnc.getCNCnumber())
                        break
                    else:
                        pass
    max_temp = 0
    for cnc2 in cncs:
        if(cnc2.getReservedTime() > max_temp):
            max_temp = cnc2.getReservedTime()
    deadline_max_time_list.append(max_temp)
    date_idx +=1
print("=====================================================================================")
#cncs = sorted(cncs,key= lambda j:j.cnc_num)
cnt =0
d_time_max = 0 # 가장 오래걸리는 cnc가 끝나기까지 걸리는 시간
for cnc in cncs:
    #cnc.printJobList()
    cnt+=cnc.getNumOfJobs()
    if(cnc.getReservedTime() > d_time_max):
        d_time_max = cnc.getReservedTime()



temp_list = copy.deepcopy(empty_cncs)

picked = 0
current_req_time = 0
final_required_time = d_time_max

brk= False
cut = False
print(d_time_max)
print(cnt)
post_cncs = copy.deepcopy(temp_list)
max_time_list = []
date_idx = 0
for jobs_ in deliverydate_sorted_jobs:
    max_time_list.append(999999999)
    pre_cncs = copy.deepcopy(post_cncs)
    for _ in range(trial):
        print(date_idx+1,"/",len(deadline_max_time_list),",",float(_/trial*100))
        temp_cncs = copy.deepcopy(pre_cncs)
        for job in jobs_:
            if(job.getRequiredTime()>30000 or float(job.getSpec()>50)):
                continue
            possible_cncs = []
            possible_cncs_prob = []
            cnc_idx = []
            idx = 0
            for cnc in temp_cncs:
                if (float(cnc.getSizeFrom()) <= float(job.getSpec()) <= float(cnc.getSizeTo())):
                    if(job.getProcessCd()=='P3'):
                        if(cnc.getShape()=='2JAW'):
                            possible_cncs.append(cnc)
                            cnc_idx.append(idx)
                    else:
                        possible_cncs.append(cnc)
                        cnc_idx.append(idx)
                else:
                    idx+=1
                    continue
                idx+=1
            reserved_time_list = []
            for cnc in possible_cncs:
                reserved_time_list.append(cnc.getReservedTime())
            possible_cncs_prob = makeProbability(reserved_time_list)
            picked = np.random.choice(len(possible_cncs_prob),1,p=copy.deepcopy(possible_cncs_prob))
            if(temp_cncs[cnc_idx[picked[0]]].getReservedTime() + job.getRequiredTime() + interval_time > max_time_list[date_idx]):
                cut=True
                break
            if(temp_cncs[cnc_idx[picked[0]]].getNumOfJobs()==0):
                temp_cncs[cnc_idx[picked[0]]].setReservedTime(temp_cncs[cnc_idx[picked[0]]].getReservedTime() + job.getRequiredTime())
            else:
                temp_cncs[cnc_idx[picked[0]]].setReservedTime(temp_cncs[cnc_idx[picked[0]]].getReservedTime() + job.getRequiredTime() + interval_time)
            temp_cncs[cnc_idx[picked[0]]].jobAssign(job)
            temp_cncs[cnc_idx[picked[0]]].increaseNumOfJobs()
        if(cut == False):
            temp_max = 0
            for cnc2 in temp_cncs:
                if(cnc2.getReservedTime() > temp_max):
                    temp_max = cnc2.getReservedTime()
            max_time_list[date_idx] = temp_max
            post_cncs = copy.deepcopy(temp_cncs)
        else:
            cut= False
    date_idx += 1
time_max =0

print(max_time_list)
print(deadline_max_time_list)
sampled_cncs = copy.deepcopy(post_cncs)
for cnc in sampled_cncs:
    if (cnc.getReservedTime() > time_max):
        time_max = cnc.getReservedTime()
final_required_time = time_max
sum_jobs = 0

###################################################################################
sum_sampled_jobs = 0
print("=============================deadline job list============================")
for cnc in cncs:
    #cnc.printJobList()
    sum_jobs += cnc.getNumOfJobs()
    cnc.printJobList()
    print(cnc.getNumOfJobs())
print("=================================sampled job list=================================")
for cnc in sampled_cncs:
    sum_sampled_jobs += cnc.getNumOfJobs()

print("=====================================================================================")

#cncs = sorted(cncs,key= lambda j:j.cnc_num)
time_h = int(d_time_max/3600)
temp = int(d_time_max%3600)
time_m = int(temp/60)
temp = int(temp%60)
time_s = int(temp)

print("========================================================")
print(d_time_max)
print("작업개수:",end="")
print(sum_jobs)
print("총 소요시간:",time_h,"시간 ",time_m,"분 ",time_s,"초")
print("작업간 간격:",interval_time,"초")

time_h = int(final_required_time/3600)
temp = int(final_required_time%3600)
time_m = int(temp/60)
temp = int(temp%60)
time_s = int(temp)

print("========================================================")
print(final_required_time)
print("작업개수:",end="")
print(sum_sampled_jobs)
print("총 소요시간:",time_h,"시간 ",time_m,"분 ",time_s,"초")
print("작업간 간격:",interval_time,"초")

print(sum_jobs)
print(date_idx)

