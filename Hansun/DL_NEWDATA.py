import xlrd
import pyodbc
import copy
import random
trial = 1000
max_time =0
interval_time = 3600 #다음작업까지 걸리는시간
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


def checkJob(cnc,job):
    if(job.getProcessCd!='P3'):
        #print(cnc.getSizeFrom(),job.getSpec(),cnc.getSizeTo())
        if(float(cnc.getSizeFrom()) <= float(job.getSpec()) <= float(cnc.getSizeTo())):
            return True
    else:
        if(cnc.getShape()=='2JAW'):
            if(float(cnc.getSizeFrom()) <= float(job.getSpec()) <= float(cnc.getSizeTo())):
                return True
    return False
def swappingJobs(asc, desc, deliverydate):
    global max_time

    difference = desc.getReservedTime() - asc.getReservedTime()
    temp_min = 0

    min = desc.getReservedTime()
    idx_a = -1
    idx_d = -1
   # print(asc.printJobList())
    #print(desc.printJobList())
    for a in reversed(list(range(len(asc.joblist)))):
        for d in reversed(list(range(len(desc.joblist)))):
            #print(asc.joblist[a].getWorkNo()," ", desc.joblist[d].getWorkNo())
            if(desc.joblist[d].getRequiredTime() > asc.joblist[a].getRequiredTime()):
                if((desc.joblist[d].getDeliveryDate() == deliverydate) and (asc.joblist[a].getDeliveryDate() == deliverydate)):
                    if( asc.getReservedTime() - asc.joblist[a].getRequiredTime() + desc.joblist[d].getRequiredTime() < min ):
                        if(checkJob(asc,desc.joblist[d]) == True and checkJob(desc,asc.joblist[a]) == True):
                            if( (desc.getReservedTime() - asc.getReservedTime()) > (desc.joblist[d].getRequiredTime() - asc.joblist[a].getRequiredTime()) *2 ):
                                min = desc.getReservedTime() - (desc.joblist[d].getRequiredTime() - asc.joblist[a].getRequiredTime())
                            else:
                                min = asc.getReservedTime() + (desc.joblist[d].getRequiredTime() - asc.joblist[a].getRequiredTime())
                            idx_a = a
                            idx_d = d

    if(idx_a >=0):

        max_time = desc.getReservedTime() - desc.joblist[idx_d].getRequiredTime() + asc.joblist[idx_a].getRequiredTime()
     #   asc.joblist[idx_a].printJob()
      #  desc.joblist[idx_d].printJob()
       # print("SWAP COMPLETEDSWAP COMPLETEDSWAP COMPLETEDSWAP COMPLETEDSWAP COMPLETEDSWAP COMPLETEDSWAP COMPLETED")
        minimum_time = asc.joblist[idx_a].getRequiredTime()+desc.joblist[idx_d].getRequiredTime()
        asc.joblist.append(desc.joblist[idx_d])
        asc.setReservedTime(asc.getReservedTime()-asc.joblist[idx_a].getRequiredTime()+desc.joblist[idx_d].getRequiredTime())
        desc.joblist.append(asc.joblist[idx_a])
        desc.setReservedTime(desc.getReservedTime()-desc.joblist[idx_d].getRequiredTime()+asc.joblist[idx_a].getRequiredTime())
        asc.joblist.pop(idx_a)
        desc.joblist.pop(idx_d)
        #asc.printJobList()
        #desc.printJobList()
        return True
    else:
        #print("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
        return False
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

numOfCNC = numOf2JAW + numOf3JAW
###################################################################################
jobs = []
worksheet = workbook.sheet_by_index(1)
rows = worksheet.nrows
cols = worksheet.ncols

for i in range(rows):
    temp = Job()
    temp.setWorkNo(worksheet.cell_value(i,0))
    temp.setProcessCd(worksheet.cell_value(i,1))
    temp.setCT(worksheet.cell_value(i,2))
    temp.setWorkDate(worksheet.cell_value(i,3))
    temp.setDeliveryDate(worksheet.cell_value(i,4))
    temp.setOrderQty(worksheet.cell_value(i,5))
    temp.setGoodCd(worksheet.cell_value(i,6))
    try:
        temp.setSpec(float(worksheet.cell_value(i,7)))
    except:
        continue
    jobs.append(temp)
###############################################################################################

"""#################################################################### DB DATA #########################################################
conn = pyodbc.connect(driver ='{SQL Server}', host = '221.161.62.124,2433', database = '',user = 'Han_Eng_Back',pwd = 'HseAdmin1991')
cursor = conn.cursor()
cursor.execute(
    
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

"""##################################################################
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
cncs_asc = []
cncs_desc = []
date = ""

brk_zero = False
current_time = 99999999
for i in range(trial):
    cncs = copy.deepcopy(empty_cncs)
    for jobs in sorted_jobs:
        random.shuffle(jobs)
        date_idx = 0
        date = jobs[0].getDeliveryDate()
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
                            break
                        else:
                            pass

        ############################################################
        cncs_asc = sorted(cncs, key = lambda c:c.reservedTime)
        cncs_desc = sorted(cncs, key = lambda c:c.reservedTime,reverse=True)
        #print("============================================================================")
        #print(len(cncs_desc))
        max_time = 0
        #print("@@@@@@@@@@@@@@@@@@@first@@@@@@@@@@@@@@@@")
        for c in cncs:
            #c.printJobList()
            if(max_time < c.getReservedTime()):
                max_time = c.getReservedTime()
        #print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        asc_idx = 0
        desc_idx =0
        brk = False
        while(asc_idx < len(cncs_asc)-1):
            #print("SWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING STARTSWAPPING START")
            brk = swappingJobs(cncs_asc[asc_idx],cncs_desc[desc_idx],date)
            if(brk == True):
                cncs_asc = sorted(cncs, key = lambda c:c.reservedTime)
                cncs_desc = sorted(cncs, key = lambda c:c.reservedTime,reverse=True)
                asc_idx = 0
                desc_idx =0
                #print("@@@@@@@@@@@@CHANGE@@@@@@@@@@@@@")
                #for asc in cncs_asc:
                    #asc.printJobList()
                #print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
            else:
                asc_idx+=1
        #print("******************************************************************************************************")
        brk_zero = False
        cncs = copy.deepcopy(cncs_asc)
        max_temp = 0
        for cnc2 in cncs:
            if(cnc2.getReservedTime() > max_temp):
                max_temp = cnc2.getReservedTime()
        deadline_max_time_list.append(max_temp)
        date_idx +=1
    d_time_max = 0
    for cnc in cncs:
        if(cnc.getReservedTime() > d_time_max):
         d_time_max = cnc.getReservedTime()
    if(d_time_max < current_time):
        min_cncs = copy.deepcopy(cncs)
        current_time = d_time_max
    print(i+1)

cnt =0
for cnc in min_cncs:
    cnc.printJobList()
    #print(cnc.getNumOfJobs())
    cnt+=cnc.getNumOfJobs()
#cncs = sorted(cncs,key= lambda j:j.cnc_num)
time_h = int(current_time/3600)
temp = int(current_time%3600)
time_m = int(temp/60)
temp = int(temp%60)
time_s = int(temp)

print("========================================================")
print(current_time)
print("작업개수:",end="")
print(cnt)
print("총 소요시간:",time_h,"시간 ",time_m,"분 ",time_s,"초")
print("작업간 간격:",interval_time,"초")



