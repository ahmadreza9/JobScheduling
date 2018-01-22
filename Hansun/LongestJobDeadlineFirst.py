import xlrd
import pyodbc
import sys
class Good:
    def __init__(self):
        self.goodcd = 0
        self.cycletime = 0
    def setGoodcd(self,goodcd):
        self.goodcd = goodcd
    def getGoodcd(self):
        return self.goodcd

class CNC:
    def __init__(self):
        #self.num_code = 0
        #self.shape_code = 0
        #self.type_code = 0
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
    def getNumCode(self):
        return self.num_code
    def getShapeCode(self):
        return self.shape_code
    def getTypeCode(self):
        return self.type_code
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
        print(cnc.getReservedTime())
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

workbook = xlrd.open_workbook("hansun2.xlsx")
worksheet = workbook.sheet_by_index(0)
rows = worksheet.nrows
cols = worksheet.ncols


cncs= []
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
            cncs.append(temp)
        else:
            pass
    except:
       #print(i)
        pass

        #temp.setSizeFrom()

cncs = sorted(cncs,key =lambda c:int(c.cnc_shape),reverse=True) # for P3,JAW2
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
		where c.workno = e.workno and e.Workdate between '20171201' and '20171210' and (processcd ='P1' or processcd = 'P2' or processcd = 'P3')
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
same_deliverydate_jobs = []
date = jobs[0].deliverydate
for job in jobs:
    if(job.deliverydate != date):
        deliverydate_sorted_jobs.append(same_deliverydate_jobs)
        same_deliverydate_jobs = []
        date = job.deliverydate
        same_deliverydate_jobs.append(job)
    else:
        same_deliverydate_jobs.append(job)
deliverydate_sorted_jobs.append(same_deliverydate_jobs)


sorted_jobs = [] # deliverydate -> required_time sorted
for job in deliverydate_sorted_jobs:
    sorted_jobs.append(sorted(job,key = lambda j:j.required_time,reverse=True))
#for a in sorted_jobs:
 #   for b in a:
  #      print(b.workno,end=" ")
   #     print(b.required_time)
    #print("----------------")




for jobs in sorted_jobs:
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
print("=====================================================================================")
cncs = sorted(cncs,key= lambda j:j.cnc_num)
cnt =0
time_max = 0
time_h = 0
time_m = 0
tims_s =0

for cnc in cncs:
    cnc.printJobList()
    cnt+=cnc.getNumOfJobs()
    if(cnc.getReservedTime() > time_max):
        time_max = cnc.getReservedTime()

time_h = int(time_max/3600)
temp = int(time_max%3600)
time_m = int(temp/60)
temp = int(temp%60)
time_s = int(temp)

print("========================================================")
print(time_max)
print("작업개수:",end="")
print(len(rows))
print("총 소요시간:",time_h,"시간 ",time_m,"분 ",time_s,"초")
print("작업간 간격:",interval_time,"초")

print(cnt)





