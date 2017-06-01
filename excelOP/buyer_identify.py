# -*- coding:utf-8 -*-
'''
Created on 2017年5月31日

@author: wycheng
'''
import xlrd
import xlsxwriter
import time,datetime
import re

class BuyerManager:
    dict_mid_data={}# 维护的一个{map客户id:[第几次下单，上次日期时间]}
    
    # 获取对应行数据的订单时间
    def getDatetime(self,row_value):
        time_str=row_value[0].strip()[:-2]
        timeArray=time.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        Y,m,d,H,M,S=timeArray[0:6]
        
        dt_current=datetime.datetime(Y,m,d,H,M,S)# 转换成datetime对象，可以直接进行比较
        return dt_current
    
    # 将所有工作表的行按照订单日期升序排序
    def getList_sorted(self,list_xl):# list_xl: Excel文件的地址list
        
        list_row=[]# 将行数据存储到list中，便于排序
        
        for exl in list_xl:  
            print u'正在打开文件 '+exl
            wb=xlrd.open_workbook(exl)
            sheet=wb.sheets()[0]
            nrows=sheet.nrows# 工作表的行数
            print u'正在插入文件 '+exl+u'的row_value'
            for i in range(nrows):
                #用正则匹配过滤掉空行和标题行
                str_date=sheet.row_values(i)[0].strip()[:-2]
                if re.match('[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}', str_date)!=None:
                    list_row.append(sheet.row_values(i))
        
        print u'正在排序……'
        list_row=sorted(list_row,key=self.getDatetime)# 将list_row排序，排序是对key进行比较，key指定的函数会作用于list中的每一个元素
        return list_row

    def process(self,list_rowValues,N):# list_rowValues: 存放所有row_value的list  N: 间隔N天内是existing
        # 遍历每一行
        line=1
        for row_value in list_rowValues:
            print u'正在处理第'+str(line)+u'行'
            line+=1
            
            dt_current=self.getDatetime(row_value)# 订单日期时间的datetime类型
            mber_id=row_value[1].strip()# 客户id
            
             # 维护一个dict，用一个dict保存，客户id作为key，[当前第几次，上次订单日期时间]作为value
             # 并且依此写入新数据到EXcel
            if mber_id in self.dict_mid_data: # 如果存在这个key，说明该顾客之前有订单记录，更新dict，同时插入新数据到row_value
                
                self.dict_mid_data[mber_id][0]+=1# 更新下单次数+1
                row_value[3]=self.dict_mid_data[mber_id][0]# 插入下单次数
                
                dt_last=self.dict_mid_data[mber_id][1]
                row_value[4]=dt_last.strftime("%Y-%m-%d %H:%M:%S")# 插入上次订单日期时间
                
                dis=abs(dt_current-dt_last)# 时间差的绝对值
                row_value[5]=str(dis)# 插入与上次订单时间的间隔时间差
                
                # 插入usertype
                if dis <= datetime.timedelta(days=N):# 如果间隔在N天内
                    row_value[6]='existing'
                else:
                    row_value[6]='new' 
                     
                if dt_current>dt_last:# 如果当前时间更近,更新dict里的上次日期时间
                    self.dict_mid_data[mber_id][1]=dt_current
            else:# 不存在这个key，直接保存初始值
                self.dict_mid_data[mber_id]=[1,dt_current]
                row_value[3]=1 # 当前是第几次订单
                row_value[4]=u'首次下单' # 当前日期时间
                row_value[5]='-' # 与上次订单间隔时间
                row_value[6]='new' # usertype
                
        return list_rowValues    
    
    # 写入Excel并保存
    def write_t_xl(self,list_rowValues,xl_addr):
        wb=xlsxwriter.Workbook(xl_addr)
        sheet=wb.add_worksheet('sheet1')
        # 写入标题
        sheet.write(0,0,'order_dt')
        sheet.write(0,1,'member_id')
        sheet.write(0,2,'member_type')
        sheet.write(0,3,'times')
        sheet.write(0,4,'last_order_dt')
        sheet.write(0,5,'interval')
        sheet.write(0,6,'user_type')
        
        # 写入处理后的数据
        len_list=len(list_rowValues)
        for i in range(len_list):
            print u'正在写入第'+str(i+1)+u'行……'
            row_value=list_rowValues[i]
            len_row=len(row_value)
            for j in range(len_row):
                sheet.write(i+1,j,row_value[j])
                   
        wb.close()
        print u'写入完毕，excel文件已生成！'
    
l=['../excel/buyer_day.xlsx']#需要输入处理的文件路径list,即可以输入多个文件进行处理
buyerManager=BuyerManager()
list_rowValues=buyerManager.getList_sorted(l)
list_rowValues_new=buyerManager.process(list_rowValues, 100)
buyerManager.write_t_xl(list_rowValues, '../excel/buyer_day_new.xlsx')

# dt1=datetime.datetime(2017,5,2,13,23,01)
# dt2=datetime.datetime(2017,3,2,12,00,00)
# dt3=datetime.datetime(2017,6,19,0,0,1)
# dt4=datetime.datetime(2017,5,21)
# print dt2
#    
# dis1=abs(dt2-dt1)# 相减返回的类型是timedelta
# dis2=dt3-dt4
# print str(dis1)
# print dis2
# print dis1>dis2
# print dis2<=datetime.timedelta(days=29)

# str1=u' 2017-01-01 09:08:58.0  '
# str1=str1.strip()[:-2]
# str2='        order_dt        '
# print re.match('[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}', str2)!=None

