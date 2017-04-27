#coding:utf-8
import io
import os
import csv
import json
import time
import traceback
from functools import wraps
from docx import Document
from decimal import Decimal
from datetime import datetime

from csvBuilder import CSVBuilder
from csvBuilderWithoutPaid import CSVBuilderWithoutPaid
from setDataToJson import SetDataToJson
from docxBuilder import DocxBuilder

"""
doc
"""

def fn_timer(function):
    @wraps(function)
    def function_timer(*args, **kwargs):
        t0 = time.time()
        result = function(*args, **kwargs)
        t1 = time.time()
        print('Total time running %s: %s seconds' %
                (function.func_name, str(t1-t0))
                )
        return result
    return function_timer


class Builder:

    def __init__(self, csvin, year_month):
        self.csvin = csvin
        self.year_month = year_month
        # filter columns
        self.columns = (
        'InvoiceID',
        'LinkedAccountId',
        'RecordType',
        'BillingPeriodStartDate',
        'BillingPeriodEndDate',
        'InvoiceDate',
        'TaxationAddress',
        'ProductCode',
        'CostBeforeTax',
        'Credits',
        'TaxAmount',
        'TotalCost')
        self.successCount = 0
        self.failingCount = 0


    def getPayment(self, com_name, linkedAccountId):
        CSVBuilder(linkedAccountId, self.columns, self.csvin).run()
        CSVBuilderWithoutPaid(self.year_month).run()
        SetDataToJson().run()
        DocxBuilder(com_name).run()
        
        
    # 返回｛公司名称：账户ID｝的键值对
    def getComs_info(self):
        coms_info = {}
        with open('coms_info.json', 'rb') as f:
            reader = json.load(f)
            reader = reader['com_info']
            for com_info in reader:
                coms_info[com_info] = reader[com_info][u'账户号码']
                #testprint
                #print com_info, coms_info[com_info]
        return coms_info
        
        
    # 生成全部公司账单
    def getAllPayments(self):
        linkedAccountId = self.getComs_info()
        for com_name in linkedAccountId:
            try:
                self.getPayment(com_name, linkedAccountId[com_name])
            except ValueError as e:
                print u'Error: Payment order {com_name} failed to create.'.format(com_name=com_name)
                print u'    Maybe this linkedAccountId:{id} not exist in the {csvin}.'.format(
                    id=linkedAccountId[com_name],
                    csvin=self.csvin
                    )
                print(traceback.format_exc())
                self.failingCount += 1
                continue
            except Exception:
                print u'Error: Payment order {com_name} failed to create.'.format(com_name=com_name)
                print(traceback.format_exc())
                self.failingCount += 1
                continue
            self.successCount += 1
    
    
    def printRunInfo(self):
        print '{0} success and {1} failing.'.format(self.successCount, self.failingCount)
    
    
def clear():
    print 'Removing middle file ...'
    os.remove('csvOut.csv')
    os.remove('csvOutFiltered.csv')
    os.remove('csvOutFilteredWithoutPaid.csv')
    os.remove('coms_info_complete.json')
    print 'middle file cleared!'
    print 'Removing .pyc file ...'
    os.remove('csvBuilder.pyc')
    os.remove('csvBuilderWithoutPaid.pyc')
    os.remove('docxBuilder.pyc')
    os.remove('setDataToJson.pyc')
    print '.pyc cleared!'
    print ''
        

@fn_timer
def run(csvin, year_month):
    builder = Builder(csvin, year_month)
    builder.getAllPayments()
    clear()
    builder.printRunInfo()
        
        
if __name__ == '__main__':
    # 只需要修改需要用到的 csv 文件名
    csvin = '412764460734-aws-cost-allocation-ACTS-2017-03.csv'
    # year_month 指出账日期（账单的次月），例如要生成3月的账单， 则出账日为4月
    year_month = datetime.now().strftime('%Y/%m')
    run(csvin, year_month)
    