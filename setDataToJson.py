#coding:utf-8
import os
import csv
import json
import io
from datetime import datetime

"""
把 csvOutFilteredWithoutPaid.csv 和 coms_info.json 中的数据整合，
生成 coms_info_complete.json 文件，该文件包含账单数据，作为 docxBuilder.py 的数据源。
"""

class SetDataToJson:
    
    def __init__(self):
        self.csvin = 'csvOutFilteredWithoutPaid.csv'
        self.jsonin = 'coms_info.json'
        self.jsonout = 'coms_info_complete.json'


    #def getTabsFromJson(self):
    #    with open(self.jsonin, 'rb') as f:
    #        reader = json.load(f)
    #       tables = reader['common']['tables']
    #    return tables

    
    def getDataFromCSV(self):
        data = {
            "hard_info": {
                "InvoiceID": "",
                "LinkedAccountId": "",
                "BillingPeriodStartDate": "",
                "BillingPeriodEndDate": "",
                "InvoiceDate": "",
                "TaxationAddress": ""
            },
            "items": [{
                "ProductCode": "",
                "CostBeforeTax": "",
                "Credits": "",
                "TaxAmount": "",
                "TotalCost": ""
            }],
            "total": {
                "CostBeforeTax": "",
                "Credits": "",
                "TaxAmount": "",
                "TotalCost": ""
            }
        }
        with open(self.csvin, 'rb') as f:
            reader = csv.reader(f)
            header = next(reader)
        
            # get(data['hard_info'])
            for row in reader:
                if row[header.index('RecordType')] == 'LinkedLineItem':
                    for col in data['hard_info']:
                        data['hard_info'][col] = row[header.index(col)]
                    break
        
            # get(data['items'])
            item = data['items'].pop()
            f.seek(0)
            for row in reader:
                if row[header.index('RecordType')] == 'LinkedLineItem' and float(row[header.index('CostBeforeTax')]) != 0:
                    for col in item:
                        item[col] = row[header.index(col)]
                    #print item
                    #print '--------------------'
                    data['items'].append(item)
                    item = dict(item)
            # debug print fomat data['items']
            #for i in data['items']:
            #    for ii in i.items():
            #        print ii
        
            # get(data['total'])
            f.seek(0)
            for row in reader:
                if row[header.index('RecordType')] == 'AccountTotal':
                    for col in data['total']:
                        data['total'][col] = row[header.index(col)]
        # test data
        #print data
        return data


    #def setDataToTabs(self):
    #    tables = self.getTabsFromJson(self.jsonin)
    #    hard_info = self.getHardInfoFromCSV(self.csvin)
    #    pass
                

    def setDataToJson(self, data):
        with open(self.jsonin, 'rb') as fin:
            reader = json.load(fin)
            tables = reader['common']['tables']
            para = reader['common']['paragraphs']['above_billing_paragraph']
            
            # set payment_overview_table
            tab = tables['payment_overview_table']['items'][0]
            tab[u'付款通知号码：'] = '{invoiceID} - {date}'.format(
                invoiceID=data['hard_info']['InvoiceID'], 
                date=datetime.strptime(data['hard_info']['BillingPeriodStartDate'], '%Y/%m/%d %H:%M:%S').strftime('%Y%m')
            )
            tab[u'付款通知日期：'] = datetime.strftime(datetime.now(), '%Y%m%d')
            tab[u'总金额 到期日期：'] = ''
            tab[u'预存总余额：'] = ''
        
            # set billing_overview_table
            tab = tables['billing_overview_table']['items'][0]
            tab[u'总服务费用'] = data['total']['TotalCost']
            tab[u'服务费用'] = data['total']['CostBeforeTax']
            tab[u'折扣'] = data['total']['Credits']
            tab[u'增值税'] = data['total']['TaxAmount']
            tab[u'本付款通知总计'] = data['total']['TotalCost']
        
            # set billing_detail_table
            tables['billing_detail_table']['items'] = []
            tab = tables['billing_detail_table']['items']
            for data_item in data['items']:
                #print data_item
                #tab_item[u'服务名'] = data_item['ProductCode']
                tab_item = dict()
                tab_item["order"] = [u"服务费用", u"折扣", u"增值税"]
                tab_item["order"].insert(0, data_item['ProductCode'])
                tab_item[data_item['ProductCode']] = data_item['TotalCost']
                tab_item[u'服务费用'] = data_item['CostBeforeTax']
                tab_item[u'折扣'] = data_item['Credits']
                tab_item[u'增值税'] = data_item['TaxAmount']
                tab.append(tab_item)
            #debug print format data['items']
            #for i in tab:
            #    for (k,v) in i.items():
            #        print k,v
            #    print '============'
            
            # set para: above_billing_paragraph
            key = u'本付款通知所包括的AWS资源的使用期间'
            if key in para:
                start_date = data['hard_info']['BillingPeriodStartDate']
                end_date = data['hard_info']['BillingPeriodEndDate']
                para[key] = '{start} - {end}'.format(start=start_date, end=end_date)
        
        
        with io.open(self.jsonout, 'w', encoding='utf-8') as fout:
            data = json.dumps(reader, ensure_ascii=False)
            fout.write(unicode(data))
    

    # clean middle file
    def clear(self):
        if self.csvin == 'csvOutFilteredWithoutPaid.csv':
            os.remove(self.csvin)
            print '{csvin} cleared!'.format(csvin=self.csvin)
    
    
    def run(self):
        # 从'csvOutFilteredWithoutPaid.csv'，提取需要的数据
        data = self.getDataFromCSV()
        # 生成'coms_info_complete.json'，作为word账单的数据源
        self.setDataToJson(data)
        # 'csvOutFilteredWithoutPaid.csv' 已经完成了它的使命，所以删掉
        #self.clear()

    
if __name__ == '__main__':
    setDataToJson = SetDataToJson()
    setDataToJson.run()
