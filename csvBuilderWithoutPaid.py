#coding:utf-8
import os
import csv
from datetime import datetime
from decimal import Decimal

"""
{
    "文档说明": "输出表格，该表格含有客户的付款信息，并对未付款项的最终账单保留两位小数（四舍五入）
                ，作为 setDataToJson.py 的数据源",
    "函数": "run(year_month)",
    "输入变量": {
        "year_month": "账单年月(格式形如：2017/04)，这个日期是出账日期，对应源表格中的 invoiceID 列"
    },
    "输出变量": {
        "csvOutFilteredWithoutPaid.csv": "输出的表格，该文件名默认不可修改"
    },
    "示例": "见主函数"
}
"""

class CSVBuilderWithoutPaid:
    
    # 由于其它文件可能会使用 'csvOutFiltered.csv' 'csvOutFilteredWithoutPaid.csv' 这两个变量名，
    # 所以保护 csvin csvout，不使用户接触 csvin csvout 变量
    def __init__(self, year_month):
        self.year_month = year_month
        self.csvin = 'csvOutFiltered.csv'
        self.csvout = 'csvOutFilteredWithoutPaid.csv'

        
    def getStrToDate(self, str):
        # return datetime
        return datetime.strptime(str, '%Y/%m/%d %H:%M:%S')

        
    def getYearMonth(self, date):
        # return string
        return date.strftime('%Y/%m')
    
    
    def colsStrToDecimal(self, header, row):
        row[header.index('CostBeforeTax')] = Decimal(row[header.index('CostBeforeTax')])
        row[header.index('Credits')] = Decimal(row[header.index('Credits')])
        row[header.index('TaxAmount')] = Decimal(row[header.index('TaxAmount')])
        row[header.index('TotalCost')] = Decimal(row[header.index('TotalCost')])
        # python float 对象计算不精确，弃用
        #row[header.index('CostBeforeTax')] = float(row[header.index('CostBeforeTax')])
        #row[header.index('Credits')] = float(row[header.index('Credits')])
        #row[header.index('TaxAmount')] = float(row[header.index('TaxAmount')])
        #row[header.index('TotalCost')] = float(row[header.index('TotalCost')])
    

    # 参数 q 表示精度，例如：q='0.00' 表示保留两位小数（四舍五入）
    def decimalQuantize(self, header, rows, q):
        for row in rows:
            if row != header:
                row[header.index('CostBeforeTax')] = row[header.index('CostBeforeTax')].quantize(Decimal(q))
                row[header.index('Credits')] = row[header.index('Credits')].quantize(Decimal(q))
                row[header.index('TaxAmount')] = row[header.index('TaxAmount')].quantize(Decimal(q))
                row[header.index('TotalCost')] = row[header.index('TotalCost')].quantize(Decimal(q))
    
    
    def rowCostSubPaidRowsCost(self, header, row, paid_rows):
        for paid_row in paid_rows:
            if paid_row != header:
                row[header.index('CostBeforeTax')] -= paid_row[header.index('CostBeforeTax')]
                row[header.index('Credits')] -= paid_row[header.index('Credits')]
                row[header.index('TaxAmount')] -= paid_row[header.index('TaxAmount')]
                row[header.index('TotalCost')] -= paid_row[header.index('TotalCost')]

            
    # {"csvin": "CSV源表名，由 csvBuilder.py 产生的 'csvOutFiltered.csv'，该表格包含已经付款的项目，所以需要对其过滤"}
    def getRowsWithoutPaid(self):
        rows = []
        paid_rows = []
        with open(self.csvin, 'rb') as f:
            reader = csv.reader(f)
            header = next(reader)
            rows.append(header)
            paid_rows.append(header)
            for row in reader:
                self.colsStrToDecimal(header, row)
                if row[header.index('RecordType')] == 'LinkedLineItem':
                    #invoiceID = row[header.index('InvoiceID')]
                    invoiceDate_str = row[header.index('InvoiceDate')]
                    invoiceDate_dt = self.getStrToDate(invoiceDate_str)
                    row_year_month_str = self.getYearMonth(invoiceDate_dt)
                    if row_year_month_str == self.year_month:
                        rows.append(row)
                    else:
                        paid_rows.append(row)
                elif row[header.index('RecordType')] == 'AccountTotal':
                    self.rowCostSubPaidRowsCost(header, row, paid_rows)
                    rows.append(row)
        # 对未付款项的最终账单保留两位小数（四舍五入）
        self.decimalQuantize(header, rows, '0.00')
        return rows

    
    def outPutCSV(self, rows):
        with open(self.csvout, 'wb') as f:
            writer = csv.writer(f)
            writer.writerows(rows)

    
    # 'csvOutFiltered.csv' 已经完成了它的使命，所以删掉
    def clear(self):
        if self.csvin == 'csvOutFiltered.csv':
            os.remove(self.csvin)
            print '{csvin} cleared!'.format(csvin=self.csvin)
            
    
    def run(self):
        # 输出去掉已付款项的csv表格
        rows = self.getRowsWithoutPaid()
        self.outPutCSV(rows)
        #self.clear()

    
if __name__ == '__main__':
    # year_month 指出账日期（账单的次月），例如要生成3月的账单， 则出账日为4月
    year_month = datetime.now().strftime('%Y/%m')
    csvbuilderwithoutPaid = CSVBuilderWithoutPaid(year_month)
    csvbuilderwithoutPaid.run()
