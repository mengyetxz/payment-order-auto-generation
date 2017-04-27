#coding:utf-8
import os
import csv
from decimal import Decimal

"""
{
    "文档说明": "通过 linkedAccountId 筛选指定客户的账单数据",
    "函数": "run(linkedAccountId, columns, csvin)",
    "输入变量": {
        "linkedAccountId": "客户的账户ID",
        "columns": "账单表中需要显示的列名",
        "csvin": "S3中每月底的原始账单文件"
    },
    "输出变量": {
        "csvOutFiltered.csv": "该表格含有客户的付款信息，作为 csvBuilderWithoutPaid.py 的数据源，文件名默认不可修改"
    },
    "示例": "run('639127770567', columns, '412764460734-aws-cost-allocation-ACTS-2017-03.csv')"
}
"""

class CSVBuilder:

    def __init__(self, linkedAccountId, columns, csvin):
        self.linkedAccountId = linkedAccountId
        self.columns = columns
        self.csvin = csvin
        self.csvoutMid = 'csvOut.csv'
        self.csvoutFin = 'csvOutFiltered.csv'
    
    def colsStrToDecimal(self, header, row):
        row[header.index('CostBeforeTax')] = Decimal(row[header.index('CostBeforeTax')])
        row[header.index('Credits')] = Decimal(row[header.index('Credits')])
        row[header.index('TaxAmount')] = Decimal(row[header.index('TaxAmount')])
        row[header.index('TotalCost')] = Decimal(row[header.index('TotalCost')])
    
    def sumCols(self, header, row, sumrow):
        sumrow[header.index('CostBeforeTax')] += Decimal(row[header.index('CostBeforeTax')])
        sumrow[header.index('Credits')] += Decimal(row[header.index('Credits')])
        sumrow[header.index('TaxAmount')] += Decimal(row[header.index('TaxAmount')])
        sumrow[header.index('TotalCost')] += Decimal(row[header.index('TotalCost')])
    
    def csvBill(self):
        with open(self.csvin, 'rb') as f:
            reader = csv.reader(f)
            next(reader)
            header = next(reader)
    
            # get required rows
            rows = []
            # sum same column of cost by the same (InvoiceID, productCode)
            row_idxs_by_productCode = {}
            for row in reader:
                if row[header.index('TotalCost')] == 0 or row[header.index('LinkedAccountId')]=='':
                    continue
                elif row[header.index('LinkedAccountId')] == self.linkedAccountId:
                    # (InvoiceID, productCode) as key and row index as value
                    if (row[header.index('InvoiceID')], row[header.index('ProductCode')]) in row_idxs_by_productCode:
                        idx = row_idxs_by_productCode[(row[header.index('InvoiceID')], row[header.index('ProductCode')])]
                        self.sumCols(header, row, rows[idx])
                    else:
                        self.colsStrToDecimal(header, row)
                        rows.append(row)
                        row_idxs_by_productCode[(row[header.index('InvoiceID')], row[header.index('ProductCode')])] = rows.index(row)

        # write csv file
        with open(self.csvoutMid, 'wb') as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerows(rows)


    # filter required columns
    def filter_cols(self, delimiter=','):
        with open(self.csvoutMid, 'rb') as fin, open(self.csvoutFin, 'wb') as fout:
            reader = csv.reader(fin, delimiter=delimiter)
            writer = csv.writer(fout, delimiter=delimiter)
            header = next(reader)
            writer.writerow(tuple(column for column in header if column in self.columns))
            idxs = [idx for idx, column in enumerate(header) if column in self.columns]
            rows = (tuple(column for idx, column in enumerate(row) if idx in idxs) for row in reader)
            writer.writerows(rows)
            
    
    # clean middle file
    def clear(self):
        if self.csvoutMid == 'csvOut.csv':
            os.remove(self.csvoutMid)
            print '{csvoutMid} cleared!'.format(csvoutMid=self.csvoutMid)

    
    # 由于其它文件可能会使用 'csvOut.csv' 'csvOutFiltered.csv' 这两个变量名，所以保护起来，不使用户接触csvout变量
    def run(self):
        # 输出初步过滤好的csv表格
        self.csvBill()
        # 输出只显示需要的列的csv表格
        self.filter_cols()
        # 'csvOut.csv' 已经完成了它的使命，所以删掉
        #self.clear()
    
if __name__ == '__main__':
    # required columns
    columns = (
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
    csvBuilder = CSVBuilder('639127770567', columns, '412764460734-aws-cost-allocation-ACTS-2017-03.csv')
    csvBuilder.run()
    