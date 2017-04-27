# payment-order-auto-generation
### Auto generate payment order from csv to docx

 - 进度：目前能批量输出相应公司账单 docx 文档。
 - 要点：对数据过滤，去掉已经付款的部分，对数据二次处理，结果导入另一个 json 文件，代码中使用 unicode 编码中文字符。
 - 说明：公司信息写入 json 模版。程序需要两个参数，年月（格式2017/04，默认自动获取）和账单文件名（*2017-03.csv，根据需要自行指定）。

#### 将公司信息写入模版文件 com_info.json ，图例说明：

![image](https://github.com/mengyetxz/payment-order-auto-generation/blob/master/screenshots/com_info_json.png)

#### 使用方法：
```shell
$ pip install python-csv python-docx
$ python run.py
```

#### 生成的付款通知：

![image](https://github.com/mengyetxz/payment-order-auto-generation/blob/master/screenshots/overview_1.png)

![image](https://github.com/mengyetxz/payment-order-auto-generation/blob/master/screenshots/overview_2.png)

![image](https://github.com/mengyetxz/payment-order-auto-generation/blob/master/screenshots/overview_3.png)

![image](https://github.com/mengyetxz/payment-order-auto-generation/blob/master/screenshots/overview_4.png)


#### PS：未填写的数据：（原因：CSV表格中未找到对应数据）
![image](https://github.com/mengyetxz/payment-order-auto-generation/blob/master/screenshots/none.png)


