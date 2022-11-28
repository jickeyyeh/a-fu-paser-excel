# 幫阿福寫的程式, 用來將業主的差勤系統吐出的Excel, 轉換成出缺勤資料

## 1. parse_excel_20220222.py

新竹台南通用版本

## 2. parse_excel_20221003.py

因為新竹廠差勤系統換廠商, 產出報表格式有變
獨立出來

2022/10/03 

(1). 更改Excel讀取欄位, 因為排序有變

(2). 更改總工作時數計算方法

總工作時數(self.work_date) ,無法相減
因為新竹場產出的報表, 為純文字
因此需要將純文字部份轉換格式為datetime格式, 才能相減


程式更動部份為

```python
                # 取得上班打卡時間C
                #self.work_start_time = self.sh.cell(row=i, column=3).value
                # 取得上班打卡時間I
                self.work_start_time = self.sh.cell(row=i, column=9).value
                self.work_start_time = datetime.datetime.strptime(self.work_start_time,"%Y-%m-%d %H:%M:%S")
                print ('上班打卡時間', self.work_start_time)
                # 取得下班打卡時間D
                #self.work_end_time = self.sh.cell(row=i, column=4).value
                # 取得下班打卡時間J
                self.work_end_time = self.sh.cell(row=i, column=10).value
                self.work_end_time = datetime.datetime.strptime(self.work_end_time,"%Y-%m-%d %H:%M:%S")
                print ('下班打卡時間', self.work_end_time)
                # 計算總工作時數
                self.work_date = (self.work_end_time - self.work_start_time)

```
