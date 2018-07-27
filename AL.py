# Data-cleaner. 
# 7-17.

import math
import xlrd
import xlwt

#####################
main_month = range(1,13)
#####################

def cdiv(i, divider = 12):
    if i % divider == 0:
        return divider
    else:
        return i % divider

def divide_year(month):
    dic = {}
    dic[0] = -1
    m_len = len(month)
    for i in range(m_len):
        lef = cdiv(month[i]+11)
        rig = cdiv(month[(i+1)%m_len]+11)
        if lef < rig or (lef == rig):
            for j in range(lef,rig):
                dic[j] = (i+1)%m_len
        else:
            for j in range(lef,13):
                dic[j] = (i+1)%m_len
            for j in range(1,rig):
                dic[j] = (i+1)%m_len
    dic[0] = -1
    return dic


def backvar(days, yie, starts = 0):
	normv = []

	for k in range(starts, len(yie)-days+1):
		ssum = 0
		nsum = 0
		biassum = 0
		for i in range(days):
			nsum += yie[starts+k+i]
			#ssum += yie[starts+i]*yie[starts+i]
			#normv.append(math.sqrt((ssum-(nsum*nsum)/days)/(days-1)))
		for i in range(days):
			biassum += (yie[starts+k+i]-nsum/days)*(yie[starts+k+i]-nsum/days)
			#print(math.sqrt(biassum/(days-1))-normv[i])
		normv.append(math.sqrt(biassum/(days-1)))
		#print(k,nsum,biassum,normv[i])
	return normv



workbook = xlrd.open_workbook(u'/Users/lin.tl/Downloads/AL price.xlsx')
datasheet = workbook.sheets()[0]

nrows = datasheet.nrows
ncols = datasheet.ncols

dic = divide_year(main_month)
r_price = []
r_yield = []
ln_yield = []
raw_mat = []
m_len = len(main_month)
for j in range(m_len):
	raw_mat.append([])

for i in range(1,nrows):
	for j in range(1,m_len+1):
		raw_mat[j-1].append(sum(datasheet.row_values(i)[j::m_len]))


for i in range(1,nrows):
	year, month, day, a, b, c = xlrd.xldate_as_tuple(datasheet.cell(i,0).value, 0)
	if i == 1 or raw_mat[dic[month]][i-2] == 0:
		r_price.append(raw_mat[dic[month]][i-1])
		r_yield.append(-12345)
		ln_yield.append(-12345)
		continue
	r_price.append(raw_mat[dic[month]][i-1])
	r_yield.append((raw_mat[dic[month]][i-1]/raw_mat[dic[month]][i-2])-1)
	ln_yield.append(math.log(1+r_yield[i-1]))

sigma30 = backvar(30, ln_yield)
sigma60 = backvar(60, ln_yield)
sigma90 = backvar(90, ln_yield)

output = xlwt.Workbook(encoding = 'ascii')
sheet1 = output.add_sheet('Sheet1',cell_overwrite_ok= True)
for i in range(1,nrows):
	sheet1.write(i,1, r_price[i-1])
	sheet1.write(i,2, r_yield[i-1])
	sheet1.write(i,3, ln_yield[i-1])
	if i >= 30:
		sheet1.write(i,4,sigma30[i-30])
	if i >= 60:
		sheet1.write(i,5,sigma60[i-60])
	if i >= 90:
		sheet1.write(i,6,sigma90[i-90])

sheet1.write(0,0,'时间')
sheet1.write(0,1,'价格')
sheet1.write(0,2,'收益率')
sheet1.write(0,3,'对数收益率')
sheet1.write(0,4,'30天')
sheet1.write(0,5,'60天')
sheet1.write(0,6,'90天')



output.save('/Users/lin.tl/Downloads/processed_AL.xls')




