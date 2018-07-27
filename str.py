import math
import xlrd
import xlwt
import numpy as np
import time
import matplotlib.pyplot as plt
from math import *

nameofop = input()
mfuture = xlrd.open_workbook(u'/Users/lin.tl/Desktop/CITIC/Strategy/Zoom_2.xls').sheets()[0]
moption = xlrd.open_workbook(u'/Users/lin.tl/Desktop/CITIC/Strategy/Moption price_'+nameofop+'.xlsx').sheets()[0]
dayvdata = xlrd.open_workbook(u'/Users/lin.tl/Desktop/CITIC/Strategy/dayvolmonths.xlsx').sheets()[0]
n1 = 4
n2 = 6
taryield = 0.01


def phi(x):
    #'Cumulative distribution function for the standard normal distribution'
    return (1.0 + erf(x / sqrt(2.0))) / 2.0


def bsform(s,x,r,sigma,tau):
    d1 = (np.log(s/x)+(r+sigma*sigma/2)*tau)/(sigma*np.sqrt(tau))
    d2 = d1 - sigma*sigma*np.sqrt(tau)
    return s*phi(d1)-x*np.exp(-r*tau)*phi(d2)

def calc_impliedvol(st, en, op_price, fut_price, tau):
    if en - st < 0.000001:
        return st
    else:
        mid = (st+en)/2
        pst = bsform(fut_price, int(nameofop), 0.033, st, tau) - op_price
        pen = bsform(fut_price, int(nameofop), 0.033, en, tau) - op_price
        pmid = bsform(fut_price, int(nameofop), 0.033, mid, tau) - op_price
        if pst*pmid <= 0:
            return calc_impliedvol(st, mid, op_price, fut_price, tau)
        if pen*pmid <= 0:
            return calc_impliedvol(mid, en, op_price, fut_price, tau)


def my_std(inplist, days):
    return [0 for i in range(days)]+[np.std(inplist[i:i+days], ddof = 1) for i in range(len(inplist)-days+1)]

def my_norm(inp):
    return np.exp(-inp*inp/2)/math.sqrt(2*math.pi)


def calc_vega(price_fut,vol,tau):
    d1 = (np.log(price_fut/int(nameofop))+(0.033+vol*vol/2)*tau)/(vol*np.sqrt(tau))
    vega = np.sqrt(tau)*price_fut*my_norm(d1)
    return vega


price_op = np.array(moption.col_values(1)[1:])
price_fut = np.array(mfuture.col_values(1)[1:])
t_date = np.array(mfuture.col_values(0)[1:])

price_st = np.array(dayvdata.col_values(1)[1:])
price_low = np.array(dayvdata.col_values(2)[1:])
tau = np.array(dayvdata.col_values(4)[1:])
day_vol = np.log(price_st/price_low)*np.log(price_st/price_low)/(4*np.log(2))

#domain = (np.array(range(-1000,1000)))/1000
#print([bsform(price_fut[5], int(nameofop), 0.033, domain[i], tau[5])-price_op[5] for i in range(len(domain))])

impvol = np.array([calc_impliedvol(-0.29,0.3, price_op[i], price_fut[i], tau[i]) for i in range(len(price_op))] )
print(impvol)
vega = calc_vega(price_fut,impvol,tau)
d1 = (np.log(price_st/int(nameofop))+(0.033+impvol*impvol/2)*tau)/(impvol*np.sqrt(tau))
delta = [phi(i) for i in d1]
sigma = vega*day_vol

pos = 0
open_pos = 0
cyield = 0.00
tyield = [0]

sigv1 =[0 for i in range(0,n2)] + [np.sum(sigma[i-n1:i])/np.sum(vega[i-n1:i]) for i in range(n2,len(sigma))]
sigv2=[0 for i in range(0,n2)] + [np.sum(sigma[i-n2:i])/np.sum(vega[i-n1:i]) for i in range(n2,len(sigma))]

for i in range(n2, len(sigma)):
    if (cyield > taryield) or (sigv1[i]>sigv1[i-1]):
        cyield = 0
        pos = 0
    if pos == 0 and sigv1[i]<sigv2[i]:
        open_pos = 1
    if pos == 0 and open_pos == 1:
        open_pos = 0
        pos = 1
        cyield += math.log(-price_op[i]+delta[i]*price_fut[i])- math.log(-price_op[i-1]+delta[i]*price_fut[i-1])
        tyield.append(tyield[len(tyield)-1]+(-price_op[i]+delta[i]*price_fut[i])-(-price_op[i-1]+delta[i]*price_st[i]))
        continue
    if pos == 1:
        cyield += math.log(-price_op[i]+delta[i]*price_fut[i])- math.log(-price_op[i-1]+delta[i]*price_fut[i-1])
        tyield.append(tyield[len(tyield)-1]+(-price_op[i]+delta[i]*price_fut[i])-(-price_op[i-1]+delta[i]*price_st[i]))
    else:
        tyield.append(tyield[len(tyield)-1])

print(tyield)
    
    

output = xlwt.Workbook(encoding = 'ascii')
sheet1 = output.add_sheet('Sheet1',cell_overwrite_ok= True)
for i in range(0, len(tyield)):
    sheet1.write(i+6,1,tyield[i])
for i in range(0, len(vega)):
    sheet1.write(i+1,2,vega[i])
    sheet1.write(i+1,3,delta[i])
    sheet1.write(i+1,4,price_fut[i])
    sheet1.write(i+1,5,price_op[i])
    sheet1.write(i+1,6,day_vol[i])
    sheet1.write(i+1,7,sigv1[i])
    sheet1.write(i+1,8,sigv2[i])
output.save('/Users/lin.tl/Downloads/fm_c'+nameofop+'imm.xls')