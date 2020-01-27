#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 13:24:10 2019

@author: bahram.vazir.nezhad@accenture.com
"""

#from mpl_toolkits import mplot3d
import pandas as pd
#import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits import mplot3d

## authenticate to the API using your service account credentials
#minTrainSizeList = [3]#[8, 13, 34]# 5, 21, test 3 too
#confThr = [0.25, 0.3, 0.35, 0.4, 0.45]# 
#inScopeIntents = ["BalanceCheck", "BillRequest", "BillPay", "DirectDebitChange", "PaymentExtend", "ContractExpiryRequest", "SimActivate", "InternetAccess"]
#bestAccuracy = 0
#bestModifiedAccuracy = 0

data = pd.ExcelFile ("results23Oct.xlsx")
results = data.parse ("Sheet1")
fig = plt.figure ()
ax = plt.axes (projection='3d')#
x, y, z = [], [], []# Modified Total Accuracy	Min acceptable train	 Threshold
for row in range (len (results)):
    x.append (int (results ["Min acceptable train"][row]))
for row in range (len (results)):
    y.append (results ["Threshold"][row])
for row in range (len (results)):
    z.append (results ["WISF"][row])
ax = plt.axes (projection = '3d')#
ax.plot_trisurf (x, y, z,cmap = 'viridis', linewidths = 0.2);
#
#X, Y = np.meshgrid(x, y)   
#
#ax.contour3D(X, Y, Z, 50, cmap='binary')
ax.set_xlabel ("Min Sample Required")
ax.set_ylabel ("Threshold")
ax.set_zlabel ("WISF");
ax.set_title ("WISF");
#fig
ax.view_init (15, 30)
