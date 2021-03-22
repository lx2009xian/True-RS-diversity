#!/usr/bin/env python
# coding: utf-8



import numpy as np
import pandas as pd



# Calculate the value of di and dj
data=pd.read_excel("C:/Users/86198/Desktop/three components/result.xlsx",header=0,index_col=0)
data_copy=pd.DataFrame(np.zeros((data.shape[0],data.shape[1])),index=data._stat_axis.values.tolist(),columns= data.columns.values.tolist())
for i in range(0,data.shape[1]):
    data_process=data.iloc[:,i]
    data_count=data_process.sum()
    data_process[data_process>0]=data_process/data_count
    data_copy.iloc[:,i]=data_process





#Create an empty dataframe to store the cumulative sum of dij of each paper
cumulative_sum_f=pd.DataFrame(np.zeros((data.shape[1],1)),index=data.columns.values.tolist())
#Create an empty dataframe to store the new diversity of each paper
True_diversity=pd.DataFrame(np.zeros((data.shape[1],1)),index=data.columns.values.tolist())





#Input the disparity matrix（use the 1-cosine as the distance measurement）
disparity_matrix=pd.read_excel("C:/Users/86198/Desktop/three components/1-cosine.xlsx",header=0,index_col=0)#input disparity matrix



#Calculate the cumulative sum of dij 
for j in range(0,data_copy.shape[1]):
    data_array=data_copy.iloc[:,j].values
    data_arrayt=np.transpose(data_array)
    cumulative_sum=np.dot(np.dot(data_arrayt,disparity_matrix),data_array)
    cumulative_sum_f.iloc[j,0]=cumulative_sum




for k in range(0,cumulative_sum_f.shape[0]):
    True_diversity.iloc[k,0]=(1/(1-cumulative_sum_f.iloc[k,0]))
True_diversity.to_excel("C:/Users/86198/Desktop/three components/True diversity.xlsx")






