import pandas as pd
import gurobipy as gb
from gurobipy import GRB
import numpy as np
from re import findall

#Read data
df = pd.read_excel('./data.xlsx',usecols="E:J")

try:
  #Read need data
  need_df = df.iloc[1:3,:5].copy()
  need_df.columns = need_df.iloc[0]
  need_df = need_df.iloc[1:].reset_index(drop=True)
  need_df = need_df.set_index('Need')
  #Read product data
  product_df = df.iloc[7:53,:].copy()
  product_df.columns =product_df.iloc[0]
  product_df = product_df.iloc[1:].reset_index(drop=True)
  product_df=product_df.set_index('Product ')
  #Re=product_dfad distances
  distances_df = df.iloc[55:61,:]
  distances_df.columns = distances_df.iloc[0]
  distances_df = distances_df.iloc[1:].reset_index(drop=True)
  distances_df  = distances_df.set_index('Distance')
  #Read other variables
  budget = df.iloc[4,1]
  distance_limit = df.iloc[63,2]
  distance_limit
  #Part c variables
  prd_i_c = 31
  prd_j_c = 25
  #Part d variables
  prd_i_d = 22
  prd_j_d =26
  #Part e
  t_d = 300

  model = gb.Model("Max_satisfaction")

  #maximize model
  model.modelSense = GRB.MAXIMIZE
  #parameters od model
  T = distance_limit
  B = budget 
  Si = product_df['Satisfaction'].values
  Ci = product_df['Price'].values
  Ai_temp = product_df['Amount per Packet'].values
  dkl = distances_df.values
  Dn = [2,2.5,0.4,0.3]  #n = 1, 2, 3, 4 (Beverages,       Carbodydrates, Cheese, Breakfast Foods)
  print(Dn)
  #Ai extraction
  Ai =[]
  for i in range(len(Ai_temp)):
    res=findall(r'(\d+)[.]?(\d*)?',Ai_temp[i])
    res = res[0][0] + '.'+res[0][1]
    res = float(res)
    Ai.append(res)
  Ai = np.asarray(Ai)
  # print('Ai is', Ai)
  # print(Ai.shape)
  #Need type of the product i as type n => Din
  temp= product_df['Need'].values

  for i in range (45):
      if(temp[i]== 'Beverages'):
          temp[i]= 0
      elif(temp[i]== 'Carbodydrates'):
          temp[i]= 1
      elif(temp[i]== 'Cheese'):
          temp[i]= 2
      else:
          temp[i]= 3
  Din=[]        
  for i in range(4):
      col=[0]*45
      for j in range(45):
          if(temp[j]== i ):
              col[j]=1
      Din.append(col)
  Din = np.asarray(Din)

  #Mik: Availability of the product i in the market k
  temp= product_df['Market'].values

  for i in range (45):
      if(temp[i]== 'A'):
          temp[i]= 1
      elif(temp[i]== 'B'):
          temp[i]= 2
      elif(temp[i]== 'C'):
          temp[i]= 3
      elif(temp[i]=='D'):
          temp[i]= 4
      else:
          temp[i]= 0
  Mik=[] 
  col=[0]*45
  Mik.append(col)
  for i in range(4):
      col=[0]*45
      for j in range(45):
          if(temp[j]== (i + 1 ) ):
              col[j]=1
      Mik.append(col)
  Mik = np.asarray(Mik)
  # print(Mik)
  #decision variables
  #Xi Amount of unit product taken from product 
  Xi=[]
  for num in range(45):
      Xi.append(model.addVar(vtype=GRB.INTEGER, name=("x" + str(num + 1)), lb=0))
      
  
  #If the market k is visited
  Vk=[]
  for k in range(5):
      Vk.append(model.addVar(vtype=GRB.BINARY, name=("V" + str(k))))
  
  #If the path from k to l is taken
  Rkl = []
  for k in range(5):
      temp=[]
      for l in range(5):
          temp.append(model.addVar(vtype=GRB.BINARY,obj= (dkl[k][l]), name='R' +str(k)+str(l)))
      Rkl.append(temp)
      
  #binary var of if
  Zk=[]
  for k in range(5):
      Zk.append(model.addVar(vtype=GRB.BINARY, name=("Z" + str(k))))
  #new binary var
  Yk=[]
  for k in range(5):
      Yk.append(model.addVar(vtype=GRB.BINARY, name=("Y" + str(k))))

  model.update()
  
  model.setObjective(np.sum(np.multiply(Si, Xi)) + (B - np.sum(np.multiply(Ci, Xi))), GRB.MAXIMIZE)
  
  #constraints
  # budget constraint
  model.addConstr(np.sum(np.multiply(Ci, Xi)) <= B)
  # demand constraint
  sum_demand = 0
  for n in range(4):
      coeff = Din[n][:]* Ai
      # print('Ai is ',Ai)
      sum_demand = np.sum(np.multiply(coeff,Xi))
      # print('here is in sum_demand ', n,' ', sum_demand)
      model.addConstr(sum_demand >= Dn[n] )
  # distance constraint
  model.addConstr(np.sum(np.sum(np.multiply(Rkl, dkl))) <= T)
  
  # path constraint
  for k in range(5):
    sum_r = 0
    sum_l = 0
    for l in range(5):
      sum_r = sum_r + Rkl[k][l]
      sum_l = sum_l + Rkl[l][k]
    model.addConstr((Vk[k]==1)>>(sum_r >= 1))
    model.addConstr((Vk[k]==0)>>(sum_r == 0))
    model.addConstr((Vk[k]==1)>>(sum_l >= 1))
    model.addConstr((Vk[k]==0)>>(sum_l == 0))
    
  #home constraint
  sum_to = 0
  sum_from = 0
  for k in range(5):
    sum_to = sum_to + Rkl[k][0]
    sum_from = sum_from + Rkl[0][k]
  model.addConstr(sum_to == 1)
  model.addConstr(sum_from == 1)
  
  #home -> node -> home constraint
  eps = 0.0001
  M = 30
  sum_visits = 0
  for k in range(5):
    sum_visits = sum_visits + Vk[k]
  model.addConstr( sum_visits >= 2 + eps - M * (1-Yk[k]))
  model.addConstr( sum_visits <= 2 +  M * Yk[k])
  for k in range(5):
    model.addConstr((Yk[k]==1)>>(Rkl[0][k] + Rkl[k][0] <= 1))
    
  # visiting constraint
  eps = 0.0001
  M = 30 
  for k in range(5):
    if (k != 0): 
      sum = 0
      for i in range(45):
        sum = sum + Mik[k][i] * Xi[i]
      #   print('M%g-%g %g' % (i, k, Mik[k][i]))
      # print(sum)
      model.addConstr( sum >= 0 + eps - M * (1-Zk[k]))
      model.addConstr( sum <= 0 +  M *Zk[k])
      
      model.addConstr((Zk[k]==1)>>(Vk[k] == 1))
      model.addConstr((Zk[k]==0)>>(Vk[k] == 0))
  
  # additional const.
      # Vhouse => 1
  model.addConstr(Vk[0]==1)
      # 00 11 22 33 44 => 0
  for k in range(5):  
    for l in range(5):
      if (k==l):
        model.addConstr(Rkl[k][l] == 0)
  
  # Bundled products constraint product 31 and product 25
  model.addConstr(Xi[30]<=Xi[24])
  
  # Xi >= 0
  for i in range(45):
    model.addConstr(Xi[i] >=0)
    
  # Optimize model
  model.optimize()
  for var in model.getVars():
    if var.X != 0:
      print('%s %g' % (var.VarName, var.X))
  
  print('\nObjective value: %g' % model.ObjVal)
  # for k in range(5):
  #   if(k != 0):
  #     print("\n")
  #     for i in range(45):
  #       print('M%g-%g %g' % (i, k, Mik[k][i]))
  
except gb.GurobiError as e:
    print('Error code ' + str(e.errno) + ': ' + str(e))

except AttributeError:
    print('Encountered an attribute error')