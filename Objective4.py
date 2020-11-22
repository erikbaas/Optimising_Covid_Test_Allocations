#Objective4: includes fixed charge, penalty for far travelling, data from excel
# people will be testes in their preferred timeslot or the one after that
# =========================================================
# objective statement
# =========================================================
#Minimise the distance traveled by people going to a Covid test facility (weight $\alpha$: 0.6)
#Minimise the cost for the test- locations in the Netherlands (weight $1-\alpha$: 0.4)
#min  Z = \alpha Z_{1} + (1-\alpha)Z_{2}
# $Z_{1}$ = Minimisation of the travelling distance (km)
# Z_{1} = \sum_{i=1}^{n}\sum_{j=1}^{m} c_{ij}x_{ij} + 1000\sum_{i=1}^n\sum_{j=1}^{m}b_{i,j}
#Z_{2}$ = Minimisation of the fixed charge cost per opening test facility and the variable cost regarding needed testkits/employees
#Z_{2} = c_{3}\sum_{j=1}^{n}x_{ij} +c_4\sum_{j=1}^{n}y_j

# =========================================================
# Importing packages
# =========================================================

from gurobipy import *
from numpy import *
from openpyxl import *
from time import *
import pandas as pd
import numpy as np

# =========================================================
# Data
# =========================================================
alpha = 0.5 #weight factor
fixedcharge = 50 #fixed charge to open up TL_i
testcost = 10 #cost per testkit/employee

#sets to determine ranges
Testlocations = ['TL1','TL2','TL3']
testlocations = range(len(Testlocations))
Livinglocations = ['LL1','LL2','LL3']
livinglocations = range(len(Livinglocations))
Timeslots = ['TS1','TS2','TS3']
timeslots = range(len(Timeslots))

#distance data from location to location
distance1 = [[0, 20, 50], #travel cost from i to j and back c1
            [20, 0, 20],
            [20, 50, 0]]

# =========================================================
# Reading worksheets from xlsx file
# =========================================================
workbook = pd.ExcelFile("testees.xlsx")
df_testees = pd.read_excel(workbook, 'RandomPopulation')
df_livloc = pd.DataFrame(df_testees, columns= ['livloc'])
df_timepref = pd.DataFrame(df_testees, columns= ['timepref'])
alltestees = df_testees.to_numpy()
df_distances= pd.read_excel(workbook, 'distancetable')
distance = df_distances.to_numpy()
#print(alltestees)

#get all testees per day
alltesteesloc = alltestees[:,0]
livloc1 = np.count_nonzero(alltesteesloc==1)
livloc2 = np.count_nonzero(alltesteesloc==2)
livloc3 = np.count_nonzero(alltesteesloc==3)
testees = [livloc1,livloc2,livloc3] #testees at loc i total

# total number of testees:
Ttot = len(alltestees)
xtot=[]

#preferences for time
alltesteestime = alltestees[:,1]
liv1tim1=0
liv1tim2=0
liv1tim3=0
liv2tim1=0
liv2tim2=0
liv2tim3=0
liv3tim1=0
liv3tim2=0
liv3tim3=0
for a in range(len(alltestees)):
    if alltesteestime[a] and alltesteesloc[a] ==1:
        liv1tim1+=1
    elif alltesteestime[a] ==2 and alltesteesloc[a] ==1:
        liv1tim2+=1
    elif alltesteestime[a] ==3 and alltesteesloc[a] ==1:
        liv1tim3+=1
    elif alltesteestime[a] ==1 and alltesteesloc[a] ==2:
        liv2tim1+=1
    elif alltesteestime[a] and alltesteesloc[a] ==2:
        liv2tim2+=1
    elif alltesteestime[a] ==3 and alltesteesloc[a] ==2:
        liv2tim3+=1
    elif alltesteestime[a] ==1 and alltesteesloc[a] ==3:
        liv3tim1+=1
    elif alltesteestime[a] ==2 and alltesteesloc[a] ==3:
        liv3tim2+=1
    elif alltesteestime[a] and alltesteesloc[a] ==3:
        liv3tim3+=1
livtimepref= [[liv1tim1,liv1tim2,liv1tim3,0],
             [liv2tim1,liv2tim2,liv2tim3,0],
             [liv3tim1,liv3tim2,liv3tim3,0]]
livtim=[]
print(livtimepref)

# location capacities
loccapt = [[15, 40, 10],
           [30, 10, 0],
           [30, 40, 0]] # test location capacity per time slot t
loccap = [90,80,0] # test location capacity total

# big number M needed for penalty and fixed charge constraint
M=10000

# ==========================================================
# Start modelling optimization problem
# ==========================================================
m = Model('objective1')

# decision variables
x = {} #number of testees travelling from i to j during t
b = {} #binary penalty travel distance
y = {} #binary fixed charge cost
z={} #penalty for late test

for j in testlocations:
    y[j] = m.addVar(obj=+(1-alpha)*fixedcharge, lb=0,vtype=GRB.BINARY) #binary fixed charge cost
    for i in livinglocations:
        b[i,j] = m.addVar(obj=+alpha*1000, lb=0,vtype=GRB.BINARY)  # plus penalty cost
        for t in timeslots:
            x[i,j,t] = m.addVar(obj = (+alpha*distance[i][j]+(1-alpha)*testcost),lb=0,vtype=GRB.INTEGER) # distance x x_ij
for i in livinglocations:
    for t in timeslots:
        z[i,t] = m.addVar(obj=+ 10, lb=0, vtype=GRB.BINARY)  # penalty for late test

m.update()
m.setObjective(m.getObjective(), GRB.MINIMIZE)  # The objective is to minimize travel distance+ delay + facility cost


for j in testlocations:
    # k2 number of testees travelling from i to j = smaller or equal to capacity of j
    m.addConstr(quicksum(x[i,j,t] for i in livinglocations for t in timeslots), GRB.LESS_EQUAL, loccap[j])
    # k3 number of people needed to open up testloc
    m.addConstr((quicksum(x[i,j,t] for i in livinglocations for t in timeslots) - M * y[j]), GRB.LESS_EQUAL, fixedcharge)
    for t in timeslots:
        # k1 number of testees per timeslot
        m.addConstr(quicksum(x[i,j,t] for i in livinglocations), GRB.LESS_EQUAL, loccapt[j][t])
    for i in livinglocations:
        # k6 soft constraint
        m.addConstr( -M*b[i,j] + distance[i][j] * quicksum(x[i,j,t] for t in timeslots), GRB.LESS_EQUAL, 10 * quicksum(x[i,j,t] for t in timeslots))
#k4 people wait max so long
for t in timeslots:
    for i in livinglocations:
        # make sure they can only be added to their time slot
        m.addConstr(quicksum(x[i, j, t] for j in testlocations), GRB.LESS_EQUAL, livtimepref[i][t] + livtimepref[i][t - 1])
        #add penalty for non preferred time slot
        m.addConstr(quicksum(x[i,j,t] for j in testlocations), GRB.GREATER_EQUAL, livtimepref[i][t] - z[i,t]*(livtimepref[i][t]-quicksum(x[i,j,t] for j in testlocations)))

#sum of all people with symptoms must equal all people that get tested
#m.addConstr(quicksum(x[i,j,t] for i in livinglocations for j in testlocations for t in timeslots), GRB.EQUAL, Ttot)
#sum of all people from i that get tested must equal sum of all people in i
for i in livinglocations:
    xtot.append(quicksum(x[i,j,t] for j in testlocations for t in timeslots))
    m.addConstr(xtot[i], GRB.EQUAL, testees[i])

m.update()
# ==========================================================
# start optimising
# ==========================================================
#m.write('test.lp')
# Set time constraint for optimization (5minutes)
# m.setParam('TimeLimit', 1 * 60)
# m.setParam('MIPgap', 0.009)
m.optimize()
# m.write("testout.sol")
status = m.status

if status == GRB.Status.UNBOUNDED:
    print('The model cannot be solved because it is unbounded')

elif status == GRB.Status.OPTIMAL or True:
    f_objective = m.objVal
    print('***** RESULTS ******')
    print('\nObjective Function Value: \t %g' % f_objective)

elif status != GRB.Status.INF_OR_UNBD and status != GRB.Status.INFEASIBLE:
    print('Optimization was stopped with status %d' % status)
    print('\nThe following constraint(s) cannot be satisfied:')
    for c in m.getConstrs():
        if c.IISConstr:
            print('%s' % c.constrName)

# ========================================================
# Print out Solutions
# ========================================================
print()
print ("Objective Function =", m.ObjVal/1.0)
for i in livinglocations:
    for j in testlocations:
        for t in timeslots:
            print( x[i,j,t].X, "people from livinglocation", Livinglocations[i], "go to testlocation", Testlocations[j], "during timeslot", Timeslots[t])

