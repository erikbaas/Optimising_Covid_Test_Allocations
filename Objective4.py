# Objective4: includes fixed charge, penalty for far travelling, data from excel
# people will be testes in their preferred timeslot or the one after that
# =========================================================
# Importing packages
# =========================================================
from gurobipy import *
from numpy import *
from openpyxl import *
from time import *
import pandas as pd
import numpy as np
import sys
#np.set_printoptions(threshold=sys.maxsize)

""" ---------------------------------- LOADING IN DATA FROM DATABASE -------------------------------------- """

# =========================================================
# Data
# =========================================================
alpha = 0.5  # weight factor
fixedcharge = 1664  # fixed charge to open up TL_j
testcost = 100  # cost per testkit/employee
minopen = 3  # min number of testees to open up a location

# sets to determine ranges
Testlocations = ['Arnhem', 'Assen', 'Den Bosch', 'Den Haag', 'Groningen', 'Haarlem', 'Leeuwarden', 'Lelystad',
                 'Maastricht', 'Middelburg', 'Utrecht', 'Zwolle']
testlocations = range(len(Testlocations))
Livinglocations = ['Arnhem', 'Assen', 'Den Bosch', 'Den Haag', 'Groningen', 'Haarlem', 'Leeuwarden', 'Lelystad',
                   'Maastricht', 'Middelburg', 'Utrecht', 'Zwolle']
livinglocations = range(len(Livinglocations))
Timeslots = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
timeslots = range(len(Timeslots))

# =========================================================
# Reading worksheets from xlsx file
# =========================================================
workbook = pd.ExcelFile("testees.xlsx")
df_testees = pd.read_excel(workbook, 'RandomPopulation')
df_livloc = pd.DataFrame(df_testees, columns=['livloc'])
df_timepref = pd.DataFrame(df_testees, columns=['timepref'])
alltestees = df_testees.to_numpy()
df_distances = pd.read_excel(workbook, 'distancetable')
distance = df_distances.to_numpy()

# get all testees per day
alltesteesloc = alltestees[:, 0]
livloc1 = np.count_nonzero(alltesteesloc == 1)
livloc2 = np.count_nonzero(alltesteesloc == 2)
livloc3 = np.count_nonzero(alltesteesloc == 3)
livloc4 = np.count_nonzero(alltesteesloc == 4)
livloc5 = np.count_nonzero(alltesteesloc == 5)
livloc6 = np.count_nonzero(alltesteesloc == 6)
livloc7 = np.count_nonzero(alltesteesloc == 7)
livloc8 = np.count_nonzero(alltesteesloc == 8)
livloc9 = np.count_nonzero(alltesteesloc == 9)
livloc10 = np.count_nonzero(alltesteesloc == 10)
livloc11 = np.count_nonzero(alltesteesloc == 11)
livloc12 = np.count_nonzero(alltesteesloc == 12)

# Final big data version
# testees = [livloc1, livloc2, livloc3, livloc4, livloc5, livloc6, livloc7, livloc8, livloc9, livloc10, livloc11,
#            livloc12]  # testees at loc i total

# Initial small data version, replace later^
testees = [12, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # testees at loc i total

# total number of testees:
Ttot = len(alltestees)

#Now, adjust manually, for the small data set:
Ttot = 12

xtot = []

# preferences for time
alltesteestime = alltestees[:, 1]
liv1tim1 = 0
liv1tim2 = 0
liv1tim3 = 0
liv1tim4 = 0
liv1tim5 = 0
liv2tim1 = 0
liv2tim2 = 0
liv2tim3 = 0
liv2tim4 = 0
liv2tim5 = 0
liv3tim1 = 0
liv3tim2 = 0
liv3tim3 = 0
liv3tim4 = 0
liv3tim5 = 0
liv4tim1 = 0
liv4tim2 = 0
liv4tim3 = 0
liv4tim4 = 0
liv4tim5 = 0
liv5tim1 = 0
liv5tim2 = 0
liv5tim3 = 0
liv5tim4 = 0
liv5tim5 = 0
liv6tim1 = 0
liv6tim2 = 0
liv6tim3 = 0
liv6tim4 = 0
liv6tim5 = 0
liv7tim1 = 0
liv7tim2 = 0
liv7tim3 = 0
liv7tim4 = 0
liv7tim5 = 0
liv8tim1 = 0
liv8tim2 = 0
liv8tim3 = 0
liv8tim4 = 0
liv8tim5 = 0
liv9tim1 = 0
liv9tim2 = 0
liv9tim3 = 0
liv9tim4 = 0
liv9tim5 = 0
liv10tim1 = 0
liv10tim2 = 0
liv10tim3 = 0
liv10tim4 = 0
liv10tim5 = 0
liv11tim1 = 0
liv11tim2 = 0
liv11tim3 = 0
liv11tim4 = 0
liv11tim5 = 0
liv12tim1 = 0
liv12tim2 = 0
liv12tim3 = 0
liv12tim4 = 0
liv12tim5 = 0

for a in range(len(alltestees)):
    if alltesteestime[a] and alltesteesloc[a] == 1:
        liv1tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 1:
        liv1tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 1:
        liv1tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 1:
        liv1tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 1:
        liv1tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 2:
        liv2tim1 += 1
    elif alltesteestime[a] and alltesteesloc[a] == 2:
        liv2tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 2:
        liv2tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 2:
        liv2tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 2:
        liv2tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 3:
        liv3tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 3:
        liv3tim2 += 1
    elif alltesteestime[a] and alltesteesloc[a] == 3:
        liv3tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 3:
        liv3tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 3:
        liv3tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 4:
        liv4tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 4:
        liv4tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 4:
        liv4tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 4:
        liv4tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 4:
        liv4tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 5:
        liv5tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 5:
        liv5tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 5:
        liv5tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 5:
        liv5tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 5:
        liv5tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 6:
        liv6tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 6:
        liv6tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 6:
        liv6tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 6:
        liv6tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 6:
        liv6tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 7:
        liv7tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 7:
        liv7tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 7:
        liv7tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 7:
        liv7tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 7:
        liv7tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 8:
        liv8tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 8:
        liv8tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 8:
        liv8tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 8:
        liv8tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 8:
        liv8tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 9:
        liv9tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 9:
        liv9tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 9:
        liv9tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 9:
        liv9tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 9:
        liv9tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 10:
        liv10tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 10:
        liv10tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 10:
        liv10tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 10:
        liv10tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 10:
        liv10tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 11:
        liv11tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 11:
        liv11tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 11:
        liv11tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 11:
        liv11tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 11:
        liv11tim5 += 1
    elif alltesteestime[a] == 1 and alltesteesloc[a] == 12:
        liv12tim1 += 1
    elif alltesteestime[a] == 2 and alltesteesloc[a] == 12:
        liv12tim2 += 1
    elif alltesteestime[a] == 3 and alltesteesloc[a] == 12:
        liv12tim3 += 1
    elif alltesteestime[a] == 4 and alltesteesloc[a] == 12:
        liv12tim4 += 1
    elif alltesteestime[a] == 5 and alltesteesloc[a] == 12:
        liv12tim5 += 1

# Big data version, use later
# livtimepref = [[liv1tim1, liv1tim2, liv1tim3, liv1tim4, liv1tim5],
#                [liv2tim1, liv2tim2, liv2tim3, liv2tim4, liv2tim5],
#                [liv3tim1, liv3tim2, liv3tim3, liv3tim4, liv3tim5],
#                [liv4tim1, liv4tim2, liv4tim3, liv4tim4, liv4tim5],
#                [liv5tim1, liv5tim2, liv5tim3, liv5tim4, liv5tim5],
#                [liv6tim1, liv6tim2, liv6tim3, liv6tim4, liv6tim5],
#                [liv7tim1, liv7tim2, liv7tim3, liv7tim4, liv7tim5],
#                [liv8tim1, liv8tim2, liv8tim3, liv8tim4, liv8tim5],
#                [liv9tim1, liv9tim2, liv9tim3, liv9tim4, liv9tim5],
#                [liv10tim1, liv10tim2, liv10tim3, liv10tim4, liv10tim5],
#                [liv11tim1, liv11tim2, liv11tim3, liv11tim4, liv11tim5],
#                [liv12tim1, liv12tim2, liv12tim3, liv12tim4, liv12tim5]]

# Small data version, delete later
                # monday, tuesday, wednesday #etc
livtimepref = [[3, 9, 0, 0, 0], # Assen
                [0, 0, 0, 0, 0], # Arnhem
                [0, 0, 0, 0, 0], # etc
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0]]

# Capacity per day per location [monday, tuesday, wednesday, thursday, friday]
loccapt = [[5, 5, 5, 5, 5],     # For test location 1
           [5, 5, 5, 5, 5],     # For test location 2
           [5, 5, 5, 5, 5],     # For test location 3
           [5, 5, 5, 5, 5],     # For test location 4
           [5, 5, 5, 5, 5],     # For test location 5
           [5, 5, 5, 5, 5],
           [5, 5, 5, 5, 5],
           [5, 5, 5, 5, 5],
           [5, 5, 5, 5, 5],
           [5, 5, 5, 5, 5],
           [5, 5, 5, 5, 5],
           [5, 5, 5, 5, 5]]  # test location capacity per time slot t

# Capacity per week per location [Assen, Arnhem, etc.]
loccap = [21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21]

# Big number M needed for penalty and fixed charge constraint
M = 99999

# Penalty value
PV_Delay = 5
PV_Distance =800

# ==========================================================
# Start modelling optimization problem
# ==========================================================
m = Model('objective')

# decision variables
x = {}  # number of testees travelling from i to j during t
b = {}  # binary penalty travel distance
y = {}  # binary fixed charge cost
O = {}  # Surplus


""" ---------------------------------- OBJECTIVE FUNCTION -------------------------------------- """

for j in testlocations:
    for i in livinglocations:
        b[i, j] = m.addVar(obj=+alpha * PV_Distance, lb=0, vtype=GRB.BINARY)  # Penalty cost for large distance
        for t in timeslots:
            y[j, t] = m.addVar(obj=+(1 - alpha) * fixedcharge, lb=0, vtype=GRB.BINARY)  # Binary fixed charge cost
            x[i, j, t] = m.addVar(obj=(+alpha * distance[i][j] + (1 - alpha) * testcost), lb=0,
                                  vtype=GRB.INTEGER)  # Penalty for distance x x_ij
            O[i, j, t] = m.addVar(obj=(+alpha * PV_Delay), lb=0,  # Penalty for delay
                                  vtype=GRB.INTEGER)

m.update()
m.setObjective(m.getObjective(), GRB.MINIMIZE)  # The objective is to minimize travel distance + delay + facility cost


""" ---------------------------------- CONSTRAINTS -------------------------------------- """

for j in testlocations:
    # k2 number of testees travelling from i to j <= to capacity of j per week (= nr of testkits)
    m.addConstr(quicksum(x[i, j, t] for i in livinglocations for t in timeslots), GRB.LESS_EQUAL, loccap[j], "Testlocation capacity per week")

    for t in timeslots:
        # k1 number of testees per day
        m.addConstr(quicksum(x[i, j, t] for i in livinglocations), GRB.LESS_EQUAL, loccapt[j][t], "Test location timeslot capacity per day")
        # k3 number of testees required to open up test location
        m.addConstr((quicksum(x[i, j, t] for i in livinglocations) - M * y[j,t]), GRB.LESS_EQUAL,
                    minopen, "Minimum people required to open up test location for a day")

    for i in livinglocations:
        # k6 soft constraint
        m.addConstr(-M * b[i, j] + distance[i][j] * quicksum(x[i, j, t] for t in timeslots), GRB.LESS_EQUAL,
                    25 * quicksum(x[i, j, t] for t in timeslots), "soft distance constraint")

# K 90: Set delay penalty per amount of people not placed on a day
for j in testlocations:
    for i in livinglocations:
        for t in range(0,len(Timeslots)):
            if t == 0:  # Monday
                m.addConstr( livtimepref[i][t] - (quicksum(x[i,j,t] for j in testlocations)), GRB.GREATER_EQUAL, 0, "Delay constraint for the first day")
            elif t == 1:  # Tuesday: include overschot van i-1
                m.addConstr( livtimepref[i][t-1] - (quicksum(x[i,j,t-1] for j in testlocations)) - O[i,j,t], GRB.LESS_EQUAL, 0, "Delay constraint for the second day")
            else: #Other days, include overschot of day before AND all before that
                m.addConstr( livtimepref[i][t-1] - (quicksum(x[i,j,t-1] for j in testlocations)) + O[i,j,t-1] - O[i,j,t], GRB.LESS_EQUAL, 0, "Delay constraint for remainder days")

# Sum of all people with symptoms must equal all people that get tested
m.addConstr(quicksum(x[i,j,t] for i in livinglocations for j in testlocations for t in timeslots), GRB.EQUAL, Ttot)

# Sum of all people from i that get tested must equal sum of all people in i
for i in livinglocations:
    xtot.append(quicksum(x[i, j, t] for j in testlocations for t in timeslots))
    m.addConstr(xtot[i], GRB.EQUAL, testees[i], "all testees get tested")

m.update()
# ==========================================================
# start optimising
# ==========================================================
# m.write('test.lp')
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

# Result in matrix format
result_array = np.array([])

print()
print("Objective Function =", m.ObjVal / 1.0)
for i in livinglocations:
    for j in testlocations:
        result_array = np.append(result_array,[Livinglocations[i], Testlocations[j]])
        for t in timeslots:
            if x[i, j, t].X != 0: # ONLY print the travel routs which are not zero!
                print(x[i, j, t].X, "people living in", Livinglocations[i], "go to a test location in", Testlocations[j],
                  "on", Timeslots[t], "to get tested")

        # for t in range(1,len(Timeslots)):
        #     print(O[i, j, t].X, "of overschot from people living in", Livinglocations[i], "go to a test location in",
        #               Testlocations[j],
        #               "on", Timeslots[t], "to get tested")


            result_array = np.append(result_array,x[i, j, t].X)

# Write to excel
result_array = pd.DataFrame(np.reshape(result_array,(-1,7)),columns=['Living Location', 'Test Location', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])
result_array.to_excel('results.xlsx')

