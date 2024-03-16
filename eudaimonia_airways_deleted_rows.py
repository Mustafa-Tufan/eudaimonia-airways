import time
from pyomo.environ import *
import pandas as pd

# Read the sheet
excel_file = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Positions')

# Add new column for position types
excel_file.loc[0:19, 'Type'] = 0
excel_file.loc[20:59, 'Type'] = 1
excel_file.loc[60:92, 'Type'] = 0
excel_file.loc[93:103, 'Type'] = 1

original_excel_file = excel_file.copy()
excel_file.drop(labels = range(38, 60), axis=0, inplace=True)
excel_file.reset_index(inplace = True)
excel_file.drop(labels = "index", axis = 1, inplace = True)


# Sort the sheet by H-arm's and create new one
sorted_excel_file_M = excel_file.iloc[:60].sort_values(by='H-arm', ascending=True)
sorted_excel_file_L = excel_file.iloc[60:].sort_values(by='H-arm', ascending=True)

sorted_excel_file_M.to_excel('END395_ProjectPartIDataset_SORTED_M.xlsx', index=False)
sorted_excel_file_L.to_excel('END395_ProjectPartIDataset_SORTED_L.xlsx', index=False)

Positions_M = pd.read_excel('END395_ProjectPartIDataset_SORTED_M.xlsx')
Positions_L = pd.read_excel('END395_ProjectPartIDataset_SORTED_L.xlsx')


print(Positions_M.to_string())
print(Positions_L.to_string())


#Getting The Main Deck Parameters
Lock1_M = Positions_M.iloc[:,1]
Lock2_M = Positions_M.iloc[:,2]
H_arm_M = Positions_M.iloc[:,3]
Max_Weight_M = Positions_M.iloc[:,4]
Cumulative_M = Positions_M.iloc[:,5]
Coefficient_M = Positions_M.iloc[:,6]
Type_M = Positions_M.iloc[:,7]

#Getting The Lower Deck Parameters
Lock1_L = Positions_L.iloc[:,1]
Lock2_L = Positions_L.iloc[:,2]
H_arm_L = Positions_L.iloc[:,3]
Max_Weight_L = Positions_L.iloc[:,4]
Cumulative_L = Positions_L.iloc[:,5]
Coefficient_L = Positions_L.iloc[:,6]
Type_L = Positions_L.iloc[:,7]

# Take the pallets data
Pallets = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Pallets1')
OriginalPallets = Pallets.copy()

Pallet_Type = Pallets.iloc[:,1]


for i in range(len(Pallet_Type)):
    if Pallet_Type.values[i].startswith("PAG"):
        Pallet_Type.values[i] = 0
    else:
        Pallet_Type.values[i] = 1


Pallet_Weight = Pallets.iloc[:,2]

model = ConcreteModel()


model.DOW = Pallets.columns[4] #TRICKYYYYYY understandable
model.DOI = Pallets.iloc[0,4]

start_time = time.time()

# Initialize ranges
model.Main_Deck_Position_Index = RangeSet(0,59)
model.Lower_Deck_Position_Index = RangeSet(0,21)
model.Position_Index = RangeSet(0,81)
model.Pallet_Index = RangeSet(0, len(Pallet_Type) - 1)

# Decision Variables
model.M = Var(model.Main_Deck_Position_Index, model.Pallet_Index, within=Binary, initialize=0)
model.L = Var(model.Lower_Deck_Position_Index, model.Pallet_Index, within=Binary, initialize=0)
model.w = Var(model.Position_Index, domain=NonNegativeIntegers, initialize=0)
model.i = Var(model.Position_Index, domain=NonNegativeReals, initialize=0)

model.W = Var(domain=NonNegativeIntegers)
model.I = Var(domain=NonNegativeReals)

# Binary decision variables that we are going to use for (and-or) or (if-then) constraints
# TODO give domains to these so we won't have millions of binary decision variables
model.y1 = Var(within=Binary)
model.y2 = Var(within=Binary)
model.y3 = Var(within=Binary)
model.y4 = Var(within=Binary)

model.z1 = Var(within=Binary)
model.z2 = Var(within=Binary)
model.z3 = Var(within=Binary)

model.z4 = Var(within=Binary)
model.z5 = Var(within=Binary)

# Objective function
model.obj = Objective(expr=sum(model.M[i,j] * Pallet_Weight.values[j] for i in model.Main_Deck_Position_Index for j in model.Pallet_Index) + sum(model.L[i,j] * Pallet_Weight.values[j] for i in model.Lower_Deck_Position_Index for j in model.Pallet_Index), sense=maximize)
model.constraints = ConstraintList()


# -----------------------------------
#        PART 1: PLACEMENT
# -----------------------------------

# at max 1 pallet to each position

for i in model.Main_Deck_Position_Index:
    model.constraints.add(sum(model.M[i,j] for j in model.Pallet_Index) <= 1)
for i in model.Lower_Deck_Position_Index:
    model.constraints.add(sum(model.L[i,j] for j in model.Pallet_Index) <= 1)


# Weight limit
for i in model.Main_Deck_Position_Index:
    model.constraints.add(sum((model.M[i,j] * Pallet_Weight.values[j]) for j in model.Pallet_Index) <= Max_Weight_M[i])

for i in model.Lower_Deck_Position_Index:
    model.constraints.add(sum((model.L[i,j] * Pallet_Weight.values[j]) for j in model.Pallet_Index) <= Max_Weight_L[i])


# At max 1 pallet to each position  
for j in model.Pallet_Index:
    model.constraints.add(sum(model.M[i,j] for i in model.Main_Deck_Position_Index) + sum(model.L[i,j] for i in model.Lower_Deck_Position_Index) <= 1)


# Pallet type control
# Added binary decision variables y1 y2 y3 y4 to the model for this part
K = 10000


# Main Deck PAG And PMC control
for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        # Check if the position is compatible with PAG pallets
        if Type_M[i] == 0:
            model.constraints.add(model.M[i,j] * Pallet_Type.values[j] <= 0)
        if Type_M[i] == 1:
            model.constraints.add(model.M[i,j] * (1 - Pallet_Type.values[j]) <= 0)
      
# Lower Deck PAG And PMC control
for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        # Check if the position is compatible with PAG pallets
        if Type_L[i] == 0:
            model.constraints.add(model.L[i,j] * Pallet_Type.values[j] <= 0)
        if Type_L[i] == 1:
            model.constraints.add(model.L[i,j] * (1 - Pallet_Type.values[j]) <= 0)


# -----------------------------------
#        PART 2: COLLISION
# -----------------------------------

#Upper Deck
for i in model.Main_Deck_Position_Index:
    for j in model.Main_Deck_Position_Index:
        if(i != j):
            if(Lock1_M[i] <= Lock2_M[j] and Lock2_M[i] >= Lock1_M[j]):
                if (Lock1_M[i] != Lock1_M[i] or Lock2_M[i] != Lock2_M[j]):
                    for k in model.Pallet_Index:
                        model.constraints.add(model.M[i,k] + model.M[j,k]  <= 1 )
             
                        
#Lower Deck
for i in model.Lower_Deck_Position_Index:
    for j in model.Lower_Deck_Position_Index:
        if(i != j):
            if(Lock1_L[i] <= Lock2_L[j] and Lock2_L[i] >= Lock1_L[j]):
                if (Lock1_L[i] != Lock1_L[i] or Lock2_L[i] != Lock2_L[j]):
                    for k in model.Pallet_Index:
                        model.constraints.add(model.L[i,k] + model.L[j,k]  <= 1 )


# -----------------------------------
#        PART 3: CUMULATIVE
# -----------------------------------

# Ensures that cumulative limit for each position in Main deck in front part is not violated
for j in model.Pallet_Index:
    for k in range(34):
        model.constraints.add(sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(34) if H_arm_M.values[i] <= H_arm_M.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12) if H_arm_L[i] <= H_arm_M[k]) <= Cumulative_M[k])

# Ensures that cumulative limit for each position in Lower deck in front part is not violated
for j in model.Pallet_Index:
    for k in range(12):
        model.constraints.add(sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(34) if H_arm_M.values[i] <= H_arm_L.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12) if H_arm_L[i] <= H_arm_L[k]) <= Cumulative_L[k])

# Ensures that cumulative limit for each position in Main deck in aft part is not violated
for j in model.Pallet_Index:
    for k in range(41,60):
        model.constraints.add(sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(41,60) if H_arm_M.values[i] <= H_arm_M.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12,22) if H_arm_L[i] <= H_arm_M[k]) <= Cumulative_M[k])

# Ensures that cumulative limit for each position in Lower deck in aft part is not violated
for j in model.Pallet_Index:
    for k in range(12,22):
        model.constraints.add(sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(41,60) if H_arm_M.values[i] <= H_arm_L.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12,22) if H_arm_L[i] <= H_arm_L[k]) <= Cumulative_L[k])


# -----------------------------------
#        PART 4: BLUE ENVELOPE
# -----------------------------------


# Weight of each position
for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        model.w[i] += model.M[i,j] * Pallet_Weight.values[j]

for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        model.w[i + 60] += (model.L[i,j] * Pallet_Weight.values[j])


# Index of each position
for i in model.Main_Deck_Position_Index:
    model.i[i] = ((H_arm_M.values[i] - 36.3495) * model.w[i]) / 2500
for i in model.Lower_Deck_Position_Index:
    model.i[i + 60] = ((H_arm_L.values[i] - 36.3495) * model.w[i + 60]) / 2500


# Total weight and Total index
model.I = sum(model.i[i] for i in model.Position_Index) + model.DOI
model.W = sum(model.w[i] for i in model.Position_Index) + model.DOW


# Blue envelope constraints
model.constraints.add(2*model.I -model.W <= 240)
model.constraints.add(model.I + model.W >= 235)
model.constraints.add(model.W >= 120)
model.constraints.add(model.W <= 180)


# Solve
solver = SolverFactory('gurobi')
solver.solve(model)
#display(model)

#Calculate the CPU time
cpu_time = time.time() - start_time


#To print values with their code names
print("Optimal Assignment:")

for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.M[i,j]) == 1:
            print("Pallet ", OriginalPallets['Code'][j] , "is assigned to Position ", Positions_M['Position'][i])

for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.L[i,j]) == 1:
            print("Pallet ", OriginalPallets['Code'][j] , "is assigned to Position ", Positions_L['Position'][i])
print("\nTotal weight:", value(model.obj))

'''
print("Optimal Assignment:")

for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.M[i,j]) == 1:
            print(f"Pallet {j} is assigned to Position {i}")
for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.L[i,j]) == 1:
            print(f"Pallet {j} is assigned to Position {i}")
print("\nTotal weight:", value(model.obj))
'''

#Print the CPU time
print("CPU Time:", cpu_time, "seconds")    


