import time
from pyomo.environ import *
import pandas as pd

# Read the sheet
excel_file = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Positions')

# Add new column for position types
excel_file.loc[0:20, 'Type'] = 0
excel_file.loc[20:60, 'Type'] = 1
excel_file.loc[60:93, 'Type'] = 0
excel_file.loc[93:104, 'Type'] = 1

# Sort the sheet by H-arm's and create new one
sorted_excel_file_M = excel_file.iloc[:82].sort_values(by='H-arm', ascending=True)
sorted_excel_file_L = excel_file.iloc[82:].sort_values(by='H-arm', ascending=True)

sorted_excel_file_M.to_excel('END395_ProjectPartIDataset_SORTED_M.xlsx', index=False)
sorted_excel_file_L.to_excel('END395_ProjectPartIDataset_SORTED_L.xlsx', index=False)

Positions_M = pd.read_excel('END395_ProjectPartIDataset_SORTED_M.xlsx')
Positions_L = pd.read_excel('END395_ProjectPartIDataset_SORTED_L.xlsx')

Lock1_M = Positions_M.iloc[:,1]
Lock2_M = Positions_M.iloc[:,2]
H_arm_M = Positions_M.iloc[:,3]
Max_Weight_M = Positions_M.iloc[:,4]
Cumulative_M = Positions_M.iloc[:,5]
Coefficient_M = Positions_M.iloc[:,6]
Type_M = Positions_M.iloc[:,7]

Lock1_L = Positions_L.iloc[:,1]
Lock2_L = Positions_L.iloc[:,2]
H_arm_L = Positions_L.iloc[:,3]
Max_Weight_L = Positions_L.iloc[:,4]
Cumulative_L = Positions_L.iloc[:,5]
Coefficient_L = Positions_L.iloc[:,6]
Type_L = Positions_L.iloc[:,7]

# Take the pallet dat
Pallets = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Pallets1')

# TODO Pallet Type'yi integer olarak çek
Pallet_Type = Pallets.iloc[:,1]
for i in range(len(Pallet_Type)):
    if Pallet_Type.values[i].startswith("PAG"):
        Pallet_Type.values[i] = 0
    else:
        Pallet_Type.values[i] = 1
Pallet_Weight = Pallets.iloc[:,2]
DOW = Pallets.columns[4] #TRICKYYYYYY
DOI = Pallets.iloc[0,4]

model = ConcreteModel()
start_time = time.time()

# Initialize ranges
model.Main_Deck_Position_Index = RangeSet(0,81)
model.Lower_Deck_Position_Index = RangeSet(0,21)
model.Position_Index = RangeSet(0,103)
model.Pallet_Index = RangeSet(0, len(Pallet_Type) - 1)

# Decision Variables
model.M = Var(model.Main_Deck_Position_Index, model.Pallet_Index, within=Binary)
model.L = Var(model.Lower_Deck_Position_Index, model.Pallet_Index, within=Binary)
model.w = Var(model.Position_Index, domain=PositiveIntegers)
model.i = Var(model.Position_Index, domain=PositiveIntegers)
model.W = 0
model.I = 0

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
model.obj = Objective(expr=sum(model.w[i] for i in model.Position_Index), sense=maximize)

model.constraints = ConstraintList()

# -----------------------------------
#        PART 1: PLACEMENT
# -----------------------------------

# at max 1 pallet to each position
# Do we really need those? 
for i in model.Main_Deck_Position_Index:
    model.constraints.add(sum(model.M[(i,j)] for j in model.Pallet_Index) <= 1)
for i in model.Lower_Deck_Position_Index:
    model.constraints.add(sum(model.L[(i,j)] for j in model.Pallet_Index) <= 1)

# Weight limit
for i in model.Main_Deck_Position_Index:
    model.constraints.add(sum(model.M[(i,j)] * Pallet_Weight.values[j] for j in model.Pallet_Index) <= Max_Weight_M[i])
for i in model.Lower_Deck_Position_Index:
    model.constraints.add(sum(model.L[(i,j)] * Pallet_Weight.values[j] for j in model.Pallet_Index) <= Max_Weight_L[i])
    
# At max 1 pallet to each position  
for j in model.Pallet_Index:
    model.constraints.add(sum(model.M[(i,j)] for j in model.Pallet_Index) + sum(model.L[(i,j)] for j in model.Pallet_Index) <= 1)
    
# Pallet type control
# Added binary decision variables y1 y2 y3 y4 to the model for this part
K = 10000

# Main PAG Control
for i in model.Main_Deck_Position_Index:
    model.constraints.add(1-Type_M.values[i] <= K * model.y1)
for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        model.constraints.add(model.M[(i,j)] * Pallet_Type.values[j] <= K * (1 - model.y1))
        
# Lower PAG Control
for i in model.Lower_Deck_Position_Index:
    model.constraints.add(1-Type_L.values[i] <= K * model.y2)
for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        model.constraints.add(model.L[(i,j)] * Pallet_Type.values[j] <= K * (1 - model.y2))

# Main PMC Control
for i in model.Main_Deck_Position_Index:
    model.constraints.add(Type_M.values[i] <= K * model.y3)
for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        model.constraints.add(model.M[(i,j)] * (1 - Pallet_Type.values[j]) <= K * (1 - model.y3))

# Lower PMC Control
for i in model.Lower_Deck_Position_Index:
    model.constraints.add(Type_L.values[i] <= K * model.y4)
for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        model.constraints.add(model.L[(i,j)] * (1 - Pallet_Type.values[j]) <= K * (1 - model.y4))   
# -----------------------------------
#        PART 2: COLLIDING
# -----------------------------------

# Main deck collision handling method
# (Lock1 and Lock2 of pallet j is lower than Lock1 and Lock2 of pallet n on main deck)

for i in model.Main_Deck_Position_Index:
    for k in model.Main_Deck_Position_Index:
        if k != i:          
            model.constraints.add(sum(model.M[(i,j)] * Lock2_M.values[i] for j in model.Pallet_Index) - sum(model.M[(k,n)] * Lock1_M.values[k] for n in model.Pallet_Index) <= K * (1 - model.z1))                       
            model.constraints.add(sum(model.M[(i,j)] * Lock2_M.values[i] for j in model.Pallet_Index) - sum(model.M[(k,n)] * Lock2_M.values[k] for n in model.Pallet_Index) <= K * (1 - model.z1))            
            model.constraints.add(sum(model.M[(k,n)] * Lock1_M.values[k] for n in model.Pallet_Index) - sum(model.M[(i,j)] * Lock1_M.values[i] for j in model.Pallet_Index) <= K * (1 - model.z2))
            model.constraints.add(sum(model.M[(k,n)] * Lock1_M.values[k] for n in model.Pallet_Index) - sum(model.M[(i,j)] * Lock2_M.values[i] for j in model.Pallet_Index) <= K * (1 - model.z2))            
            model.constraints.add(sum(model.M[(k,n)] * Lock1_M.values[k] for n in model.Pallet_Index) - sum(model.M[(i,j)] * Lock1_M.values[i] for j in model.Pallet_Index) == K * (1 - model.z3))            
            model.constraints.add(sum(model.M[(k,n)] * Lock2_M.values[k] for n in model.Pallet_Index) - sum(model.M[(i,j)] * Lock2_M.values[i] for j in model.Pallet_Index) == K * (1 - model.z3))
    print("1")
    
# At least one of the above should be satisfied
model.constraints.add(model.z1 + model.z2 + model.z3 >= 1)            
                       
# Lower deck collision handling method
# (Lock1 and Lock2 of pallet j is lower than Lock1 and Lock2 of pallet n on lower deck)
for i in model.Lower_Deck_Position_Index:
    for k in model.Lower_Deck_Position_Index:
        if k != i:           
            model.constraints.add(sum(model.L[(i,j)] * Lock2_L.values[i] for j in model.Pallet_Index) - sum(model.L[(k,n)] * Lock1_L.values[k] for n in model.Pallet_Index) <= K * (1 - model.z4))                       
            model.constraints.add(sum(model.L[(i,j)] * Lock2_L.values[i] for j in model.Pallet_Index) - sum(model.M[(k,n)] * Lock2_L.values[k] for n in model.Pallet_Index) <= K * (1 - model.z4))            
            model.constraints.add(sum(model.L[(k,n)] * Lock1_L.values[k] for n in model.Pallet_Index) - sum(model.L[(i,j)] * Lock1_L.values[i] for j in model.Pallet_Index) <= K * (1 - model.z5))
            model.constraints.add(sum(model.L[(k,n)] * Lock1_L.values[k] for n in model.Pallet_Index) - sum(model.L[(i,j)] * Lock2_L.values[i] for j in model.Pallet_Index) <= K * (1 - model.z5))
# At least one of the above should be satisfied
model.constraints.add(model.z4 + model.z5 >= 1)
# -----------------------------------
#        PART 3: CUMULATIVE
# -----------------------------------
print("3")
# Ensures that cumulative limit for each position in Main deck in front part is not violated
for j in model.Pallet_Index:
    for k in range(34):
        model.constraints.add(sum(model.M[(i,j)] * model.Pallet_Weight[j] * Coefficient_M.values[i] for i in range(k) if H_arm_M.values[i]) + sum(H_arm_M.values[k] + model.L[(i,j)] * model.Pallet_Weight[j] * Coefficient_L.values[i] for i in range(12) if H_arm_L[i] <= H_arm_M[k]) <= Cumulative_M[k])

# Ensures that cumulative limit for each position in Lower deck in front part is not violated
for j in model.Pallet_Index:
    for k in range(12):
        model.constraints.add(sum(model.M[(i,j)] * model.Pallet_Weight[j] * Coefficient_M.values[i] for i in range(34) if H_arm_M.values[i] <= H_arm_L.values[k]) + sum(model.L[(i,j)] * model.Pallet_Weight[j] * Coefficient_L.values[i] for i in range(k) if H_arm_L[i] <= H_arm_L[k]) <= Cumulative_L[k])

# Ensures that cumulative limit for each position in Main deck in aft part is not violated
for j in model.Pallet_Index:
    for k in range(41,82):
        model.constraints.add(sum(model.M[(i,j)] * model.Pallet_Weight[j] * Coefficient_M.values[i] for i in range(k,82) if H_arm_M.values[i] <= H_arm_M.values[k]) + sum(model.L[(i,j)] * model.Pallet_Weight[j] * Coefficient_L.values[i] for i in (12,22) if H_arm_L[i] <= H_arm_M[k]) <= Cumulative_M[k])

# Ensures that cumulative limit for each position in Lower deck in aft part is not violated
for j in model.Pallet_Index:
    for k in range(12,22):
        model.constraints.add(sum(model.M[(i,j)] * model.Pallet_Weight[j] * Coefficient_M.values[i] for i in range(41,82) if H_arm_M.values[i] <= H_arm_L.values[k]) + sum(model.L[(i,j)] * model.Pallet_Weight[j] * Coefficient_L.values[i] for i in (k,22) if H_arm_L[i] <= H_arm_L[k]) <= Cumulative_L[k])
print("4")
# -----------------------------------
#        PART 4: BLUE ENVELOPE
# -----------------------------------

# Weight of each position
for i in model.Main_Deck_Position_Index:
    model.w[i] = sum(model.M[(i,j)] * Pallet_Weight.values[j] for j in model.Pallet_Index)
for i in model.Lower_Deck_Position_Index:
    model.w[i + 82] = sum(model.L[(i,j)] * Pallet_Weight.values[j] for j in model.Pallet_Index)

# Index of each position
for i in model.Main_Deck_Position_Index:
    model.i[i] = (H_arm_M.values[i] - 36.3495) * model.w[i] / 2500
for i in model.Lower_Deck_Position_Index:
    model.i[i + 82] = (H_arm_L.values[i] - 36.3495) * model.w[i + 82] / 2500
print("5")
# Total weight and Total index
model.W = sum(model.w[i] + model.DOW)
model.I = sum(model.i[i] + model.DOı)

# Blue envelope constraints
model.constraints.add(2*model.I - model.W <= 240)
model.constraints.add(model.I + model.W >= 235)
model.constraints.add(model.W >= 120)
model.constraints.add(model.W <= 180)

# Solve
solver = SolverFactory('cplex_direct')
solver.solve(model)
display(model)

#Calculate the CPU time
cpu_time = time.time() - start_time

print("Optimal Assignment:")
for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.M[(i,j)]) > 0:
            print(f"Pallet {i} is assigned to Position {j}")
for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.L[(i,j)]) > 0:
            print(f"Pallet {i} is assigned to Position {j}")
print("\nTotal weight:", value(model.obj))

#Print the CPU time
print("CPU Time:", cpu_time, "seconds")    
