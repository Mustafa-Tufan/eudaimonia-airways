import time
from pyomo.environ import *
import pandas as pd

# Reads The Sheet
excel_file = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Positions')

# Adds New Column For Position Types. 0 for PAG, 1 for PMC, and 2 for both PAG PMC
excel_file.loc[0:19, 'Type'] = 0
excel_file.loc[20:37, 'Type'] = 1
excel_file.loc[38:81, 'Type'] = 2
excel_file.loc[82:92, 'Type'] = 0
excel_file.loc[93:103, 'Type'] = 1

# Drops Same Rows (e.g. FHR(88) and FHR(96) belongs to same place in plane)  
original_excel_file = excel_file.copy()
excel_file.drop(labels = range(38, 60), axis=0, inplace=True)
excel_file.reset_index(inplace = True)
excel_file.drop(labels = "index", axis = 1, inplace = True)

# Divides sheet into Main and Lower deck. Then sorts the sheets by H-arm values.
sorted_excel_file_M = excel_file.iloc[:60].sort_values(by='H-arm', ascending=True)
sorted_excel_file_L = excel_file.iloc[60:].sort_values(by='H-arm', ascending=True)

# -----------------------------------
#            PARAMETERS
# -----------------------------------

# Assigns the main and lower deck positions to corresponded dataframe
Positions_M = sorted_excel_file_M.reset_index(drop=True)
Positions_L = sorted_excel_file_L.reset_index(drop=True)

#Getting The Main Deck Parameters
Lock1_M = Positions_M.iloc[:,1]
Lock2_M = Positions_M.iloc[:,2]
H_arm_M = Positions_M.iloc[:,3]
Max_Weight_M = Positions_M.iloc[:,4]   #Stores the weight limits for positions of Main deck
Cumulative_M = Positions_M.iloc[:,5]
Coefficient_M = Positions_M.iloc[:,6]
Type_M = Positions_M.iloc[:,7]   #Stores the Type (PAG or PMC) of positions in the Main deck

#Getting The Lower Deck Parameters
Lock1_L = Positions_L.iloc[:,1]
Lock2_L = Positions_L.iloc[:,2]
H_arm_L = Positions_L.iloc[:,3]
Max_Weight_L = Positions_L.iloc[:,4]   #Stores the weight limits for positions of Lower deck
Cumulative_L = Positions_L.iloc[:,5]
Coefficient_L = Positions_L.iloc[:,6]
Type_L = Positions_L.iloc[:,7]   #Stores the Type (PAG or PMC) of positions in the Lower deck

# Take the Pallets' data
Pallets = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Pallets1')
OriginalPallets = Pallets.copy()

Pallet_Weight = Pallets.iloc[:,2]
Pallet_Type = Pallets.iloc[:,1]

# Assigns 0 if Pallet's Type is PAG and Assigns 1 if Pallet's Type is PAG and 
for i in range(len(Pallet_Type)):
    if Pallet_Type.values[i].startswith("PAG"):
        Pallet_Type.values[i] = 0
    else:
        Pallet_Type.values[i] = 1


DOW = Pallets.columns[4]
DOI = Pallets.iloc[0,4]

model = ConcreteModel()
start_time = time.time()

# Initializes Ranges
model.Main_Deck_Position_Index = RangeSet(0,59)
model.Lower_Deck_Position_Index = RangeSet(0,21)
model.Position_Index = RangeSet(0,81)
model.Pallet_Index = RangeSet(0, len(Pallet_Type) - 1)

# -----------------------------------
#        DECISION VARIABLES
# -----------------------------------

model.M = Var(model.Main_Deck_Position_Index, model.Pallet_Index, within=Binary)
model.L = Var(model.Lower_Deck_Position_Index, model.Pallet_Index, within=Binary)

model.W = Var(domain=NonNegativeIntegers)
model.I = Var(domain=Reals)

weight = [0] * len(model.Position_Index)
index = [0] * len(model.Position_Index)

# Binary decision variables for (either-or) and (if-then) constraints
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

# Ensures That At Max 1 Pallet Could Be Assigned To Each Position.
for i in model.Main_Deck_Position_Index:
    model.constraints.add(sum(model.M[i,j] for j in model.Pallet_Index) <= 1)
for i in model.Lower_Deck_Position_Index:
    model.constraints.add(sum(model.L[i,j] for j in model.Pallet_Index) <= 1)


# Ensures That Pallets Do Not Exceed The Weight Limit For Each Position.
for i in model.Main_Deck_Position_Index:
    model.constraints.add(sum((model.M[i,j] * Pallet_Weight.values[j]) for j in model.Pallet_Index) <= Max_Weight_M[i])
for i in model.Lower_Deck_Position_Index:
    model.constraints.add(sum((model.L[i,j] * Pallet_Weight.values[j]) for j in model.Pallet_Index) <= Max_Weight_L[i])


# Ensures That A Pallet Could Be Assigned Only One Positions
for j in model.Pallet_Index:
    model.constraints.add(sum(model.M[i,j] for i in model.Main_Deck_Position_Index) + sum(model.L[i,j] for i in model.Lower_Deck_Position_Index) <= 1)


# This Part Ensures That PAG Pallets Assigned To PAG Places and PMC Pallets Assigned To PMC Places. (Uses y1, y2, y3 and y4 as binary variables)

# Main Deck PAG And PMC control
for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        # Check if the position is compatible with PAG pallets
        if Type_M[i] == 0:
            model.constraints.add(model.M[i,j] * Pallet_Type.values[j] <= 0)
        # Check if the position is compatible with PMC pallets
        if Type_M[i] == 1:
            model.constraints.add(model.M[i,j] * (1 - Pallet_Type.values[j]) <= 0)

# Lower Deck PAG And PMC control
for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        # Check if the position is compatible with PAG pallets
        if Type_L[i] == 0:
            model.constraints.add(model.L[i,j] * Pallet_Type.values[j] <= 0)
        # Check if the position is compatible with PMC pallets
        if Type_L[i] == 1:
            model.constraints.add(model.L[i,j] * (1 - Pallet_Type.values[j]) <= 0)
            
# -----------------------------------
#        PART 2: COLLISION
# -----------------------------------

# Main Deck Collision Checking
for i in range(0, len(model.Main_Deck_Position_Index) - 1):
    for j in range(i + 1, len(model.Main_Deck_Position_Index)):
        #If positions are side by side. Pass.
        if Lock1_M[i] == Lock1_M[j] and Lock2_M[i] == Lock2_M[j]:
            pass
        #If one position located in another positions inside. Adds constraint.
        elif Lock1_M[i] <= Lock2_M[j] and Lock2_M[i] >= Lock1_M[j]:
            for k in model.Pallet_Index:
                for l in model.Pallet_Index:
                    if k != l:  
                        model.constraints.add(model.M[i,k] + model.M[j,l] <= 1)

# Lower Deck Collision Checking
for i in range(0, len(model.Lower_Deck_Position_Index) - 1):
    for j in range(i + 1, len(model.Lower_Deck_Position_Index)):
        #If one position located in another positions inside. Adds constraint.
        if Lock1_L[i] <= Lock2_L[j] and Lock2_L[i] >= Lock1_L[j]:
            for k in model.Pallet_Index:
                for l in model.Pallet_Index:
                    if k != l:
                        model.constraints.add(model.L[i,k] + model.L[j,l] <= 1)
                         
# -----------------------------------
#        PART 3: CUMULATIVE
# -----------------------------------

# Ensures that cumulative limit for each position in Main deck in front part is not violated
for k in range(24):
    model.constraints.add(sum((sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(24) if H_arm_M.values[i] <= H_arm_M.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12) if H_arm_L.values[i] <= H_arm_M.values[k])) for j in model.Pallet_Index) <=  Cumulative_M[k])

# Ensures that cumulative limit for each position in Lower deck in front part is not violated
for k in range(12):
    model.constraints.add(sum((sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(24) if H_arm_M.values[i] <= H_arm_L.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12) if H_arm_L.values[i] <= H_arm_L.values[k])) for j in model.Pallet_Index ) <= Cumulative_L[k])

# Ensures that cumulative limit for each position in Main deck in aft part is not violated
for k in range(29,60):
    model.constraints.add(sum((sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(29,60) if H_arm_M.values[i] >= H_arm_M.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12,22) if H_arm_L.values[i] >= H_arm_M.values[k])) for j in model.Pallet_Index) <= Cumulative_M[k])

# Ensures that cumulative limit for each position in Lower deck in aft part is not violated
for k in range(12,22):
    model.constraints.add(sum((sum(model.M[i,j] * Pallet_Weight.values[j] * Coefficient_M.values[i] for i in range(29,60) if H_arm_M.values[i] >= H_arm_L.values[k]) + sum(model.L[i,j] * Pallet_Weight.values[j] * Coefficient_L.values[i] for i in range(12,22) if H_arm_L.values[i] >= H_arm_L.values[k])) for j in model.Pallet_Index) <= Cumulative_L[k])

# -----------------------------------
#        PART 4: BLUE ENVELOPE
# -----------------------------------

# Calculates Weights Of Each Position.
for i in model.Main_Deck_Position_Index:
    weight[i] = sum(model.M[i,j] * Pallet_Weight.values[j] for j in model.Pallet_Index)
for i in model.Lower_Deck_Position_Index:
    weight[i+60] = sum(model.L[i,j] * Pallet_Weight.values[j] for j in model.Pallet_Index)

# Calculates Indices Of Each Position.
for i in model.Main_Deck_Position_Index:
    index[i] = sum(model.M[i,j] * (((H_arm_M[i] - 36.3495) * Pallet_Weight.values[j]) / 2500) for j in model.Pallet_Index)
for i in model.Lower_Deck_Position_Index:
    index[i+60] = sum(model.L[i,j] * (((H_arm_L[i] - 36.3495) * Pallet_Weight.values[j]) / 2500) for j in model.Pallet_Index)


# Calculates W
model.constraints.add(model.W >=  sum(weight[i] for i in model.Position_Index) + DOW)
model.constraints.add(model.W <=  sum(weight[i] for i in model.Position_Index) + DOW)

# Calculates I
model.constraints.add(model.I >=  sum(index[i] for i in model.Position_Index) + DOI)
model.constraints.add(model.I <=  sum(index[i] for i in model.Position_Index) + DOI)


# Ensures That W and I Are In Blue Envelope.
model.constraints.add(2 * model.I - (model.W/1000) <= 240)
model.constraints.add(model.I + (model.W/1000) >= 235)
model.constraints.add((model.W/1000) >= 120)
model.constraints.add((model.W/1000) <= 180)

# Solves The Model
solver = SolverFactory('gurobi')
solver.solve(model)
    
# Calculates the CPU time
cpu_time = time.time() - start_time

# Assigning PAG/PMC To Main Deck Side-By-Side Positions According To Placed Pallet Types.
for i in model.Main_Deck_Position_Index:
    if len(Positions_M['Position'][i]) == 7:
        for j in model.Pallet_Index:
            if value(model.M[i,j]) == 1:
                if Pallet_Type[j] == 1:
                    Positions_M.loc[i, 'Position'] = Positions_M.loc[i, 'Position'].replace("88", "96")
                                

# Prints The Optinmal Assignments
print("\nOptimal Assignment:")

for i in model.Main_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.M[i,j]) == 1:
            print(f"Pallet {OriginalPallets['Code'][j].ljust(5)} is assigned to Position {Positions_M['Position'][i]}")
            #print("Pallet ", OriginalPallets['Code'][j] , "is assigned to Position ", Positions_M['Position'][i])
for i in model.Lower_Deck_Position_Index:
    for j in model.Pallet_Index:
        if value(model.L[i,j]) == 1:
            print(f"Pallet {OriginalPallets['Code'][j].ljust(5)} is assigned to Position {Positions_L['Position'][i]}")
            
print("\nObjective Function Value (Total Weight):", value(model.obj))

#Print the CPU time
print("CPU Time:", cpu_time, "seconds")    
