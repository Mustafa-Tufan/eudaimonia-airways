import time
from pyomo.environ import *
import pandas as pd

# Reads The Sheet
excel_file = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Positions')
excel_file2 = pd.read_excel('END395_ProjectPartIIDataset.xlsx', sheet_name='Position-CG')

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

# Divides The Sheet Into Main And Lower Deck. Then Sorts The Sheets By H-arm Values.
sorted_excel_file_M = excel_file.iloc[:60].sort_values(by='H-arm', ascending=True)
sorted_excel_file_L = excel_file.iloc[60:].sort_values(by='H-arm', ascending=True)

# Extracts Lower Deck Positions from Dataset II
slice1 = excel_file2.loc[2:3]
slice2 = excel_file2.loc[10:11]
slice3 = excel_file2.loc[18:19]
slice4 = excel_file2.loc[26:27]
slice5 = excel_file2.loc[30:31]
slice6 = excel_file2.loc[38:39]
slice7 = excel_file2.loc[67:68]
slice8 = excel_file2.loc[75:76]
slice9 = excel_file2.loc[83:84]
slice10 = excel_file2.loc[90:91]
slice11 = excel_file2.loc[94:95]

sliced_excel_file_L = pd.concat([slice1, slice2, slice3, slice4, slice5, slice6, slice7, slice8, slice9, slice10, slice11], ignore_index=True)
sliced_excel_file_L.reset_index(drop=True, inplace=True)

# Extracts Main Deck Positions from Dataset II
excel_file2.drop(slice1.index, inplace=True)
excel_file2.drop(slice2.index, inplace=True)
excel_file2.drop(slice3.index, inplace=True)
excel_file2.drop(slice4.index, inplace=True)
excel_file2.drop(slice5.index, inplace=True)
excel_file2.drop(slice6.index, inplace=True)
excel_file2.drop(slice7.index, inplace=True)
excel_file2.drop(slice8.index, inplace=True)
excel_file2.drop(slice9.index, inplace=True)
excel_file2.drop(slice10.index, inplace=True)
excel_file2.drop(slice11.index, inplace=True)

excel_file2.drop(excel_file2.loc[4:5].index, inplace=True)
excel_file2.drop(excel_file2.loc[12:13].index, inplace=True)
excel_file2.drop(excel_file2.loc[21:22].index, inplace=True)
excel_file2.drop(excel_file2.loc[32:33].index, inplace=True)
excel_file2.drop(excel_file2.loc[40:41].index, inplace=True)
excel_file2.drop(excel_file2.loc[47:48].index, inplace=True)
excel_file2.drop(excel_file2.loc[53:54].index, inplace=True)
excel_file2.drop(excel_file2.loc[61:62].index, inplace=True)
excel_file2.drop(excel_file2.loc[69:70].index, inplace=True)
excel_file2.drop(excel_file2.loc[77:78].index, inplace=True)
excel_file2.drop(excel_file2.loc[86:87].index, inplace=True)

sliced_excel_file_M = excel_file2
sliced_excel_file_M.reset_index(inplace = True)

#sliced_excel_file_M.to_excel('2_M.xlsx', index=False)
#sliced_excel_file_L.to_excel('2_L.xlsx', index=False)

results = []
outputs = [[],[],[],[]]

for cg_interval in [1,0,2,3]:
    try:
        # -----------------------------------
        #            PARAMETERS
        # -----------------------------------

        # Assigns The Main And Lower Deck Positions To Corresponded DataFrame
        Positions_M = sorted_excel_file_M.reset_index(drop=True)
        Positions_L = sorted_excel_file_L.reset_index(drop=True)

        #Getting The Main Deck Parameters
        Lock1_M = Positions_M.iloc[:,1]
        Lock2_M = Positions_M.iloc[:,2]
        H_arm_M = Positions_M.iloc[:,3]
        Max_Weight_M = Positions_M.iloc[:,4]   #Stores the weight limits for positions of Main deck
        #Cumulative_M = Positions_M.iloc[:,5]
        Coefficient_M = Positions_M.iloc[:,6]
        Type_M = Positions_M.iloc[:,7]   #Stores the Type (PAG or PMC) of positions in the Main deck

        #Getting The Lower Deck Parameters
        Lock1_L = Positions_L.iloc[:,1]
        Lock2_L = Positions_L.iloc[:,2]
        H_arm_L = Positions_L.iloc[:,3]
        Max_Weight_L = Positions_L.iloc[:,4]   #Stores the weight limits for positions of Lower deck
        #Cumulative_L = Positions_L.iloc[:,5]
        Coefficient_L = Positions_L.iloc[:,6]
        Type_L = Positions_L.iloc[:,7]   #Stores the Type (PAG or PMC) of positions in the Lower deck
 
        Cumulative_M = sliced_excel_file_M.iloc[:,cg_interval+2]
        Cumulative_L = sliced_excel_file_L.iloc[:,cg_interval+2]

        # Take the Pallets' data
        Pallets = pd.read_excel('END395_ProjectPartIIDataset.xlsx', sheet_name='Pallets4')
        OriginalPallets = Pallets.copy()

        Pallet_Weight = Pallets.iloc[:,2]
        Pallet_Type = Pallets.iloc[:,1]

        # Assigns 0 if Pallet's Type is PAG and Assigns 1 if Pallet's Type is PAG. 
        for i in range(len(Pallet_Type)):
            if Pallet_Type.values[i].startswith("PAG"):
                Pallet_Type.values[i] = 0
            else:
                Pallet_Type.values[i] = 1

        #Takes DOW and DOI Values.
        DOW = Pallets.columns[5]
        DOI = Pallets.iloc[0,5]

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

        Weight = [0] * len(model.Position_Index)
        Index = [0] * len(model.Position_Index)

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

        # -----------------------------------
        #         OBJECTIVE FUNCTION
        # -----------------------------------
        model.obj = Objective(expr=sum(model.M[i,j] * Pallet_Weight.values[j] for i in model.Main_Deck_Position_Index for j in model.Pallet_Index) + sum(model.L[i,j] * Pallet_Weight.values[j] for i in model.Lower_Deck_Position_Index for j in model.Pallet_Index), sense=maximize)

        # -----------------------------------
        #            CONSTRAINTS
        # -----------------------------------
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

        # Calculates Weights And Indices Of Each Position.
        for i in model.Main_Deck_Position_Index:
            Weight[i] = sum(model.M[i,j] * Pallet_Weight.values[j] for j in model.Pallet_Index)
            Index[i] = sum(model.M[i,j] * (((H_arm_M[i] - 36.3495) * Pallet_Weight.values[j]) / 2500) for j in model.Pallet_Index)
        for i in model.Lower_Deck_Position_Index:
            Weight[i+60] = sum(model.L[i,j] * Pallet_Weight.values[j] for j in model.Pallet_Index)
            Index[i+60] = sum(model.L[i,j] * (((H_arm_L[i] - 36.3495) * Pallet_Weight.values[j]) / 2500) for j in model.Pallet_Index)


        # Calculates W
        #model.constraints.add(model.W >=  sum(Weight[i] for i in model.Position_Index) + DOW)
        #model.constraints.add(model.W <=  sum(Weight[i] for i in model.Position_Index) + DOW)
        model.constraints.add(model.W ==  sum(Weight[i] for i in model.Position_Index) + DOW)

        # Calculates I
        #model.constraints.add(model.I >=  sum(Index[i] for i in model.Position_Index) + DOI)
        #model.constraints.add(model.I <=  sum(Index[i] for i in model.Position_Index) + DOI)
        model.constraints.add(model.I ==  sum(Index[i] for i in model.Position_Index) + DOI)


        # Ensures That W and I Are In The Blue Envelope.
        model.constraints.add(2 * model.I - (model.W/1000) <= 240)
        model.constraints.add(model.I + (model.W/1000) >= 235)
        model.constraints.add((model.W/1000) >= 120)
        model.constraints.add((model.W/1000) <= 180)

        Coefficient_M = Positions_M.iloc[:,6]

        # Solves The Model
        solver = SolverFactory('cplex')
        solver.solve(model)
        results.append((cg_interval,value(model.obj)))
        # Assigning PAG/PMC To Main Deck Side-By-Side Positions According To Placed Pallet Types.
        for i in model.Main_Deck_Position_Index:
            if len(Positions_M['Position'][i]) == 7:
                for j in model.Pallet_Index:
                    if value(model.M[i,j]) == 1:
                        if Pallet_Type[j] == 1:
                            Positions_M.loc[i, 'Position'] = Positions_M.loc[i, 'Position'].replace("88", "96")
                                        
        output = []
        # Prints The Optinmal Assignments
        output += ["\nOptimal Assignment:"]

        for i in model.Main_Deck_Position_Index:
            for j in model.Pallet_Index:
                if value(model.M[i,j]) == 1:
                    output += [f"Pallet {OriginalPallets['Code'][j].ljust(5)} is assigned to Position {Positions_M['Position'][i]}"]
        for i in model.Lower_Deck_Position_Index:
            for j in model.Pallet_Index:
                if value(model.L[i,j]) == 1:
                    output += [f"Pallet {OriginalPallets['Code'][j].ljust(5)} is assigned to Position {Positions_L['Position'][i]}"]
                    
        output += [f"\nObjective Function Value (Total Weight):{value(model.obj)}"]
        
        outputs[cg_interval] += output
        
    except Exception as e: 
        print(f"CG Interval {cg_interval+1} is Infeasible")


# Calculates the CPU time
cpu_time = time.time() - start_time


# Find the tuple with the maximum second element
max_tuple = max(results, key=lambda x: x[1])
print(f"{max_tuple[0]+1} is optimal. Obj: {max_tuple[1]}")
for line in outputs[max_tuple[0]]:
    print(line)


#Print the CPU time
print("CPU Time:", cpu_time, "seconds")
