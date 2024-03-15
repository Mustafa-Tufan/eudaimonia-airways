from pyomo.environ import *
import pandas as pd

# Take the pallet dat
Pallets = pd.read_excel('END395_ProjectPartIDataset.xlsx', sheet_name='Pallets1')

model = ConcreteModel()
model.constraints = ConstraintList()

# Initialize ranges
model.Main_Deck_Position_Index = RangeSet(0,81)
model.Pallet_Index = RangeSet(0,31)

# Decision Variables
model.M = Var(model.Main_Deck_Position_Index, model.Pallet_Index, within=Binary)

for i in model.Main_Deck_Position_Index:
    model.constraints.add((model.M[i,0]) <= 1)
    model.constraints.add((model.M[i,0]) >= 1)

model.obj = Objective(expr=sum(model.M[i,j] for i in model.Main_Deck_Position_Index for j in model.Pallet_Index), sense=maximize)

# Solve
solver = SolverFactory('cplex_direct')
solver.solve(model)
display(model)

print("\nTotal weight:", value(model.obj))
