# psse_python
Python code for running simulations in PSS/E

# Python version
The version of psspy I'm using here requires Python 2 (I use Python 2.7.18).
The plotting functions in plot_psse_output.py, however, use Python 3.

# Utility functions
psse_utility_functions.py contains utility functions for PSS/E-Python and should be imported at the start of a simulation workflow. It imports the psspy library from its file location and provides various functions for running power flow and dynamic simulations in PSS/E.

# Dynamic simulation workflow
psse_dynamic_simulation_workflow.py runs a dynamic simulation for a single machine, infinite bus system as described in https://bit.ly/3FqbfPZ.
In the future, this workflow will be expanded for a more general case.

# Plotting PSS/E output
plot_psse_output.py plots the dynamic simulation output variables from an Excel file. Currently, the generator angle, electrical power, and terminal voltage are plotted, but in the future, this file will be expanded for a more general case.