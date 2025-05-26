# REQUIREMENTS: Python 2.7.18 32-bit, PSS/E 34.0
# This script automates the process of running a dynamic simulation in PSS/E using Python.
# It sets up the environment, loads a case, runs a dynamic simulation, and saves the results in an Excel file.

import os
import psse_utility_functions as psse_utils # see this Python file for utility functions

# File paths and names
SAV_LOCATION = r"C:\\Users\\bmcgo\\OneDrive\\Documents\\work related\\job search - oct 2024\\self-learning\\power system studies\\smib"
SAV_FILE = r"smib-basecase.sav"
CONV_FILE = r"smib-transient-stability-converted.sav"
DYR_FILE = r"smib-dynamic-model.dyr"
OUT_FILE = r"smib-dynamic-result.out"

# Open save case file
os.chdir(SAV_LOCATION)
psse_utils.open_sav(SAV_FILE)

# Run load flow analysis
converged = psse_utils.run_load_flow()
# Error handling for load flow convergence
if not converged:
    raise RuntimeError("Load flow did not converge. Please check the case setup.")

# Prepare for dynamic simulation
# Convert case for switching studies, model generators as classical, and set output channels
psse_utils.convert_for_switching_studies(CONV_FILE)
psse_utils.model_generators_classical(DYR_FILE, [3, 3], [1, 1])
psse_utils.set_output_channels()

# Run dynamic simulation and export results to an Excel file
psse_utils.run_dynamic_simulation(OUT_FILE)
psse_utils.export_results_to_excel(OUT_FILE)
