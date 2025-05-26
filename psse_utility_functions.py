# REQUIREMENTS: Python 2.7.18 32-bit, PSS/E 34.0
# Imports psspy from file location and provides utility functions for PSS/E power flow and dynamic simulations.
import os, sys

PSSE_LOCATION = r"C:\\Program Files (x86)\\PTI\\PSSEXplore34\\PSSPY27"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' + PSSE_LOCATION
import pssexplore34
import psspy
import dyntools

psspy.psseinit() # Initialize PSSE Python application

def open_sav(sav_file):
    """
    Open a PSS/E save case file.
    
    Parameters:
    sav_file (str): Path to the save case file.
    """
    psspy.case(sav_file)

def run_load_flow():
    """
    Perform a load flow analysis using the full Newton-Raphson method.
    This function is typically called after opening a save case file.
    Return true if the load flow converged, otherwise false.
    """
    psspy.fnsl([0,0,0,1,1,1,99,0])
    ival = psspy.solved()  # Check whether last solution reached tolerance

    # # Other useful outputs
    # print(psspy.sysmsm(), psspy.sysmva())  # Total system MVA mismatch, system base MVA
    # print(psspy.systot('LOAD'))  # Total system load

    return ival == 0  # Return true if converged

def convert_for_switching_studies(conv_file):
    """
    Convert the PSS/E case for switching studies by applying necessary conversions.
    This function saves the case after conversion.

    Parameters:
    conv_file (str): Path to the converted save case file.
    """
    psspy.save(conv_file)
    psspy.cong(0)
    psspy.conl(0,1,1,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.conl(0,1,3,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.ordr(0)
    psspy.fact()
    psspy.tysl(1)
    psspy.save(conv_file)

def model_generators_classical(dyr_file, H, D):
    """
    Create a classical generator model with specified inertia and damping.

    Parameters:
    dyr_file (str): Path to the dynamic model file.
    H (list of floats): Inertia constants of the generators. List same length as number of generators.
    D (list of floats): Damping factor of the generator. List same length as number of generators.
    """
    psspy.add_plant_model(1,r"""1""",1,r"""GENCLS""",0,"",0,[],[],2,[0.0,0.0])
    psspy.change_plmod_con(1,r"""1""",r"""GENCLS""",1, H[0])
    psspy.change_plmod_con(1,r"""1""",r"""GENCLS""",2, D[0])
    psspy.add_plant_model(5,r"""1""",1,r"""GENCLS""",0,"",0,[],[],2,[0.0,0.0])
    psspy.change_plmod_con(5,r"""1""",r"""GENCLS""",1, H[1])
    psspy.change_plmod_con(5,r"""1""",r"""GENCLS""",2, D[1])
    psspy.dyda(0,1,[2,2,0],0,dyr_file)
    psspy.fact()
    psspy.tysl(1)

def set_output_channels():
    """
    Set output channels for the dynamic simulation results.
    """
    psspy.chsb(0,1,[-1,-1,-1,1,1,0])
    psspy.chsb(0,1,[-1,-1,-1,1,2,0])
    psspy.chsb(0,1,[-1,-1,-1,1,4,0])
    psspy.set_chnfil_type(0)

def run_dynamic_simulation(out_file):
    """
    Run a dynamic simulation with specified output file.

    Parameters:
    out_file (str): Path to the output file where results will be saved.
    """
    psspy.strt_2([0,0],out_file)
    psspy.run(0, 0.1,0,1,0) # Run simulation for 0.1 seconds
    psspy.dist_bus_fault(2,1,0.0,[ 0.1E+13,0.0]) # Apply a fault at bus 2
    psspy.change_channel_out_file(out_file)
    psspy.run(0, 0.2,0,1,0) # Run simulation for 0.2 seconds
    psspy.dist_clear_fault(1) # Clear the fault
    psspy.change_channel_out_file(out_file)
    psspy.run(0, 10.0,0,1,0) # Run simulation for 10 seconds

def export_results_to_excel(out_file, xlsx_file='out.xlsx'):
    """
    Export PSS/E .out file as .xlsx for easier plotting and data management.

    Parameters:
    out_file (str): Path to the output file where results will be saved.
    xlsx_file (str): Path to the output Excel file. Default is 'out.xlsx'.
    """
    dyntools.CHNF.xlsout(dyntools.CHNF(out_file), channels='', show='True', xlsfile = xlsx_file, sheet = '', overwritesheet=True)
