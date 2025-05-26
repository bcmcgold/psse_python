# REQUIREMENTS: Python 3
# This script reads the output from a PSS/E dynamic simulation and plots the results using matplotlib.

import pandas as pd
import matplotlib.pyplot as plt

import pandas as pd
df = pd.read_excel('smib/out.xlsx', sheet_name='Sheet1', skiprows=2, header=1)
t = df['Time(s)']
a1 = df['ANGL 1[BUS 1 24.000]1']
a2 = df['ANGL 5[BUS 5 230.00]1']
p1 = df['POWR 1[BUS 1 24.000]1']
p2 = df['POWR 5[BUS 5 230.00]1']
e1 = df['ETRM 1[BUS 1 24.000]1']
e2 = df['ETRM 5[BUS 5 230.00]1']
plt.figure(1, figsize=(15, 5))
plt.subplot(1, 3, 1)
plt.plot(t, a1, t, a2)
plt.xlabel('Time (s)')
plt.ylabel('Angle (deg)')
plt.title('Phase angle vs. time')
plt.legend(['BUS 1', 'BUS 5'])
plt.subplot(1, 3, 2)
plt.plot(t, p1, t, p2)
plt.xlabel('Time (s)')
plt.ylabel('Electrical power (pu)')
plt.title('Electrical power vs. time')
plt.legend(['BUS 1', 'BUS 5'])
plt.subplot(1, 3, 3)
plt.plot(t, e1, t, e2)
plt.xlabel('Time (s)')
plt.ylabel('Terminal voltage (pu)')
plt.title('Terminal voltage vs. time')
plt.legend(['BUS 1', 'BUS 5'])
plt.show()