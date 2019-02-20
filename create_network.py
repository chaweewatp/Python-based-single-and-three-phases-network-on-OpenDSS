import win32com.client

# ****************************************************
# * Initialize OpenDSS
# ****************************************************
# Instantiate the OpenDSS Object
try:
    DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
except:
    print ("Unable to start the OpenDSS Engine")
    raise SystemExit
# Set up the Text, Circuit, and Solution Interfaces
DSSText = DSSObj.Text
DSSCircuit = DSSObj.ActiveCircuit
DSSSolution = DSSCircuit.Solution


DSSText.Command = "Clear"
DSSText.Command = "Set DefaultBaseFrequency=60"

#14 bus transmission system
DSSText.Command = "new circuit.14bus basekv=69 pu=1.06 phases=3 bus1=SourceBus Angle=0 MVASC3=210000"

#line definetions
DSSText.Command = "New Line.1 Phases=3 Bus1=2 Bus2=5 R1=2.7114 X1=8.2784 B1=0.0007"
DSSText.Command = "New Line.2 Phases=3 Bus1=7 Bus2=9 X1=0.2095"
DSSText.Command = "New Line.3 Phases=3 Bus1=SourceBus Bus2=2 R1=0.9227 X1=2.8171 B1=0.0011"
DSSText.Command = "New Line.5 Phases=3 Bus1=3 Bus2=2 R1=2.2372 X1=9.4254 B1=0.00091"
DSSText.Command = "New Line.6 Phases=3 Bus1=3 Bus2=4 R1=3.1903 X1=8.1427 B1=0.00073"
DSSText.Command = "New Line.7 Phases=3 Bus1=SourceBus Bus2=5 R1=2.5724 X1=10.6189 B1=0.00103"
DSSText.Command = "New Line.8 Phases=3 Bus1=5 Bus2=4 R1=0.6356 X1=2.0049 B1=0.000269"
DSSText.Command = "New Line.9 Phases=3 Bus1=2 Bus2=4 R1=2.7666 X1=8.3946 B1=0.000786"
DSSText.Command = "New Line.10 Phases=3 Bus1=6 Bus2=12 R1=0.2341 X1=0.4872"
DSSText.Command = "New Line.11 Phases=3 Bus1=12 Bus2=13 R1=0.4207 X1=0.3807"
DSSText.Command = "New Line.12 Phases=3 Bus1=6 Bus2=13 R1=0.1260 X1=0.2481"
DSSText.Command = "New Line.13 Phases=3 Bus1=6 Bus2=11 R1=0.1809 X1=0.3788"
DSSText.Command = "New Line.14 Phases=3 Bus1=11 Bus2=10 R1=0.1563 X1=0.3658"
DSSText.Command = "New Line.15 Phases=3 Bus1=9 Bus2=10 R1=0.0606 X1=0.1609"
DSSText.Command = "New Line.16 Phases=3 Bus1=9 Bus2=14 R1=0.2421 X1=0.5149"
DSSText.Command = "New Line.17 Phases=3 Bus1=14 Bus2=13 R1=0.3255 X1=0.6628"



#Transmission system TRANSFORMER DEFINITION
DSSText.Command = "New Transformer.t2 Phases=3 Windings=2 XHL=25.2020 tap=0.932"
DSSText.Command = "~ wdg=1 bus=5 kv=69 kva=100000 "
DSSText.Command = "~ wdg=2 bus=6 kv=13.8 kva=100000"

DSSText.Command = "New Transformer.t3 Phases=3 Windings=2 XHL=55.6180 tap=0.969"
DSSText.Command = "~ wdg=1 bus=4 kv=69 kva=100000 "
DSSText.Command = "~ wdg=2 bus=9 kv=13.8 kva=100000"

DSSText.Command = "New Transformer.t4 Phases=3 Windings=2 XHL=20.9120 tap=0.978"
DSSText.Command = "~ wdg=1 bus=4 kv=69 kva=100000"
DSSText.Command = "~ wdg=2 bus=7 kv=13.8 kva=100000"

DSSText.Command = "New Transformer.t5 Phases=3 Windings=2 XHL=17.6150 tap=0.98"
DSSText.Command = "~ wdg=1 bus=8 kv=18 kva=100000"
DSSText.Command = "~ wdg=2 bus=7 kv=13.8 kva=100000"

#transmission system LOAD DEFINITIONS
DSSText.Command = "New Load.11 Bus1=11 Phases=3 Model=1 kV=13.8 kW=4900 kvar=2520 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.13 Bus1=13 Phases=3 Model=1 kV=13.8 kW=18900 kvar=8120 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.3 Bus1=3 Phases=3 Model=1 kV=69 kW=131880 kvar=26600 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.5 Bus1=5 Phases=3 Model=1 kV=69 kW=10640 kvar=2240 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.2 Bus1=2 Phases=3 Model=1 kV=69 kW=30380 kvar=17780 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.6 Bus1=6 Phases=3 Model=1 kV=13.8 kW=15680 kvar=10500 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.4 Bus1=4 Phases=3 Model=1 kV=69 kW=66920 kvar=5600 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.14 Bus1=14 Phases=3 Model=1 kV=13.8 kW=20860 kvar=7000 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.12 Bus1=12 Phases=3 Model=1 kV=13.8 kW=8540 kvar=2240 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.9 Bus1=9 Phases=3 Model=1 kV=13.8 kW=41300 kvar=23240 Vmaxpu=1.2 Vminpu=0.8"
DSSText.Command = "New Load.10 Bus1=10 Phases=3 Model=1 kV=13.8 kW=3570 kvar=1730 Vmaxpu=1.2 Vminpu=0.8"



DSSText.Command = "Set Voltagebases=[69, 18, 13.8]"
DSSText.Command = "CalcVoltageBases"

DSSText.Command = "Solve"
DSSText.Command = "Show Voltage LN Nodes"

DSSText.Command = "Export Voltages FileVabc14_end.CSV"
DSSText.Command = "Export Currents FileIabc14_end.CSV"
