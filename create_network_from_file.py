import pandas as pd


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
line_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE14Bus.xlsx",
                            sheet_name='line', dtype={"line_name": str, "phases": int, "from_bus": str, "to_bus": str,
                            "R1": float, "X1": float, "B1": float})
for _, line in line_df.iterrows():
    # print("New Line.{} Phases={} Bus1={} Bus2={} R1={} X1={} B1={}".format(line.line_name, line.phases, line.from_bus, line.to_bus, line.R1, line.X1, line.B1))
    DSSText.Command = "New Line.{} Phases={} Bus1={} Bus2={} R1={} X1={} B1={}".format(line.line_name, line.phases, line.from_bus, line.to_bus, line.R1, line.X1, line.B1)


#Transmission system TRANSFORMER DEFINITION
tranfo_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE14Bus.xlsx",
                            sheet_name='tranfo', dtype={"tranfo_name": str, "phases": int, "widings": int, "XHL": float,
                            "tap": float, "pri_wdg": int, "pri_bus": str, "pri_kv": float, "pri_kva": float,
                            "sec_wdg": int, "sec_bus": str, "sec_kv": float, "sec_kva": float})
for _, tranfo in tranfo_df.iterrows():
    # print("New Transformer.{} Phases={} Windings={} XHL={} tap={}".format(tranfo.tranfo_name, tranfo.phases, tranfo.widings, tranfo.XHL, tranfo.tap))
    # print("~ wdg={} bus={} kv={} kva={}".format(tranfo.pri_wdg, tranfo.pri_bus, tranfo.pri_kv, tranfo.pri_kva))
    # print("~ wdg={} bus={} kv={} kva={}".format(tranfo.sec_wdg, tranfo.sec_bus, tranfo.sec_kv, tranfo.sec_kva))
    DSSText.Command = "New Transformer.{} Phases={} Windings={} XHL={} tap={}".format(tranfo.tranfo_name, tranfo.phases, tranfo.widings, tranfo.XHL, tranfo.tap)
    DSSText.Command = "~ wdg={} bus={} kv={} kva={}".format(tranfo.pri_wdg, tranfo.pri_bus, tranfo.pri_kv, tranfo.pri_kva)
    DSSText.Command = "~ wdg={} bus={} kv={} kva={}".format(tranfo.sec_wdg, tranfo.sec_bus, tranfo.sec_kv, tranfo.sec_kva)


#transmission system LOAD DEFINITIONS
load_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE14Bus.xlsx",
                            sheet_name='load', dtype={"load_name": str, "bus": str, "phases": int, "model": int,
                            "kV": float, "kW": float, "kvar": float, "Vmaxnpu": float, "Vminpu": float})
for _, load in load_df.iterrows():
    # print("New Load.{} Bus1={} Phases={} Model={} kV={} kW={} kvar={} Vmaxpu={} Vminpu={}".format(load.load_name, load.bus, load.phases, load.model, load.kV, load.kW, load.kvar, load.Vmaxpu, load.Vminpu))
    DSSText.Command = "New Load.{} Bus1={} Phases={} Model={} kV={} kW={} kvar={} Vmaxpu={} Vminpu={}".format(load.load_name, load.bus, load.phases, load.model, load.kV, load.kW, load.kvar, load.Vmaxpu, load.Vminpu)

# transmission system generator DEFINITIONS

gen_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE14Bus.xlsx",
                            sheet_name='gen', dtype={"gen_name": str, "bus": str, "kV": float, "kW": float,
                            "kVA": float, "model": int, "Vpu": float, "Maxkvar": float, "Minkvar": float,
                            "Vmaxpu": float, "Vminpu": float})
for _, gen in gen_df.iterrows():
    # print("New Generator.{} Bus1={} kV={} MVA={} Model={} Vpu={} Maxkvar={} Minkvar={} Vmaxpu={} Vminpu={}".format(gen.gen_name, gen.bus, gen.kV, gen.kW, gen.MVA, gen.Model, gen.Vpu, gen.Maxkvar, gen.Minkvar, gen.Vmaxpu, gen.Vminpu))
    DSSText.Command = "New Generator.{} Bus1={} kV={} kVA={} Model={} Vpu={} Maxkvar={} Minkvar={} Vmaxpu={} Vminpu={}".format(gen.gen_name, gen.bus, gen.kV, gen.kW, gen.kVA, gen.model, gen.Vpu, gen.Maxkvar, gen.Minkvar, gen.Vmaxpu, gen.Vminpu)



DSSText.Command = "Set Voltagebases=[69, 18, 13.8]"
DSSText.Command = "CalcVoltageBases"

DSSText.Command = "Solve"
DSSText.Command = "Show Voltage LN Nodes"

DSSText.Command = "Export Voltages FileVabc14_end.CSV"
DSSText.Command = "Export Currents FileIabc14_end.CSV"
