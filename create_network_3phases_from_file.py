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

# ! INSTANTIATE A NEW CIRCUIT AND DEFINE A STIFF 4160V SOURCE
# ! The new circuit is called "ieee123"
# ! This creates a Vsource object connected to "sourcebus". This is now the active circuit element, so
# ! you can simply continue to edit its property value.
# ! The basekV is redefined to 4.16 kV. The bus name is changed to "150" to match one of the buses in the test feeder.
# ! The source is set for 1.0 per unit and the Short circuit impedance is set to a small value (0.0001 ohms)
# ! The ~ is just shorthad for "more" for the New or Edit commands

DSSText.Command = "New object=circuit.ieee123"
DSSText.Command = "~ basekv=4.16 Bus1=150 pu=1.00 R1=0 X1=0.0001 R0=0 X0=0.0001"

# ! 3-PHASE GANGED REGULATOR AT HEAD OF FEEDER (KERSTING ASSUMES NO IMPEDANCE IN THE REGULATOR)
# ! the first line defines the 3-phase transformer to be controlled by the regulator control.
# ! The 2nd line defines the properties of the regulator control according to the test case
reg_3p_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='reg_3p', dtype={"reg_name": str, "phases": int, "windings": int, "buses": list,
                            "conns": list, "kvs": list, "kvas": list, "xhl": float, "percent_loadloss": float, "ppm": float})
# print(reg_3p_df)
for _, reg_3p in reg_3p_df.iterrows():
    # print("new transformer.{} phases={} windings={} buses={} conns={} kvs={} kvas={} XHL={} %LoadLoss={} ppm={}".format(reg_3p.reg_name, reg_3p.phases, reg_3p.windings, reg_3p.buses, reg_3p.conns, reg_3p.kvs, reg_3p.kvs, reg_3p.xhl, reg_3p.percent_loadloss, reg_3p.ppm))
    DSSText.Command = "new transformer.{} phases={} windings={} buses={} conns={} kvs={} kvas={} XHL={} %LoadLoss={} ppm={}".format(reg_3p.reg_name, reg_3p.phases, reg_3p.windings, reg_3p.buses, reg_3p.conns, reg_3p.kvs, reg_3p.kvs, reg_3p.xhl, reg_3p.percent_loadloss, reg_3p.ppm)

creg_3p_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='creg_3p', dtype={"creg_name": str, "transformer": str, "windings": int, "vreg": float,
                            "band": int, "ptration": int, "ctprim": int, "r": float, "x": float})
# print(creg_3p_df)
for _, creg_3p in creg_3p_df.iterrows():
    # print("new regcontrol.{} transformer={} winding={} vreg={} band={} ptratio={} ctprim={} R={} X={}".format(creg_3p.creg_name, creg_3p.transformer, creg_3p.winding, creg_3p.vreg, creg_3p.band, creg_3p.ptration, creg_3p.ctprim, creg_3p.r, creg_3p.x))
    DSSText.Command = "new regcontrol.{} transformer={} winding={} vreg={} band={} ptratio={} ctprim={} R={} X={}".format(creg_3p.creg_name, creg_3p.transformer, creg_3p.winding, creg_3p.vreg, creg_3p.band, creg_3p.ptration, creg_3p.ctprim, creg_3p.r, creg_3p.x)



# ! REDIRECT INPUT STREAM TO FILE CONTAINING DEFINITIONS OF LINECODES
# ! This file defines the line impedances is a similar manner to the description in the test case.
DSSText.Command = "Redirect IEEELineCodes.DSS"

#line definetions
line_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='line', dtype={"line_name": str, "phases": int, "from_bus": str, "to_bus": str,
                            "linecode": int, "length": float, "units": str})
for _, line in line_df.iterrows():
    # print("New Line.{} Phase={} Bus1={} Bus2={} LineCode={} Length={} units={}".format(line.line_name, line.phases, line.from_bus, line.to_bus, line.linecode, line.length, line.units))
    DSSText.Command = "New Line.{} Phase ={} Bus1={} Bus2={} LineCode={} Length={} units={}".format(line.line_name, line.phases, line.from_bus, line.to_bus, line.linecode, line.length, line.units)

# ! NORMALLY CLOSED SWITCHES ARE DEFINED AS SHORT LINES
# ! Could also be defned by setting the Switch=Yes property
# ! NORMALLY OPEN SWITCHES; DEFINED AS SHORT LINE TO OPEN BUS SO WE CAN SEE OPEN POINT VOLTAGES.
# ! COULD ALSO BE DEFINED AS DISABLED OR THE TERMINCAL COULD BE OPENED AFTER BEING DEFINED
switch_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='switch', dtype={"switch_name": str, "phases": int, "from_bus": str, "to_bus": str,
                            "r1": float, "r0": float, "x1": float, "x0": float, "c1": float, "c0": float, "length": float})
for _, switch in switch_df.iterrows():
    # print("New Line.{} Phase ={} Bus1={} Bus2={} r1={} r0={} x1={} x0={} c1={} c0={} Length={}".format(switch.switch_name, switch.phases, switch.from_bus, switch.to_bus, switch.r1, switch.r0, switch.x1, switch.x0, switch.c1, switch.c0, switch.length))
    DSSText.Command = "New Line.{} Phase ={} Bus1={} Bus2={} r1={} r0={} x1={} x0={} c1={} c0={} Length={}".format(switch.switch_name, switch.phases, switch.from_bus, switch.to_bus, switch.r1, switch.r0, switch.x1, switch.x0, switch.c1, switch.c0, switch.length)

#Transmission system Load TRANSFORMER DEFINITION
load_tranfo_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='load_tranfo', dtype={"load_tranfo_name": str, "phases": int, "windings": int, "xhl": float,
                            "pri_wdg": int, "pri_bus": str, "pri_conn": str, "pri_kv": float, "pri_kva": float, "pri_percent_r": float,
                            "sec_wdg": int, "sec_bus": str, "sec_conn": str, "sec_kv": float, "sec_kva": float, "sec_percent_r": float})
for _, load_tranfo in load_tranfo_df.iterrows():
    # print("New Transformer.{} Phase ={} Windings={} Xhl={}".format(load_tranfo.load_tranfo_name, load_tranfo.phases, load_tranfo.windings, load_tranfo.xhl))
    # print("~ wdg={} bus={} conn={} kv={} kva={} %r={}".format(load_tranfo.pri_wdg, load_tranfo.pri_bus, load_tranfo.pri_conn, load_tranfo.pri_kv, load_tranfo.pri_kva, load_tranfo.pri_percent_r))
    # print("~ wdg={} bus={} conn={} kv={} kva={} %r={}".format(load_tranfo.sec_wdg, load_tranfo.sec_bus, load_tranfo.sec_conn, load_tranfo.sec_kv, load_tranfo.sec_kva, load_tranfo.sec_percent_r))
    DSSText.Command = "New Transformer.{} Phase ={} Windings={} Xhl={}".format(load_tranfo.load_tranfo_name, load_tranfo.phases, load_tranfo.windings, load_tranfo.xhl)
    DSSText.Command = "~ wdg={} bus={} conn={} kv={} kva={} %r={}".format(load_tranfo.pri_wdg, load_tranfo.pri_bus, load_tranfo.pri_conn, load_tranfo.pri_kv, load_tranfo.pri_kva, load_tranfo.pri_percent_r)
    DSSText.Command = "~ wdg={} bus={} conn={} kv={} kva={} %r={}".format(load_tranfo.sec_wdg, load_tranfo.sec_bus, load_tranfo.sec_conn, load_tranfo.sec_kv, load_tranfo.sec_kva, load_tranfo.sec_percent_r)


# ! CAPACITORS
# ! Capacitors are 2-terminal devices. The 2nd terminal (Bus2=...) defaults to all phases
# ! connected to ground (Node 0). Thus, it need not be specified if a Y-connected or L-N connected
# ! capacitor is desired
cap_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='cap', dtype={"cap_name": str, "bus": str, "phases": int, "kVAR": float, "kV": float})
for _, cap in cap_df.iterrows():
    # print("New Capapcitor.{} Bus1={} Phases={} kVAR={} kV=4.16".format(cap.cap_name, cap.bus, cap.phases, cap.kVAR, cap.kV))
    DSSText.Command = "New Capacitor.{} Bus1={} Phases={} kVAR={} kV=4.16".format(cap.cap_name, cap.bus, cap.phases, cap.kVAR, cap.kV)

# create single phase regulation
reg_1p_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='reg_1p', dtype={"reg_name": str, "phases": int, "windings": int, "bank": str, "buses": list,
                            "conns": list, "kvs": list, "kvas": list, "xhl": float, "percent_loadloss": float, "ppm": float})
# print(reg_3p_df)
for _, reg_1p in reg_1p_df.iterrows():
    # print("new transformer.{} phases={} windings={} bank={} buses={} conns={} kvs={} kvas={} XHL={} %LoadLoss={} ppm={}".format(reg_1p.reg_name, reg_1p.phases, reg_1p.windings, reg_1p.bank, reg_1p.buses, reg_1p.conns, reg_1p.kvs, reg_1p.kvs, reg_1p.xhl, reg_1p.percent_loadloss, reg_1p.ppm))
    DSSText.Command = "new transformer.{} phases={} windings={} bank={} buses={} conns={} kvs={} kvas={} XHL={} %LoadLoss={} ppm={}".format(reg_1p.reg_name, reg_1p.phases, reg_1p.windings, reg_1p.bank, reg_1p.buses, reg_1p.conns, reg_1p.kvs, reg_1p.kvs, reg_1p.xhl, reg_1p.percent_loadloss, reg_1p.ppm)

crep_1p_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='creg_1p', dtype={"creg_name": str, "transformer": str, "windings": int, "vreg": float,
                            "band": int, "ptration": int, "ctprim": int, "r": float, "x": float})
# print(creg_3p_df)
for _, creg_1p in crep_1p_df.iterrows():
    # print("new regcontrol.{} transformer={} winding={} vreg={} band={} ptratio={} ctprim={} R={} X={}".format(creg_1p.creg_name, creg_1p.transformer, creg_1p.winding, creg_1p.vreg, creg_1p.band, creg_1p.ptration, creg_1p.ctprim, creg_1p.r, creg_1p.x))
    DSSText.Command = "new regcontrol.{} transformer={} winding={} vreg={} band={} ptratio={} ctprim={} R={} X={}".format(creg_1p.creg_name, creg_1p.transformer, creg_1p.winding, creg_1p.vreg, creg_1p.band, creg_1p.ptration, creg_1p.ctprim, creg_1p.r, creg_1p.x)





# transmission system LOAD DEFINITIONS
load_df = pd.read_excel("C:\\Users\\ece\\Desktop\\Pornchai\\OpenDSS\\Pornchai cases\\python codes\\IEEE_data\\data_IEEE123Bus.xlsx",
                            sheet_name='load', dtype={"load_name": str, "bus": str, "phases": int, "conn": str,
                            "model": int, "kV": float, "kW": float, "kvar": float})
for _, load in load_df.iterrows():
    # print("New Load.{} Bus1={} Phases={} Conn={} Model={} kV={} kW={} kvar={}".format(load.load_name, load.bus, load.phases, load.conn, load.model, load.kV, load.kW, load.kvar))
    DSSText.Command = "New Load.{} Bus1={} Phases={} Conn={} Model={} kV={} kW={} kvar={}".format(load.load_name, load.bus, load.phases, load.conn, load.model, load.kV, load.kW, load.kvar)



DSSText.Command = "Set Voltagebases=[4.16, 0.48]"  #! ARRAY OF VOLTAGES IN KV

DSSText.Command = "CalcVoltageBases"   #! PERFORMS ZERO LOAD POWER FLOW TO ESTIMATE VOLTAGE BASES

#
DSSText.Command = "Solve"
DSSText.Command = "Show Voltage LN Nodes"
