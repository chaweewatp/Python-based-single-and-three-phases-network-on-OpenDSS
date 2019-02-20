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


# Load in an example circuit
DSSText.Command = r"Compile 'C:\Users\ece\Desktop\Pornchai\OpenDSS\Pornchai cases\python codes\IEEE14Master.dss'"

# Solve DSS Text Interface
#DSSText.Command = "solve"
DSSSolution.Solve()

if DSSSolution.Converged:
    print("The Circuit Solved Successfully.")


# print result to text file
DSSText.Command = "show voltages LN Nodes"


#see file OpenDSSManual.pdf page 37 for more export item

#print result to csv file
DSSText.Command = "Export Voltages"
Filename = DSSText.Result
print("File save to: " + Filename)


#print result to csv file
DSSText.Command = "Export Powers"
Filename = DSSText.Result
print("File save to: " + Filename)



#print result to csv file
DSSText.Command = "Export Currents"
Filename = DSSText.Result
print("File save to: " + Filename)
