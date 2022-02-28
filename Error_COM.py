""

from __future__ import print_function
import os
# COM-Server
import win32com.client as com


## Connecting the COM Server => Open a new Vissim Window:
Vissim = com.gencache.EnsureDispatch("Vissim.Vissim") #
# Vissim = com.Dispatch("Vissim.Vissim") # once the cache has been generated, its faster to call Dispatch which also creates the connection to Vissim.
# If you have installed multiple Vissim Versions, you can open a specific Vissim version adding the version number
# Vissim = com.gencache.EnsureDispatch("Vissim.Vissim.10") # Vissim 10
# Vissim = com.gencache.EnsureDispatch("Vissim.Vissim.22") # Vissim 2022


### for advanced users, with this command you can get all Constants from PTV Vissim with this command (not required for the example)
##import sys
##Constants = sys.modules[sys.modules[Vissim.__module__].__package__].constants

Path_of_COM_Basic_Commands_network = 'C:\\Users\\Public\\Documents\\PTV Vision\\PTV Vissim 2022\\Examples Training\\COM\\Basic Commands\\'

## Load a Vissim Network:
Filename               = os.path.join(Path_of_COM_Basic_Commands_network, 'COM Basic Commands.inpx')
flag_read_additionally = False # you can read network(elements) additionally, in this case set "flag_read_additionally" to true
Vissim.LoadNet(Filename, flag_read_additionally)

## Load a Layout:
Filename = os.path.join(Path_of_COM_Basic_Commands_network, 'COM Basic Commands.layx')
Vissim.LoadLayout(Filename)

# Set a signal controller program:
SC_number = 1 # SC = SignalController
SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number)
new_signal_programm_number = 2
SignalController.SetAttValue('ProgNo', new_signal_programm_number)

# Set the state of a signal controller:
# Note: Once a state of a signal group is set, the attribute "ContrByCOM" is automatically set to True. Meaning the signal group will keep this state until another state is set by COM or the end of the simulation
# To switch back to the defined signal controller, set the attribute signal "ContrByCOM" to False (example see below).
SC_number = 1 # SC = SignalController
SG_number = 2 # SG = SignalGroup
SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number)
SignalGroup = SignalController.SGs.ItemByKey(SG_number)
new_state = "GREEN" # possible values 'GREEN', 'RED', 'AMBER', 'REDAMBER' and more, see COM Help: SignalizationState Enumeration
SignalGroup.SetAttValue("SigState", new_state)


#____________________________________________________________________________________________
# GetMultipleAttributes     Read multiple attributes of all objects:
Attributes1 = ("Name", "Length2D")
Name_Length_of_Links = Vissim.Net.Links.GetMultipleAttributes(Attributes1)
print(Name_Length_of_Links)

SignalGroup.GetMultiAttValues("SigState", 3)


Vissim.Net.SignalControllers.ItemByKey(SC_number).SGs.ItemByKey(SG_number).GetMultiAttValues()
SignalGroup = 
new_state = "GREEN" # possible values 'GREEN', 'RED', 'AMBER', 'REDAMBER' and more, see COM Help: SignalizationState Enumeration
SignalGroup


All_Vehicles = Vissim.Net.Vehicles.GetAll() 
