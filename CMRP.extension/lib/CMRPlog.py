import datetime
import os

from pyrevit import revit
from CMRP import *
from CMRPsettings import *

# ====================================================================
# Export Functions
# ====================================================================
def excel_transfer(starttime, data, unit, name):
    """This function transfers the output data to a csv-file
    Args:
        starttime (_timestamp_): the timestamp at te start of the running script
        doc (_document_): the document in wich the output data is collected
        data (_list_): a list with the output data of the model
        unit (_list_): a list with the units of the output data
        name (_list_): a list which indicates the name/type of the output data
    """

    # Retrieve file data
    # -----------------------------------------------------------------------
    doc = revit.doc
    docTitle = doc.Title
    script = os.path.basename(__file__)
    scriptname = str(script)

    # fucntion om de kpi-data te combineren
    # --------------------------------------------------------------------
    outputdata = ""
    for x,y,z in zip(data,unit,name):
        new = str(x)+";"+ str(y) +";"+ str(z) +";"
        outputdata = str(outputdata) + new

    # check revit version
    # --------------------------------------------------------------------
    hostapp = pyrevit._HostApplication()
    if hostapp.is_newer_than(2017):
        fullVersion = hostapp.subversion
        settings_set('revit_version', fullVersion)
    else:
        fullVersion = str(hostapp.subversion)
        settings_set('revit_version', fullVersion)

    # get properties for writing the log file
    # --------------------------------------------------------------------
    dateStamp   = datetime.datetime.today().strftime("%d/%m/%y %H:%M:%S")
    userName    = os.environ.get('USERNAME')
    userProfile = os.environ.get('USERPROFILE')

    # determine the relevant path if provided
    # --------------------------------------------------------------------
    if os.path.exists("H:\_GROEPEN\CAD\DEV\Tools\CMRP\DataExportCMRP"):
        myPath = "H:\_GROEPEN\CAD\DEV\Tools\CMRP\DataExportCMRP" + "\\"
        result = "server export"
    else:
        if os.path.exists(userProfile + "\OneDrive - b-b.be\Bureaublad\BB_ErrorReport\Data"):
            myPath = userProfile + "\OneDrive - b-b.be\Bureaublad\BB_ErrorReport\Data\\"     
            result = "lokale export - map bestaat"
        else:
            os.makedirs(userProfile + "\OneDrive - b-b.be\Bureaublad\BB_ErrorReport\Data")
            myPath = userProfile + "\OneDrive - b-b.be\Bureaublad\BB_ErrorReport\Data\\"    
            result = "lokale export - map aangemaakt"

    # Error catch
    # --------------------------------------------------------------------
    #if scriptError:
        #errors = "TRUE"
    #else:
        #errors = "FALSE"

    # Calculate Runtime
    # --------------------------------------------------------------------
    endtime = datetime.datetime.today()
    runtime = endtime - starttime
    strRuntime = str(runtime)

    # Generate data to write
    # --------------------------------------------------------------------
    myLog   = myPath + "CMRP_Log.csv"
    dataRow = dateStamp + ";" + userName + ";" + docTitle +".rvt" + ";" + fullVersion + ";" + scriptname + ";" + strRuntime + ";" + outputdata

    # Adds new line to log file or creates one if doesn't exist
    # --------------------------------------------------------------------
    try:
        with open(myLog, "a") as file:
            file.writelines(dataRow + "\n")
        result = dataRow
    except:
        result = "export mislukt"

    return result