import arcpy
import pandas as pd
class Toolbox(object):

    def __init__(self):
        self.label = "Toolbox"
        self.alias = "toolbox"
        self.tools = [TableToExcel]

class TableToExcel(object):

    def __init__(self):
        self.label = "Table To Excel"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        
        Table = arcpy.Parameter(
            displayName = "Enter table path",
            name = "Table",
            datatype = "GPTableView",
            parameterType = "Required",
            direction = "Input")
        excel = arcpy.Parameter(
            displayName = "Enter excel path",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        excel.filter.list = ["xlsx"]
        return [Table,excel]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        return

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        Table = parameters[0].valueAsText
        excel = parameters[1].valueAsText

        # read rows in to pandas
        fields = [f.name for f in arcpy.ListFields(Table)]
        data = [f for f in arcpy.da.SearchCursor(Table,fields)]
        # write to excel
        # writer = pd.ExcelWriter(excel)


        df = pd.DataFrame(data,columns=fields)
        with pd.ExcelWriter(excel,engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="data") 
        writer.save()

            # df2.to_excel(writer, sheet_name="Sheet2")  
        # # read to pandas
        # fields = [f.name for f in arcpy.ListFields(Table)]
        # data = [f for f in arcpy.da.SearchCursor(Table,fields)]
        # df = pd.DataFrame(data,columns=fields)
        # # write to excel
        # writer = pd.ExcelWriter(excel)
        # df.to_excel(writer,'Sheet1',index=False)

        return