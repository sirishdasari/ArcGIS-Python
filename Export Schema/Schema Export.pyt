# -*- coding: utf-8 -*-

import arcpy
import pandas as pd
import os

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
        self.tools = [ExportSchema,ExportSubTypes]


class ExportSchema(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Export Schema"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        Layer = arcpy.Parameter(
            displayName = "Enter GDB path",
            name = "Layer",
            datatype = "DEWorkspace",
            parameterType = "Required",
            direction = "Input")
        Layer.filter.list = ['Local Database','Remote Database']
        excel = arcpy.Parameter(
            displayName = "Enter excel output",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        excel.filter.list = ['xlsx','xls']  

        return [Layer,excel]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        # os walk

        gdbWalk = arcpy.da.Walk(parameters[0].valueAsText,datatype="FeatureClass")
        data = []
        for dirpath, dirnames, filenames in gdbWalk:
            for file in filenames:
                arcpy.AddMessage(file)
                featurePath = os.path.join(dirpath,file)
                for f in arcpy.ListFields(featurePath):
                    data.append((file,f.name,f.aliasName,f.type,f.isNullable,f.domain,f.defaultValue,f.length))

        dfFields = pd.DataFrame(data,columns=["Layer","Fieldname","Alias","Type","isNullable","domain","defaultValuetable","length"])
        # dfFields.to_excel(parameters[1].valueAsText,sheet_name ="Schema",index=False)

        domainDataCoded = []
        domainDataRange = []
        for domain in arcpy.da.ListDomains(parameters[0].valueAsText):
            # arcpy.AddMessage(domain)
            if domain.domainType == 'CodedValue':
                coded_values = domain.codedValues
                for val, desc in coded_values.items():
                    domainDataCoded.append((domain.name,"CodedValue",val, desc))
            elif domain.domainType == "Range":
                domainDataRange.append((domain.name,"Range",domain.range[0],domain.range[1]))
        dfDomains = pd.DataFrame(domainDataCoded,columns=["Domain","Domain Type","Code","Description"])
        dfRange = pd.DataFrame(domainDataRange,columns=["Domain","Domain Type","Min","Max"])
        # dfCode.to_excel(parameters[1].valueAsText,sheet_name ='Domains Coded',index=False)
        # Excel Writer

        with pd.ExcelWriter(parameters[1].valueAsText) as writer:
            dfFields.to_excel(writer,sheet_name ="Fields",index=False)
            dfDomains.to_excel(writer,sheet_name ="DomainsCoded",index=False)
            dfRange.to_excel(writer,sheet_name ="DomainsRange",index=False)

        # data = []
        # # with pd.ExcelWriter() as writer:
        # for layer in parameters[0].value:
        #     arcpy.AddMessage(layer[0].name)
        #     for f in arcpy.ListFields(layer[0].name):
        #         data.append((layer[0].name,f.name,f.aliasName,f.type,f.isNullable,f.domain,f.defaultValue,f.length))
        # df = pd.DataFrame(data,columns=["Layer","Fieldname","Alias","Type","isNullable","domain","defaultValuetable","length"])
        # df.to_excel(parameters[1].valueAsText,sheet_name ="Schema",index=False)

class ExportFields(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Fields Schema"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        Layer = arcpy.Parameter(
            displayName = "Enter Feature Layer",
            name = "Layer",
            datatype = "GPTableView",
            parameterType = "Required",
            direction = "Input",
            multiValue=True)
        Layer.columns = [['GPTableView', 'Layer']]
        excel = arcpy.Parameter(
            displayName = "Enter Output Folder for TR Filter excels",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        excel.filter.list = ['xlsx','xls']  

        return [Layer,excel]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        data = []
        # with pd.ExcelWriter() as writer:
        for layer in parameters[0].value:
            arcpy.AddMessage(layer[0].name)
            for f in arcpy.ListFields(layer[0].name):
                data.append((layer[0].name,f.name,f.aliasName,f.type,f.isNullable,f.domain,f.defaultValue,f.length))
        df = pd.DataFrame(data,columns=["Layer","Fieldname","Alias","Type","isNullable","domain","defaultValuetable","length"])
        df.to_excel(parameters[1].valueAsText,sheet_name ="Schema",index=False)

class ExportDomains(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Domains Schema"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        gdb = arcpy.Parameter(
            displayName = "Enter gdb",
            name = "gdb",
            datatype = "DEWorkspace",
            parameterType = "Required",
            direction = "Input")
        gdb.filter.list = ['gdb']  
        excel = arcpy.Parameter(
            displayName = "Enter Output Folder for TR Filter excels",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        excel.filter.list = ['xlsx','xls']  

        return [gdb,excel]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        # domainDataRange = []
        domainDataCoded = []
        for domain in arcpy.da.ListDomains(parameters[0].valueAsText):
            arcpy.AddMessage(domain)
            arcpy.AddMessage('Domain name: {0}'.format(domain.name))
            if domain.domainType == 'CodedValue':
                coded_values = domain.codedValues
                for val, desc in coded_values.items():
                    domainDataCoded.append((domain.name,val, desc))
        dfCode = pd.DataFrame(domainDataCoded,columns=["Domain","Code","Description"])
        dfCode.to_excel(parameters[1].valueAsText,sheet_name ='Domains Coded',index=False)
                # if domain.domainType == 'Range':
                #     domainDataRange((domain.name,domain.range[0], domain.range[1]))
        # dfRange = pd.DataFrame(domainDataRange,columns=["Domain","Minimum","Maximum"])
        # dfRange.to_excel(writer,sheet_name ='Domains Range',index=False)
            # print('Max: {0}'.format(domain.range[1]))‍‍‍‍‍‍‍‍‍‍‍‍‍
        # arcpy.AddMessage(domain)


class ExportSubTypes(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "SubTypes Schema"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        Layer = arcpy.Parameter(
            displayName = "Enter Feature Layer",
            name = "Layer",
            datatype = "GPTableView",
            parameterType = "Required",
            direction = "Input",
            multiValue=True)
        Layer.columns = [['GPTableView', 'Layer']]
        excel = arcpy.Parameter(
            displayName = "Enter Output Folder for TR Filter excels",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Optional",
            direction = "Output")
        excel.filter.list = ['xlsx','xls']  

        return [Layer,excel]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        data = []
        # with pd.ExcelWriter() as writer:
        for layer in parameters[0].value:
            arcpy.AddMessage(layer[0].name)
            subtypes = arcpy.da.ListSubtypes(layer[0].name)

            for stcode, stdict in list(subtypes.items()):
                # arcpy.AddMessage(f"Code: {stcode}")
                for stkey in list(stdict.keys()):
                    if stkey == "Name":
                        data.append([layer[0].name,stcode,stdict[stkey]])
                        arcpy.AddMessage([layer[0].name,stcode,stdict[stkey]])
        df = pd.DataFrame(data,columns=["Layer","SubTypeCode","Description"])
        df.to_excel(parameters[1].valueAsText,sheet_name ="Schema",index=False)
                        # arcpy.AddMessage(f"Name: {stdict[stkey]}")
                        # fields = stdict[stkey]
                        # for field, fieldvals in list(fields.items()):
                        #     arcpy.AddMessage(f" --Field name: {field}")
                        #     arcpy.AddMessage(f" --Field default value: {fieldvals[0]}")
                        #     if not fieldvals[1] is None:
                        #         arcpy.AddMessage(f" --Domain name: {fieldvals[1].name}")
            #         else:
            #             arcpy.AddMessage(f"{stkey}: {stdict[stkey]}")
