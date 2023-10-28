# -*- coding: utf-8 -*-

import arcpy
import os
import pandas as pd
from decimal import Decimal

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
        self.tools = [ExportSchema,CompareGDB,DomainCheck,DecimalCheckExcel]

class ExportSchema(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "1. Export Schema"
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

class CompareGDB(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "2. Compare GDB"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        GDB1 = arcpy.Parameter(
            displayName = "Latest GDB",
            name = "GDB1",
            datatype = "DEWorkspace",
            parameterType = "Required",
            direction = "Input")
        GDB2 = arcpy.Parameter(
            displayName = "Old GDB",
            name = "GDB2",
            datatype = "DEWorkspace",
            parameterType = "Required",
            direction = "Input")
        excel = arcpy.Parameter(
            displayName = "Enter Excel Output",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        excel.filter.list = ['xlsx','xls']      

        mapping = arcpy.Parameter(
            displayName='Layer/Table Mapping for changed names',
            name='Mapping',
            datatype='GPValueTable',
            parameterType='Optional',
            direction='Input')
        
        mapping.columns = [['GPString', 'New'], ['GPString', 'Old']]
        return [GDB1,GDB2,excel,mapping]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        if parameters[0].altered and parameters[1].altered:
            new_walk = arcpy.da.Walk(parameters[0].valueAsText,datatype=["FeatureClass","Table"])
            old_walk = arcpy.da.Walk(parameters[1].valueAsText,datatype=["FeatureClass","Table"])
            new_FC = []
            old_FC = []
            for dirpath, dirnames, filenames in new_walk:
                for filename in filenames:
                    new_FC.append(os.path.join(dirpath,filename))
            for dirpath, dirnames, filenames in old_walk:
                for filename in filenames:
                    old_FC.append(os.path.join(dirpath,filename))
            parameters[3].filters[0].list = new_FC
            parameters[3].filters[1].list = old_FC
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        # if parameters[3].value !=None:
        #     arcpy.AddMessage(parameters[3].value)
        

        with pd.ExcelWriter(parameters[2].valueasText) as writer:
            new_walk = arcpy.da.Walk(parameters[0].valueAsText,datatype=["FeatureClass","Table"])
            old_walk = arcpy.da.Walk(parameters[1].valueAsText,datatype=["FeatureClass","Table"])

            new_FC = []
            old_FC = []

            for dirpath, dirnames, filenames in new_walk:
                for filename in filenames:
                    if arcpy.Describe(os.path.join(dirpath,filename)).dataType == "Table":
                        new_FC.append([os.path.basename(parameters[0].valueAsText),arcpy.Describe(os.path.join(dirpath, filename)).dataType,filename,None])
                    elif arcpy.Describe(os.path.join(dirpath,filename)).dataType == "FeatureClass":
                        new_FC.append([os.path.basename(parameters[0].valueAsText),arcpy.Describe(os.path.join(dirpath,filename)).shapeType,filename,None])
                    
            for dirpath, dirnames, filenames in old_walk:
                for filename in filenames:
                    if arcpy.Describe(os.path.join(dirpath,filename)).dataType == "Table":
                        old_FC.append([os.path.basename(parameters[1].valueAsText),arcpy.Describe(os.path.join(dirpath, filename)).dataType,filename,None])
                    elif arcpy.Describe(os.path.join(dirpath,filename)).dataType == "FeatureClass":
                        old_FC.append([os.path.basename(parameters[1].valueAsText),arcpy.Describe(os.path.join(dirpath, filename)).shapeType,filename,None])

            new = new_FC
            old = old_FC
            
            for f in new:
                if f[2] in [i[2] for i in old]:
                    f[3] = "Matched"
                if f[2] not in [i[2] for i in old]:
                    f[3] = "Added"
                if f[2].lower() in [i[2].lower() for i in old] and f[3] == "Added":
                    f[3] = "Modified"

            for f in old:
                if f[2] in [i[2] for i in new]:
                    f[3] = "Matched"
                if f[2] not in [i[2] for i in new]:
                    f[3] = "Removed"
                if f[2].lower() in [i[2].lower() for i in new] and f[3] == "Removed":
                    f[3] = "Modified"
            
            df1= pd.DataFrame(new,columns=["GDB","Feature Type","Feature Class","Status"])
            df2= pd.DataFrame(old,columns=["GDB","Feature Type","Feature Class","Status"])

            df1.append(df2).to_excel(writer,sheet_name ='Layers Report',index=False)
            matched_fields=[f[2] for f in new_FC if f[2] in [i[2] for i in old_FC]]# and f[3]!="Table"]
            
            # arcpy.AddMessage(matched_fields)
            # arcpy.AddMessage(new_FC)
            # arcpy.AddMessage(old_FC)

            # arcpy.AddMessage("Modified")
            modified_fields = []
            for f in new_FC:
                for o in old_FC:
                    if f[2].lower() == o[2].lower() and f[2] != o[2]:
                        modified_fields.append([f[2],o[2]])
            # arcpy.AddMessage(modified_fields)
            # arcpy.AddMessage(len(modified_fields))
            
            if len(modified_fields)>0:
                arcpy.AddWarning(f"Check for mapping fields{modified_fields}")

            arcpy.AddMessage("======================= Compare Field Name, Field Length and Field Domain ====================")
            totalReport = []
        
            if parameters[3].value !=None:
                arcpy.AddMessage([os.path.basename(f[0])+' => '+os.path.basename(f[1]) for f in parameters[3].value])
                for mlayer in parameters[3].value:
                    newfields = {f.name:[f.type,f.length,f.domain] for f in arcpy.ListFields(mlayer[0])}
                    oldfields = {f.name:[f.type,f.length,f.domain] for f in arcpy.ListFields(mlayer[1])}
                    for f in newfields.keys():
                        if newfields[f][0] != "String":
                            newfields[f][1] =None 
                    for f in oldfields.keys():
                        if oldfields[f][0] != "String":
                            oldfields[f][1] =None 
                    def compareFields(newFieldNames,oldFieldNames):
                        # Fields Data Type
                        for f in newFieldNames.keys():
                            if f in oldFieldNames.keys() and newFieldNames[f][0] != oldFieldNames[f][0]:
                                totalReport.append([os.path.basename(parameters[0].valueAsText),os.path.basename(mlayer[0]),"Added",f,newFieldNames[f][0],"",""])
                            if f not in oldFieldNames.keys():
                                totalReport.append([os.path.basename(parameters[0].valueAsText),os.path.basename(mlayer[0]),"Added",f,"","",""])
                                
                        for f in oldFieldNames.keys():
                            if f in newFieldNames.keys() and newFieldNames[f][0] != oldFieldNames[f][0]:
                                totalReport.append([os.path.basename(parameters[1].valueAsText),os.path.basename(mlayer[1]),"Removed",f,oldFieldNames[f][0],"",""])
                            if f not in newFieldNames.keys():
                                totalReport.append([os.path.basename(parameters[1].valueAsText),os.path.basename(mlayer[1]),"Removed",f,"","",""])

                        # Fields Length
                        for f in newFieldNames.keys():
                            if f in oldFieldNames.keys() and newFieldNames[f][1] != oldFieldNames[f][1]:
                                totalReport.append([os.path.basename(parameters[0].valueAsText),os.path.basename(mlayer[0]),"Text Changed",f,"",newFieldNames[f][1],""])
                                
                        for f in oldFieldNames.keys():
                            if f in newFieldNames.keys() and newFieldNames[f][1] != oldFieldNames[f][1]:
                                totalReport.append([os.path.basename(parameters[1].valueAsText),os.path.basename(mlayer[1]),"Text Changed",f,"",oldFieldNames[f][1],""])

                        # Fields Domain
                        for f in newFieldNames.keys():
                            if f in oldFieldNames.keys() and newFieldNames[f][2] != oldFieldNames[f][2]:
                                totalReport.append([os.path.basename(parameters[0].valueAsText),os.path.basename(mlayer[0]),"Domain Changed",f,"","",newFieldNames[f][2]])
                                
                        for f in oldFieldNames.keys():
                            if f in newFieldNames.keys() and newFieldNames[f][2] != oldFieldNames[f][2]:
                                totalReport.append([os.path.basename(parameters[1].valueAsText),os.path.basename(mlayer[1]),"Domain Changed",f,"","",oldFieldNames[f][2]])                    

                    compareFields(newfields,oldfields)
     
            for checkfield in matched_fields:
                oldlyrpath = ""
                newlyrpath = ""
                
                for dirpath, dirnames, filenames in arcpy.da.Walk(parameters[0].valueAsText,datatype=["FeatureClass","Table"]):
                    for file in filenames:
                        if checkfield == file:
                            newlyrpath = os.path.join(dirpath,file)
                            
                for dirpath, dirnames, filenames in arcpy.da.Walk(parameters[1].valueAsText,datatype=["FeatureClass","Table"]):
                    for file in filenames:
                        if checkfield == file:
                            oldlyrpath = os.path.join(dirpath,file)
                            
                arcpy.AddMessage(checkfield)
                # arcpy.AddMessage(len(checkfield)*"-")
                
                newfields = {f.name:[f.type,f.length,f.domain] for f in arcpy.ListFields(newlyrpath)}
                oldfields = {f.name:[f.type,f.length,f.domain] for f in arcpy.ListFields(oldlyrpath)}
                
                for f in newfields.keys():
                    if newfields[f][0] != "String":
                       newfields[f][1] =None 
                for f in oldfields.keys():
                    if oldfields[f][0] != "String":
                       oldfields[f][1] =None 
                        

                def compareFields(newFieldNames,oldFieldNames):
                    # Fields Data Type
                    for f in newFieldNames.keys():
                        if f in oldFieldNames.keys() and newFieldNames[f][0] != oldFieldNames[f][0]:
                            totalReport.append([os.path.basename(parameters[0].valueAsText),checkfield,"Added",f,newFieldNames[f][0],"",""])
                        if f not in oldFieldNames.keys():
                            totalReport.append([os.path.basename(parameters[0].valueAsText),checkfield,"Added",f,"","",""])
                            
                    for f in oldFieldNames.keys():
                        if f in newFieldNames.keys() and newFieldNames[f][0] != oldFieldNames[f][0]:
                            totalReport.append([os.path.basename(parameters[1].valueAsText),checkfield,"Removed",f,oldFieldNames[f][0],"",""])
                        if f not in newFieldNames.keys():
                            totalReport.append([os.path.basename(parameters[1].valueAsText),checkfield,"Removed",f,"","",""])

                    # Fields Length
                    for f in newFieldNames.keys():
                        if f in oldFieldNames.keys() and newFieldNames[f][1] != oldFieldNames[f][1]:
                            totalReport.append([os.path.basename(parameters[0].valueAsText),checkfield,"Text Changed",f,"",newFieldNames[f][1],""])
                            
                    for f in oldFieldNames.keys():
                        if f in newFieldNames.keys() and newFieldNames[f][1] != oldFieldNames[f][1]:
                            totalReport.append([os.path.basename(parameters[1].valueAsText),checkfield,"Text Changed",f,"",oldFieldNames[f][1],""])

                    # Fields Domain
                    for f in newFieldNames.keys():
                        if f in oldFieldNames.keys() and newFieldNames[f][2] != oldFieldNames[f][2]:
                            totalReport.append([os.path.basename(parameters[0].valueAsText),checkfield,"Domain Changed",f,"","",newFieldNames[f][2]])
                            
                    for f in oldFieldNames.keys():
                        if f in newFieldNames.keys() and newFieldNames[f][2] != oldFieldNames[f][2]:
                            totalReport.append([os.path.basename(parameters[1].valueAsText),checkfield,"Domain Changed",f,"","",oldFieldNames[f][2]])                    

                compareFields(newfields,oldfields)
                # changemodified = []
                # for uniqueLayers in set([f[1] for f in totalReport]):
                #     for changeLayer in [i for i in totalReport if uniqueLayers == i[1]]:
                #         if changeLayer[3].lower() in [m[3].lower() for m in totalReport if changeLayer == f]:
                #             f[2] = "Modified"

            fieldReportDF= pd.DataFrame(totalReport,columns=["GDB","Feature Class","Status","Field Name","Field Type","Field Length","Field Domain"])
            fieldReportDF.to_excel(writer,sheet_name ='Field Names Report',index=False)
            


            newdomains = [f.name for f in arcpy.da.ListDomains(parameters[0].valueAsText)]
            olddomains = [f.name for f in arcpy.da.ListDomains(parameters[1].valueAsText)]
            
            domainsdata = []

            for f in newdomains:
                if f not in olddomains:
                    domainsdata.append([os.path.basename(parameters[0].valueAsText),"Added",f]) 
            for f in olddomains:
                if f not in newdomains:
                    domainsdata.append([os.path.basename(parameters[1].valueAsText),"Removed",f]) 
            domainmatched = []

            for f in newdomains:
                if f in olddomains:
                    domainsdata.append([os.path.basename(parameters[0].valueAsText),"Matching",f]) 
                    domainmatched.append(f.lower())

            for f in olddomains:
                if f in newdomains:
                    domainsdata.append([os.path.basename(parameters[1].valueAsText),"Matching",f]) 
                    domainmatched.append(f.lower())

            for f in newdomains:
                if f.lower() in [f.lower() for f in olddomains] and f.lower() not in domainmatched:
                    domainsdata.append([os.path.basename(parameters[0].valueAsText),"Modified",f]) 

            for f in olddomains:
                if f.lower() in [f.lower() for f in newdomains] and f.lower() not in domainmatched:
                    domainsdata.append([os.path.basename(parameters[1].valueAsText),"Modified",f]) 


            domainReportDF = pd.DataFrame(domainsdata,columns=["GDB","Status","Field Domain"])
            domainReportDF.to_excel(writer,sheet_name ='Domains Report',index=False)
            
            domainlist = []

            arcpy.AddMessage("======================= Compare Domain Name ==============================================")
            for newdomainlist in arcpy.da.ListDomains(parameters[0].valueAsText):
                def compareDomainList(new,old,domain):
                    # newDescr = {f.lower():new[f].lower() for f in new}
                    # oldDescr = {f.lower():old[f].lower() for f in old}
                    for f in new:
                        # Added Both
                        if f not in old.keys() and old.get(f)!=new[f]:
                            domainlist.append([os.path.basename(parameters[0].valueAsText),domain,"Added Both",f,new[f]])
                        else:
                            # Added
                            if f not in old.keys():
                                domainlist.append([os.path.basename(parameters[0].valueAsText),domain,"Added Code",f,new[f]])
                            # Removed
                            if new[f] != old.get(f):
                                domainlist.append([os.path.basename(parameters[0].valueAsText),domain,"Added Description",f,new[f]])

                    for f in old:
                        # Added Both
                        if f not in new.keys() and new.get(f)!=old[f]:
                            domainlist.append([os.path.basename(parameters[1].valueAsText),domain,"Removed Both",f,old[f]])
                        else:
                            # Added
                            if f not in new.keys():
                                domainlist.append([os.path.basename(parameters[1].valueAsText),domain,"Removed Code",f,old[f]])
                            # Removed
                            if old[f] != new.get(f):
                                domainlist.append([os.path.basename(parameters[1].valueAsText),domain,"Removed Description",f,old[f]])

                newDL = {}
                oldDL = []
                if newdomainlist.name in [f for f in newdomains if f in olddomains]:
                    if newdomainlist.domainType == "CodedValue":
                        newDL = newdomainlist.codedValues
                    elif newdomainlist.domainType == "Range":
                        newDL = newdomainlist.range
                    for olddomainlist in arcpy.da.ListDomains(parameters[1].valueAsText):
                        if newdomainlist.name == olddomainlist.name:
                            if olddomainlist.domainType == "CodedValue":
                                oldDL = olddomainlist.codedValues
                            elif olddomainlist.domainType == "Range":
                                oldDL = olddomainlist.range
                    if newDL!= oldDL:
                        arcpy.AddMessage(newdomainlist.name)
                        compareDomainList(newDL,oldDL,newdomainlist.name)

            domainListReportDF = pd.DataFrame(domainlist,columns=["GDB","Domain","Status","Code","Description"])
            domainListReportDF.to_excel(writer,sheet_name ='Domains List Report',index=False)

class DomainCheck(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "3. Domain check"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        GDB = arcpy.Parameter(
            displayName = "Enter Geodatabase to check",
            name = "GDB",
            datatype = "DEWorkspace",
            parameterType = "Required",
            direction = "Input")
        GDB.filter.list = ["Local Database"] 
        excel = arcpy.Parameter(
            displayName = "Enter Excel output",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        excel.filter.list = ['csv']
        return [GDB,excel]

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
        gdb = parameters[0].valueAsText
        listdomains = {f.name:f.codedValues for f in arcpy.da.ListDomains(gdb)}
        alldata = []
        def checkLayerField(gdb,name):
            for f in arcpy.ListFields(name):
                if f.domain !='':
                    data = []
                    with arcpy.da.SearchCursor(name,["OID@",f.name]) as cursor:
                        for row in cursor:
                            try:
                                if row[1] not in list(listdomains[f.domain].keys()):
                                    data.append([row[0],row[1]]) 
                            except:
                                pass
                    if len(data)!=0:
                        for val in list(set([f[1] for f in data])):
                            if val == None:
                                val = 'None'
                            alldata.append((os.path.basename(name),f.name,val))
                            arcpy.AddMessage((os.path.basename(name),f.name,val))
        
        featuresGDB = arcpy.da.Walk(parameters[0].valueAsText,datatype=["FeatureClass","Table"])
        for dirpath, dirnames, filenames in featuresGDB:
            for file in filenames:
                checkLayerField(gdb,os.path.join(dirpath,file))
        df = pd.DataFrame(alldata,columns=['Name','Domain','Value'])
        df.to_csv(parameters[1].valueAsText,index=False)

class DecimalCheckExcel(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "4. Schema Check with Excel"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        gdb = arcpy.Parameter(
            displayName = "Enter Geodatabase path",
            name = "gdb",
            datatype = "DEWorkspace",
            parameterType = "Required",
            direction = "Input")

        excel = arcpy.Parameter(
            displayName = "Enter excel  input",
            name = "excel",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Input")
        excel.filter.list = ['xlsx']

        export = arcpy.Parameter(
            displayName = "Enter export path",
            name = "export",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        export.filter.list = ['xlsx']        
        return [gdb,excel,export]

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


        path = parameters[0].valueAsText
        data = parameters[1].valueAsText

        # Get Features from GDB
        def FeatureClasses(path):
            fc=[]
            walk = arcpy.da.Walk(path, datatype="FeatureClass")
            for dirpath, dirnames, filenames in walk:
                for filename in filenames:
                    fc.append(os.path.join(dirpath, filename))
            return fc
        
        #excel Writer
        with pd.ExcelWriter(parameters[2].valueAsText) as writer:
            ErrorReport = []
            domains = arcpy.da.ListDomains(path)
            gdbDomains = {}
            
            # Get Domains
            for domain in domains:
                if domain.domainType == "CodedValue":
                    gdbDomains[domain.name] = {val:desc for val, desc in domain.codedValues.items() if domain.domainType == "CodedValue"}

            fieldslength = {}
            linkDomain = {}

            fc =FeatureClasses(path)

            sheets = pd.ExcelFile(data).sheet_names

            # Find Domain Error Data
            arcpy.AddMessage(f'==========Domain Error==========')
            pfldata = []
            for l in fc:
                if os.path.basename(l) in sheets:
                    gisfields = [f.name for f in arcpy.ListFields(l) if f.domain!='']
                    arcpy.AddMessage(f'==========={os.path.basename(l)}===========')
                    arcpy.AddMessage(f'Fields containing domains: {gisfields}')
                    for g in gisfields:
                        ErrorReport.append(['Domain',os.path.basename(l),'Compared',g])
                    linkDomain[os.path.basename(l)] = {f.name:f.domain for f in arcpy.ListFields(l) if f.domain!=''}
                    fieldslength[os.path.basename(l)] = {f.name:f.length for f in arcpy.ListFields(l) if f.domain!=''}
                    if len(gisfields)!=0:
                        df = pd.read_excel(data,sheet_name=os.path.basename(l)).fillna('')
                        if len([f for f in gisfields if f not in df.columns])>0:
                            for g in [f for f in gisfields if f not in df.columns]:
                                ErrorReport.append(['Domain',os.path.basename(l),'Ignored',g])
                            arcpy.AddMessage(f'Field name not matching with excel: {[f for f in gisfields if f not in df.columns]}')
                        newdf = df[[f for f in gisfields if f in df.columns]]
                        for i in newdf.index:
                            for f in newdf.columns:
                                if newdf.iloc[i][list(newdf.columns).index(f)] not in gdbDomains[linkDomain[os.path.basename(l)][f]].keys():
                                    pfldata.append([os.path.basename(l),f,newdf.iloc[i][list(newdf.columns).index(f)]])
                else:
                    arcpy.AddMessage(f'=========== Layer Not Checked: {os.path.basename(l)}')
            df = pd.DataFrame(pfldata,columns=['layer','Field','Value'])
            # arcpy.AddMessage(df)
            newdf = df.fillna('Blank').groupby(['layer','Field'])['Value'].unique().apply(list).reset_index()
            newdf.to_excel(writer,sheet_name='DomainError',index=False)

            # Find Null Error Data
            arcpy.AddMessage(f'==========Null Error==========')
            NullData = []
            for f in fc:
                if os.path.basename(f) in sheets:
                    arcpy.AddMessage(f'========{os.path.basename(f)}========')
                    gisNull = [f.name for f in arcpy.ListFields(f) if f.isNullable==False and f.name!='OBJECTID']
                    df = pd.read_excel(data,sheet_name=os.path.basename(f))
                    for g in gisNull:
                        ErrorReport.append(['Null',os.path.basename(f),'Compared',g])
                    if len([f for f in gisNull if f not in list(df.columns)])>0:
                        for g in [f for f in gisNull if f not in list(df.columns)]:
                            ErrorReport.append(['Null',os.path.basename(f),'Ignored',g])
                        
                    arcpy.AddMessage(f'Considered Fields: {[f for f in gisNull if f in list(df.columns)]}')
                    arcpy.AddMessage(f'Ignored Fields: {[f for f in gisNull if f not in list(df.columns)]}')
                    if len([f for f in gisNull if f in list(df.columns)])>0:
                        for field in [f for f in gisNull if f in list(df.columns)]:
                            # arcpy.AddMessage(set(pd.isnull(df[field]).tolist()))
                            if True in set(pd.isnull(df[field]).tolist()):
                                NullData.append([os.path.basename(f),field])
                else:
                    arcpy.AddMessage(f"Absent: {os.path.basename(f)}")
            df = pd.DataFrame(NullData,columns=['layer','field'])
            df.to_excel(writer,sheet_name='NullError',index=False)

            #length Check
            arcpy.AddMessage(f'==========Length Errorr==========')
            LengthData = []
            for f in fc:
                if os.path.basename(f) in sheets:
                    arcpy.AddMessage(f'========{os.path.basename(f)}========')
                    gisFields = [f.name for f in arcpy.ListFields(f) if f.type=='String' and f.domain=='']
                    gisLength = [f.length for f in arcpy.ListFields(f) if f.type=='String' and f.domain=='']

                            
                    df = pd.read_excel(data,sheet_name=os.path.basename(f)).fillna('')
                    arcpy.AddMessage(f'Considered Fields: {[f for f in gisFields if f in list(df.columns)]}')
                    arcpy.AddMessage(f'Ignored Fields: {[f for f in gisFields if f not in list(df.columns)]}')
                    for g in [f for f in gisFields if f in list(df.columns)]:
                        ErrorReport.append(['Length',os.path.basename(f),'Compared',g])
                    if len([f for f in gisFields if f not in list(df.columns)])>0:
                        for g in [f for f in gisFields if f not in list(df.columns)]:
                            ErrorReport.append(['Length',os.path.basename(f),'Ignored',g])

                    if len([f for f in gisFields if f in list(df.columns)])>0:
                        newdf = df[[f for f in gisFields if f in list(df.columns)]]
                        for i in newdf.index:
                            for field in newdf.columns:
                                if isinstance(newdf.iloc[i][field], float)==True:
                                    
                                    if abs(Decimal(str(newdf.iloc[i][field])).as_tuple().exponent)>2:
                                        LengthData.append([os.path.basename(f),field,2,newdf.iloc[i][field]])
                                    if abs(Decimal(str(newdf.iloc[i][field])).as_tuple().exponent)>4:
                                        LengthData.append([os.path.basename(f),field,4,newdf.iloc[i][field]])
                                else:
                                    if len(str(newdf.iloc[i][field]))>gisLength[gisFields.index(field)]:
                                        LengthData.append([os.path.basename(f),field,gisLength[gisFields.index(field)],newdf.iloc[i][field]])
                else:
                    arcpy.AddMessage(f"Absent: {os.path.basename(f)}")
            df = pd.DataFrame(LengthData,columns=['layer','field','CharLength','value'])
            df.to_excel(writer,sheet_name='LengthError',index=False)

            #Double Check
            arcpy.AddMessage(f'==========Double Errorr==========')
            LengthData = []
            for f in fc:
                if os.path.basename(f) in sheets:
                    arcpy.AddMessage(f'========{os.path.basename(f)}========')
                    gisFields = [f.name for f in arcpy.ListFields(f) if f.type=='Double']
                    # gisLength = [f.length for f in arcpy.ListFields(f) if f.type=='Double']

                            
                    df = pd.read_excel(data,sheet_name=os.path.basename(f)).fillna('')
                    arcpy.AddMessage(f'Considered Fields: {[f for f in gisFields if f in list(df.columns)]}')
                    arcpy.AddMessage(f'Ignored Fields: {[f for f in gisFields if f not in list(df.columns)]}')

                    for g in [f for f in gisFields if f in list(df.columns)]:
                        ErrorReport.append(['Double',os.path.basename(f),'Compared',g])
                    if len([f for f in gisFields if f not in list(df.columns)])>0:
                        for g in [f for f in gisFields if f not in list(df.columns)]:
                            ErrorReport.append(['Double',os.path.basename(f),'Ignored',g])

                    if len([f for f in gisFields if f in list(df.columns)])>0:
                        newdf = df[[f for f in gisFields if f in list(df.columns)]]
                        for i in newdf.index:
                            for field in newdf.columns:
                                if isinstance(newdf.iloc[i][field], float)==True:
                                    arcpy.AddMessage(type(newdf.iloc[i][field]))
                                # arcpy.AddMessage(newdf.iloc[i][field])
                                # arcpy.AddMessage(str(newdf.iloc[i][field]))
                                # arcpy.AddMessage(Decimal(str(newdf.iloc[i][field])).as_tuple().exponent)
                                # if newdf.iloc[i][field] !=None:
                                    if abs(Decimal(str(newdf.iloc[i][field])).as_tuple().exponent)>2:
                                        arcpy.AddMessage([newdf.iloc[i][field],Decimal(str(newdf.iloc[i][field])).as_tuple().exponent])
                                        LengthData.append([os.path.basename(f),field,2,newdf.iloc[i][field]])
                                    if abs(Decimal(str(newdf.iloc[i][field])).as_tuple().exponent)>4:
                                        LengthData.append([os.path.basename(f),field,4,newdf.iloc[i][field]])
                                # else:
                                #     if len(str(newdf.iloc[i][field]))>gisLength[gisFields.index(field)]:
                                #         LengthData.append([os.path.basename(f),field,gisLength[gisFields.index(field)],newdf.iloc[i][field]])
                else:
                    arcpy.AddMessage(f"Absent: {os.path.basename(f)}")
            arcpy.AddMessage(LengthData)
            arcpy.AddMessage(ErrorReport)
            if len(LengthData)>0:
                df = pd.DataFrame(LengthData,columns=['layer','field','CharLength','value'])
                df.to_excel(writer,sheet_name='Double',index=False)
            if len(ErrorReport)>0:
                df = pd.DataFrame(ErrorReport,columns=['Check','Layer','Status','Fields'])
                df.to_excel(writer,sheet_name='Report',index=False)
  