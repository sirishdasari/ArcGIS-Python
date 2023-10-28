# -*- coding: utf-8 -*-

import arcpy
import pandas as pd
from operator import itemgetter

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
        self.tools = [PFL,GISIntegrationAll,InsertData,PFL2,InsertData2]
class PFL(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "GIS to PFL"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        pipe = arcpy.Parameter(
            displayName = "Enter Pipe Feature",
            name = "pipe",
            datatype = "GPFeatureLayer",
            parameterType = "Required",
            direction = "Input")
        # pipe.value = "P_Pipes_INSV_All"
        lineLayers = arcpy.Parameter(
            displayName = "Enter Line Layers",
            name = "line",
            datatype = "GPFeatureLayer",
            parameterType = "Optional",
            direction = "Input",
            multiValue=True)
        lineLayers.columns = [['GPFeatureLayer', 'Layer']]
        # lineLayers.values = [['P_PipeCasing_INSV_All']]

        pointLayers = arcpy.Parameter(
            displayName = "Enter Point Layers",
            name = "point",
            datatype = "GPFeatureLayer",
            parameterType = "Optional",
            direction = "Input",
            multiValue=True)
        pointLayers.columns = [['GPFeatureLayer', 'Layer']]
        # pointLayers.values = [['P_ControllableFitting_INSV_All'],['P_NonControllableFitting_INSV_All'],['P_Valve_INSV_All']]

        FieldMapping = arcpy.Parameter(
            displayName = "Enter Field Mappings",
            name = "FieldMapping",
            datatype = "DETable",
            parameterType = "Optional",
            direction = "Input")
        # FieldMapping.value = "Field_Mapping_Template"

        FMTemplate = arcpy.Parameter(
            displayName = "Field mapping Template",
            name = "FMTemplate",
            datatype = "GPValueTable",
            parameterType = "Optional",
            direction = "Input")
        FMTemplate.columns = [['GPString', 'Layer'], ['GPString', 'GIS'], ['GPString', 'PFL']]
        # FMTemplate.values = [['FeatureLayer', 'GIS', 'PFL']]

        pflTemplate = arcpy.Parameter(
            displayName = "Enter MAOP table template",
            name = "pflTemplate",
            datatype = "DETable",
            parameterType = "Optional",
            direction = "Input")
        # pflTemplate.value = "MAOP_Gas_PFLTemplate_V12"
        pflFields = arcpy.Parameter(
            displayName = "Enter PFL Fields given below",
            name = "pflFields",
            datatype = "GPValueTable",
            parameterType = "Optional",
            direction = "Input")
        pflFields.columns = [["GPString","From Measure"],["GPString","To Measure"],["GPString","Route/TR"]]
        # pflFields.values = [['From_Measure', 'To_Measure', 'TR']]
        table = arcpy.Parameter(
            displayName = "Enter output for table",
            name = "table",
            datatype = "DETable",
            parameterType = "Required",
            direction = "Output")
        # layerMapping.columns = [["GPString","P_Pipe"],["GPString","P_Casing"]]
        # # layerMapping.values = [['P_Pipes INSV_All', 'P_PipeCasing_INSV_All']]
        return [pipe,lineLayers,pointLayers,FieldMapping,FMTemplate,pflTemplate,pflFields,table]

    def updateParameters(self, parameters):
        fieldMapping = parameters[3]
        fieldMappingsList = parameters[4]
        assignFields = parameters[5]
        assignFieldsList = parameters[6]


        if fieldMapping.altered:
            fieldMappingsList.filters[0].list = [f.name for f in arcpy.ListFields(fieldMapping.valueAsText)]
            fieldMappingsList.filters[1].list = [f.name for f in arcpy.ListFields(fieldMapping.valueAsText)]
            fieldMappingsList.filters[2].list = [f.name for f in arcpy.ListFields(fieldMapping.valueAsText)]
        if assignFields.altered:
            assignFieldsList.filters[0].list = [f.name for f in arcpy.ListFields(assignFields.valueAsText)]
            assignFieldsList.filters[1].list = [f.name for f in arcpy.ListFields(assignFields.valueAsText)]
            assignFieldsList.filters[2].list = [f.name for f in arcpy.ListFields(assignFields.valueAsText)]
        return
    def updateMessages(self, parameters):
        # if parameters[1].altered and parameters[2].altered and parameters[3].altered and parameters[1].value!=None and parameters[2].value!=None and parameters[4].value!=None:
        #     layers = [f[0].name for f in parameters[0].value]
        #     newdata = []
        #     templateNames = [f.aliasName for f in arcpy.ListFields(parameters[4].valueAsText)]
        #     for layer in layers:
        #         gis = {row[1]:row[2] for row in arcpy.da.SearchCursor(parameters[2].valueAsText,parameters[3].value[0]) if row[0] == layer and row[2] not in ['',None] }
        #         for k in gis.values():
        #             if k not in templateNames:
        #                 newdata.append(f"{k} field not in {layer}")
        #     if len(newdata)!=0:
        #         parameters[3].setWarningMessage(newdata)
        #     if len(parameters[2].valueAsText.split(" ")) != 4:
        #         parameters[2].setWarningMessage('Remove space if present in field names')
        #     else:
        #         parameters[2].clearMessage()
        return

    def execute(self, parameters, messages):
        # arcpy.AddMessage("Processing ...")
        Pipe = parameters[0].valueAsText
        lineLayer = parameters[1].value
        pointLayer =  parameters[2].value
        fieldMapping =  parameters[3].valueAsText
        fieldMappingList = parameters[4].value
        MAOPtemplate = parameters[5].valueAsText
        MAOPtemplateFields = parameters[6].value
        ExportOutput = parameters[7].valueAsText

        # PFL Fields dictionary
        arcpy.management.CopyRows(MAOPtemplate, "in_memory/PFL")
        # arcpy.management.CopyRows(MAOPtemplate, "in_memory/Final")
        arcpy.management.CopyRows(MAOPtemplate, "in_memory/Line")
        arcpy.management.CopyRows(MAOPtemplate, "in_memory/Point")

        # Dict PFL aliasName:name 
        pfl = {f.aliasName:f.name for f in arcpy.ListFields("in_memory/PFL")}

        # All Field Names including ObjectID
        allPFLfields = [f.name for f in arcpy.ListFields("in_memory/PFL")]
        insertToPFL = arcpy.da.InsertCursor("in_memory/PFL",allPFLfields)

        # Get Layers names from list
        if parameters[2].value !=None:
            pointLayers = [f[0].name for f in pointLayer]
        if parameters[1].value !=None:
            lineLayers = [f[0].name for f in lineLayer]

        # Get index with list as input
        def pflIndex(fields):
            return [allPFLfields.index(l) for l in fields]


        fromIndex = pflIndex(MAOPtemplateFields[0])[0]
        toIndex = pflIndex(MAOPtemplateFields[0])[1]
        trIndex = pflIndex(MAOPtemplateFields[0])[2]
        
        def fieldsGIStoPFL(layer):
            if parameters[1].value==None:
                processlayers=pointLayers+[Pipe]
                arcpy.AddMessage(processlayers)
            else:
                processlayers=lineLayers+pointLayers+[Pipe]
                arcpy.AddMessage(processlayers)
            for l in processlayers:
                if l==layer:
                    GISfields = []
                    PFLfields = []

                    GIStoPFL = {row[1]:row[2] for row in arcpy.da.SearchCursor(fieldMapping,fieldMappingList[0]) if row[0] == layer and row[2] not in ['',None]}
                    # arcpy.AddMessage(GIStoPFL)
                    for f in GIStoPFL.keys():
                        GISfields.append(f)
                        PFLfields.append(pfl[GIStoPFL[f]])
                    return GISfields, PFLfields
        
        def processPipe(lname):

            arcpy.AddMessage(fieldsGIStoPFL(lname))
            insert = arcpy.da.InsertCursor("in_memory/PFL",fieldsGIStoPFL(lname)[1])
            data = [list(row) for row in arcpy.da.SearchCursor(lname,fieldsGIStoPFL(lname)[0])]
            for d in data:
                insert.insertRow(d)
            del insert

            # count = len(data)
            # arcpy.SetProgressor("step", "Processing...",0, count, 1)
            with arcpy.da.UpdateCursor("in_memory/PFL", allPFLfields) as cursor:
                # n=0
                for row in cursor:
                    # n=n+1
                    # arcpy.SetProgressorLabel(f"Convert {parameters[0].valueAsText} to PFL {0} of {1}".format(n,count))
                    row[fromIndex] = round(row[fromIndex],2)
                    row[toIndex] = round(row[toIndex],2)
                    cursor.updateRow(row)
                    # arcpy.SetProgressorPosition()
                # arcpy.ResetProgressor()


 
        LayerInmemory = []
        def processLineLayers(lname):
            for layer in lname:
                arcpy.management.CopyRows(MAOPtemplate, f"in_memory/{layer}")
                LayerInmemory.append(f"in_memory/{layer}")
                insertLine = arcpy.da.InsertCursor(f"in_memory/{layer}",fieldsGIStoPFL(layer)[1])
                data = [list(row) for row in arcpy.da.SearchCursor(layer,fieldsGIStoPFL(layer)[0])]
                for d in data:
                    insertLine.insertRow(d)
                del insertLine

                with arcpy.da.UpdateCursor(f"in_memory/{layer}", allPFLfields) as cursor:
                    for row in cursor:
                        row[fromIndex] = round(row[fromIndex],2)
                        row[toIndex] = round(row[toIndex],2)
                        cursor.updateRow(row)

        def processPointLayers(lname):
            for layer in lname:
                insertPoint = arcpy.da.InsertCursor("in_memory/Point",fieldsGIStoPFL(layer)[1])
                data = [list(row) for row in arcpy.da.SearchCursor(layer,fieldsGIStoPFL(layer)[0])]
                for d in data:
                    insertPoint.insertRow(d)
                del insertPoint


            with arcpy.da.UpdateCursor("in_memory/Point", allPFLfields) as cursor:
                for row in cursor:
                    row[fromIndex] = round(row[fromIndex],2)
                    row[toIndex] = row[fromIndex]
                    cursor.updateRow(row)
        
        def linePoints(layer,id):
            pflSplit = []
            with arcpy.da.SearchCursor(layer,allPFLfields) as cursor:
                for row in cursor:
                    if row[trIndex] == id:
                        pflSplit.append(row[1])
                        pflSplit.append(row[2])
            return sorted(set(pflSplit))
        
        def getPoints(layer,id):
            pflSplit = []
            with arcpy.da.SearchCursor(layer,allPFLfields) as cursor:
                for row in cursor:
                    if row[trIndex] == id:
                        pflSplit.append(row[1])
            return sorted(set(pflSplit))

        processPipe(Pipe)        
        processPointLayers(pointLayers) 
        if parameters[1].value!=None:
            processLineLayers(lineLayers) 

        for layer in LayerInmemory:
            count = int(arcpy.management.GetCount("in_memory/PFL").getOutput(0))
            arcpy.SetProgressor("step", "Processing...",0, count, 1)
            with arcpy.da.UpdateCursor("in_memory/PFL",allPFLfields) as cursor:
                n=0
                for row in cursor:
                    n=n+1
                    arcpy.SetProgressorLabel(f"Splitting PFL from {layer[10:]}: {n} of {count}")
                    points = [f for f in linePoints(layer,row[trIndex]) if row[fromIndex]<f<row[toIndex]]
                    if len(points)!=0:
                        insert = [row[fromIndex]]+points+[row[toIndex]]
                        for i in range(len(insert)-1):
                            row[fromIndex] = insert[i]
                            row[toIndex] = insert[i+1]
                            insertToPFL.insertRow([int(arcpy.management.GetCount("in_memory/PFL").getOutput(0))]+row[1:])
                        cursor.deleteRow()
                    arcpy.SetProgressorPosition()
                arcpy.ResetProgressor()
            del cursor

        for layer in LayerInmemory:
            count = int(arcpy.management.GetCount("in_memory/PFL").getOutput(0))
            arcpy.SetProgressor("step", "Processing...",0, count, 1)
            with arcpy.da.UpdateCursor("in_memory/PFL",allPFLfields) as cursor:
                n=0
                for pipes in cursor:
                    n=n+1
                    arcpy.SetProgressorLabel(f"Updating PFL fields from {layer[10:]}: {n} of {count}")
                    otherShort  =   sorted([row for row in arcpy.da.SearchCursor(layer,allPFLfields) if pipes[fromIndex]<=row[fromIndex]<=row[toIndex]<=pipes[toIndex]],key=itemgetter(fromIndex,toIndex))
                    for short in otherShort:
                        for new in [f for f in fieldsGIStoPFL(layer[10:])[1] if f not in MAOPtemplateFields[0]]:
                            pipes[pflIndex([new])[0]] = short[pflIndex([new])[0]]
                        cursor.updateRow(pipes)
                    arcpy.SetProgressorPosition()
                arcpy.ResetProgressor()
            del cursor

        with arcpy.da.UpdateCursor("in_memory/PFL",allPFLfields) as cursor:
            count = int(arcpy.management.GetCount("in_memory/PFL").getOutput(0))
            arcpy.SetProgressor("step", "Processing...",0, count, 1)
            n=0
            for row in cursor:
                n=n+1
                arcpy.SetProgressorLabel(f"Splitting PFL from Point Layer: {n} of {count}")
                # arcpy.AddMessage(row[trIndex])
                points = []
                # f for f in getPoints("in_memory/Point",row[trIndex]) if row[fromIndex]<f<row[toIndex]and row[fromIndex]!=None and row[toIndex]!=None and f!=None
                for pt in getPoints("in_memory/Point",row[trIndex]):
                    if row[fromIndex]!=None and row[toIndex]!=None and pt!=None:
                        if row[fromIndex]<pt<row[toIndex]:
                            points.append(pt)
                # arcpy.AddMessage([row[1],row[2],"=========",points])
                if len(points)!=0:
                    insert = [row[fromIndex]]+points+[row[toIndex]]
                    for i in range(len(insert)-1):
                        row[fromIndex] = insert[i]
                        row[toIndex] = insert[i+1]
                        insertToPFL.insertRow([int(arcpy.management.GetCount("in_memory/PFL").getOutput(0))]+row[1:])
                    cursor.deleteRow()
                arcpy.SetProgressorPosition()
            arcpy.ResetProgressor()
        del cursor

        count = int(arcpy.management.GetCount("in_memory/Point").getOutput(0))
        arcpy.SetProgressor("step", "Processing...",0, count, 1)
        n=0
        for f in [row for row in arcpy.da.SearchCursor("in_memory/Point",allPFLfields)]:
            n=n+1
            arcpy.SetProgressorLabel(f"Updating Point to PFL: {n} of {count}")
            insertToPFL.insertRow([int(arcpy.management.GetCount("in_memory/PFL").getOutput(0))]+list(f[1:]))
            arcpy.SetProgressorPosition()
        arcpy.ResetProgressor()
        arcpy.management.CopyRows("in_memory/PFL", ExportOutput)
class GISIntegrationAll(object):

    def __init__(self):
        self.label = "PFL to GIS"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        pfl = arcpy.Parameter(
            displayName = "Enter PFL path",
            name = "pfl",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Input")
        pfl.filter.list = ["xlsx","xls"]
        sheet = arcpy.Parameter(
            displayName = "Enter sheet name",
            name = "sheet",
            datatype = "GPString",
            parameterType = "Required",
            direction = "Input")
        
        pflFields = arcpy.Parameter(
            displayName = "Enter feature field",
            name = "pflFields",
            datatype = "GPString",
            parameterType = "Required",
            direction = "Input")
        # pflFields.columns = [['GPString', 'Feature']]
        # pflFields.parameterDependencies = [pfl.name]

        fieldMapping = arcpy.Parameter(
            displayName = "Enter field mapping path",
            name = "fieldMapping",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Input")
        # fieldMapping.filter.list = ["xlsx","xls"]
        fieldMappingFields = arcpy.Parameter(
            displayName = "Enter field mapping fields",
            name = "fieldMappingFields",
            datatype = "GPValueTable",
            parameterType = "Required",
            direction = "Input")
        fieldMappingFields.columns = [['GPString', 'Layer'],['GPString', 'GIS Field'], ['GPString', 'PFL Field']]
        fieldMappingFields.parameterDependencies = [fieldMapping.name]

        Feature = arcpy.Parameter(
            displayName = "Enter Feature Mapping Path",
            name = "Feature",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Input")
        
        FeatureMapping = arcpy.Parameter(
            displayName = "Enter Feature Mapping Fields",
            name = "FeatureMapping",
            datatype = "GPValueTable",
            parameterType = "Required",
            direction = "Input")
        FeatureMapping.columns = [['GPString', 'Layer'],['GPString', 'Feature']]

        export = arcpy.Parameter(
            displayName = "Enter export path",
            name = "export",
            datatype = "DEFolder",
            parameterType = "Required",
            direction = "Input")
        export.filter.list = ["xlsx","xls"]
        # export.filter.list = ["xlsx","xls"]
        return [pfl,sheet,pflFields,fieldMapping,fieldMappingFields,Feature,FeatureMapping,export]
        # return [pfl]
        
    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        if parameters[0].altered:
            fields = pd.ExcelFile(parameters[0].valueAsText).sheet_names
            parameters[1].filter.list = fields
        if parameters[1].altered:
            fields = pd.read_excel(parameters[0].valueAsText,sheet_name=parameters[1].valueAsText).head().columns.tolist()
            parameters[2].filter.list = fields
        if parameters[3].altered:
            fields = [f.name for f in arcpy.ListFields(parameters[3].valueAsText)]
            parameters[4].filters[0].list = fields
            parameters[4].filters[1].list = fields
            parameters[4].filters[2].list = fields
        if parameters[5].altered:
            fields = [f.name for f in arcpy.ListFields(parameters[5].valueAsText)]
            parameters[6].filters[0].list = fields
            parameters[6].filters[1].list = fields
        return

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        # arcpy.AddMessage("Add Message")

        # arcpy.AddMessage(list(set(pd.read_csv(parameters[5].valueAsText).fillna('')[parameters[6].value[0][0]].tolist())))
        fieldMapping = parameters[4].value[0]
        featureMapping = parameters[6].value[0]
        # read feature mapping
        featureMap = pd.read_csv(parameters[5].valueAsText).fillna('')[featureMapping].groupby(featureMapping[0])[featureMapping[1]].apply(list).to_dict()

        # read field mapping
        fieldMapGIS = pd.read_csv(parameters[3].valueAsText).fillna('')[fieldMapping].groupby(fieldMapping[0])[fieldMapping[1]].apply(list).to_dict()
        fieldMapPFL = pd.read_csv(parameters[3].valueAsText).fillna('')[fieldMapping].groupby(fieldMapping[0])[fieldMapping[2]].apply(list).to_dict()
        fieldMerge = pd.read_csv(parameters[3].valueAsText)[fieldMapping].dropna().groupby([fieldMapping[0]])[fieldMapping[1]].apply(list).to_dict()
        arcpy.AddMessage(f"Performing Dissolve on:- {list(fieldMerge.keys())}")
        for l in featureMap.keys():
            if l in fieldMapGIS:
                arcpy.AddMessage(l)
                # arcpy.AddMessage([fieldMapGIS[l][fieldMapPFL[l].index(f)] for f in parameters[2].value[0][-2:]])#fieldMapGIS[l][fieldMapPFL[l].index(f)] for f in parameters[2].value[0][-2:])
                fm = {}
                for gis in fieldMapGIS[l]:
                    if fieldMapPFL[l][fieldMapGIS[l].index(gis)] != '':
                        fm[gis] = fieldMapPFL[l][fieldMapGIS[l].index(gis)]
                mappingFields = {fm[f]:f for f in fm.keys()}
                pfl = pd.read_excel(parameters[0].valueAsText,sheet_name=parameters[1].valueAsText).fillna('')
                pflFields = [f for f in fieldMapPFL[l] if f != '']
                gisFields = list(set([f for f in fieldMapGIS[l] if f != '']))
                pflData = pfl[pfl[parameters[2].valueAsText].isin(featureMap[l])][pflFields]
                
                for pfl in pflData.columns.tolist():
                    if pfl in fieldMapPFL[l]:
                        pflData.rename(columns=mappingFields,inplace=True)
                
                data = []
                for fieldPFL in pflData.values.tolist():
                    fieldsData = []
                    for gis in gisFields:
                        if gis in pflData.columns.tolist():
                            fieldsData.append(fieldPFL[pflData.columns.tolist().index(gis)])
                        else:
                            fieldsData.append("")
                    data.append(fieldsData)
                mitigationActions = pd.DataFrame(data,columns=gisFields)
                # arcpy.AddMessage(f'{parameters[7].valueAsText}\{l}.csv')
                mitigationActions[fieldMapGIS[l]].to_excel(f'{parameters[7].valueAsText}\{l}.xlsx',index=False)

class InsertData(object):

    def __init__(self):
        self.label = "With ID Insert data"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        input = arcpy.Parameter(
            displayName = "Enter input",
            name = "input",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Input")
        input.filter.list = ['xlsx']
        sheet = arcpy.Parameter(
            displayName = "Enter sheet",
            name = "sheet",
            datatype = "GPString",
            parameterType = "Required",
            direction = "Input")
        idField = arcpy.Parameter(
            displayName = "Enter ID Field",
            name = "idField",
            datatype = "GPString",
            parameterType = "Required",
            direction = "Input")    
        layer = arcpy.Parameter(
            displayName = "Enter layer",
            name = "layer",
            datatype = "GPTableView",
            parameterType = "Required",
            direction = "Input")
        idLayer = arcpy.Parameter(
            displayName = "Enter Layer Facility ID",
            name = "idLayer",
            datatype = "GPString",
            parameterType = "Required",
            direction = "Input") 
        fillFields = arcpy.Parameter(
            displayName = "Enter Fill Fields",
            name = "fillFields",
            datatype = "GPValueTable",
            parameterType = "Required",
            direction = "Input")
        # fillFields.parameterDependencies = [layer.name]
        fillFields.columns = [['GPString','PFL'],['GPString','GIS']]
        export = arcpy.Parameter(
            displayName = "Enter Export table",
            name = "export",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output")
        export.filter.list = ['xlsx']
        return [input,sheet,idField,layer,idLayer,fillFields,export]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        #sheet
        if parameters[0].altered:
            parameters[1].filter.list = pd.ExcelFile(parameters[0].valueAstext).sheet_names
        #PFL ID
        if parameters[1].altered:
            parameters[2].filter.list = sorted(pd.read_excel(parameters[0].valueAsText,sheet_name = parameters[1].valueAsText).columns.tolist())
            parameters[5].filters[0].list = sorted(pd.read_excel(parameters[0].valueAsText,sheet_name = parameters[1].valueAsText).columns.tolist())
        if parameters[3].altered:
            parameters[4].filter.list = sorted([f.name for f in arcpy.ListFields(parameters[3].valueAstext)])
            parameters[5].filters[1].list = sorted([f.name for f in arcpy.ListFields(parameters[3].valueAstext)])

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        pflfields = [f[0] for f in parameters[5].value]
        gisfields = [f[1] for f in parameters[5].value]

        df = pd.read_excel(parameters[0].valueAsText,sheet_name = parameters[1].valueAsText).fillna("")

        fields = [f.name for f in arcpy.ListFields(parameters[3].valueAsText)]
        data = [row for row in arcpy.da.SearchCursor(parameters[3].valueAsText,fields)]
        dfgis = pd.DataFrame(data,columns = fields)[[parameters[4].valueAsText]+gisfields].dropna().set_index(parameters[4].valueAsText).to_dict()

        col = df.columns.to_list()
        for i in df.index:
            for f in pflfields:
                df.iloc[i,col.index(f)] = dfgis[gisfields[pflfields.index(f)]].get(df.iloc[i,col.index(parameters[4].valueAstext)])
    
        writer = pd.ExcelWriter(parameters[6].valueAsText)
        df.to_excel(writer,sheet_name=parameters[1].valueAsText,index=False)
        writer.close()

class InsertData2(object):

    def __init__(self):
        self.label = "With ID Insert data 2"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        input = arcpy.Parameter(
            displayName = "Enter input",
            name = "input",
            datatype = "DEFile",
            parameterType = "Optional",
            direction = "Input")
        input.filter.list = ['xlsx']
        sheet = arcpy.Parameter(
            displayName = "Enter sheet",
            name = "sheet",
            datatype = "GPString",
            parameterType = "Optional",
            direction = "Input")
        idField = arcpy.Parameter(
            displayName = "Enter ID Field",
            name = "idField",
            datatype = "GPString",
            parameterType = "Optional",
            direction = "Input")    
        layer = arcpy.Parameter(
            displayName = "Enter layer",
            name = "layer",
            datatype = "GPTableView",
            parameterType = "Optional",
            direction = "Input")
        idLayer = arcpy.Parameter(
            displayName = "Enter Layer Facility ID",
            name = "idLayer",
            datatype = "GPString",
            parameterType = "Optional",
            direction = "Input") 
        fieldMapping = arcpy.Parameter(
            displayName = "Enter GIS to GIS Field Mapping",
            name = "fieldMapping",
            datatype = "DEFile",
            parameterType = "Optional",
            direction = "Input")
        fieldMapping.filter.list = ['csv']
        fillFields = arcpy.Parameter(
            displayName = "Enter Fill Fields",
            name = "fillFields",
            datatype = "GPValueTable",
            parameterType = "Optional",
            direction = "Input")
        # fillFields.parameterDependencies = [layer.name]
        fillFields.columns = [['GPString','Layer'],['GPString','CSV'],['GPString','GIS']]
        LayerMapField = arcpy.Parameter(
            displayName = "Enter mapping layer",
            name = "LayerMapField",
            datatype = "GPString",
            parameterType = "Optional",
            direction = "Input")
        export = arcpy.Parameter(
            displayName = "Enter Export table",
            name = "export",
            datatype = "DEFile",
            parameterType = "Optional",
            direction = "Output")
        export.filter.list = ['xlsx']
        return [input,sheet,idField,layer,idLayer,fieldMapping,fillFields,LayerMapField,export]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        #sheet
        if parameters[0].altered:
            parameters[1].filter.list = pd.ExcelFile(parameters[0].valueAstext).sheet_names
        #PFL ID
        if parameters[1].altered:
            parameters[2].filter.list = sorted(pd.read_excel(parameters[0].valueAsText,sheet_name = parameters[1].valueAsText).columns.tolist())
        if parameters[3].altered:
            parameters[4].filter.list = sorted([f.name for f in arcpy.ListFields(parameters[3].valueAstext)])
            
        if parameters[5].altered:
            parameters[6].filters[0].list = pd.read_csv(parameters[5].valueAsText).columns.tolist()
            parameters[6].filters[1].list = pd.read_csv(parameters[5].valueAsText).columns.tolist()
            parameters[6].filters[2].list = pd.read_csv(parameters[5].valueAsText).columns.tolist()
        if parameters[6].altered:
            parameters[7].filter.list = list(set(pd.read_csv(parameters[5].valueAsText)[parameters[6].value[0][0]]))


    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        dataframe = pd.read_csv(parameters[5].valueAsText).fillna("")
        # arcpy.AddMessage(parameters[6].value[0][1])
        # arcpy.AddMessage(list(set(dataframe[parameters[6].value[0][0]])))
        df2 = dataframe[(dataframe[parameters[6].value[0][0]]==parameters[7].valueAsText)]
        fieldList = [f for f in df2[[parameters[6].value[0][1],parameters[6].value[0][2]]].values.tolist() if f[1] !=""]

        csvgisfields = [f[0] for f in fieldList]
        gisfields = [f[1] for f in fieldList]

        # arcpy.AddMessage(csvgisfields)

        df = pd.read_excel(parameters[0].valueAsText,sheet_name = parameters[1].valueAsText).fillna("")
        data = {row[gisfields.index(parameters[2].valueAsText)]:row for row in arcpy.da.SearchCursor(parameters[3].valueAsText,gisfields)}
        # arcpy.AddMessage(df.iloc[1,list(df.columns).index(parameters[2].valueAsText)])
        for i in df.index:
            # arcpy.AddMessage(df.iloc[i])
            if data.get(df.iloc[i,list(df.columns).index(parameters[2].valueAsText)])!=None:
                enterData = data.get(df.iloc[i,list(df.columns).index(parameters[2].valueAsText)])
                # arcpy.AddMessage([len(list(df.columns)),len(enterData)])
                # arcpy.AddMessage([len(csvgisfields),len(enterData),len(df.columns)])
                for col in list(df.columns):
                    # arcpy.AddMessage(col)
                    if col in csvgisfields:
                        df.iloc[i,list(df.columns).index(col)] = enterData[csvgisfields.index(col)]
                        # arcpy.AddMessage(enterData[csvgisfields.index(col)])
                
                # for k,c in enumerate(csvgisfields):
                #     # arcpy.AddMessage(k)
                #     arcpy.AddMessage([list(df.columns).index(c),enterData[list(df.columns).index(c)-1]])
                #     df.iloc[i,list(df.columns).index(c)] = enterData[list(df.columns).index(c)-1]

            else:
                arcpy.AddMessage(df.iloc[i,list(df.columns).index(parameters[2].valueAsText)])
        
            

            # arcpy.AddMessage(df.iloc[i,csvgisfields.index(parameters[2].valueAsText)])
            # if f in data.keys():
                

        # arcpy.AddMessage(data)
        # arcpy.AddMessage(csvgisfields[csvgisfields.index(parameters[2].valueAsText)])
        
        
        # fields = [f.name for f in arcpy.ListFields(parameters[3].valueAsText)]
        # data = [row for row in arcpy.da.SearchCursor(parameters[3].valueAsText,fields)]
        # dfgis = pd.DataFrame(data,columns = fields)[[parameters[4].valueAsText]+gisfields].dropna().set_index(parameters[4].valueAsText).to_dict()

        # col = df.columns.to_list()
        # for i in df.index:
        #     for f in pflfields:
        #         df.iloc[i,col.index(f)] = dfgis[gisfields[pflfields.index(f)]].get(df.iloc[i,col.index(parameters[4].valueAstext)])
        # arcpy.AddMessage(df)
    
        writer = pd.ExcelWriter(parameters[8].valueAsText)
        df.to_excel(writer,sheet_name=parameters[7].valueAsText,index=False)
        writer.close()







            #     if len(data)!=0:
            #             # mitigationActions.to_csv(f"C:\Users\SirishDasari\OneDrive - Total Infrastructure Management Solutions LLC\Desktop\sirish\AIC\GISIntegrationFiles\test\{l}.csv")
            #             arcpy.AddMessage(mitigationActions)
            #             mitigationActions.to_excel(writer,sheet_name=l,index=False)
            # arcpy.AddMessage('Add Message')
# class GISIntegrationAll(object):

#     def __init__(self):
#         self.label = "PFL to GIS Integration"
#         self.description = ""
#         self.canRunInBackground = False

#     def getParameterInfo(self):
#         pfl = arcpy.Parameter(
#             displayName = "Enter PFL path",
#             name = "pfl",
#             datatype = "DEFile",
#             parameterType = "Required",
#             direction = "Input")
#         pfl.filter.list = ["xlsx","xls"]
#         sheet = arcpy.Parameter(
#             displayName = "Enter sheet name",
#             name = "sheet",
#             datatype = "GPString",
#             parameterType = "Required",
#             direction = "Input")
     
#         fieldMapping = arcpy.Parameter(
#             displayName = "Enter field mapping path",
#             name = "fieldMapping",
#             datatype = "DEFile",
#             parameterType = "Required",
#             direction = "Input")
#         # fieldMapping.filter.list = ["xlsx","xls"]
#         fieldMappingFields = arcpy.Parameter(
#             displayName = "Enter field mapping fields",
#             name = "fieldMappingFields",
#             datatype = "GPValueTable",
#             parameterType = "Required",
#             direction = "Input")
#         fieldMappingFields.columns = [['GPString', 'Layer'],['GPString', 'GIS Field'], ['GPString', 'PFL Field']]
#         fieldMappingFields.parameterDependencies = [fieldMapping.name]

#         layer = arcpy.Parameter(
#             displayName = "Enter field mapping Layer",
#             name = "layer",
#             datatype = "GPString",
#             parameterType = "Required",
#             direction = "Input")
#         # layer.columns = [['GPString','Layer'],['GPString','PFL Field']]
        
#         PFLField = arcpy.Parameter(
#             displayName = "Enter PFL Field",
#             name = "PFLField",
#             datatype = "GPString",
#             parameterType = "Required",
#             direction = "Input")
#         featureName = arcpy.Parameter(
#             displayName = "Enter Feature Names",
#             name = "featureName",
#             datatype = "GPValueTable",
#             parameterType = "Required",
#             direction = "Input")
#         featureName.columns = [['GPString','Names']]
#         FeatureMapping = arcpy.Parameter(
#             displayName = "Enter Feature Mapping Fields",
#             name = "FeatureMapping",
#             datatype = "GPValueTable",
#             parameterType = "Required",
#             direction = "Input")
#         FeatureMapping.columns = [['GPString', 'Layer']]
        
#         export = arcpy.Parameter(
#             displayName = "Enter export path",
#             name = "export",
#             datatype = "DEFile",
#             parameterType = "Required",
#             direction = "Output")
#         export.filter.list = ["xlsx","xls"]
#         # export.filter.list = ["xlsx","xls"]
#         return [pfl,sheet,fieldMapping,fieldMappingFields,layer,PFLField,featureName,FeatureMapping,export]
#         # return [pfl]
        
#     def isLicensed(self):
#         return True

#     def updateParameters(self, parameters):
#         if parameters[0].altered:
#             fields = pd.ExcelFile(parameters[0].valueAsText).sheet_names
#             parameters[1].filter.list = fields
#         if parameters[2].altered:
#             fields = [f.name for f in arcpy.ListFields(parameters[2].valueAsText)]
#             parameters[3].filters[0].list = fields
#             parameters[3].filters[1].list = fields
#             parameters[3].filters[2].list = fields
#         if parameters[3].altered:
#             parameters[4].filter.list = list(set([f[0] for f in arcpy.da.SearchCursor(parameters[2].valueAsText,parameters[3].value[0][0])]))
#         if parameters[1].altered:
#             parameters[5].filter.list = list(pd.read_excel(parameters[0].valueAsText,sheet_name=parameters[1].valueAsText).columns)
#         if parameters[5].altered:
#             parameters[6].filter.list = [f for f in list(set(pd.read_excel(parameters[0].valueAsText,sheet_name=parameters[1].valueAsText)[parameters[6].valueAsText].values.tolist()))if f !=None]
#         # if parameters[6].altered:
#         #     fields = list(set(pd.read_csv(parameters[5].valueAsText).fillna('')[parameters[6].value[0][0]].tolist()))
#         #     parameters[7].filters[0].list = fields
#         # if parameters[6].altered:
#         #     fields = list(set(pd.read_csv(parameters[5].valueAsText).fillna('')[parameters[6].value[0][0]].tolist()))
#         #     parameters[8].filters[0].list = fields
#         return

#     def updateMessages(self, parameters):
#         return

#     def execute(self, parameters, messages):
#         arcpy.AddMessage(list(set(pd.read_csv(parameters[5].valueAsText).fillna('')[parameters[6].value[0][0]].tolist())))
#         fieldMapping = parameters[4].value[0]
#         featureMapping = parameters[6].value[0]
#         # read feature mapping
#         featureMap = pd.read_csv(parameters[5].valueAsText).fillna('')[featureMapping].groupby(featureMapping[0])[featureMapping[1]].apply(list).to_dict()

#         # read field mapping
#         fieldMapGIS = pd.read_csv(parameters[3].valueAsText).fillna('')[fieldMapping].groupby(fieldMapping[0])[fieldMapping[1]].apply(list).to_dict()
#         fieldMapPFL = pd.read_csv(parameters[3].valueAsText).fillna('')[fieldMapping].groupby(fieldMapping[0])[fieldMapping[2]].apply(list).to_dict()
#         fieldMerge = pd.read_csv(parameters[3].valueAsText)[fieldMapping].dropna().groupby([fieldMapping[0]])[fieldMapping[1]].apply(list).to_dict()

#         for l in featureMap.keys():
#             if l in fieldMapGIS:
#                 arcpy.AddMessage(l)
#                 fm = {}
#                 for gis in fieldMapGIS[l]:
#                     if fieldMapPFL[l][fieldMapGIS[l].index(gis)] != '':
#                         fm[gis] = fieldMapPFL[l][fieldMapGIS[l].index(gis)]
#                 mappingFields = {fm[f]:f for f in fm.keys()}
#                 pfl = pd.read_excel(parameters[0].valueAsText,sheet_name=parameters[1].valueAsText).fillna('')
#                 pflFields = [f for f in fieldMapPFL[l] if f != '']
#                 gisFields = list(set([f for f in fieldMapGIS[l] if f != '']))
#                 pflData = pfl[pfl[parameters[2].value[0][3]].isin(featureMap[l])][pflFields]
                
#                 for pfl in pflData.columns.tolist():
#                     if pfl in fieldMapPFL[l]:
#                         pflData.rename(columns=mappingFields,inplace=True)
                
#                 data = []
#                 for fieldPFL in pflData.values.tolist():
#                     fieldsData = []
#                     for gis in gisFields:
#                         if gis in pflData.columns.tolist():
#                             fieldsData.append(fieldPFL[pflData.columns.tolist().index(gis)])
#                         else:
#                             fieldsData.append("")
#                     data.append(fieldsData)
#                 mitigationActions = pd.DataFrame(data,columns=gisFields)
#                 mitigationActions.to_csv(parameters[6].valueAstext)


class PFL2(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "GIS to PFL V2"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        pipe = arcpy.Parameter(
            displayName = "Enter Pipe Feature",
            name = "pipe",
            datatype = "GPFeatureLayer",
            parameterType = "Required",
            direction = "Input")
        # pipe.value = "P_Pipes_INSV_All"
        # lineLayers = arcpy.Parameter(
        #     displayName = "Enter Line Layers",
        #     name = "line",
        #     datatype = "GPFeatureLayer",
        #     parameterType = "Optional",
        #     direction = "Input",
        #     multiValue=True)
        # lineLayers.columns = [['GPFeatureLayer', 'Layer']]
        # lineLayers.values = [['P_PipeCasing_INSV_All']]

        pointLayers = arcpy.Parameter(
            displayName = "Enter Point Layers",
            name = "point",
            datatype = "GPFeatureLayer",
            parameterType = "Optional",
            direction = "Input",
            multiValue=True)
        pointLayers.columns = [['GPFeatureLayer', 'Layer']]
        # pointLayers.values = [['P_ControllableFitting_INSV_All'],['P_NonControllableFitting_INSV_All'],['P_Valve_INSV_All']]

        FieldMapping = arcpy.Parameter(
            displayName = "Enter Field Mappings",
            name = "FieldMapping",
            datatype = "DETable",
            parameterType = "Optional",
            direction = "Input")
        # FieldMapping.value = "Field_Mapping_Template"

        FMTemplate = arcpy.Parameter(
            displayName = "Field mapping Template",
            name = "FMTemplate",
            datatype = "GPValueTable",
            parameterType = "Optional",
            direction = "Input")
        FMTemplate.columns = [['GPString', 'Layer'], ['GPString', 'GIS'], ['GPString', 'PFL']]
        # FMTemplate.values = [['FeatureLayer', 'GIS', 'PFL']]

        pflTemplate = arcpy.Parameter(
            displayName = "Enter MAOP table template",
            name = "pflTemplate",
            datatype = "DETable",
            parameterType = "Optional",
            direction = "Input")
        # pflTemplate.value = "MAOP_Gas_PFLTemplate_V12"
        pflFields = arcpy.Parameter(
            displayName = "Enter PFL Fields given below",
            name = "pflFields",
            datatype = "GPValueTable",
            parameterType = "Optional",
            direction = "Input")
        pflFields.columns = [["GPString","From Measure"],["GPString","To Measure"],["GPString","Route/TR"]]
        # pflFields.values = [['From_Measure', 'To_Measure', 'TR']]
        table = arcpy.Parameter(
            displayName = "Enter output for table",
            name = "table",
            datatype = "DETable",
            parameterType = "Required",
            direction = "Output")
        # layerMapping.columns = [["GPString","P_Pipe"],["GPString","P_Casing"]]
        # # layerMapping.values = [['P_Pipes INSV_All', 'P_PipeCasing_INSV_All']]
        return [pipe,pointLayers,FieldMapping,FMTemplate,pflTemplate,pflFields,table]

    def updateParameters(self, parameters):
        fieldMapping = parameters[2]
        fieldMappingsList = parameters[3]
        assignFields = parameters[4]
        assignFieldsList = parameters[5]


        if fieldMapping.altered:
            fieldMappingsList.filters[0].list = [f.name for f in arcpy.ListFields(fieldMapping.valueAsText)]
            fieldMappingsList.filters[1].list = [f.name for f in arcpy.ListFields(fieldMapping.valueAsText)]
            fieldMappingsList.filters[2].list = [f.name for f in arcpy.ListFields(fieldMapping.valueAsText)]
        if assignFields.altered:
            assignFieldsList.filters[0].list = [f.name for f in arcpy.ListFields(assignFields.valueAsText)]
            assignFieldsList.filters[1].list = [f.name for f in arcpy.ListFields(assignFields.valueAsText)]
            assignFieldsList.filters[2].list = [f.name for f in arcpy.ListFields(assignFields.valueAsText)]
        return
    def updateMessages(self, parameters):
        # if parameters[1].altered and parameters[2].altered and parameters[3].altered and parameters[1].value!=None and parameters[2].value!=None and parameters[4].value!=None:
        #     layers = [f[0].name for f in parameters[0].value]
        #     newdata = []
        #     templateNames = [f.aliasName for f in arcpy.ListFields(parameters[4].valueAsText)]
        #     for layer in layers:
        #         gis = {row[1]:row[2] for row in arcpy.da.SearchCursor(parameters[2].valueAsText,parameters[3].value[0]) if row[0] == layer and row[2] not in ['',None] }
        #         for k in gis.values():
        #             if k not in templateNames:
        #                 newdata.append(f"{k} field not in {layer}")
        #     if len(newdata)!=0:
        #         parameters[3].setWarningMessage(newdata)
        #     if len(parameters[2].valueAsText.split(" ")) != 4:
        #         parameters[2].setWarningMessage('Remove space if present in field names')
        #     else:
        #         parameters[2].clearMessage()
        return

    def execute(self, parameters, messages):
        # arcpy.AddMessage("Processing ...")
        Pipe = parameters[0].valueAsText
        # lineLayer = parameters[1].value
        pointLayer =  parameters[1].value
        fieldMapping =  parameters[2].valueAsText
        fieldMappingList = parameters[3].value
        MAOPtemplate = parameters[4].valueAsText
        MAOPtemplateFields = parameters[5].value
        ExportOutput = parameters[6].valueAsText


        # # PFL Fields dictionary
        arcpy.management.CopyRows(MAOPtemplate, "in_memory/CenterLine")
        arcpy.management.CopyRows(MAOPtemplate, "in_memory/PFL")
        arcpy.management.CopyRows(MAOPtemplate, "in_memory/Point")

        # Dict PFL aliasName:name 
        pfl = {f.aliasName:f.name for f in arcpy.ListFields("in_memory/CenterLine")}

        # All Field Names including ObjectID
        allPFLfields = [f.name for f in arcpy.ListFields("in_memory/CenterLine")]
        # arcpy.AddMessage(allPFLfields)
        insertToPFL = arcpy.da.InsertCursor("in_memory/PFL",allPFLfields)

        # Get Layers names from list
        if parameters[1].value !=None:
            pointLayers = [f[0].name for f in pointLayer]

        # Get index with list as input
        def pflIndex(fields):
            return [allPFLfields.index(l) for l in fields]


        fromIndex = pflIndex(MAOPtemplateFields[0])[0]
        toIndex = pflIndex(MAOPtemplateFields[0])[1]
        trIndex = pflIndex(MAOPtemplateFields[0])[2]
        
        def fieldsGIStoPFL(layer):
            if parameters[1].value==None:
                processlayers=pointLayers+[Pipe]
                # arcpy.AddMessage(processlayers)
            else:
                processlayers=pointLayers+[Pipe]
                # arcpy.AddMessage(processlayers)
            for l in processlayers:
                if l==layer:
                    GISfields = []
                    PFLfields = []

                    GIStoPFL = {row[1]:row[2] for row in arcpy.da.SearchCursor(fieldMapping,fieldMappingList[0]) if row[0] == layer and row[2] not in ['',None]}
                    # arcpy.AddMessage(GIStoPFL)
                    for f in GIStoPFL.keys():
                        GISfields.append(f)
                        PFLfields.append(pfl[GIStoPFL[f]])
                    return GISfields, PFLfields
        
        def processPipe(lname):
            # arcpy.AddMessage(fieldsGIStoPFL(lname))
            insert = arcpy.da.InsertCursor("in_memory/CenterLine",fieldsGIStoPFL(lname)[1])
            data = [list(row) for row in arcpy.da.SearchCursor(lname,fieldsGIStoPFL(lname)[0])]
            for d in data:
                insert.insertRow(d)
            del insert

            with arcpy.da.UpdateCursor("in_memory/CenterLine", allPFLfields) as cursor:
                for row in cursor:
                    row[fromIndex] = round(row[fromIndex],2)
                    row[toIndex] = round(row[toIndex],2)
                    cursor.updateRow(row)
 
        # LayerInmemory = []
        # def processLineLayers(lname):
        #     for layer in lname:
        #         arcpy.management.CopyRows(MAOPtemplate, f"in_memory/{layer}")
        #         LayerInmemory.append(f"in_memory/{layer}")
        #         insertLine = arcpy.da.InsertCursor(f"in_memory/{layer}",fieldsGIStoPFL(layer)[1])
        #         data = [list(row) for row in arcpy.da.SearchCursor(layer,fieldsGIStoPFL(layer)[0])]
        #         for d in data:
        #             insertLine.insertRow(d)
        #         del insertLine

        #         with arcpy.da.UpdateCursor(f"in_memory/{layer}", allPFLfields) as cursor:
        #             for row in cursor:
        #                 row[fromIndex] = round(row[fromIndex],2)
        #                 row[toIndex] = round(row[toIndex],2)
        #                 cursor.updateRow(row)

        def processPoint(lname):
            for layer in lname:
                insertPoint = arcpy.da.InsertCursor("in_memory/Point",fieldsGIStoPFL(layer)[1])
                data = [list(row) for row in arcpy.da.SearchCursor(layer,fieldsGIStoPFL(layer)[0])]
                for d in data:
                    insertPoint.insertRow(d)
                del insertPoint


            with arcpy.da.UpdateCursor("in_memory/Point", allPFLfields) as cursor:
                for row in cursor:
                    row[fromIndex] = round(row[fromIndex],2)
                    row[toIndex] = row[fromIndex]
                    cursor.updateRow(row)
        
        def linePoints(layer,id):
            pflSplit = []
            with arcpy.da.SearchCursor(layer,allPFLfields) as cursor:
                for row in cursor:
                    if row[trIndex] == id:
                        pflSplit.append(row[1])
                        pflSplit.append(row[2])
            return sorted(set(pflSplit))
        
        def getPoints(layer,id):
            pflSplit = []
            with arcpy.da.SearchCursor(layer,allPFLfields) as cursor:
                for row in cursor:
                    if row[trIndex] == id:
                        pflSplit.append(row[1])
            return sorted(set(pflSplit))

        processPipe(Pipe)        
        processPoint(pointLayers) 
        with arcpy.da.UpdateCursor('in_memory/Point',[MAOPtemplateFields[0][0],MAOPtemplateFields[0][1]]) as cursor:
            for row in cursor:
                row[0]=round(row[0],2)
                row[1]=round(row[1],2)
                cursor.updateRow(row)
        trs = set([f[0] for f in arcpy.da.SearchCursor("in_memory/Point",[MAOPtemplateFields[0][2]])])
        tr = {}
        for f in trs:
            tr[f]=[]
        for row in arcpy.da.SearchCursor("in_memory/Point",[MAOPtemplateFields[0][0],MAOPtemplateFields[0][2]]):
            tr[row[1]].append(row[0])
        # pointsList = [f for f in arcpy.da.SearchCursor('in_memory/Points',[MAOPtemplateFields[0][0],MAOPtemplateFields[0][2]])]
        # arcpy.AddMessage(pointsList)
        n=0
        with arcpy.da.UpdateCursor('in_memory/CenterLine',allPFLfields) as cursor:
            for row in cursor:
                # arcpy.AddMessage([row[fromIndex],tr[row[trIndex]],row[toIndex]])
                inside = [f for f in tr[row[trIndex]] if row[fromIndex]<f<row[toIndex]]
                if len(inside)>0:
                    insert = sorted([row[fromIndex]]+inside+[row[toIndex]])
                    arcpy.AddMessage(insert)
                    for i in range(len(insert)-1):
                        row[fromIndex] = insert[i]
                        row[toIndex] = insert[i+1]
                        # arcpy.AddMessage([row[fromIndex],row[toIndex]])
                        insertToPFL.insertRow([n]+row[1:])
                        n=n+1
                else:
                    # arcpy.AddMessage([row[fromIndex],row[toIndex]])
                    insertToPFL.insertRow([n]+row[1:])
                    n=n+1
        
        # with arcpy.da.SearchCursor("in_memory/PFL")
        for f in [row for row in arcpy.da.SearchCursor("in_memory/Point",allPFLfields)]:
            n=n+1
            insertToPFL.insertRow([n]+list(f[1:]))


        # for layer in LayerInmemory:
        #     count = int(arcpy.management.GetCount("in_memory/CenterLine").getOutput(0))
        #     arcpy.SetProgressor("step", "Processing...",0, count, 1)
        #     with arcpy.da.UpdateCursor("in_memory/CenterLine",allPFLfields) as cursor:
        #         n=0
        #         for row in cursor:
        #             n=n+1
        #             arcpy.SetProgressorLabel(f"Splitting PFL from {layer[10:]}: {n} of {count}")
        #             points = [f for f in linePoints(layer,row[trIndex]) if row[fromIndex]<f<row[toIndex]]
        #             if len(points)!=0:
        #                 insert = [row[fromIndex]]+points+[row[toIndex]]
        #                 for i in range(len(insert)-1):
        #                     row[fromIndex] = insert[i]
        #                     row[toIndex] = insert[i+1]
        #                     insertToPFL.insertRow([int(arcpy.management.GetCount("in_memory/CenterLine").getOutput(0))]+row[1:])
        #                 cursor.deleteRow()
        #             arcpy.SetProgressorPosition()
        #         arcpy.ResetProgressor()
        #     del cursor

        # for layer in LayerInmemory:
        #     count = int(arcpy.management.GetCount("in_memory/CenterLine").getOutput(0))
        #     arcpy.SetProgressor("step", "Processing...",0, count, 1)
        #     with arcpy.da.UpdateCursor("in_memory/CenterLine",allPFLfields) as cursor:
        #         n=0
        #         for pipes in cursor:
        #             n=n+1
        #             arcpy.SetProgressorLabel(f"Updating PFL fields from {layer[10:]}: {n} of {count}")
        #             otherShort  =   sorted([row for row in arcpy.da.SearchCursor(layer,allPFLfields) if pipes[fromIndex]<=row[fromIndex]<=row[toIndex]<=pipes[toIndex]],key=itemgetter(fromIndex,toIndex))
        #             for short in otherShort:
        #                 for new in [f for f in fieldsGIStoPFL(layer[10:])[1] if f not in MAOPtemplateFields[0]]:
        #                     pipes[pflIndex([new])[0]] = short[pflIndex([new])[0]]
        #                 cursor.updateRow(pipes)
        #             arcpy.SetProgressorPosition()
        #         arcpy.ResetProgressor()
        #     del cursor

        # with arcpy.da.UpdateCursor("in_memory/CenterLine",allPFLfields) as cursor:
        #     count = int(arcpy.management.GetCount("in_memory/CenterLine").getOutput(0))
        #     arcpy.SetProgressor("step", "Processing...",0, count, 1)
        #     n=0
        #     for row in cursor:
        #         n=n+1
        #         arcpy.SetProgressorLabel(f"Splitting PFL from Point Layer: {n} of {count}")
        #         # arcpy.AddMessage(row[trIndex])
        #         points = []
        #         # f for f in getPoints("in_memory/Point",row[trIndex]) if row[fromIndex]<f<row[toIndex]and row[fromIndex]!=None and row[toIndex]!=None and f!=None
        #         for pt in getPoints("in_memory/Point",row[trIndex]):
        #             if row[fromIndex]!=None and row[toIndex]!=None and pt!=None:
        #                 if row[fromIndex]<pt<row[toIndex]:
        #                     points.append(pt)
        #         # arcpy.AddMessage([row[1],row[2],"=========",points])
        #         if len(points)!=0:
        #             insert = [row[fromIndex]]+points+[row[toIndex]]
        #             for i in range(len(insert)-1):
        #                 row[fromIndex] = insert[i]
        #                 row[toIndex] = insert[i+1]
        #                 insertToPFL.insertRow([int(arcpy.management.GetCount("in_memory/CenterLine").getOutput(0))]+row[1:])
        #             cursor.deleteRow()
        #         arcpy.SetProgressorPosition()
        #     arcpy.ResetProgressor()
        # del cursor

        # count = int(arcpy.management.GetCount("in_memory/Point").getOutput(0))
        # arcpy.SetProgressor("step", "Processing...",0, count, 1)
        # n=0
        # for f in [row for row in arcpy.da.SearchCursor("in_memory/Point",allPFLfields)]:
        #     n=n+1
        #     arcpy.SetProgressorLabel(f"Updating Point to PFL: {n} of {count}")
        #     insertToPFL.insertRow([int(arcpy.management.GetCount("in_memory/CenterLine").getOutput(0))]+list(f[1:]))
        #     arcpy.SetProgressorPosition()
        # arcpy.ResetProgressor()
        # arcpy.management.CopyRows("in_memory/CenterLine", ExportOutput)
        arcpy.management.CopyRows("in_memory/PFL", ExportOutput)
        # arcpy.management.CopyRows("in_memory/Point", 'Points')