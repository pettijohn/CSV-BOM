# This file does all the hard work of translating a list of objects into a variety of CSV formats.
# This file is separate to enable unit testing the logic of CSV generation from the Fusion 360 Addin runtime. 

import csv

class BomItem:
    def __init__(self, name, quantity, description, physicalAttributes, component=None):
        self.Name = name
        self.Quantity = quantity
        self.Description = description
        self.Component = component
        self.PhysicalAttributes = physicalAttributes

class PhysicalAttributes:
    def __init__(self, dimensions, volume, area, mass, density, material):
        self.Volume = volume
        self.Dimensions = dimensions
        self.Area = area
        self.Mass = mass
        self.Density = density
        self.Material = material

class Dimensions:
    def __init__(self, x, y, z):
        self.X = x
        self.Y = y
        self.Z = z

class Helper:
    def GetCsvString():
        csvStr = ''
        defaultUnit = design.fusionUnitsManager.defaultLengthUnits
        csvHeader = ["Part name", "Quantity"]
        if prefs["incVol"]:
            csvHeader.append("Volume cm^3")
        if prefs["incBoundDims"]:
            if prefs["splitDims"]:
                csvHeader.append("Width " + defaultUnit)
                csvHeader.append("Length " + defaultUnit)
                csvHeader.append("Height " + defaultUnit)
            else:
                csvHeader.append("Dimension " + defaultUnit)
        if prefs["incArea"]:
            csvHeader.append("Area cm^2")
        if prefs["incMass"]:
            csvHeader.append("Mass kg")
        if prefs["incDensity"]:
            csvHeader.append("Density kg/cm^2")
        if prefs["incMaterial"]:
            csvHeader.append("Material")
        if prefs["incDesc"]:
            csvHeader.append("Description")
        for k in csvHeader:
            csvStr += '"' + k + '",'
        csvStr += '\n'
        for item in bom:
            dims = ''
            name = self.filterFusionCompNameInserts(item["name"])
            if prefs["ignoreUnderscorePrefComp"] is False and prefs["underscorePrefixStrip"] is True and name[0] == '_':
                name = name[1:]
            csvStr += '"' + name + '","' + self.replacePointDelimterOnPref(prefs["useComma"], item["instances"]) + '",'
            if prefs["incVol"]:
                csvStr += '"' + self.replacePointDelimterOnPref(prefs["useComma"], item["volume"]) + '",'
            if prefs["incBoundDims"]:
                dim = 0
                footInchDispFormat = app.preferences.unitAndValuePreferences.footAndInchDisplayFormat
                
                for k in item["boundingBox"]:
                    dim += item["boundingBox"][k]
                if dim > 0:
                    if footInchDispFormat == 0:
                        dimX = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["x"], defaultUnit, False))
                        dimY = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["y"], defaultUnit, False))
                        dimZ = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["z"], defaultUnit, False))
                        if prefs["sortDims"]:
                            dimSorted = sorted([dimX, dimY, dimZ])
                            bbZ = "{0:.3f}".format(dimSorted[0])
                            bbX = "{0:.3f}".format(dimSorted[1])
                            bbY = "{0:.3f}".format(dimSorted[2])
                        else:
                            bbX = "{0:.3f}".format(dimX)
                            bbY = "{0:.3f}".format(dimY)
                            bbZ = "{0:.3f}".format(dimZ)
    
                        if prefs["splitDims"]:
                            csvStr += '"' + self.replacePointDelimterOnPref(prefs["useComma"], bbX) + '",'
                            csvStr += '"' + self.replacePointDelimterOnPref(prefs["useComma"], bbY) + '",'
                            csvStr += '"' + self.replacePointDelimterOnPref(prefs["useComma"], bbZ) + '",'
                        else:
                            dims += '"' + self.replacePointDelimterOnPref(prefs["useComma"], bbX) + ' x '
                            dims += self.replacePointDelimterOnPref(prefs["useComma"], bbY) + ' x '
                            dims += self.replacePointDelimterOnPref(prefs["useComma"], bbZ)
                            csvStr += dims + '",'
                    else:
                        dimX = design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["x"], defaultUnit, False)
                        dimY = design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["y"], defaultUnit, False)
                        dimZ = design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["z"], defaultUnit, False)

                        if prefs["splitDims"]:
                            csvStr += '"' + dimX.replace('"', '""') + '",'
                            csvStr += '"' + dimY.replace('"', '""') + '",'
                            csvStr += '"' + dimZ.replace('"', '""') + '",'
                        else:
                            dims += '"' + dimX.replace('"', '""') + ' x '
                            dims += dimY.replace('"', '""') + ' x '
                            dims += dimZ.replace('"', '""')
                            csvStr += dims + '",'                        
                else:
                    if prefs["splitDims"]:
                        csvStr += "0" + ','
                        csvStr += "0" + ','
                        csvStr += "0" + ','
                    else:
                        csvStr += "0" + ','
            if prefs["incArea"]:
                csvStr += '"' + self.replacePointDelimterOnPref(prefs["useComma"], "{0:.2f}".format(item["area"])) + '",'
            if prefs["incMass"]:
                csvStr += '"' + self.replacePointDelimterOnPref(prefs["useComma"], "{0:.5f}".format(item["mass"])) + '",'
            if prefs["incDensity"]:
                csvStr += '"' + self.replacePointDelimterOnPref(prefs["useComma"], "{0:.5f}".format(item["density"])) + '",'
            if prefs["incMaterial"]:
                csvStr += '"' + item["material"] + '",'
            if prefs["incDesc"]:
                csvStr += '"' + item["desc"] + '",'
            csvStr += '\n'

    def GetCutlistGaryDarbyString():
        # defaultUnit may be fractional inches, which breaks the sorting functionality (necessary to determine material thickness)
        defaultUnit = design.fusionUnitsManager.defaultLengthUnits
        # Get a decimal unit mm too 
        internalUnit = design.fusionUnitsManager.internalUnits

        # Init CutList Header
        cutListStr = 'V2\n'
        if prefs["useComma"]:
            cutListStr += 'FormatSettings.decimalseparator,\n'
        else:
            cutListStr += 'FormatSettings.decimalseparator.\n'
        cutListStr += '\n'
        cutListStr += 'Required\n'

        #add parts:
        for item in bom:
            name = self.filterFusionCompNameInserts(item["name"])
            if prefs["ignoreUnderscorePrefComp"] is False and prefs["underscorePrefixStrip"] is True and name[0] == '_':
                name = name[1:]
            # dimensions:
            dim = 0
            for k in item["boundingBox"]:
                dim += item["boundingBox"][k]
            if dim > 0:
                # Formatted units may be strings when fractional inches, e.g., 18 1/2"
                # Get the internal unit as well, which is decimal, for sorting
                axises = ["x", "y", "z"]
                dimensions = {}
                for axis in axises:
                    dimensions[axis] = [item["boundingBox"][axis], design.fusionUnitsManager.formatInternalValue(item["boundingBox"][axis], defaultUnit, False)]
                
                # Sort on the first value, which is decimal
                sortedDimensions =  collections.OrderedDict(sorted(dimensions.items(), key=lambda d: d[1]))

                # Cutlist requires whole inch measurements to have 0/0 fractions
                for axis in axises:
                    if ('/' in dimensions[axis][1]) or ('.' in dimensions[axis][1] and not prefs["useComma"]) or (',' in dimensions[axis][1] and prefs["useComma"]):
                        pass #easier than all the nots
                    else:
                        dimensions[axis][1] = dimensions[axis][1] + " 0/0"
                        
                if prefs["sortDims"]:
                    dims = list(map(lambda d: d[1][1], sortedDimensions.items()))
                else:
                    if type(dimensions["z"][1]) == str:
                        # String units (fractional inches) don't need to be formatted
                        dims = [dimensions["z"][1], dimensions["x"][1], dimensions["y"][1]]
                    else:
                        # decimal units need to be formatted
                        f = "{0:.4f}"
                        dims = [f.format(dimensions["z"][1]), f.format(dimensions["x"][1]), f.format(dimensions["y"][1])]

                partStr = ' '  # leading space
                partStr += self.replacePointDelimterOnPref(prefs["useComma"], dims[1]).ljust(12)  # width, x
                partStr += self.replacePointDelimterOnPref(prefs["useComma"], dims[2]).ljust(12)  # length, y

                partStr += name
                partStr += ' (thickness: ' + self.replacePointDelimterOnPref(prefs["useComma"], dims[0]) + defaultUnit + ')'
                partStr += '\n'

            else:
                partStr = ' 0        0      ' + name + '\n'

            # add all instances of the component to the CutList:
            quantity = int(item["instances"])
            for i in range(0, quantity):
                cutListStr += partStr

        # empty entry for available materials (sheets):
        cutListStr += '\n' + "Available" + '\n'