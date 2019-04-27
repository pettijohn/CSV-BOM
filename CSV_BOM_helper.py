# This file does all the hard work of translating a list of objects into a variety of CSV formats.
# This file is separate to enable unit testing the logic of CSV generation from the Fusion 360 Addin runtime. 

import csv
import io
import re
import collections
from typing import List


class Dimensions:
    
    """ Internal values should be floats and sortable (mm), while Formatted are strings (incl. fractional inches) """
    def __init__(self, xInternal: float, yInternal: float, zInternal: float, xFormatted: str, yFormatted: str, zFormatted: str):
        self._X = [xInternal, xFormatted]
        self._Y = [yInternal, yFormatted]
        self._Z = [zInternal, zFormatted]

    def GetArray(self):
        return [self._X, self._Y, self._Z]

    def GetSortedTuples(self):
        # Sort the tuples on [0]
        return sorted(self.GetArray(), key=lambda d: d[0])

    def GetSortedInternal(self):
        return list(map(lambda t: t[0], self.GetSortedTuples()))

    def GetSortedFormatted(self):
        return list(map(lambda t: t[1], self.GetSortedTuples()))

    def GetUnsortedFormatted(self):
        return list(map(lambda t: t[1], self.GetArray()))


class PhysicalAttributes:
    def __init__(self, dimensions: Dimensions, volume, area, mass, density, material):
        self.Volume = volume
        self.Dimensions = dimensions
        self.Area = area
        self.Mass = mass
        self.Density = density
        self.Material = material

class BomItem:
    def __init__(self, name, quantity, description, physicalAttributes: PhysicalAttributes, component=None):
        self.Name = name
        self.Quantity = quantity
        self.Description = description
        self.Component = component
        self.PhysicalAttributes = physicalAttributes

class Helper:
    def __init__(self):
        pass

    def filterFusionCompNameInserts(self, name):
        name = re.sub("\([0-9]+\)$", '', name)
        name = name.strip()
        name = re.sub("v[0-9]+$", '', name)
        return name.strip()

    def replacePointDelimterOnPref(self, pref: bool, value: str):
        """ Replace decimal point in str(number) with a comma """
        if (pref):
            return str(value).replace(".", ",")
        return str(value)

    def SaveCsv(self, filename, bom: List[BomItem], prefs):
        with open(filename, 'w', newline='') as csvFile:
            self.WriteCsv(csvFile, bom, prefs)

    def WriteCsv(self, f, bom: List[BomItem], prefs):
        #TODO
        defaultUnit = "Inches"

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

        # Extras Action = Ignore means that when items in the dict are encountered that
        #  are not present in the header, they are ignored. This lets us perform all the
        #  conditional logic of which fields to include just once up front when 
        #  building the header.
        writer = csv.DictWriter(f, fieldnames=csvHeader, extrasaction='ignore')
        writer.writeheader()

        
        for item in bom:
            csvRow = {}
            name = self.filterFusionCompNameInserts(item.Name)
            if prefs["ignoreUnderscorePrefComp"] is False and prefs["underscorePrefixStrip"] is True and name[0] == '_':
                name = name[1:]
                
            csvRow["Part name"] = name
            csvRow["Quantity"] = item.Quantity
            
            csvRow["Volume cm^3"] = self.replacePointDelimterOnPref(prefs["useComma"], item.PhysicalAttributes.Volume)
            
            #############
            if prefs["sortDims"]:
                dimensions = item.PhysicalAttributes.Dimensions.GetSortedFormatted()
            else:
                dimensions = item.PhysicalAttributes.Dimensions.GetUnsortedFormatted()
            if prefs["splitDims"]:
                csvRow["Width " + defaultUnit] = dimensions[0]
                csvRow["Length " + defaultUnit] = dimensions[1]
                csvRow["Height " + defaultUnit] = dimensions[2]
            else:
                csvRow["Dimension " + defaultUnit] = dimensions[0] + " x " + dimensions[1] + " x " + dimensions[2]

            csvRow["Area cm^2"] = self.replacePointDelimterOnPref(prefs["useComma"], "{0:.2f}".format(item.PhysicalAttributes.Area))
            csvRow["Mass kg"] = self.replacePointDelimterOnPref(prefs["useComma"], "{0:.5f}".format(item.PhysicalAttributes.Mass))
            csvRow["Density kg/cm^2"] = self.replacePointDelimterOnPref(prefs["useComma"], "{0:.5f}".format(item.PhysicalAttributes.Density))
            csvRow["Material"] = item.PhysicalAttributes.Material
            csvRow["Description"] = item.Description
            
            writer.writerow(csvRow)


    def WriteCutlistGaryDarby(self, stream: io.IOBase, bom: List[BomItem], prefs):
        # Init CutList Header
        stream.write('V2\n')
        if prefs["useComma"]:
            stream.write('FormatSettings.decimalseparator,\n')
        else:
            stream.write('FormatSettings.decimalseparator.\n')
        stream.write('\n')
        stream.write('Required\n')

        for item in bom:
            #add parts:
            name = self.filterFusionCompNameInserts(item.Name)
            if prefs["ignoreUnderscorePrefComp"] is False and prefs["underscorePrefixStrip"] is True and name[0] == '_':
                name = name[1:]
            # dimensions:
            # dim = 0
            # for k in item["boundingBox"]:
            #     dim += item["boundingBox"][k]
            # if dim > 0:
            #     # Formatted units may be strings when fractional inches, e.g., 18 1/2"
            #     # Get the internal unit as well, which is float, for sorting
            #     axises = ["x", "y", "z"]
            #     dimensions = {}
            #     for axis in axises:
            #         dimensions[axis] = [item["boundingBox"][axis], design.fusionUnitsManager.formatInternalValue(item["boundingBox"][axis], defaultUnit, False)]
                
            #     # Sort on the first value, which is decimal
            #     sortedDimensions =  collections.OrderedDict(sorted(dimensions.items(), key=lambda d: d[1]))

            #     # Cutlist requires whole inch measurements to have 0/0 fractions
            #     for axis in axises:
            #         if ('/' in dimensions[axis][1]) or ('.' in dimensions[axis][1] and not prefs["useComma"]) or (',' in dimensions[axis][1] and prefs["useComma"]):
            #             pass #easier than all the nots
            #         else:
            #             dimensions[axis][1] = dimensions[axis][1] + " 0/0"
                        
            #     if prefs["sortDims"]:
            #         dims = list(map(lambda d: d[1][1], sortedDimensions.items()))
            #     else:
            #         if type(dimensions["z"][1]) == str:
            #             # String units (fractional inches) don't need to be formatted
            #             dims = [dimensions["z"][1], dimensions["x"][1], dimensions["y"][1]]
            #         else:
            #             # decimal units need to be formatted
            #             f = "{0:.4f}"
            #             dims = [f.format(dimensions["z"][1]), f.format(dimensions["x"][1]), f.format(dimensions["y"][1])]

            #     partStr = ' '  # leading space
            #     partStr += self.replacePointDelimterOnPref(prefs["useComma"], dims[1]).ljust(12)  # width, x
            #     partStr += self.replacePointDelimterOnPref(prefs["useComma"], dims[2]).ljust(12)  # length, y

            #     partStr += name
            #     partStr += ' (thickness: ' + self.replacePointDelimterOnPref(prefs["useComma"], dims[0]) + defaultUnit + ')'
            #     partStr += '\n'

            # else:
            #     partStr = ' 0        0      ' + name + '\n'

            if prefs["sortDims"]:
                dims = item.PhysicalAttributes.Dimensions.GetSortedFormatted()
            else:
                dims = item.PhysicalAttributes.Dimensions.GetUnsortedFormatted()
            
            partStr = " {0} {1} {2} (thickness: {3})\n".format(dims[0], dims[1], name, dims[2])

            # add all instances of the component to the CutList:
            for i in range(0, item.Quantity):
                stream.write(partStr)

        # empty entry for available materials (sheets):
        stream.write('\n' + "Available" + '\n')