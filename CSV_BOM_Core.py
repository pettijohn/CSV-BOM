# This file does all the hard work of translating a list of objects into a variety of CSV formats.
# This file is separate to enable unit testing the logic of CSV generation from the Fusion 360 Addin runtime. 

import csv
import io
import re
import collections
import json
from typing import List

class CsvBomPrefs:
    # Following serialization pattern from https://stackoverflow.com/a/33270983/435368
    def __init__(self, onlySelectedComponents=False, sortDimensions=True, 
        ignoreUnderscorePrefixedComponents=True, stripUnderscorePrefix=False,
        ignoreCompWoBodies=True, ignoreLinkedComponents=True,
        ignoreVisibleState=True, useCommaDecimal=False, useQuantity=True, **kwargs):
        self.onlySelectedComponents = onlySelectedComponents
        self.sortDimensions=sortDimensions
        self.ignoreUnderscorePrefixedComponents=ignoreUnderscorePrefixedComponents
        self.stripUnderscorePrefix = stripUnderscorePrefix
        self.ignoreCompWoBodies=ignoreCompWoBodies
        self.ignoreLinkedComponents=ignoreLinkedComponents
        self.ignoreVisibleState=ignoreVisibleState
        self.useCommaDecimal=useCommaDecimal
        self.useQuantity=useQuantity

    @classmethod
    def from_json(cls, json_str):
        json_dict = json.loads(json_str)
        return cls(**json_dict)

    def to_json(self):
        return json.dumps(self.__dict__)
    

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
        return sorted(self.GetArray(), key=lambda d: d[0], reverse=True)

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
        name = re.sub(r"\([0-9]+\)$", '', name)
        name = name.strip()
        name = re.sub(r"v[0-9]+$", '', name)
        return name.strip()

    def replacePointDelimterOnPref(self, pref: bool, value: str):
        """ Replace decimal point in str(number) with a comma """
        if (pref):
            return str(value).replace(".", ",")
        return str(value)

    def SaveCsv(self, filename, bom: List[BomItem], prefs):
        with open(filename, 'w', newline='') as csvFile:
            self.WriteCsv(csvFile, bom, prefs)

    def WriteCsv(self, f, bom: List[BomItem], prefs: CsvBomPrefs):
        #TODO
        defaultUnit = "Inches"

        csvHeader = ["Part name"]
        
        if(prefs.useQuantity):
            csvHeader.append("Quantity")

        csvHeader.append("Volume cm^3")
        csvHeader.append("Width " + defaultUnit)
        csvHeader.append("Length " + defaultUnit)
        csvHeader.append("Height " + defaultUnit)
        csvHeader.append("Area cm^2")
        csvHeader.append("Mass kg")
        csvHeader.append("Density kg/cm^2")
        csvHeader.append("Material")
        csvHeader.append("Description")

        # Extras Action = Ignore means that when items in the dict are encountered that
        #  are not present in the header, they are ignored. This lets us perform all the
        #  conditional logic of which fields to include just once up front when 
        #  building the header.
        writer = csv.DictWriter(f, fieldnames=csvHeader, extrasaction='ignore')
        writer.writeheader()

        
        for item in bom:
            # If we don't use a quanitity flag, then repeat the row
            repeat = 1
            if not prefs.useQuantity:
                repeat = item.Quantity
                
            for r in range(repeat):
                csvRow = {}
                name = self.filterFusionCompNameInserts(item.Name)
                if prefs.ignoreUnderscorePrefixedComponents is False and prefs.stripUnderscorePrefix is True and name[0] == '_':
                    name = name[1:]
                    
                csvRow["Part name"] = name
                csvRow["Quantity"] = item.Quantity
                
                csvRow["Volume cm^3"] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, item.PhysicalAttributes.Volume)
                
                #############
                if prefs.sortDimensions:
                    dimensions = item.PhysicalAttributes.Dimensions.GetSortedFormatted()
                else:
                    dimensions = item.PhysicalAttributes.Dimensions.GetUnsortedFormatted()
                csvRow["Width " + defaultUnit] = dimensions[0]
                csvRow["Length " + defaultUnit] = dimensions[1]
                csvRow["Height " + defaultUnit] = dimensions[2]
            
                csvRow["Area cm^2"] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, "{0:.2f}".format(item.PhysicalAttributes.Area))
                csvRow["Mass kg"] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, "{0:.5f}".format(item.PhysicalAttributes.Mass))
                csvRow["Density kg/cm^2"] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, "{0:.5f}".format(item.PhysicalAttributes.Density))
                csvRow["Material"] = item.PhysicalAttributes.Material
                csvRow["Description"] = item.Description
                
                writer.writerow(csvRow)


    def WriteCutlistGaryDarby(self, stream: io.IOBase, bom: List[BomItem], prefs):
        # Init CutList Header
        stream.write('V2\n')
        if prefs.useCommaDecimal:
            stream.write('FormatSettings.decimalseparator,\n')
        else:
            stream.write('FormatSettings.decimalseparator.\n')
        stream.write('\n')
        stream.write('Required\n')

        for item in bom:
            #add parts:
            name = self.filterFusionCompNameInserts(item.Name)
            if prefs.ignoreUnderscorePrefixedComponents is False and prefs.stripUnderscorePrefix is True and name[0] == '_':
                name = name[1:]
            
            if prefs.sortDimensions:
                dims = item.PhysicalAttributes.Dimensions.GetSortedFormatted()
            else:
                dims = item.PhysicalAttributes.Dimensions.GetUnsortedFormatted()
            
            partStr = " {0} {1} {2} (thickness: {3})\n".format(dims[0], dims[1], name, dims[2])

            # add all instances of the component to the CutList:
            for i in range(0, item.Quantity):
                stream.write(partStr)

        # empty entry for available materials (sheets):
        stream.write('\n' + "Available" + '\n')