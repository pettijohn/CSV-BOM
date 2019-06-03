# This file does all the hard work of translating a list of objects into a variety of CSV formats.
# This file is separate to enable unit testing the logic of CSV generation from the Fusion 360 Addin runtime. 

import csv
import io
import re
import collections
import json
from typing import List

class OutputFormats:
    GaryDarby = "Cutlist (Gary Darby)"
    FullCsv="Full CSV (All properties)"
    FullCsvTemplate = """Part name,Quantity,Volume cm^3,Width {},Length {},Height {},Area cm^2,Mass kg,Density kg/cm^2,Material,Description
Name,Quantity,Volume,Width,Length,Height,Area,Mass,Density,Material,Description"""
    MinimalCsvTemplate = """Part name,Quantity,Width {},Length {},Height {}
Name,Quantity,Width,Length,Height"""
    MaxcutTemplate = """Name,Length,Width,Quantity,Notes,Can Rotate (0 = No / 1 = Yes / 2 = Same As Material),Material,Edging Length 1,Edging Length 2,Edging Width 1,Edging Width 2,Include Edging Thickness,Note 1,Note 2,Note 3,Note 4,Group,Report Tags,Import ID,Parent ID,Library Item Name,Holes Length 1,Holes Length 2,Holes Width 1,Holes Width 2,Grooving Length 1,Grooving Length 2,Grooving Width 1,Grooving Width 2,Material Tag,Edging Length 1 Tag,Edging Length 2 Tag,Edging Width 1 Tag,Edging Width 2 Tag
Name,Length,Width,Quantity,Thickness,SameAsSheet,Material,,,,,,,,,,,,,,,,,,,,,,,,,,,""" 
    
    all = {
        FullCsv: FullCsvTemplate,
        "Minimal CSV (Dimensions and Name only)": MinimalCsvTemplate,
        "Cutlist (Maxcut)": MaxcutTemplate,
        #"Cutlist (CutList Plus fx)": "" ,
        GaryDarby: ""
    }
    

class CsvBomPrefs:
    # Following serialization pattern from https://stackoverflow.com/a/33270983/435368
    def __init__(self, onlySelectedComponents=False, sortDimensions=True, 
        ignoreUnderscorePrefixedComponents=True, stripUnderscorePrefix=False,
        ignoreCompWoBodies=True, ignoreLinkedComponents=True,
        ignoreVisibleState=True, useCommaDecimal=False, useQuantity=True, lengthUnitString="", 
        outputFormat=OutputFormats.FullCsv, **kwargs):
        self.onlySelectedComponents = onlySelectedComponents
        self.sortDimensions=sortDimensions
        self.ignoreUnderscorePrefixedComponents=ignoreUnderscorePrefixedComponents
        self.stripUnderscorePrefix = stripUnderscorePrefix
        self.ignoreCompWoBodies=ignoreCompWoBodies
        self.ignoreLinkedComponents=ignoreLinkedComponents
        self.ignoreVisibleState=ignoreVisibleState
        self.useCommaDecimal=useCommaDecimal
        self.useQuantity=useQuantity
        self.lengthUnitString=lengthUnitString
        self.outputFormat=outputFormat

    @classmethod
    def from_json(cls, json_str):
        json_dict = json.loads(json_str)
        return cls(**json_dict)

    def to_json(self):
        return json.dumps(self.__dict__)
    

class Dimensions:
    """ Internal values should be floats and sortable (cm), while Formatted are strings (incl. fractional inches) """

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

    def replacePointDelimterOnPref(self, useComma: bool, value: str):
        """ Replace decimal point in str(number) with a comma """
        if (useComma):
            return str(value).replace(".", ",")
        return str(value)

    def SaveFile(self, filename, bom: List[BomItem], prefs: CsvBomPrefs):
        """ Determine the correct output option from prefs.outputFormat"""
        with open(filename, 'w', newline='') as csvFile:
            if prefs.outputFormat == OutputFormats.GaryDarby:
                self.WriteCutlistGaryDarby(csvFile, bom, prefs)
            else:
                template = self.ParseCsvTemplate(prefs, OutputFormats.all[prefs.outputFormat])
                self.WriteCsvFromTemplate(csvFile, bom, prefs, template)

    def ParseCsvTemplate(self, prefs: CsvBomPrefs, template) -> collections.OrderedDict:
        """Take a two-line CSV template and parse it into a dict for writing to arbitrary CSVs"""
        f = io.StringIO(template)

        d = collections.OrderedDict()
        reader = csv.reader(f)
        for row in reader:
            if reader.line_num == 1:
                fields = list(map(lambda x: x.format(prefs.lengthUnitString), row))
            if reader.line_num == 2:
                i = 0
                for k in fields:
                    d[k] = row[i].format(prefs.lengthUnitString)
                    i+=1
                break
        return d



        


    def WriteCsvFromTemplate(self, f, bom: List[BomItem], prefs: CsvBomPrefs, template: collections.OrderedDict):
        # Valid fields:
        fieldNames = ["Name","Quantity","Volume","Width","Length","Height","Area","Mass","Density","Material","Description"] 

        # Construct an inverted template, because we'll want it for mapping values
        # Find the index of the value, find the corresponding key, add to ordered dict
        invertedTemplate = collections.OrderedDict()
        for v in template.values():
            if v in fieldNames:
                i = list(template.values()).index(v)
                invertedTemplate[v] = list(template.keys())[i]
               
        if(not prefs.useQuantity and "Quantity" in template.values()):
            del template[invertedTemplate["Quantity"]]

        csvKeys = list(template.keys())
        valueKeys = list(invertedTemplate.keys())
            
        # Extras Action = Ignore means that when items in the dict are encountered that
        #  are not present in the header, they are ignored. This lets us perform all the
        #  conditional logic of which fields to include just once up front when 
        #  building the header.
        writer = csv.DictWriter(f, fieldnames=csvKeys, extrasaction='ignore')
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
                
                if "Name" in valueKeys:
                    csvRow[invertedTemplate["Name"]] = name
                if "Quantity" in valueKeys:
                    csvRow[invertedTemplate["Quantity"]] = item.Quantity
                if "Volume" in valueKeys:
                    csvRow[invertedTemplate["Volume"]] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, item.PhysicalAttributes.Volume)
                
                #############
                if prefs.sortDimensions:
                    dimensions = item.PhysicalAttributes.Dimensions.GetSortedFormatted()
                else:
                    dimensions = item.PhysicalAttributes.Dimensions.GetUnsortedFormatted()
                if "Width" in valueKeys:                
                    csvRow[invertedTemplate["Width"]] = dimensions[0]
                if "Length" in valueKeys:
                    csvRow[invertedTemplate["Length"]] = dimensions[1]
                if "Height" in valueKeys:
                    csvRow[invertedTemplate["Height"]] = dimensions[2]
            
                # Fusion 360 API doesn't make it easy to convert area, mass, or density.
                # This code works:
                #  design.fusionUnitsManager.convert(1.0, "in * in * in / lbmass", "cm * cm * cm / kg") 
                # But the units manager doesn't expose user preferences other than lenght/distance units
                #  http://help.autodesk.com/view/fusion360/ENU/?guid=GUID-40dda15b-8dec-4122-b0fa-cbd604cd35b5
                if "Area" in valueKeys:
                    csvRow[invertedTemplate["Area"]] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, "{0:.2f}".format(item.PhysicalAttributes.Area))
                if "Mass" in valueKeys:
                    csvRow[invertedTemplate["Mass"]] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, "{0:.5f}".format(item.PhysicalAttributes.Mass))
                if "Density" in valueKeys:
                    csvRow[invertedTemplate["Density"]] = self.replacePointDelimterOnPref(prefs.useCommaDecimal, "{0:.5f}".format(item.PhysicalAttributes.Density))
                if "Material" in valueKeys:
                    csvRow[invertedTemplate["Material"]] = item.PhysicalAttributes.Material
                if "Description" in valueKeys:
                    csvRow[invertedTemplate["Description"]] = item.Description
                
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

        useFractions = None
        for item in bom:
            # Look at all dimensions for a decimal or fraction separator to determine integer handling
            for dim in item.PhysicalAttributes.Dimensions.GetUnsortedFormatted():
                if '/' in dim:
                    useFractions = True
                    break
                if (not prefs.useCommaDecimal and '.' in dim) or (prefs.useCommaDecimal and ',' in dim):
                    useFractions = False
            if useFractions is not None:
                break
        if useFractions is None:
            useFractions = False

        for item in bom:
            #add parts:
            name = self.filterFusionCompNameInserts(item.Name)
            if prefs.ignoreUnderscorePrefixedComponents is False and prefs.stripUnderscorePrefix is True and name[0] == '_':
                name = name[1:]
            
            if prefs.sortDimensions:
                dims = item.PhysicalAttributes.Dimensions.GetSortedFormatted()
            else:
                dims = item.PhysicalAttributes.Dimensions.GetUnsortedFormatted()

            for i in range(len(dims)):
                # GD Cutlist requires integer legths to end with "0/0" when using fractions
                if useFractions and '/' not in dims[i]:
                    dims[i] += " 0/0"
            
            partStr = " {0}\t{1}\t{2} (thickness: {3})\n".format(dims[0], dims[1], name, dims[2])

            # add all instances of the component to the CutList:
            for i in range(0, item.Quantity):
                stream.write(partStr)

        # empty entry for available materials (sheets):
        stream.write('\n' + "Available" + '\n')