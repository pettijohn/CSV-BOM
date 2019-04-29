import io
import os
import sys
import unittest
import CSV_BOM_Core as Core
# from . import CSV_BOM_Core as Core
# from CSV_BOM_Core import BomItem, PhysicalAttributes, Dimensions, Helper

# https://www.autodesk.com/autodesk-university/class/Testing-Strategies-Python-Fusion-360-Add-Ins-2017#video
# https://github.com/theJenix/au2017/tree/master/test
# https://github.com/bommerio/f360mock
#sys.path.append(os.path.abspath('adsk-lib/defs'))3*

#import adsk
#import CSV_BOM

class MyTest(unittest.TestCase):

    def test_Dimensions_sort(self):
        d = Core.Dimensions(3, 2, 1, "3", "2", "1")
        s = d.GetSortedTuples()
        assert s[2] == [1, "1"]
        assert s[1] == [2, "2"]
        assert s[0] == [3, "3"]

        i = d.GetSortedInternal()
        assert i == [3, 2, 1]

        f = d.GetSortedFormatted()
        assert f == ["3", "2", "1"]

    def getDefaultBom(self):
        return Core.BomItem("My component name", 1, "This is my mocked component",
            Core.PhysicalAttributes(
                Core.Dimensions(3, 4, 5, "3 0/0", "4 0/0", "5 0/0"),
                3*4*5,
                4*5,
                3*4*5,
                1,
                "Water"
            ))

    def getDefaultPrefs(self):
        return {
                "onlySelComp": False,
                "incBoundDims": True,
                "splitDims": True,
                "sortDims": True,
                "ignoreUnderscorePrefComp": True,
                "underscorePrefixStrip": False,
                "ignoreCompWoBodies": True,
                "ignoreLinkedComp": True,
                "ignoreVisibleState": True,
                "incVol": False,
                "incArea": False,
                "incMass": False,
                "incDensity": False,
                "incMaterial": False,
                "generateCutList": False,
                "incDesc": False,
                "useComma": False
        }

    def test_CsvWrite(self):
        bomItem = self.getDefaultBom()

        prefs = self.getDefaultPrefs()
        
        h = Core.Helper()
        f = io.StringIO(newline='')
        h.WriteCsv(f, [bomItem], prefs)
    
        val = f.getvalue()
        expected = """Part name,Quantity,Width Inches,Length Inches,Height Inches
My component name,1,5 0/0,4 0/0,3 0/0
"""
        #self.assertMultiLineEqual(val == expected) #Fails due to \r\n and \n inconsistencies 
        self.assertEqual(val.splitlines(), expected.splitlines()) #Works as it compares contents of the array, each line of the string


    def test_cutlistGaryDarby(self):
        bomItem = self.getDefaultBom()
        prefs = self.getDefaultPrefs()

        h = Core.Helper()
        f = io.StringIO(newline='')
        # with open("./testfile.cutlist.txt", 'w') as f:
        h.WriteCutlistGaryDarby(f, [bomItem], prefs)
        
        expected = """V2
FormatSettings.decimalseparator.

Required
 5 0/0 4 0/0 My component name (thickness: 3 0/0)

Available
"""
        val = f.getvalue()
        self.assertEqual(val.splitlines(), expected.splitlines())


        