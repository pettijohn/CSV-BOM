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
        return Core.BomItem("My component name", 2, "This is my mocked component",
            Core.PhysicalAttributes(
                Core.Dimensions(3, 4, 5, "3 0/0", "4 0/0", "5 0/0"),
                3*4*5,
                4*5,
                3*4*5,
                1,
                "Water"
            ))

    def test_prefs(self):
        start = Core.CsvBomPrefs()
        # Override two defaults
        start.useCommaDecimal = True
        start.useQuantity = False
        j = start.to_json()

        finish = Core.CsvBomPrefs.from_json(j)
        assert finish.useCommaDecimal
        assert not start.useQuantity

    def test_CsvWrite(self):
        bomItem = self.getDefaultBom()

        prefs = Core.CsvBomPrefs()
        
        h = Core.Helper()
        f = io.StringIO(newline='')
        h.WriteCsv(f, [bomItem], prefs)
    
        val = f.getvalue()
        expected = """Part name,Quantity,Volume cm^3,Width Inches,Length Inches,Height Inches,Area cm^2,Mass kg,Density kg/cm^2,Material,Description
My component name,2,60,3 0/0,4 0/0,5 0/0,20.00,60.00000,1.00000,Water,This is my mocked component
"""
        #self.assertMultiLineEqual(val == expected) #Fails due to \r\n and \n inconsistencies 
        self.assertEqual(val.splitlines(), expected.splitlines()) #Works as it compares contents of the array, each line of the string

    def test_CsvWrite_noQuantity(self):
        bomItem = self.getDefaultBom()

        prefs = Core.CsvBomPrefs(useQuantity=False)
        
        h = Core.Helper()
        f = io.StringIO(newline='')
        h.WriteCsv(f, [bomItem], prefs)
    
        val = f.getvalue()
        expected = """Part name,Volume cm^3,Width Inches,Length Inches,Height Inches,Area cm^2,Mass kg,Density kg/cm^2,Material,Description
My component name,60,3 0/0,4 0/0,5 0/0,20.00,60.00000,1.00000,Water,This is my mocked component
My component name,60,3 0/0,4 0/0,5 0/0,20.00,60.00000,1.00000,Water,This is my mocked component
"""
        #self.assertMultiLineEqual(val == expected) #Fails due to \r\n and \n inconsistencies 
        self.assertEqual(val.splitlines(), expected.splitlines()) #Works as it compares contents of the array, each line of the string


    def test_cutlistGaryDarby(self):
        bomItem = self.getDefaultBom()
        prefs = Core.CsvBomPrefs()
        prefs.sortDimensions = True

        h = Core.Helper()
        f = io.StringIO(newline='')
        # with open("./testfile.cutlist.txt", 'w') as f:
        h.WriteCutlistGaryDarby(f, [bomItem], prefs)
        
        expected = """V2
FormatSettings.decimalseparator.

Required
 5 0/0 4 0/0 My component name (thickness: 3 0/0)
 5 0/0 4 0/0 My component name (thickness: 3 0/0)

Available
"""
        val = f.getvalue()
        self.assertEqual(val.splitlines(), expected.splitlines())


        