import io
import os
import sys
import unittest
from CSV_BOM_helper import BomItem, PhysicalAttributes, Dimensions, Helper

# https://www.autodesk.com/autodesk-university/class/Testing-Strategies-Python-Fusion-360-Add-Ins-2017#video
# https://github.com/theJenix/au2017/tree/master/test
# https://github.com/bommerio/f360mock
#sys.path.append(os.path.abspath('adsk-lib/defs'))3*

#import adsk
#import CSV_BOM

class MyTest(unittest.TestCase):

    def test_InstantiateBomItem(self):
        z = BomItem("My component name", 1, "This is my mocked component",
            PhysicalAttributes(
                Dimensions(3, 4, 5, "3 0/0", "4 0/0", "5 0/0"),
                3*4*5,
                4*5,
                3*4*5,
                1,
                "Water"
            ))
        pass

    def test_Dimensions_sort(self):
        d = Dimensions(3, 2, 1, "3", "2", "1")
        s = d.GetSortedTuples()
        assert s[0] == [1, "1"]
        assert s[1] == [2, "2"]
        assert s[2] == [3, "3"]

        i = d.GetSortedInternal()
        assert i == [1, 2, 3]

        f = d.GetSortedFormatted()
        assert f == ["1", "2", "3"]

    def test_CsvWrite(self):
        bomItem = BomItem("My component name", 1, "This is my mocked component",
            PhysicalAttributes(
                Dimensions(3, 4, 5, "3 0/0", "4 0/0", "5 0/0"),
                3*4*5,
                4*5,
                3*4*5,
                1,
                "Water"
            ))

        prefs = {
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
        
        h = Helper()
        f = io.StringIO()
        h.WriteCsv(f, [bomItem], prefs)
    
        print(f.getvalue())
        assert(len(f.getvalue()) > 0)