import sys
import os
import unittest
from CSV_BOM_helper import BomItem, PhysicalAttributes, Dimensions

# https://www.autodesk.com/autodesk-university/class/Testing-Strategies-Python-Fusion-360-Add-Ins-2017#video
# https://github.com/theJenix/au2017/tree/master/test
# https://github.com/bommerio/f360mock
#sys.path.append(os.path.abspath('adsk-lib/defs'))3*

#import adsk
#import CSV_BOM

class MyTest(unittest.TestCase):

    def test_pass(self):
        z = BomItem("My component name", 1, "This is my mocked component",
            PhysicalAttributes(
                Dimensions(3, 4, 5),
                3*4*5,
                4*5,
                3*4*5,
                1,
                "Water"
            ))
        pass