import os
import sys
import time
import subprocess
import uiautomation as auto
from Mongo_Client import Mongo_Client
from BWField import BWField


class UIScrapy():
    def __init__(self):
        self.client = Mongo_Client()
        self.OTSVFields = self.client.db["OTSV"]

        pass

    def GetFirstChild(self,item: auto.Control):
        if isinstance(item, auto.TreeItemControl):
            ecpt = item.GetExpandCollapsePattern()
            if ecpt and ecpt.ExpandCollapseState == auto.ExpandCollapseState.Expanded:
                child = None
                tryCount = 0
                #some tree items need some time to finish expanding
                while not child:
                    tryCount += 1
                    child = item.GetFirstChildControl()
                    if child or tryCount > 20:
                        break
                    time.sleep(0.05)
                return child
        else:
            return item.GetFirstChildControl()


    def GetNextSibling(self,item: auto.Control):
        return item.GetNextSiblingControl()


    def ExpandTreeItem(self, treeItem: auto.TreeItemControl):
        lst = []
        for item, depth in auto.WalkTree(treeItem, getFirstChild=self.GetFirstChild, getNextSibling=self.GetNextSibling, includeTop=True, maxDepth=1):
            # or item.ControlType == auto.ControlType.TreeItemControl
            if isinstance(item, auto.TreeItemControl):
                print(item.Name)
                arr = item.Name.split(']')
                field = BWField(arr[0][1:].strip(), arr[1].strip())
                print(field)
                self.OTSVFields.insert_one(field.__dict__)


    def ReadTechnicalNameFromBexQuery(self):
        note = auto.WindowControl(searchDepth=1, RegexName = "BEx Query Designer*")
        #note.SetActive()
        #note.SetTopmost()

        RCColumn = note.PaneControl(searchDepth=1, Name='Rows/Columns')
        for c, d in auto.WalkControl(RCColumn):
            if isinstance(c, auto.TreeControl):
               self.ExpandTreeItem(c)

    


if __name__ == "__main__":
    uiscrapy = UIScrapy()
    uiscrapy.ReadTechnicalNameFromBexQuery()
    
