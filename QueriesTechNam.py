import os
import sys
import time
import subprocess
import uiautomation as auto
from Mongo_Client import Mongo_Client
from BWField import BWField
from constans import QueryList

class UIScrapy():
    def __init__(self):
        self.client = Mongo_Client()
        self.CurrentQueryName = ''
        self.db = None
        self.QueryWindow = None
        self.dbAllinOne = self.client.db["allinone"]
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
                    if child or tryCount > 200:
                        break
                    time.sleep(1)
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
                if item.Name != 'Key Figures':
                    arr = item.Name.split(']')
                    field = BWField(arr[0][1:].strip(), arr[1].strip(),self.CurrentQueryName)
                    print(field)
                    self.db.insert_one(field.__dict__)
                    self.dbAllinOne.insert_one(field.__dict__)


    def ReadTechnicalNameFromBexQuery(self):
        if self.QueryWindow is None:
            self.QueryWindow = auto.WindowControl(searchDepth=1, RegexName = "BEx Query Designer*")

        RCColumn = self.QueryWindow.PaneControl(
            searchDepth=1, Name='Rows/Columns')
        for c, d in auto.WalkControl(RCColumn):
            if isinstance(c, auto.TreeControl):
               self.ExpandTreeItem(c)

    def OpenQueryByName(self, queryName):
        if self.QueryWindow is None:
            self.QueryWindow = auto.WindowControl(searchDepth=1, RegexName="BEx Query Designer*")
        
        self.QueryWindow.SetActive()
        #self.QueryWindow.SetTopmost()

        b = self.QueryWindow.ToolBarControl(Name="Standard")

        ctrlOpen = b.MenuItemControl(Name="Open...")

        if ctrlOpen is None:
            return
        
        ctrlOpen.Click(10)

        dlgOpen = auto.WindowControl(RegexName="Open Query*")

        find = dlgOpen.ListItemControl(Name="Find")
        if find:
            find.Click(10)

        ctrlEdit = dlgOpen.EditControl(Name="Search Term")
        if ctrlEdit:
            ctrlEdit.SendKeys(self.CurrentQueryName,0.05)

        btnFind = dlgOpen.ButtonControl(Name="Find")
        if btnFind:
            btnFind.Click(5)

        btnOpen = dlgOpen.ButtonControl(AutomationId = "btnOK")
        if btnOpen:
            btnOpen.Click(10)

        trycnt = 1
        while trycnt <10:
            try:
                self.QueryWindow.SetActive()
                trycnt = 11
            except:
                print("screen not response")
                time.sleep(5)
                trycnt = trycnt+1

        
        pnlInfoProvider = self.QueryWindow.PaneControl(Name="InfoProvider")

        if pnlInfoProvider:
            trycnt = 1
            while trycnt < 10:
                try:
                    for c, d in auto.WalkControl(pnlInfoProvider):
                        if isinstance(c, auto.TreeControl):
                            trycnt = 11
                            break
                                    
                except:
                    print("screen no response! try again")
                    time.sleep(5)
                    trycnt = trycnt+1                    
        

        # click Technical name
        v = self.QueryWindow.ToolBarControl(Name="View")

        techname = v.MenuControl(Name="Technical Names")
        if techname:
            ischecked = False
            while not ischecked:
                techname.Click()
                for c in techname.GetChildren():
                    if c.Name == "[Key] Text":
                        value = c.GetLegacyIAccessiblePattern().State
                        if value == 1048592:
                            ischecked = True
                            break


    def GetTechnicalNameFromQueries(self):
        for queryName in QueryList:
            print(queryName)
            self.CurrentQueryName = queryName
            self.db = self.client.db[queryName]

            #open query
            self.OpenQueryByName(queryName)
            #read technical name from query
            self.ReadTechnicalNameFromBexQuery()

if __name__ == "__main__":
    uiscrapy = UIScrapy()
    uiscrapy.GetTechnicalNameFromQueries()
    #uiscrapy.ReadTechnicalNameFromBexQuery()
    
