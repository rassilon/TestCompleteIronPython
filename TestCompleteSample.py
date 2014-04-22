import System
import System.Runtime.InteropServices
import time
import clr
clr.AddReferenceToFileAndPath(r"C:\Program Files (x86)\SmartBear\TestComplete 9\Connected Apps\.NET\AutomatedQA.TestComplete.CSConnectedApp.dll")
from AutomatedQA.TestComplete import Connect

import pprint

def RGB(red, green, blue):
    result = (min(red, 255) << 16) + (min(green, 255) << 8) + min(blue, 255)
    # Deal with converting Python Long to 32bit signed integer range:
    if result & 0x80000000:
        result += -0x100000000
    return result

class TestComplete:
    def __init__(self):
        self.tc = Connect
        self.TestedApps = WrappedCollection(Connect.TestedApps)
        self.Sys = WrappedCollection(Connect.Sys)
        self.aqConvert = WrappedObj(Connect.aqConvert, None)
        self.aqString = WrappedObj(Connect.aqString, None)
        self.Runner = WrappedObj(Connect.Runner, None)
        self.Log = WrappedObj(Connect.Log, None)
        self.Project = WrappedObj(Connect.Project, None)

    def RunTest(self, logName, projectName, suiteName):
        self.tc.RunTest(logName, projectName, suiteName)
    def StopTest(self):
        self.tc.StopTest()

class WrappedObj(object):
    def __init__(self, obj, name):
        self.____obj = obj
        self.____name = name
    def __getattr__(self, key):
        return WrappedObj(self.____obj[key], key)
    def __setattr__(self, key, value):
        # Prevent infinite recursion:
        if key.startswith('_WrappedObj____'):
            object.__setattr__(self, key, value)
        else:
            self.____obj[key] = value
    def __getitem__(self, key):
        return WrappedObj(self.____obj[key], key)
    def __call__(self, *args):
        newArgs = []
        for arg in args:
            if isinstance(arg, WrappedObj):
                newArgs.append(arg.____obj)
            else:
                newArgs.append(arg)
        return WrappedObj(apply(self.____obj, newArgs), None)
    def __str__(self):
        value = self.____obj.UnWrap()
        return str(value)
    def __nonzero__(self):
        value = self.____obj.UnWrap()
        return value == True
    def __coerce__(self, other):
        value = self.____obj.UnWrap()
        if isinstance(other, WrappedObj):
            return (value, other.____obj.UnWrap())
        return (value, other)
    def __add__(self, other):
        if isinstance(other, WrappedObj):
            return self.____obj.UnWrap() + other.____obj.UnWrap()
        return self.____obj.UnWrap() + other
    def __lt__(self, other):
        value = self.____obj.UnWrap()
        if isinstance(other, WrappedObj):
            return value < other.____obj.UnWrap()
        return self.____obj.UnWrap() < other.____obj.UnWrap()
    def __gt__(self, other):
        value = self.unwrap()
        if isinstance(other, WrappedObj):
            return value > other.____obj.UnWrap()
        return value > other
    def __eq__(self, other):
        value = self.unwrap()
        if isinstance(other, WrappedObj):
            return value == other.unwrap()
        return value == other
    def __ne__(self, other):
        return not self.__eq__(other)
    def __le__(self, other):
        value = self.____obj.UnWrap()
        if isinstance(other, WrappedObj):
            return value <= other.unwrap()
        return value <= other
    def __ge__(self, other):
        value = self.____obj.UnWrap()
        if isinstance(other, WrappedObj):
            return value >= other.unwrap()
        return value >= other
    def __repr__(self):
        return repr(self.____obj.UnWrap())
    def unwrap(self):
        return self.____obj.UnWrap()

class WrappedCollection:
    def __init__(self, obj):
        self.obj = obj
    def __getattr__(self, key):
        try:
            return WrappedObj(self.obj[key], key)
        except AttributeError as e:
            return getattr(self.obj, key)
    def __getitem__(self, key):
        return WrappedObj(self.obj[key], key)

#/*
#  The example demonstrates the main operations that can be executed over objects: 
#  clicks, dragging, key presses.
#  The sample works only under the English version of Windows.  
#*/

def DrawString(wMain, x, y, strDraw, strFontName):
  #var wndFonts, wPicture, wFonts;
    wPicture = wMain.Window("AfxFrameOrView*", "", 1).Window("Afx*:8", "");
    wPicture.Drag(x, y, 100, 100);
  
    wFonts = wMain.Parent.Window("Afx*", "Fonts").Window("#*", "");
    if (not wFonts.Exists or (not wFonts.VisibleOnScreen)):
        wMain.MainMenu.Check("View|Text Toolbar", true);
    
    wndFonts = wMain.Parent.WaitWindow("Afx*:8*", "Fonts");
    if (wndFonts.Exists): 
        wndFonts.Position(60, 40, wndFonts.Width, wndFonts.Height);

    if (wFonts.Window("ComboBox", "", 1).wText != strFontName):    
        wFonts.Window("ComboBox", "", 1).ClickItem(strFontName);
    if (wFonts.Window("ComboBox", "", 2).wText != "72"):
        wFonts.Window("ComboBox", "", 2).ClickItem("72");
    wPicture.Window("AfxWnd*", "").Window("Edit").Keys(strDraw);
    wPicture.Click(x - 5, y - 5);

def DrawString_w7(wPicture, wRibbon, x, y, strDraw, strFontName):
  #var wPropertyPage, wToolBar, wEdit;
  wPicture.Click(x, y);
  wPropertyPage = wRibbon.Pane("Lower Ribbon").Client(0).PropertyPage("Text");
  wToolBar = wPropertyPage.ToolBar("Font");
  
  if (wToolBar.ComboBox("Font family").Edit("*").Text != strFontName):
    wToolBar.ComboBox("Font family").Button("Open").Click(7, 9);
    wToolBar.ComboBox("Font family").Button("Close").Click(7, 9);
    wEdit = wPicture.Parent.Parent.Window("UIRibbonCommandBarDock", "UIRibbonDockTop").Window("UIRibbonCommandBar", "Ribbon").Window("UIRibbonWorkPane", "Ribbon").Window("NUIPane").PropertyPage("Ribbon").Window("NetUICtrlNotifySink", "", 1).Window("RICHEDIT50W");
    wEdit.Click();
    wEdit.wText = strFontName;
    wEdit.Keys("[Enter]");
  if (wToolBar.ComboBox("Font size").Edit("*").Text != "72"):
    wToolBar.ComboBox("Font size").Button("Open").Click(8, 9);
    wToolBar.ComboBox("Font size").Button("Close").Click(7, 9);
    wEdit = wPicture.Parent.Parent.Window("UIRibbonCommandBarDock", "UIRibbonDockTop").Window("UIRibbonCommandBar", "Ribbon").Window("UIRibbonWorkPane", "Ribbon").Window("NUIPane").PropertyPage("Ribbon").Window("NetUICtrlNotifySink", "", 2).Window("RICHEDIT50W");
    wEdit.Click();
    wEdit.wText = "72";  
    wEdit.Keys("[Enter]");
    
  wToolBar = wPropertyPage.ToolBar("Background");
  wToolBar.Button("Transparent").ClickButton();
  
  wPicture.Window("AfxWnd42u").Window("RICHEDIT50W").Keys(strDraw);  
  wPicture.Click(x - 15, y);


def SetColor(Paint, r, g, b, w7):
  #var dlgEditColors;

  if (w7):
    Paint.Window("MSPaintApp", "*").Window("UIRibbonCommandBarDock", "UIRibbonDockTop").Window("UIRibbonCommandBar", "Ribbon").Window("UIRibbonWorkPane", "Ribbon").Window("NUIPane").PropertyPage("Ribbon").Pane("Lower Ribbon").Client(0).PropertyPage("Home").ToolBar("Colors").Button("Edit colors").ClickButton();
  else:
    Paint.Window("MSPaintApp", "*").MainMenu.Click("Colors|Edit Colors...");
    
  dlgEditColors = Paint.Window("#32770", "Edit Colors");
  if (dlgEditColors.Window("Button", "&Define Custom Colors >>").Enabled):
    dlgEditColors.Window("Button", "&Define Custom Colors >>").ClickButton();
  dlgEditColors.Window("Edit", "", 4).wText = r;
  dlgEditColors.Window("Edit", "", 5).wText = g;
  dlgEditColors.Window("Edit", "", 6).wText = b;
  dlgEditColors.Window("Button", "OK").ClickButton();

def Hello_w7(tc, mspaint):
  #var wMain, wRibbon, wDialog, wColors, wTools, wPicture, strFontName;

  strFontName = "Arial";
  wMain   = mspaint.WaitWindow("MSPaintApp", "*", 1, 5000);
  wMain.Maximize();

  # Create a new image 800x600
  wRibbon = wMain.Window("UIRibbonCommandBarDock", "UIRibbonDockTop", -1).Window("UIRibbonCommandBar", "Ribbon", -1).Window("UIRibbonWorkPane", "Ribbon", -1).Window("NUIPane", "*", -1)
  wRibbon = wRibbon.PropertyPage("Ribbon");
  #print "wRibbon Exists: " + wRibbon.Exists
  wRibbon.Button(1).GridDropDownButton("Application menu").Click();
  mspaint.Window("Net UI Tool Window Layered").Window("NetUIHWND").Pane(0).Client(0).Grouping(0).MenuItem("New").Click();
  #mspaint.Find("NetUIHWND").Pane(0).Client(0).Grouping(0).MenuItem("New").Click()

  wDialog = mspaint.WaitWindow("#*", "*", 1, 2000);
  if (wDialog.Exists):
    wDialog.Keys("n");
  
  wRibbon.Button(1).GridDropDownButton("Application menu").Click();
  mspaint.Window("Net UI Tool Window Layered").Window("NetUIHWND").Pane(0).Client(0).Grouping(3).MenuItem("Properties").Click();
  wDialog = mspaint.Window("#32770", "Image Properties");
  wDialog.Window("Button", "&Pixels").ClickButton();
  wDialog.Window("Button", "Co&lor").ClickButton();
  wDialog.Window("Edit", "", 1).wText = 800;
  wDialog.Window("Edit", "", 2).wText = 600;
  wDialog.Window("Button", "OK").ClickButton();
   
  wColors = wRibbon.Pane("Lower Ribbon").Client(0).PropertyPage("Home").ToolBar("Colors").Grouping(0);
  wTools = wRibbon.Pane("Lower Ribbon").Client(0).PropertyPage("Home").ToolBar("Tools");
  wPicture = wMain.Window("MSPaintView").Window("Afx:*:8");
  
  # Background color
  SetColor(mspaint, 0, 0, 0, True);
  wTools.Button("Fill with color").ClickButton();       
  wPicture.Click(10, 10);

  # Choose "Text"
  wTools.Button("Text").ClickButton();       

  # Shade for T - E - S - T
  SetColor(mspaint, 100, 100, 100, True);
  DrawString_w7(wPicture, wRibbon, 152, 152, " ", strFontName); # Enlarge the edit box after changing the font size
  DrawString_w7(wPicture, wRibbon, 152, 152, "T", strFontName);
  DrawString_w7(wPicture, wRibbon, 252, 152, "E", strFontName);
  DrawString_w7(wPicture, wRibbon, 352, 152, "S", strFontName);
  DrawString_w7(wPicture, wRibbon, 452, 152, "T", strFontName);
  
  # T - E - S - T
  SetColor(mspaint, 255, 0, 0, True);
  DrawString_w7(wPicture, wRibbon, 150, 150, "T", strFontName);
  SetColor(mspaint, 0, 255, 0, True);
  DrawString_w7(wPicture, wRibbon, 250, 150, "E", strFontName);
  SetColor(mspaint, 0, 0, 255, True);
  DrawString_w7(wPicture, wRibbon, 350, 150, "S", strFontName);
  SetColor(mspaint, 255, 0, 255, True);
  DrawString_w7(wPicture, wRibbon, 450, 150, "T", strFontName);

def Hello(tc, mspaint):
    #var wndFonts, bIsOldPaint,
    #    bIsWinXP, bIsWinNT, bIsWin2K, bHasToolChild,
    #    wColors, wMain, wTools, wPicture,
    #    strOSName, strFontName;

    strOSName = tc.Sys.OSInfo.Name;
    bIsWinNT = strOSName == "WinNT";
    bIsWin2K = strOSName == "Win2000";
    bIsWinXP = strOSName == "WinXP";
    bHasToolChild = bIsWin2K;
  
    strFontName = "Arial";
  
    wndFonts = mspaint.WaitWindow("Afx*:8*", "Fonts");
    if (wndFonts.Exists): 
        wndFonts.Position(60, 40, wndFonts.Width, wndFonts.Height);
    
    wMain   = mspaint.WaitWindow("MSPaintApp", "*", 1, 5000);
    wColors = wMain.Window("AfxControlBar*", "Colors").Window("AfxWnd*", "Colors");
    wTools  = wMain.Window("AfxControlBar*", "Tools").Window("AfxWnd*", "Tools");

    # Create a new image 800x600
    wMain.Maximize();
    wMain.MainMenu.Check("View|Tool Box", true);
    wMain.MainMenu.Check("View|Color Box", true);
   
    wMain.MainMenu.Click("File|New");
    wndFonts = mspaint.WaitWindow("#*", "*", 1, 1000);
    if (wndFonts.Exists):
        wndFonts.Keys("n");
    
    wMain.MainMenu.Click("Image|Attributes...");
    wndFonts = mspaint.WaitWindow("#*", "Attributes", 1, 2000);
    if (bIsWinNT):
        wndFonts.Window("Button", "&Pels").ClickButton();
    else:
        wndFonts.Window("Button", "&Pixels").ClickButton();
    wndFonts.Window("Button", "Co&lors").ClickButton();
    wndFonts.Window("Edit", "", 1).wText = "800";
    wndFonts.Window("Edit", "", 2).wText = "600";
    wndFonts.Window("Button", "OK").ClickButton();

    # Background color
    SetColor(mspaint, 0, 0, 0, False);
  
    if (bHasToolChild):
        wTools.Window("ToolChild", "", 4).Click(13, 12)
    else:
        wTools.Click(43, 36);
      
    wMain.Window("AfxFrameOrView*", "", 1).Window("Afx*:8", "").Click(10, 10);
  
    # Choose "Text"
    if (bHasToolChild):
        wTools.Window("ToolChild", "", 10).Click(10, 9);    
    else:
        wTools.Click(43, 110);  
  
    wTools.Click(25, 252);
  
    # Shade for T - E - S - T
    SetColor(mspaint, 100, 100, 100, False);
    DrawString(wMain, 150, 150, " ", strFontName); # Enlarge the edit box after changing the font size
  
    DrawString(wMain, 152, 152, "T", strFontName);
    DrawString(wMain, 252, 152, "E", strFontName);
    DrawString(wMain, 352, 152, "S", strFontName);
    DrawString(wMain, 452, 152, "T", strFontName);

    # T - E - S - T
    SetColor(mspaint, 255, 0, 0, False);
    DrawString(wMain, 150, 150, "T", strFontName);
    SetColor(mspaint, 0, 255, 0, False);
    DrawString(wMain, 250, 150, "E", strFontName);
    SetColor(mspaint, 0, 0, 255, False);
    DrawString(wMain, 350, 150, "S", strFontName);
    SetColor(mspaint, 255, 0, 255, False);
    DrawString(wMain, 450, 150, "T", strFontName);
   
def paint(tc):
  # Run the tested application
  tc.TestedApps.mspaint_NT.Run()

  System.Threading.Thread.CurrentThread.Join(5000)
  mspaint = tc.Sys.WaitProcess("mspaint", 5000)

  #pprint.pprint(dir(mspaint))
  #print mspaint.name
  #if not mspaint.Exists:
  #  tc.Log.Error("The \"MSPaint\" process was not found.",
  #  "Make sure the path was specified correctly in the list of tested applications.")
  #  return
  
  OsVersion = tc.aqConvert.StrToInt(tc.aqString.GetChar(tc.Sys.OSInfo.Version, 0) + tc.aqString.GetChar(tc.Sys.OSInfo.Version, 2));
  if (OsVersion > 60):
    Hello_w7(tc, mspaint)
  else:
    Hello(tc, mspaint)
  
  mspaint.Close();
  tc.Sys.Keys("n");
  while (mspaint.Exists()):
     time.sleep(0.500);

def Open(tc):
  tc.TestedApps.orders.Run();

def Close(tc):
  #var oProcess, oMainForm, oMsgBox;
  oProcess = tc.Sys.Process('Orders');
  oMainForm = oProcess.WaitWinFormsObject('MainForm', 5000);
  
  oMainForm.Close(0); 
  oMsgBox = tc.Sys.Process('orders').Window('#32770', 'Confirmation');
  if (oMsgBox.Exists):
    oMsgBox.Window('Button', '&No').ClickButton();

  #// Waiting until the Orders process is terminated	
  while (oProcess.Exists):
    time.sleep(0.1)

def LoadMyTable(tc):
  #var oProcess, oMainForm, oToolBar, oOpenDlg, oOrderDlg;
  oProcess = tc.Sys.Process('Orders');
  oMainForm = oProcess.WaitWinFormsObject('MainForm', 5000);
  oToolBar = oMainForm.WinFormsObject('ToolBar');
  oToolBar.ClickItem(1, False);
   
  oOpenDlg = tc.Sys.Process('Orders').Window('#32770', 'Open');
  if (not oOpenDlg.Exists):
    raise BaseException('Open dialog not found!')

  oOpenDlg.Window('ComboBoxEx32', '', 1).wText = Project.Path + '..\\..\\..\\MyTable.tbl';
  oOpenDlg.Window('Button', '&Open').ClickButton();

def ChangeRecord(tc):
  #var oProcess, oMainForm, oToolBar, oOpenDlg, oOrderDlg;
  oProcess = tc.Sys.Process('Orders');
  oMainForm = oProcess.WaitWinFormsObject('MainForm', 5000);
  oToolBar = oMainForm.WinFormsObject('ToolBar');
  oToolBar.ClickItem(1, False);
   
  oOpenDlg = tc.Sys.Process('Orders').Window('#32770', 'Open');
  if (not oOpenDlg.Exists):
    raise BaseException('Open dialog not found!')

  oOpenDlg.Window('ComboBoxEx32', '', 1).wText = tc.Project.Path + '..\\..\\..\\MyTable.tbl';
  oOpenDlg.Window('Button', '&Open').ClickButton();
    
  oMainForm.WinFormsObject('OrdersView').ClickItem('Samuel Clemens', 0);
  oToolBar.ClickItem(5, False);
  
  oOrderDlg = tc.Sys.Process('Orders').WaitWinFormsObject('OrderForm', 5000); 
  if (not oOrderDlg.Exists):
    raise BaseException('Open dialog not found!')

  oOrderDlg.WinFormsObject('Group').WinFormsObject('ProductNames').ClickItem('FamilyAlbum');
  strValue = oProcess.AppDomain('Orders.exe').dotNET.System.String.Copy('123123123123');
  ctrl = oOrderDlg.WinFormsObject('Group').WinFormsObject('CardNo')
  ctrl.set_Text(strValue)
  oOrderDlg.WinFormsObject('ButtonOK').ClickButton();

def AddRecord(tc):
 
  oProcess = tc.Sys.Process('Orders');
  oMainForm = oProcess.WaitWinFormsObject('MainForm', 5000);
  oToolBar = oMainForm.WinFormsObject('ToolBar');

  oToolBar.ClickItem(4, False);
  oOrderDlg = tc.Sys.Process('Orders').WaitWinFormsObject('OrderForm', 5000); 
  if (not oOrderDlg.Exists):
    raise BaseException('Open dialog not found!')

  oOrderDlg.WinFormsObject('Group').WinFormsObject('ProductNames').set_SelectedIndex(2);
  oOrderDlg.WinFormsObject('Group').WinFormsObject('Quantity').set_Text('2');
  oOrderDlg.WinFormsObject('Group').WinFormsObject('Date').wDate = '11/02/1999';
  oOrderDlg.WinFormsObject('Group').WinFormsObject('Customer').set_Text('David Waters');
  oOrderDlg.WinFormsObject('Group').WinFormsObject('Street').set_Text('3, Music Street');
  oOrderDlg.WinFormsObject('Group').WinFormsObject('City').set_Text('Liverpool');
  oOrderDlg.WinFormsObject('Group').WinFormsObject('State').set_Text('Great Britain');
  oOrderDlg.WinFormsObject('Group').WinFormsObject('MasterCard').set_Checked(True);
  oOrderDlg.WinFormsObject('Group').WinFormsObject('CardNo').set_Text('1357902468');
  oOrderDlg.WinFormsObject('Group').WinFormsObject('ExpDate').wDate = '03/03/2008';
  oOrderDlg.WinFormsObject('ButtonOK').ClickButton();

def TestOrders(tc):
  try:
    Open(tc);
    ChangeRecord(tc);
    AddRecord(tc);
    Close(tc);
  except Exception as exception:
    tc.Log.Error("Exception", exception.ToString());
    raise

tc = TestComplete()
# MSPaint test:
#try:
#    tc.RunTest("Empty", <valid projectName>, r"C:\Users\Public\Documents\TestComplete 9 Samples\Hello\Scripts\Hello.pjs") 
#    paint(tc)
#finally:
#    tc.StopTest()

# C# Orders test:
try:
    prjName = "Orders_C#_JScript"
    tc.RunTest(prjName, prjName, r"C:\Users\Public\Documents\TestComplete 9 Samples\Open Applications\C#\TCProject\Orders.pjs") 
    TestOrders(tc)
finally:
    tc.StopTest()
