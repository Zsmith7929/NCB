#!/usr/bin/env python

import csv
import sys
import datetime
import calendar
import shelve
import string
import time
import random
import os
import shutil
import subprocess


try:
    import wx
except ImportError:
    print 
    'Hey, you need WX to run this. http://www.wxpython.org/download.php'

# Local module
import bcompiler
import lookup
import reports

class Nframe(wx.Frame):
    """This is the main application to handle inventory control."""
    
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(600,400))
    
        print '='*20
        print datetime.datetime.now()
        print '='*20
        
        # Frame Bindings
        self.Bind(wx.EVT_CLOSE, self.onClose)
        
        # Application Menu
        # File
        self.fileMenu = wx.Menu()
        self.menuAboot = self.fileMenu.Append(wx.ID_ABOUT,'&About',
                            ' Information about this program')
        self.Bind(wx.EVT_MENU, self.aboot, self.menuAboot)
        self.fileMenu.AppendSeparator()
        self.menuClose = self.fileMenu.Append(wx.ID_EXIT, 'E&xit', 
                                                ' Terminate the program.')
        self.Bind(wx.EVT_MENU, self.menuExit, self.menuClose)
        
        # Menu Bar
        self.menuBar = wx.MenuBar()
        self.menuBar.Append(self.fileMenu, '&File')
        self.SetMenuBar(self.menuBar)
        self.CreateStatusBar()
        
        # Panel
        self.panel = wx.Panel(self, wx.ID_ANY, pos=(0,0))
        
        # Attributes
        self.scanner = True
        
        # Column titles: Barcode | Reports
        f = wx.Font(32, wx.ROMAN, wx.FONTSTYLE_NORMAL, wx.BOLD, False)
        
        self.barcodeText = wx.StaticText(self.panel, wx.ID_ANY, 'Barcodes',
                    size=(200,50), style=wx.ST_NO_AUTORESIZE|wx.ALIGN_CENTRE)
        self.barcodeText.SetFont(f)
        
        self.reportsText = wx.StaticText(self.panel, wx.ID_ANY, 'Reports',
                    size=(200,50), style=wx.ST_NO_AUTORESIZE|wx.ALIGN_CENTRE)
        self.reportsText.SetFont(f)
        
        # Line seperators
        self.lineV = wx.StaticLine(self.panel, wx.ID_ANY, size=(2,600), 
                                        style=wx.LI_VERTICAL)
                                        
        self.lineH1 = wx.StaticLine(self.panel, wx.ID_ANY, size=(200,2))
        self.lineH3 = wx.StaticLine(self.panel, wx.ID_ANY, size=(200,2))
        
        # Buttons
        self.syncButton = wx.Button(self.panel, wx.ID_ANY, 'Sync Database',
                                    size=(200,50))
        self.Bind(wx.EVT_BUTTON, self.syncDatabase, self.syncButton)
        
        self.enterButton = wx.Button(self.panel, wx.ID_ANY, 'Enter Barcodes',
                                    size=(200,50))
        self.Bind(wx.EVT_BUTTON, self.enterBarcodes, self.enterButton)
        
        self.lookupButton = wx.Button(self.panel, wx.ID_ANY, 'Lookup product',
                                        size=(200,50))
        self.Bind(wx.EVT_BUTTON, self.lookup, self.lookupButton)
        
        self.gReportButton = wx.Button(self.panel, wx.ID_ANY, 
                                        'Generate Reports', size=(200,50))
        self.Bind(wx.EVT_BUTTON, self.gReport, self.gReportButton)
        
        self.printReportButton = wx.Button(self.panel, wx.ID_ANY, 
                                        'Print Reports', size=(200,50))
        self.Bind(wx.EVT_BUTTON, self.printReport, self.printReportButton)
        
        self.printOrdersButton = wx.Button(self.panel, wx.ID_ANY,
                                        'Print Order Forms', size=(200,50))
        self.Bind(wx.EVT_BUTTON, self.printOrders, self.printOrdersButton)
        
        # Sizers   
        self.boxH = wx.BoxSizer(wx.HORIZONTAL)
        self.boxV1 = wx.BoxSizer(wx.VERTICAL)
        self.boxV2 = wx.BoxSizer(wx.VERTICAL)
        self.boxV3 = wx.BoxSizer(wx.VERTICAL)
        
        self.boxH.AddStretchSpacer()
        self.boxH.Add(self.boxV1, 0, wx.TOP, 25)
        self.boxH.AddStretchSpacer()
        self.boxH.Add(self.boxV2, 0)
        self.boxH.AddStretchSpacer()
        self.boxH.Add(self.boxV3, 0, wx.TOP, 25)
        self.boxH.AddStretchSpacer()

        # Column One
        self.boxV1.Add(self.barcodeText)
        self.boxV1.Add(self.lineH1, 0, wx.TOP, 5)
        self.boxV1.Add(self.syncButton, 0, wx.TOP, 15)
        self.boxV1.Add(self.enterButton, 0, wx.TOP, 15)
        self.boxV1.Add(self.lookupButton, 0, wx.TOP, 15)
        
        # Column Two
        self.boxV2.Add(self.lineV)
        
        #Column Three
        self.boxV3.Add(self.reportsText)
        self.boxV3.Add(self.lineH3, 0, wx.TOP, 5)
        self.boxV3.Add(self.gReportButton, 0, wx.TOP, 15)
        self.boxV3.Add(self.printReportButton, 0, wx.TOP, 15)
        self.boxV3.Add(self.printOrdersButton, 0, wx.TOP, 15)
        self.panel.SetSizer(self.boxH)
        
        # Build the report databases for accessing
        self.checkScanner()
        
        # Show Frame
        self.Centre()
        self.Show(True)
                
#---Methods---#FFFFFF#000000------------------------------------------
    
    def onClose(self, event):
        """Destroy itself and all other instances of the compiler, etc."""  
              
        self.Destroy()
        return
            
    def aboot(self, event):
        """Menu --> About"""
        
        self.awesome = wx.MessageDialog(
            self, 'By Zacheriah Smith: smith7929@gmail.com', 
            'About', wx.OK|wx.ICON_QUESTION)
                
        self.awesome.ShowModal()
        self.awesome.Destroy()
    
    def menuExit(self, event):
        """Menu --> File --> Exit"""
        
        self.onClose(event)
        return
        
    def checkScanner(self):
        """Checks for the scanner and passes it to buildReports(self)"""
        
        now = datetime.datetime.now()
        mystamp = '%.2d-%.2d-%.4d_%.2d_%.2d' % (now.month, now.day, now.year,
                                            now.hour, now.minute)
        try:
            with open('I:\\Scanned Barcodes\\BARCODES.txt') as f:
                s = ('Database\\Backups\\BARCODES\\BARCODES')
                backup = open(s+mystamp+'.txt','w')
                shutil.copy('I:\\Scanned Barcodes\\BARCODES.txt','%s' % 
                            (s+mystamp+'.txt'))
                backup.close()
                
        except IOError:
            # Make a dialog asking if they want to try again
            ioDiag = wx.MessageDialog(self, 
                                'Couldn\'t detect the scanner.'+
                                ' Would you like to retry?', 'Scanner Error',
                                wx.STAY_ON_TOP|wx.ICON_QUESTION|wx.YES_NO)
            
            result = ioDiag.ShowModal()
            ioDiag.Destroy()
            # If yes...
            if result == 5103:
                scannerDiag = wx.MessageDialog(self, 
                    'Please plug the scanner in now.', 'Confirm',
                    wx.STAY_ON_TOP|wx.OK|wx.ICON_HAND)
                    
                scannerDiag.ShowModal()
                scannerDiag.Destroy()
                return self.checkScanner()
                    
            # If no...    
            if result == 5104:
                noScannerDiag = wx.MessageDialog(self,
                    'Without the scanner plugged in, many features will be'+
                    ' disabled.', 'Warning', 
                    wx.STAY_ON_TOP|wx.OK|wx.ICON_EXCLAMATION)
                    
                noScannerDiag.ShowModal()
                noScannerDiag.Destroy()
                self.scanner = False
        
        finally:
            self.setAvailable()
        
    def setAvailable(self):
        """This will disable widgets associated with the reports database
        in the case that the reports database cannot be updated via scanner"""
        
        if self.scanner == False:
            self.syncButton.Disable()
            return
        
        self.printOrdersButton.Disable()
            
    def buildReports(self):
        """This builds the reports from the databases"""
        pass
                
    def syncDatabase(self, event):
        """This takes the barcodes from BARCODES.txt and puts them into
        databases filed by date scanned"""
            
        # Check if empty/already synced
        try:
            self.file_len('I:\\Scanned Barcodes\\BARCODES.txt')
        except UnboundLocalError:
            error = wx.MessageDialog(self, 
                            'Scanner is empty or has already been synced.',
                            'Error', style=wx.STAY_ON_TOP|wx.OK|wx.ICON_ERROR)
            error.ShowModal()
            error.Destroy()
            
            return
        
        with open('I:\\Scanned Barcodes\\BARCODES.txt') as f:
            unknown = set()
            try:
                m = shelve.open('Database\\master\\master.db')
                reader = csv.reader(f)
                length = self.file_len('I:\\Scanned Barcodes\\BARCODES.txt')
                progress = wx.ProgressDialog('Please wait',
                                    'Syncing Database, Please Wait...', 
                                    length, None, style=wx.PD_REMAINING_TIME|
                                    wx.PD_SMOOTH|wx.PD_AUTO_HIDE|
                                    wx.PD_CAN_ABORT)
                keepGoing = True
                count = 0
                for line in reader:
                    count += 1
                    date = line[0].replace('/','-')
                    if m.has_key(line[3]):
                        n = shelve.open('Database\\Dailies\\Daily'+date+'.db',
                                        writeback=True)
                        if n.has_key(line[3]):
                            n[line[3]] += 1
                        
                        else:
                            n.setdefault(line[3], 1)
                        n.close()
                        
                    else:
                        unknown.add(line[3])
                        o = shelve.open('Database\\Unknown\\unknown'
                                        +date+'.db', writeback=True)
                        
                        if o.has_key(line[3]):
                            o[line[3]] += 1
                        
                        else:
                            o.setdefault(line[3], {1, date})
                        o.close()
                
                    (keepGoing, skip) = progress.Update(count)
                    if not keepGoing:
                        progress.Destroy()
                
                progress.Destroy()    
                m.close()
            
            except:
                print 'Unexpected error:', sys.exc_info()[0]
                raise
            
            if len(unknown) >= 1:
                s = 'During sync, the compiler found unknown barcodes'
                ' Would you like to manually enter them into the database?'
                            
                q = wx.MessageDialog(self, s, 'Notice', 
                                    wx.STAY_ON_TOP|wx.ICON_QUESTION|wx.YES_NO)
                
                result = q.ShowModal()
                
                if result == 5103:
                    compiler = bcompiler.Compiler(None, 'Barcode Entry')
                    compiler.iterationStation()
                    
                else:
                    pass
        
        with open('I:\\Scanned Barcodes\\BARCODES.txt','w') as q:
            q.write('')
        
        return
        
    def enterBarcodes(self,event):
        """Instantiates a barcode compiler from bcompiler.py"""
        
        compiler = bcompiler.Compiler(self, 'Barcode Entry')
        compiler.Centre()
        
    def lookup(self, event):
        """Instatiates the lookup app from lookup.py"""
        
        look = lookup.Lookup(None, 'Query Database')
        look.Centre()
        
    def gReport(self, event):
        """This is a proxy for reports.createReports()"""
        
        reports.createReports()
        
        
        
    def printReport( self, event):
        """This function prints generated reports"""
        
        choiceListMonths = []
        choiceListDays = []
        choiceListWeeks = ["Last Week"]
        
        for root, dir, files in os.walk('Database\\Reports'):
            for file in files:
                if self.monthName(file[:2]) not in choiceListMonths:
                    choiceListMonths.append(self.monthName(file[:2]))
                choiceListDays.append(file)
                
        lastWeek = self.getLastWeek()
        
                
        choiceListMaster = choiceListMonths+choiceListWeeks+choiceListDays
                
        
        choice = wx.MultiChoiceDialog(None, 
                                        'What days would you like to print?', 
                                        'Print Reports', choiceListMaster)
        choice.Centre()
        choice.ShowModal()
        result = choice.GetSelections()
        
        
        for selection in result:
            
            # Get all daily reports for month if True
            if choiceListMaster[selection] in choiceListMonths:
                for root, dir, files in os.walk('Database\\Reports'):
                    for file in files:
                        if self.monthName(file[:2]) == \
                        choiceListMaster[selection]:
                            subprocess.call('"C:\\Program Files (x86)\\'
                                            'OpenOffice.org 3\\program\\'
                                            'soffice.exe" -p Database\\'
                                            'Reports\\'+file)
                                            
            elif str(choiceListMaster[selection]) in choiceListWeeks:
                for root, dir, files in os.walk('Database\\Reports'):
                    for file in files:
                        if datetime.date(
                            int(file[6:10].lstrip('0')),
                            int(file[:2].lstrip('0')),
                            int(file[3:5].lstrip('0'))) in lastWeek:
                            
                            subprocess.call('"C:\\Program Files (x86)\\'
                                            'OpenOffice.org 3\\program\\'
                                            'soffice.exe" -p Database\\'
                                            'Reports\\'+file)
                            
            else:
                subprocess.call('"C:\\Program Files (x86)\\'
                                            'OpenOffice.org 3\\program\\'
                                            'soffice.exe" -p Database\\'
                                            'Reports\\'+
                                            choiceListMaster[selection])
  
        choice.Destroy()
        
                                                    
    def printOrders(self, event):
        """This will generate and print order forms"""
        
        pass
        
    
    def getLastWeek(self):
        
        _now = datetime.date.today()
        lastweek = []
    
        for i in range(20):
            if (_now - datetime.timedelta(days=i)).isocalendar()[1] == \
            (_now - datetime.timedelta(weeks=1)).isocalendar()[1]:
                lastweek.append((_now - datetime.timedelta(days=i)))
        print "This is last week:", lastweek
        return lastweek
    
    
        
 #---Misc Methods---#000000#FFFFFF-----------------------------------------------       
    def file_len(self, fname):
        with open(fname) as f:
            for i, l in enumerate(f):
                pass
        return i + 1
    
    def monthName(self, num):
        """ Simple month number --> Name conversion """
        
        monthNames = {
            '1': 'January',
            '2': 'February',
            '3': 'March',
            '4': 'April',
            '5': 'May',
            '6': 'June',
            '7': 'July',
            '8': 'August',
            '9': 'September',
            '10': 'October',
            '11': 'November',
            '12': 'December'}
            
        return monthNames[num.lstrip('0')]
        
#---SCRIPT---#FFFFFF#000000---------------------------------------------

if __name__ == "__main__":
    app = wx.App(False)
    frame = Nframe(None, 'Nature\'s Corner')
    app.MainLoop()