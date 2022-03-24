#!/usr/bin/env python3


from tkinter import Tk, Label, Button, StringVar, Entry,NONE, END,HORIZONTAL,N, W, E, S, Checkbutton,Radiobutton, IntVar, Radiobutton, Scrollbar, Listbox, LEFT, BOTH, Spinbox, Menu, Text, NORMAL
import tkinter as tk
from tkinter import ttk

import math
# to create a dialog interface for the input/output file
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile

# to manage a different font
from tkinter import font

# need this to check if a a file exists
import os
import sys

# to read the XLS file
import xlrd

# to create an XLS file
import xlsxwriter
# to rewrite an xls file without delete it
from xlutils.copy import copy

# to manage date
import datetime
import time

# to manage SEA FLOOR format
import csv

# to manage images
from PIL import Image, ImageTk

# to manage plots
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

import numpy as np

# to map the surveys on GoogleMaps using the browser
import webbrowser

# for the web scraping on BODC VOCABS
from bs4 import BeautifulSoup
import requests


'''
To create an exe (Linux/Windows):

https://pypi.org/project/auto-py-to-exe/
pip install auto-py-to-exe
auto-py-to-exe


To include the NODC logo use the following option inside auto-py-to-exe:
--hidden-import='PIL._tkinter_finder'


FOR WINDOWS ONLY with ANACONDA:
--exclude-module scikit-learn,PyQt5,PyQt4,2to3,IPython,Jinja2,pycparser,scipy

'''

class MarineLitterManager:


    NUMBERS_ARRAY = []

    for n in range(150):
        NUMBERS_ARRAY.append(n)



    TMP_LETTERS_ARRAY = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "V",
        "W",
        "X",
        "Y",
        "Z",
    ]

    LETTERS_ARRAY = TMP_LETTERS_ARRAY

    howmanynumbers = len(NUMBERS_ARRAY)-1
    if howmanynumbers > 25:
        letters_cicles=int(float(howmanynumbers)/25)
        for c in range(letters_cicles):
            for n in range(26):
                LETTERS_ARRAY.append(TMP_LETTERS_ARRAY[c]+TMP_LETTERS_ARRAY[n])
    LETTERS_ARRAY.insert(0, " ")




    FIELDBEACHES=['BeachCode',
    'BeachName',
    'Country',
    'BeachInfoAmmendment',
    'FillingDate',
    'FillingName',
    'FillingPhone',
    'FillingMail',
    'FillingInstitute',
    'UrbanizationDegree',
    'Reference beach',
    'BeachWidthLow',
    'BeachWidthHigh',
    'BeachLength',
    'BeachLatitude',
    'BeachLongitude',
    'CoordinateSystem',
    'BeachBack',
    'BeachBackOther',
    'BeachBackDevelopment',
    'DevelopmentDescription',
    'PositionMeasurementDate',
    'CurrentsDirection',
    'WindsDirection',
    'BeachOrientation',
    'BeachMaterial',
    'BeachTopography',
    'Obstacles',
    'Usage1',
    'Usage1Seasonality',
    'Usage2',
    'Usage2Seasonality',
    'Usage3',
    'Usage3Seasonality',
    'BeachAccess',
    'BeachCleaningSeasonality',
    'SeasonalityMonths',
    'CleaningFrequency',
    'OtherDescription',
    'CleaningMethod',
    'CleaningResponsible',
    'Notes',]

    FIELDSURVEYS=['BeachCode',
    'SurveyCode',
    'SurveyType',
    'DataPolicy',
    'SurveyDate',
    'Originator',
    'Collator',
    'ProjectCode',
    'SurveyStartLatitude',
    'SurveyStartLongitude',
    'SurveyEndLatitude',
    'SurveyEndLongitude',
    'CoordinateSystem',
    'SurveyLength',
    'SurveyWidth',
    'Surveyor1Name',
    'Surveyor1Phone',
    'Surveyor1Mail',
    'Surveyor2Name',
    'Surveyor2Phone',
    'Surveyor2Mail',
    'TownName',
    'TownDistance',
    'TownPosition',
    'TownPopulation',
    'WinterTourists',
    'SpringTourists',
    'SummerTourists',
    'AutumnTourists',
    'FoodOutlets',
    'FoodOutletsDistance',
    'FoodOutletsSeasonality',
    'SeasonalityMonths',
    'FoodOutletsPosition',
    'ShippingLaneDistance',
    'ShippingLaneTraffic',
    'ShippingLaneTypes',
    'ShippingLanePosition',
    'HarbourName',
    'HarbourDistance',
    'HarbourPosition',
    'HarbourType',
    'HarbourSize',
    'RiverName',
    'RiverDistance',
    'RiverPosition',
    'WasteWaterDischarges',
    'WasteWaterDistance',
    'WasteWaterPosition',
    'LitterPresence',
    'LastCleaning',
    'WeatherConditions',
    'WeatherConditionsOther',
    'AnimalsFound',
    'AnimalsNumber',
    'Circumstances',
    'Events',
    'Notes',]



    FIELDANIMALS=['SurveyCode',
    'Animal',
    'State',
    'Sex',
    'Age',
    'Entanglement',
    'EntanglementNature',
    'Comments',]


    FIELDLITTER=['SurveyCode',
    'LitterReferenceList',
    'ItemCode',
    'ItemName',
    'ParameterOriginalName',
    'NoItems',
    'Notes',]



    FIELDSURVEYSSEAFLOOR=['SurveyName',
    'ProjectCode',
    'DataPolicy',
    'Date',
    'Ship',
    'Gear',
    'Country',
    'Originator',
    'Collator',
    'StNo',
    'HaulNo',
    'CoordRefSys',
    'ShootLat',
    'ShootLong',
    'HaulLat',
    'HaulLong',
    'Depth',
    'Distance',
    'GroundSpeed',
    'WingSpread',
    'DoorSpread',
    'WarpLength',
    'Shot_timestamp',
    'HaulDur']
    


    FIELDLITTERSEAFLOOR=['LTREF',
    'PARAM',
    'LTSZC',
    'LTSRC',
    'TYPPL',
    'LTPRP',
    'UnitWgt',
    'LT_Weight',
    'UnitItem',
    'LT_Items',
    'Shot_timestamp',
    'HaulDur',
    'StationNumber',]   # The last field has been added only to have a match with the surveys when we create the CSV output file




    checkmylistScrollListSurveyParamsSF=0


    def __init__(self, master):
        self.master = master
        master.title("Marine Litter Manager")

        '''
        THE FOLLOWING THREE ROWS EXPLAIN HOW TO CHANGE THE STATE OF A TAB
        DISABLED: the tab is visible but not active
        NORMAL: the tab is in nomral state
        HIDDEN: the tab is not visible
        these options could be useful if the software must manage different
        input/output formats and it's necessary show only some tabs
        N.B. in the example the tab is 2 (self.surveyBL)
        '''
        #self.nb.tab(2, state="disabled")
        #self.nb.tab(2, state="normal")
        #self.nb.tab(2, state="hidden")

        
        
        frameFont = ttk.Style()
        frameFont.configure('new.TFrame', family='Verdana', size=8, weight='bold', underline=1)
        
        # Defines and places the notebook widget
        self.nb = ttk.Notebook(self.master)
        self.nb.grid(row=1, column=0, columnspan=50, rowspan=49, sticky='NESW')

        # Adds tab of the notebook
        self.formats = ttk.Frame(self.nb, style='new.TFrame')
        self.nb.add(self.formats, text='FORMATS')

        # Adds tab of the notebook
        self.infoBL = ttk.Frame(self.nb, style='new.TFrame')
        self.nb.add(self.infoBL, text='BEACH LITTER')

        # Adds tab of the notebook
        self.beachesBL = ttk.Frame(self.nb)
        self.nb.add(self.beachesBL, text='BEACHES')
         
        # Adds tab of the notebook
        self.surveyBL = ttk.Frame(self.nb)
        self.nb.add(self.surveyBL, text='SURVEYS')

        # Adds tab of the notebook
        self.animalLitterBL = ttk.Frame(self.nb)
        self.nb.add(self.animalLitterBL, text='ANIMALS & LITTER')

        # Adds tab of the notebook
        self.plotBL = ttk.Frame(self.nb)
        self.nb.add(self.plotBL, text='SURVEYS PLOT')

        # Adds tab of the notebook
        self.scatterBL = ttk.Frame(self.nb)
        self.nb.add(self.scatterBL, text='PARAMS PLOT')

        # Adds tab of the notebook
        self.infoSF = ttk.Frame(self.nb)
        self.nb.add(self.infoSF, text='SEA FLOOR')

       # Adds tab of the notebook
        self.surveySF = ttk.Frame(self.nb)
        self.nb.add(self.surveySF, text='SURVEYS')

       # Adds tab of the notebook
        self.litterSF = ttk.Frame(self.nb)
        self.nb.add(self.litterSF, text='LITTER')

       # Adds tab of the notebook
        self.plotSF = ttk.Frame(self.nb)
        self.nb.add(self.plotSF, text='SURVEYS PLOT')

       # Adds tab of the notebook
        self.scatterSF = ttk.Frame(self.nb)
        self.nb.add(self.scatterSF, text='PARAMS PLOT')

        # Adds tab of the notebook
        self.infoCML = ttk.Frame(self.nb)
        self.nb.add(self.infoCML, text='COASTAL MACRO LITTER')
        
        # Adds tab of the notebook
        self.infoOSML = ttk.Frame(self.nb)
        self.nb.add(self.infoOSML, text='OPEN SEA MACRO LITTER')

        # Adds tab of the notebook
        self.dictionary = ttk.Frame(self.nb)
        self.nb.add(self.dictionary, text='DICTIONARY')

        # Adds tab of the notebook
        self.links = ttk.Frame(self.nb)
        self.nb.add(self.links, text='LINKS')
        


        self.nb.tab(1, state="hidden")
        self.nb.tab(2, state="hidden")
        self.nb.tab(3, state="hidden")
        self.nb.tab(4, state="hidden")
        self.nb.tab(5, state="hidden")
        self.nb.tab(6, state="hidden")
        self.nb.tab(7, state="hidden")
        self.nb.tab(8, state="hidden")
        self.nb.tab(9, state="hidden")
        self.nb.tab(10, state="hidden")
        self.nb.tab(11, state="hidden")
        self.nb.tab(12, state="hidden")
        self.nb.tab(13, state="hidden")
        self.nb.tab(14, state="hidden")
        self.nb.tab(15, state="hidden")




        #Common parts we have to wrap each command
        xlsGrid = master.register(self.checkGridXls)
        xlsGridSurvey = master.register(self.checkGridXlsSurvey)
        xlsGridSurveySF = master.register(self.checkGridXlsSurveySF)
        xlsGridAnimals = master.register(self.checkGridXlsAnimals)
        xlsGridLitter = master.register(self.checkGridXlsLitter)
        xlsGridLitterSF = master.register(self.checkGridXlsLitterSF)
        createOutput = master.register(self.createXlsOutput)
        createOutputModel = master.register(self.createXlsOutputModel)
        createOutputModelSF = master.register(self.SaveOutputFileModelSF)
        loadInputModel = master.register(self.loadModel)
        loadInputModelSF = master.register(self.loadModelSF)
        vcmd = master.register(self.validatenumber)
        openfile = master.register(self.OpenInputFile)
        openfileSF = master.register(self.OpenInputFileSF)
        openfilesurveyplot = master.register(self.OpenInputFilePlotSurvey)
        openfilesurveyplotSF = master.register(self.OpenInputFilePlotSurveySF)
        openfileparamsplot = master.register(self.OpenInputFilePlotParams)
        openfileparamsplotSF = master.register(self.OpenInputFilePlotParamsSF)
        savefilexls = master.register(self.SaveOutputFileXls)
        savefilecsvSF = master.register(self.SaveOutputFileCsvSF)
        savefilemodel = master.register(self.SaveOutputFileModel)
        savefilemodelSF = master.register(self.SaveOutputFileModelSF)
        openmodelfile = master.register(self.OpenModelInputFile)
        openmodelfileSF = master.register(self.OpenModelInputFileSF)
        searchLegendafile = master.register(self.SearchLegendaTermFile)
        checkForPlots = master.register(self.checkPlots)
        checkForPlotsSF = master.register(self.checkPlotsSF)
        plotMySurvey = master.register(self.executePlot)
        plotMySurveySF = master.register(self.executePlotSF)
        checkForPlotsParams = master.register(self.checkPlotsParams)
        checkForPlotsParamsSF = master.register(self.checkPlotsParamsSF)
        plotMyParams = master.register(self.executePlotParams)
        plotMyParamsSF = master.register(self.executePlotParamsSF)
        find_resource_path = master.register(self.resource_path)
        showbeachlitter = master.register(self.ShowBeachLitterTabs)
        executeLinkButtonA = master.register(self.LinkButtonA)
        executeLinkButtonB = master.register(self.LinkButtonB)
        executeLinkButtonC = master.register(self.LinkButtonC)
        executeLinkButtonD = master.register(self.LinkButtonD)
        executeLinkButtonE = master.register(self.LinkButtonE)
        executeLinkButtonF = master.register(self.LinkButtonF)
        showseafloorlitter = master.register(self.ShowSeaFloorLitterTabs)
        showcoastalmacrolitter = master.register(self.ShowCoastalMacroLitterTabs)
        showopenseamacrolitter = master.register(self.ShowOpenSeaMacroLitterTabs)
        showutilities = master.register(self.ShowUtilitiesTabs)
        formathidealltabs = master.register(self.HideAllTabs)
        changeLabelSep = master.register(self.changeLabelSeparator)


        mytext=''
        self.create=0



        '''
        START here we add a cascading menu
        '''
        self.emptymenu = Menu(self.master)
        self.menuBL = Menu(self.master)
        self.menuSF = Menu(self.master)

        # Items for Beach Litter MENU
        new_itemFile = Menu(self.menuBL)
        new_itemModel = Menu(self.menuBL)
        new_itemFile.add_command(label='Load Litter Input File', command=(openfile))
        new_itemFile.add_separator()
        new_itemFile.add_command(label='Save Litter Output File', command=(savefilexls))
        new_itemFile.add_separator()
        new_itemModel.add_command(label='Load Model', command=(openmodelfile))
        new_itemModel.add_separator()
        new_itemModel.add_command(label='Save Model', command=(savefilemodel))
        new_itemModel.add_separator()
        self.menuBL.add_cascade(label='Beach Litter Files', menu=new_itemFile)
        self.menuBL.add_cascade(label='Beach Litter Models', menu=new_itemModel)

        # Items for Beach Litter MENU
        new_itemFileSF = Menu(self.menuSF)
        new_itemModelSF = Menu(self.menuSF)
        new_itemFileSF.add_command(label='Load Litter Input File', command=(openfileSF))
        new_itemFileSF.add_separator()
        new_itemFileSF.add_command(label='Save Litter Output File', command=(savefilecsvSF))
        new_itemFileSF.add_separator()
        new_itemModelSF.add_command(label='Load Model', command=(openmodelfileSF))
        new_itemModelSF.add_separator()
        new_itemModelSF.add_command(label='Save Model', command=(savefilemodelSF))
        new_itemModelSF.add_separator()
        self.menuSF.add_cascade(label='Sea Floor Files', menu=new_itemFileSF)
        self.menuSF.add_cascade(label='Sea Floor Models', menu=new_itemModelSF)

        self.master.config(menu=self.emptymenu)

        '''
        END cascading menu
        '''
        # END cascading menu




        '''
        START FORMATS
        '''
        appHighlightFont = font.Font(family='helvetica', size=12, weight='bold', underline=1)
        
        #This is only a reminder abut the available fonts
        #
        #        print(font.families())
        #        AVAILABLE FONTS on TKINTER
        #        ('fangsong ti', 
        #         'fixed', 
        #         'clearlyu alternate glyphs', 
        #         'charter', 
        #         'lucidatypewriter', 
        #         'courier 10 pitch', 
        #         'lucidabright', 
        #         'times', 
        #         'open look glyph', 
        #         'bitstream charter', 
        #         'song ti', 'helvetica', 
        #         'open look cursor', 
        #         'newspaper', 
        #         'clearlyu ligature', 
        #         'mincho', 
        #         'clearlyu devangari extra', 
        #         'clearlyu pua', 
        #         'courier', 
        #         'clearlyu', 
        #         'lucida', 
        #         'clean', 
        #         'nil', 
        #         'clearlyu arabic', 
        #         'clearlyu devanagari', 
        #         'terminal', 
        #         'symbol', 
        #         'gothic', 
        #         'new century schoolbook', 
        #         'clearlyu arabic extra')

        self.labelOGCNODC = Label(self.formats, text=mytext, bg="SkyBlue2", fg="black", font=appHighlightFont, height=3, width=76)
        self.labelOGCNODC['text'] = 'MARINE LITTER MANAGER developed by NODC\nNational Oceanographic Data Center - OGS https://nodc.ogs.trieste.it'
        self.labelOGCNODC.grid(row=0, column=0, columnspan=10, rowspan=10, padx=25, pady=55)

        self.FormatBeachLitterButton = Button(self.formats, text="BEACH LITTER FORMAT", width=103, font=('helvetica','9','bold'),background = 'white', command=(showbeachlitter))
        self.FormatBeachLitterButton.grid(row=11, column=0, sticky=W)

        self.FormatBeachLitterButtonSF = Button(self.formats, text="SEA FLOOR LITTER FORMAT", width=103, font=('helvetica','9','bold'),background = 'white', command=(showseafloorlitter))
        self.FormatBeachLitterButtonSF.grid(row=12, column=0, sticky=W)

        self.UtilitiesButton = Button(self.formats, text="UTILITIES", width=103, font=('helvetica','9','bold'),background = 'white', command=(showutilities))
        self.UtilitiesButton.grid(row=15, column=0, sticky=W)

        self.HideAllButton = Button(self.formats, text="HIDE ALL", width=103, font=('helvetica','9','bold'),background = 'white', command=(formathidealltabs))
        self.HideAllButton.grid(row=16, column=0, sticky=W)


        self.path = self.resource_path('logo.png')
        #Creates a Tkinter-compatible photo image, which can be used everywhere Tkinter expects an image object.
        self.img = ImageTk.PhotoImage(Image.open(self.path))
        #The Label widget is a standard Tkinter widget used to display a text or image on the screen.
        self.panel = ttk.Label(self.formats, image = self.img)
        self.panel.grid(row=26, column=0, columnspan=4, rowspan=10, padx=25, pady=55)



        '''
        END FORMATS
        '''



        '''
        START LINKS
        '''
        self.LinkButtonB = Button(self.links, text="Guidelines and forms for gathering marine litter data (PDF file)", width=105, font=('helvetica','9','bold'),background = 'white', command=(executeLinkButtonB))
        self.LinkButtonB.grid(row=5, column=0, sticky=W)

        self.LinkButtonA = Button(self.links, text="Beach format template (ZIP file)", width=105, font=('helvetica','9','bold'),background = 'white', command=(executeLinkButtonA))
        self.LinkButtonA.grid(row=6, column=0, sticky=W)

        self.LinkButtonC = Button(self.links, text="Seafloor format template (ZIP file)", width=105, font=('helvetica','9','bold'),background = 'white', command=(executeLinkButtonC))
        self.LinkButtonC.grid(row=7, column=0, sticky=W)


        self.LinkButtonD = Button(self.links, text="Beach, seafloor data available through EMODnet Chemistry Data Discovery and Access Service (web page)", width=105, font=('helvetica','9','bold'),background = 'white', command=(executeLinkButtonD))
        self.LinkButtonD.grid(row=8, column=0, sticky=W)


        self.LinkButtonE = Button(self.links, text="Marine Litter Visualization Products (web page)", width=105, font=('helvetica','9','bold'),background = 'white', command=(executeLinkButtonE))
        self.LinkButtonE.grid(row=9, column=0, sticky=W)


        self.LinkButtonF = Button(self.links, text="Aggregated collections of unrestricted data for beach and seafloor litter: Sextant Catalogue Service (web page)", width=105, font=('helvetica','9','bold'),background = 'white', command=(executeLinkButtonF))
        self.LinkButtonF.grid(row=10, column=0, sticky=W)
        '''
        END LINKS
        '''


        '''
        START INFO BEACH LITTER
        '''
                
        self.labelEntryInfoBeaches = Label(self.infoBL, text=mytext)
        self.labelEntryInfoBeaches['text'] = 'The sheet for BEACHES: ' + str(mytext)
        self.labelEntryInfoBeaches.grid(row=1, column=0, sticky=E)

        self.entryInfoBeachesVars = tk.IntVar()
        self.entryInfoBeaches = Spinbox(self.infoBL, from_=1, to=10, textvariable= self.entryInfoBeachesVars, width=2, validate="key", validatecommand=(vcmd, '%P'))
        self.entryInfoBeaches.grid(row=1, column=1, sticky=W)

        self.labelEntryInfoSurveys = Label(self.infoBL, text=mytext)
        self.labelEntryInfoSurveys['text'] = 'The sheet for SURVEYS: ' + str(mytext)
        self.labelEntryInfoSurveys.grid(row=1, column=2, sticky=E)

        self.entryInfoSurveysVars = tk.IntVar()
        self.entryInfoSurveys = Spinbox(self.infoBL, from_=1, to=10, textvariable= self.entryInfoSurveysVars, width=2, validate="key", validatecommand=(vcmd, '%P'))
        self.entryInfoSurveys.grid(row=1, column=3, sticky=W)

        self.labelEntryInfoAnimals = Label(self.infoBL, text=mytext)
        self.labelEntryInfoAnimals['text'] = 'The sheet for ANIMALS: ' + str(mytext)
        self.labelEntryInfoAnimals.grid(row=2, column=0, sticky=E)

        self.entryInfoAnimalsVars = tk.IntVar()
        self.entryInfoAnimals = Spinbox(self.infoBL, from_=1, to=10, textvariable= self.entryInfoAnimalsVars, width=2, validate="key", validatecommand=(vcmd, '%P'))
        self.entryInfoAnimals.grid(row=2, column=1, sticky=W)

        self.labelEntryInfoLitter = Label(self.infoBL, text=mytext)
        self.labelEntryInfoLitter['text'] = 'The sheet for LITTER: ' + str(mytext)
        self.labelEntryInfoLitter.grid(row=2, column=2, sticky=E)

        self.entryInfoLitterVars = tk.IntVar()
        self.entryInfoLitter = Spinbox(self.infoBL, from_=1, to=10, textvariable= self.entryInfoLitterVars, width=2, validate="key", validatecommand=(vcmd, '%P'))
        self.entryInfoLitter.grid(row=2, column=3, sticky=W)

        self.openInputFileButton = Button(self.infoBL, text="Load Litter Input File", command=(openfile))
        self.openInputFileButton.grid(row=3, column=2, sticky=E)

        self.entryInfoInputFile = Entry(self.infoBL, width=20, validate="key")
        self.entryInfoInputFile.grid(row=3, column=3, sticky=W)

        self.labelInfoOutputFile = Label(self.infoBL, text=mytext)
        self.labelInfoOutputFile['text'] = 'Output file name (.xls): ' + str(mytext)
        self.labelInfoOutputFile.grid(row=4, column=2, sticky=E)

        self.entryInfoOutputFile = Entry(self.infoBL, width=20,state="readonly", validate="key")
        self.entryInfoOutputFile.grid(row=4, column=3, sticky=W)

        self.openInputModelFileButton = Button(self.infoBL, text="Load Model", command=(openmodelfile))
        self.openInputModelFileButton.grid(row=5, column=2, sticky=E)

        self.entryInfoModelInputFile = Entry(self.infoBL, width=20, validate="key")
        self.entryInfoModelInputFile.grid(row=5, column=3, sticky=W)

        self.labelInfoOutputModelFile = Label(self.infoBL, text=mytext)
        self.labelInfoOutputModelFile['text'] = 'Model file name (.csv): ' + str(mytext)
        self.labelInfoOutputModelFile.grid(row=6, column=2, sticky=E)

        self.entryInfoOutputModelFile = Entry(self.infoBL, width=20,state="readonly", validate="key")
        self.entryInfoOutputModelFile.grid(row=6, column=3, sticky=W)

        self.infoBLarea = Text(self.infoBL, height=45, width=106)
        self.infoBLarea.grid(row=59, column=0, columnspan=18, rowspan=30)
        self.infoBLarea.insert(END, "Marine Litter Manager Infobox:")

        self.labellegenda = Label(self.dictionary, text=mytext)
        self.labellegenda['text'] = 'Search term or param: ' + str(mytext)
        self.labellegenda.grid(row=1, column=0, sticky=E)

        self.entrylegendaTerm = Entry(self.dictionary, width=20, validate="key")
        self.entrylegendaTerm.grid(row=1, column=1, sticky=W)

        self.openlegendaButton = Button(self.dictionary, text="SEARCH", command=(searchLegendafile))
        self.openlegendaButton.grid(row=1, column=2, sticky=E)

        self.varVocabBODClvLegenda = IntVar(value=1)
        self.buttonvarVocabBODClvLegenda = Checkbutton(self.dictionary, text="EMBEDDED DICTIONARY SEARCH", variable=self.varVocabBODClvLegenda)
        self.buttonvarVocabBODClvLegenda.grid(row=2, column=0,columnspan=10, sticky=W)

        self.varVocabBODClvA = IntVar()
        self.buttonvarVocabBODClvA = Checkbutton(self.dictionary, text="H01 BODC VOCAB - EMODnet micro-litter types (WEB SCRAPING)", variable=self.varVocabBODClvA)
        self.buttonvarVocabBODClvA.grid(row=3, column=0,columnspan=10, sticky=W)

        self.varVocabBODClvB = IntVar()
        self.buttonvarVocabBODClvB = Checkbutton(self.dictionary, text="H02 BODC VOCAB - EMODnet micro-litter shapes (WEB SCRAPING)", variable=self.varVocabBODClvB)
        self.buttonvarVocabBODClvB.grid(row=4, column=0,columnspan=10, sticky=W)

        self.varVocabBODClvC = IntVar()
        self.buttonvarVocabBODClvC = Checkbutton(self.dictionary, text="H03 BODC VOCAB - EMODnet micro-litter size classes (WEB SCRAPING)", variable=self.varVocabBODClvC)
        self.buttonvarVocabBODClvC.grid(row=5, column=0,columnspan=10, sticky=W)

        self.varVocabBODClvD = IntVar()
        self.buttonvarVocabBODClvD = Checkbutton(self.dictionary, text="H04 BODC VOCAB - EMODnet micro-litter colour classes (WEB SCRAPING)", variable=self.varVocabBODClvD)
        self.buttonvarVocabBODClvD.grid(row=6, column=0,columnspan=10, sticky=W)

        self.varVocabBODClvE = IntVar()
        self.buttonvarVocabBODClvE = Checkbutton(self.dictionary, text="H05 BODC VOCAB - EMODnet micro-litter polymer type (WEB SCRAPING)", variable=self.varVocabBODClvE)
        self.buttonvarVocabBODClvE.grid(row=7, column=0,columnspan=10, sticky=W)

        self.varVocabBODClvF = IntVar()
        self.buttonvarVocabBODClvF = Checkbutton(self.dictionary, text="P01 BODC VOCAB - BODC Parameter Usage Vocabulary (WEB SCRAPING). ATTENTION: TIME-CONSUMING SEARCH!", variable=self.varVocabBODClvF)
        self.buttonvarVocabBODClvF.grid(row=8, column=0,columnspan=10, sticky=W)

        self.legendaarea = Text(self.dictionary, height=45, width=106)
        self.legendaarea.grid(row=10, column=0, columnspan=18, rowspan=30, sticky=W)
        self.legendaarea.insert(END, "")

        root.update()

        '''
        END INFO BEACH LITTER
        '''


        '''
        START INFO SEA FLOOR
        '''

        self.labelEntryInfoSurveysSF = Label(self.infoSF, text=mytext)
        self.labelEntryInfoSurveysSF['text'] = 'The sheet for SURVEYS: ' + str(mytext)
        self.labelEntryInfoSurveysSF.grid(row=1, column=0, sticky=E)

        self.entryInfoSurveysVarsSF = tk.IntVar()
        self.entryInfoSurveysSF = Spinbox(self.infoSF, from_=1, to=10, textvariable= self.entryInfoSurveysVarsSF, width=2, validate="key", validatecommand=(vcmd, '%P'))
        self.entryInfoSurveysSF.grid(row=1, column=1, sticky=W)

        self.labelEntryInfoLitterSF = Label(self.infoSF, text=mytext)
        self.labelEntryInfoLitterSF['text'] = 'The sheet for LITTER: ' + str(mytext)
        self.labelEntryInfoLitterSF.grid(row=1, column=2, sticky=E)

        self.entryInfoLitterVarsSF = tk.IntVar()
        self.entryInfoLitterSF = Spinbox(self.infoSF, from_=1, to=10, textvariable= self.entryInfoLitterVarsSF, width=2, validate="key", validatecommand=(vcmd, '%P'))
        self.entryInfoLitterSF.grid(row=1, column=3, sticky=W)

        self.openInputFileButtonSF = Button(self.infoSF, text="Load Litter Input File", command=(openfileSF))
        self.openInputFileButtonSF.grid(row=3, column=0, sticky=E)

        self.entryInfoInputFileSF = Entry(self.infoSF, width=20, validate="key")
        self.entryInfoInputFileSF.grid(row=3, column=1, columnspan=3, sticky=W)

        self.labelInfoOutputFileSF = Label(self.infoSF, text=mytext)
        self.labelInfoOutputFileSF['text'] = 'Output file name (.xls): ' + str(mytext)
        self.labelInfoOutputFileSF.grid(row=4, column=0, sticky=E)

        self.entryInfoOutputFileSF = Entry(self.infoSF, width=20,state="readonly", validate="key")
        self.entryInfoOutputFileSF.grid(row=4, column=1, columnspan=3, sticky=W)

        self.openInputModelFileButtonSF = Button(self.infoSF, text="Load Model", command=(openmodelfileSF))
        self.openInputModelFileButtonSF.grid(row=5, column=0, sticky=E)

        self.entryInfoModelInputFileSF = Entry(self.infoSF, width=20, validate="key")
        self.entryInfoModelInputFileSF.grid(row=5, column=1, columnspan=3, sticky=W)

        self.labelInfoOutputModelFileSF = Label(self.infoSF, text=mytext)
        self.labelInfoOutputModelFileSF['text'] = 'Model file name (.csv): ' + str(mytext)
        self.labelInfoOutputModelFileSF.grid(row=6, column=0, sticky=E)

        self.entryInfoOutputModelFileSF = Entry(self.infoSF, width=20,state="readonly", validate="key")
        self.entryInfoOutputModelFileSF.grid(row=6, column=1, columnspan=3, sticky=W)

        self.labelRadioSeparatorCSV = Label(self.infoSF, text=mytext)
        self.labelRadioSeparatorCSV['text'] = 'Define the CSV output/plot separator' + str(mytext)
        self.labelRadioSeparatorCSV.grid(row=6, column=2, sticky=E)

        self.varRadioOutputCSV = tk.IntVar()
        self.varRadioOutputCSV.set(1)
        self.R1Output = Radiobutton(self.infoSF, text="Tab", variable=self.varRadioOutputCSV, value=1, command=(changeLabelSep))
        self.R1Output.grid(row=6, column=3, sticky=E)

        self.R2Output = Radiobutton(self.infoSF, text="Comma", variable=self.varRadioOutputCSV, value=2, command=(changeLabelSep))
        self.R2Output.grid(row=6, column=4, sticky=E)

        self.infoSFarea = Text(self.infoSF,wrap=NONE, height=45, width=106)
        self.infoSFarea.grid(row=59, column=0, columnspan=16, rowspan=30)
        self.infoSFarea.insert(END, "Marine Litter Manager Infobox:")

        '''
        END INFO SEA FLOOR
        '''


        '''
        START BEACHES BEACH LITTER
        '''
        self.OnlyLabel = Label(self.beachesBL, text=mytext)
        self.OnlyLabel['text'] = 'Name of the field'
        self.OnlyLabel.grid(row=44, column=0, columnspan=8, sticky=W+E)
        root.update()

        self.labelsBeaches = []
        self.entriesBeachesRow = []
        self.entriesBeachesCol = []
        self.buttonsBeaches = []
        tmprow=0
        tmpcolumn=4
        self.entriesBeachesRowVars = []
        self.entriesBeachesColVars = []
        for i in range(42):

            mytext=self.FIELDBEACHES[i]
            if tmpcolumn == 4:
                tmpcolumn=0
            else:
                tmpcolumn=4

            if tmprow == 0:
                tmprow=1
            else:
                tmprow=0

            self.lbBeaches = Label(self.beachesBL, text=mytext)
            self.lbBeaches['text'] = str(mytext)+'('+str(i)+'): '
            self.lbBeaches.grid(row=i+tmprow, column=0+tmpcolumn, sticky='E')
            self.labelsBeaches.append(self.lbBeaches)

            tempentriesBeachesRowVars = tk.IntVar()
            self.enBeachesRow = Spinbox(self.beachesBL, from_=1, to=10, width=2, textvariable= tempentriesBeachesRowVars, validate="key", validatecommand=(vcmd, '%P'))
            self.entriesBeachesRowVars.append(tempentriesBeachesRowVars)
            self.enBeachesRow.grid(row=i+tmprow, column=1+tmpcolumn)
            self.entriesBeachesRow.append(self.enBeachesRow)

            tempentriesBeachesColVars = tk.IntVar()
            self.enBeachesCol = Spinbox(self.beachesBL, values=self.LETTERS_ARRAY, textvariable= tempentriesBeachesColVars, width=3)
            self.entriesBeachesColVars.append(tempentriesBeachesColVars)
            self.enBeachesCol.grid(row=i+tmprow, column=2+tmpcolumn)
            self.entriesBeachesCol.append(self.enBeachesCol)

            self.btBeaches = Button(self.beachesBL, text="Check", command=(xlsGrid, int(i)))
            self.btBeaches.grid(row=i+tmprow, column=3+tmpcolumn, sticky=E)
            self.buttonsBeaches.append(self.btBeaches)

        '''
        END BEACHES BEACH LITTER
        '''




        '''
        START SURVEYS BEACH LITTER
        '''
        self.OnlyLabelSurveys = Label(self.surveyBL, text=mytext)
        self.OnlyLabelSurveys['text'] = 'Name of the field'
        self.OnlyLabelSurveys.grid(row=58, column=0, columnspan=8, sticky=W+E)
        root.update()

        self.labelsSurveys = []
        self.entriesSurveysRow = []
        self.entriesSurveysCol = []
        self.buttonsSurveys = []
        tmprow=0
        tmpcolumn=4
        self.entriesSurveysRowVars = []
        self.entriesSurveysColVars = []
        for i in range(58):

            mytext=self.FIELDSURVEYS[i]
            if tmpcolumn == 4:
                tmpcolumn=0
            else:
                tmpcolumn=4

            if tmprow == 0:
                tmprow=1
            else:
                tmprow=0

            self.lbSurveys = Label(self.surveyBL, text=mytext)
            self.lbSurveys['text'] = str(mytext)+'('+str(i)+'): '
            self.lbSurveys.grid(row=i+tmprow, column=0+tmpcolumn, sticky='E')
            self.labelsSurveys.append(self.lbSurveys)

            tempentriesSurveysRowVars = tk.IntVar()
            self.enSurveysRow = Spinbox(self.surveyBL, from_=1, to=10, textvariable= tempentriesSurveysRowVars, width=2, validate="key", validatecommand=(vcmd, '%P'))
            self.entriesSurveysRowVars.append(tempentriesSurveysRowVars)
            self.enSurveysRow.grid(row=i+tmprow, column=1+tmpcolumn)
            self.entriesSurveysRow.append(self.enSurveysRow)

            tempentriesSurveysColVars = tk.IntVar()
            self.enSurveysCol = Spinbox(self.surveyBL, values=self.LETTERS_ARRAY, textvariable= tempentriesSurveysColVars, width=3)
            self.entriesSurveysColVars.append(tempentriesSurveysColVars)
            self.enSurveysCol.grid(row=i+tmprow, column=2+tmpcolumn)
            self.entriesSurveysCol.append(self.enSurveysCol)

            self.btSurveys = Button(self.surveyBL, text="Check", command=(xlsGridSurvey, int(i)))
            self.btSurveys.grid(row=i+tmprow, column=3+tmpcolumn, sticky=E)
            self.buttonsSurveys.append(self.btSurveys)

        '''
        END SURVEYS BEACH LITTER
        '''




        '''
        START SURVEYS SEA FLOOR
        '''
        '''
        New section to manage  timestamp (shot_timestamp) and 
        haul duration (haul_dur) survey tab area
        '''
        self.OnlyLabelSurveysSF = Label(self.surveySF, text=mytext)
        self.OnlyLabelSurveysSF['text'] = 'Name of the field'
        self.OnlyLabelSurveysSF.grid(row=58, column=0, columnspan=8, sticky=W+E)
        root.update()

        self.labelsSurveysSF = []
        self.entriesSurveysRowSF = []
        self.entriesSurveysColSF = []
        self.buttonsSurveysSF = []
        tmprow=0
        tmpcolumn=4
        self.entriesSurveysRowVarsSF = []
        self.entriesSurveysColVarsSF = []

        for i in range(24): 

            mytext=self.FIELDSURVEYSSEAFLOOR[i]
            if tmpcolumn == 4:
                tmpcolumn=0
            else:
                tmpcolumn=4

            if tmprow == 0:
                tmprow=1
            else:
                tmprow=0
                
            
            '''
            New section to manage  timestamp (shot_timestamp) and 
            haul duration (haul_dur) survey tab area
            '''    
            '''
            With new optional fields we change background and text
            '''
            if i in (22,23):
                self.lbSurveysSF = Label(self.surveySF, text=mytext,fg="#FF0000")
            else:
                self.lbSurveysSF = Label(self.surveySF, text=mytext)
            self.lbSurveysSF['text'] = str(mytext)+'('+str(i)+'): '
            self.lbSurveysSF.grid(row=i+tmprow, column=0+tmpcolumn, sticky='E')
            self.labelsSurveysSF.append(self.lbSurveysSF)

            tempentriesSurveysRowVarsSF = tk.IntVar()
            self.enSurveysRowSF = Spinbox(self.surveySF, from_=1, to=10, textvariable= tempentriesSurveysRowVarsSF, width=2, validate="key", validatecommand=(vcmd, '%P'))
            self.entriesSurveysRowVarsSF.append(tempentriesSurveysRowVarsSF)
            self.enSurveysRowSF.grid(row=i+tmprow, column=1+tmpcolumn)
            self.entriesSurveysRowSF.append(self.enSurveysRowSF)

            tempentriesSurveysColVarsSF = tk.IntVar()
            self.enSurveysColSF = Spinbox(self.surveySF, values=self.LETTERS_ARRAY, textvariable= tempentriesSurveysColVarsSF, width=3)
            self.entriesSurveysColVarsSF.append(tempentriesSurveysColVarsSF)
            self.enSurveysColSF.grid(row=i+tmprow, column=2+tmpcolumn)
            self.entriesSurveysColSF.append(self.enSurveysColSF)

            self.btSurveysSF = Button(self.surveySF, text="Check", command=(xlsGridSurveySF, int(i)))
            self.btSurveysSF.grid(row=i+tmprow, column=3+tmpcolumn, sticky=E)
            self.buttonsSurveysSF.append(self.btSurveysSF)
            
            if i == 23 :
                self.varCheckException = IntVar()
                self.R1Except = Checkbutton(self.surveySF, text="Use these fields",fg="#FF0000", variable=self.varCheckException)
                self.R1Except.grid(row=i+tmprow, column=3+tmpcolumn+1, sticky=E)
            
            
       

        '''
        END SURVEYS BEACH LITTER
        '''




        '''
        START LITTER SEA FLOOR
        '''

        root.update()

        self.labelsLitterSF = []
        self.entriesLitterRowSF = []
        self.entriesLitterColSF = []
        self.buttonsLitterSF = []
        tmprow=0
        tmpcolumn=4
        self.entriesLitterRowVarsSF = []
        self.entriesLitterColVarsSF = []
        for i in range(13):

            mytext=self.FIELDLITTERSEAFLOOR[i]
            if tmpcolumn == 4:
                tmpcolumn=0
            else:
                tmpcolumn=4

            if tmprow == 0:
                tmprow=1
            else:
                tmprow=0

            self.lbLitterSF = Label(self.litterSF, text=mytext)
            self.lbLitterSF['text'] = str(mytext)+'('+str(i)+'): '
            self.lbLitterSF.grid(row=i+tmprow+18, column=0+tmpcolumn, sticky='E')
            self.labelsLitterSF.append(self.lbLitterSF)

            tempentriesLitterRowVarsSF = tk.IntVar()
            self.enLitterRowSF = Spinbox(self.litterSF, from_=1, to=10, textvariable= tempentriesLitterRowVarsSF, width=2, validate="key", validatecommand=(vcmd, '%P'))
            self.entriesLitterRowVarsSF.append(tempentriesLitterRowVarsSF)
            self.enLitterRowSF.grid(row=i+tmprow+18, column=1+tmpcolumn)
            self.entriesLitterRowSF.append(self.enLitterRowSF)

            tempentriesLitterColVarsSF = tk.IntVar()
            self.enLitterColSF = Spinbox(self.litterSF, values=self.LETTERS_ARRAY, textvariable= tempentriesLitterColVarsSF, width=3)
            self.entriesLitterColVarsSF.append(tempentriesLitterColVarsSF)
            self.enLitterColSF.grid(row=i+tmprow+18, column=2+tmpcolumn)
            self.entriesLitterColSF.append(self.enLitterColSF)

            self.btLitterSF = Button(self.litterSF, text="Check", command=(xlsGridLitterSF, int(i)))
            self.btLitterSF.grid(row=i+tmprow+18, column=3+tmpcolumn, sticky=E)
            self.buttonsLitterSF.append(self.btLitterSF)



        self.OnlyLabelLitterSF = Label(self.litterSF, text=mytext)
        self.OnlyLabelLitterSF['text'] = 'Name of the field'
        self.OnlyLabelLitterSF.grid(row=i+tmprow+34, column=0, columnspan=8, sticky=W+E)

        self.SurveyNamesListSF = []
        self.ParamsNamesListSF = []


        '''
        END LITTER SEA FLOOR
        '''



        '''
        START ANIMALS BEACH LITTER
        '''
        self.OnlyLabelAnimals = Label(self.animalLitterBL, text=mytext)
        self.OnlyLabelAnimals['text'] = 'Name of the field'
        self.OnlyLabelAnimals.grid(row=8, column=0, columnspan=8, sticky=W+E)
        root.update()

        self.labelsAnimals = []
        self.entriesAnimalsRow = []
        self.entriesAnimalsCol = []
        self.buttonsAnimals = []
        tmprow=0
        tmpcolumn=4
        self.entriesAnimalsRowVars = []
        self.entriesAnimalsColVars = []
        for i in range(7):

            mytext=self.FIELDANIMALS[i]
            if tmpcolumn == 4:
                tmpcolumn=0
            else:
                tmpcolumn=4

            if tmprow == 0:
                tmprow=1
            else:
                tmprow=0

            self.lbAnimals = Label(self.animalLitterBL, text=mytext)
            self.lbAnimals['text'] = str(mytext)+'('+str(i)+'): '
            self.lbAnimals.grid(row=i+tmprow, column=0+tmpcolumn, sticky='E')
            self.labelsAnimals.append(self.lbAnimals)

            tempentriesAnimalsRowVars = tk.IntVar()
            self.enAnimalsRow = Spinbox(self.animalLitterBL, from_=1, to=10, width=2, textvariable= tempentriesAnimalsRowVars, validate="key", validatecommand=(vcmd, '%P'))
            self.entriesAnimalsRowVars.append(tempentriesAnimalsRowVars)
            self.enAnimalsRow.grid(row=i+tmprow, column=1+tmpcolumn)
            self.entriesAnimalsRow.append(self.enAnimalsRow)

            tempentriesAnimalsColVars = tk.IntVar()
            self.enAnimalsCol = Spinbox(self.animalLitterBL, values=self.LETTERS_ARRAY, textvariable= tempentriesAnimalsColVars, width=3)
            self.entriesAnimalsColVars.append(tempentriesAnimalsColVars)
            self.enAnimalsCol.grid(row=i+tmprow, column=2+tmpcolumn)
            self.entriesAnimalsCol.append(self.enAnimalsCol)

            self.btAnimals = Button(self.animalLitterBL, text="Check", command=(xlsGridAnimals, int(i)))
            self.btAnimals.grid(row=i+tmprow, column=3+tmpcolumn, sticky=E)
            self.buttonsAnimals.append(self.btAnimals)

        '''
        END ANIMALS BEACH LITTER
        '''


        '''
        START LITTER BEACH LITTER
        '''
        self.OnlyLabelLitter = Label(self.animalLitterBL, text=mytext)
        self.OnlyLabelLitter['text'] = 'Name of the field'
        self.OnlyLabelLitter.grid(row=27, column=0, columnspan=8, sticky=W+E)
        root.update()

        self.labelsLitter = []
        self.entriesLitterRow = []
        self.entriesLitterCol = []
        self.buttonsLitter = []
        tmprow=0
        tmpcolumn=4
        self.entriesLitterRowVars = []
        self.entriesLitterColVars = []
        for i in range(7):

            mytext=self.FIELDLITTER[i]
            if tmpcolumn == 4:
                tmpcolumn=0
            else:
                tmpcolumn=4

            if tmprow == 0:
                tmprow=1
            else:
                tmprow=0

            self.lbLitter = Label(self.animalLitterBL, text=mytext)
            self.lbLitter['text'] = str(mytext)+'('+str(i)+'): '
            self.lbLitter.grid(row=i+tmprow+18, column=0+tmpcolumn, sticky='E')
            self.labelsLitter.append(self.lbLitter)

            tempentriesLitterRowVars = tk.IntVar()
            self.enLitterRow = Spinbox(self.animalLitterBL, from_=1, to=10, textvariable= tempentriesLitterRowVars, width=2, validate="key", validatecommand=(vcmd, '%P'))
            self.entriesLitterRowVars.append(tempentriesLitterRowVars)
            self.enLitterRow.grid(row=i+tmprow+18, column=1+tmpcolumn)
            self.entriesLitterRow.append(self.enLitterRow)

            tempentriesLitterColVars = tk.IntVar()
            self.enLitterCol = Spinbox(self.animalLitterBL, values=self.LETTERS_ARRAY, textvariable= tempentriesLitterColVars, width=3)
            self.entriesLitterColVars.append(tempentriesLitterColVars)
            self.enLitterCol.grid(row=i+tmprow+18, column=2+tmpcolumn)
            self.entriesLitterCol.append(self.enLitterCol)

            self.btLitter = Button(self.animalLitterBL, text="Check", command=(xlsGridLitter, int(i)))
            self.btLitter.grid(row=i+tmprow+18, column=3+tmpcolumn, sticky=E)
            self.buttonsLitter.append(self.btLitter)



        self.SurveyNamesList = []
        self.BeachesNamesListUniq = []
        self.ParamsNamesList = []

        #When the user has a different sheet for the params definitions
        self.var1ParDesc = IntVar()
        self.ParDescSheet = Checkbutton(self.animalLitterBL, text="Params definitions in another sheet", variable=self.var1ParDesc)
        self.ParDescSheet.grid(row=i+tmprow+35, column=0, sticky=E)

        self.lbWichSheet = Label(self.animalLitterBL, text=mytext)
        self.lbWichSheet['text'] = 'Wich sheet? (0 is the first sheet)'
        self.lbWichSheet.grid(row=i+tmprow+36, column=0, sticky=E)

        self.enWichSheetVars = tk.IntVar()
        self.enWichSheet = Spinbox(self.animalLitterBL, from_=0, to=10, width=2, textvariable= self.enWichSheetVars, validate="key", validatecommand=(vcmd, '%P'))
        self.enWichSheet.grid(row=i+tmprow+36, column=1, sticky=W)


        self.lbWichSheetRow = Label(self.animalLitterBL, text=mytext)
        self.lbWichSheetRow['text'] = 'Wich ID row? (0 is the first row)'
        self.lbWichSheetRow.grid(row=i+tmprow+37, column=0, sticky=E)

        self.enWichSheetRowVars = tk.IntVar()
        self.enWichSheetRow = Spinbox(self.animalLitterBL, from_=0, to=10, width=2, textvariable= self.enWichSheetRowVars, validate="key", validatecommand=(vcmd, '%P'))
        self.enWichSheetRow.grid(row=i+tmprow+37, column=1, sticky=W)

        self.lbWichSheetCol = Label(self.animalLitterBL, text=mytext)
        self.lbWichSheetCol['text'] = 'Wich ID col?'
        self.lbWichSheetCol.grid(row=i+tmprow+38, column=0, sticky=E)

        self.enWichSheetColVars = tk.IntVar()
        self.enWichSheetCol = Spinbox(self.animalLitterBL, values=self.LETTERS_ARRAY, textvariable= self.enWichSheetColVars, width=3)
        self.enWichSheetCol.grid(row=i+tmprow+38, column=1, sticky=W)

        self.lbWichSheetNameCol = Label(self.animalLitterBL, text=mytext)
        self.lbWichSheetNameCol['text'] = 'Wich Name col?'
        self.lbWichSheetNameCol.grid(row=i+tmprow+39, column=0, sticky=E)

        self.enWichSheetNameColVars = tk.IntVar()
        self.enWichSheetNameCol = Spinbox(self.animalLitterBL, values=self.LETTERS_ARRAY, textvariable= self.enWichSheetNameColVars, width=3)
        self.enWichSheetNameCol.grid(row=i+tmprow+39, column=1, sticky=W)

        self.lbWichSheetOriginalNameCol = Label(self.animalLitterBL, text=mytext)
        self.lbWichSheetOriginalNameCol['text'] = 'Wich Original Name col?'
        self.lbWichSheetOriginalNameCol.grid(row=i+tmprow+40, column=0, sticky=E)

        self.enWichSheetOriginalNameColVars = tk.IntVar()
        self.enWichSheetOriginalNameCol = Spinbox(self.animalLitterBL, values=self.LETTERS_ARRAY, textvariable= self.enWichSheetOriginalNameColVars, width=3)
        self.enWichSheetOriginalNameCol.grid(row=i+tmprow+40, column=1, sticky=W)



        '''
        END LITTER BEACH LITTER
        '''

        
        self.createOutputModel = Button(self.animalLitterBL, text="Save Model", command=(savefilemodel))
        self.createOutputModel.grid(row=i+tmprow+30, column=3+tmpcolumn, sticky=E)

        self.createOutput = Button(self.animalLitterBL, text="Create litter XLS", command=(savefilexls))
        self.createOutput.grid(row=i+tmprow+30, column=4+tmpcolumn, sticky=E)


        '''
        START PLOTS BEACH LITTER
        '''

        self.checkPlotsButton = Button(self.plotBL, text="Load Beach Litter Survey Plot File", command=(openfilesurveyplot))
        self.checkPlotsButton.grid(row=1, column=1, sticky=E)

        self.entryInfoInputFilePlotSurvey = Entry(self.plotBL, width=20, validate="key")
        self.entryInfoInputFilePlotSurvey.grid(row=1, column=2, sticky=W)

        self.ExecutePlotsButton = Button(self.plotBL, text="Execute", command=(plotMySurvey))
        self.ExecutePlotsButton.grid(row=1, column=5, sticky=E)

        self.varMoreInfoPlot = IntVar()
        self.buttonMoreInfoPlot = Checkbutton(self.plotBL, text="Params complete description (this could deform the plot)", variable=self.varMoreInfoPlot)
        self.buttonMoreInfoPlot.grid(row=2, column=1,columnspan=10, sticky=W)
        
        self.varPiePlot = IntVar()
        self.buttonvarPiePlot = Checkbutton(self.plotBL, text="Pie Plot", variable=self.varPiePlot)
        self.buttonvarPiePlot.grid(row=3, column=1,columnspan=10, sticky=W)
        
        self.varHBarPlot = IntVar()
        self.buttonvarHBarPlot = Checkbutton(self.plotBL, text="Horizontal Bar Plot", variable=self.varHBarPlot)
        self.buttonvarHBarPlot.grid(row=4, column=1,columnspan=10, sticky=W)
        
        self.varVBarPlot = IntVar()
        self.buttonvarVBarPlot = Checkbutton(self.plotBL, text="Vertical Bar Plot", variable=self.varVBarPlot)
        self.buttonvarVBarPlot.grid(row=5, column=1,columnspan=10, sticky=W)    

        self.varMapPlot = IntVar()
        self.buttonMapPlot = Checkbutton(self.plotBL, text="Show START coordinates of the survey (a web page will be opened on your browser)", variable=self.varMapPlot)
        self.buttonMapPlot.grid(row=6, column=1,columnspan=10, sticky=W)    

        self.varMapPlotEnd = IntVar()
        self.buttonMapPlotEnd = Checkbutton(self.plotBL, text="Show END coordinates of the survey (a web page will be opened on your browser)", variable=self.varMapPlotEnd)
        self.buttonMapPlotEnd.grid(row=7, column=1,columnspan=10, sticky=W)  
        
        self.mylistScrollListSurvey = Listbox(self.plotBL,height=40, width=80,selectmode='multiple')



        '''
        END PLOTS BEACH LITTER
        '''



        '''
        START PLOTS SEA FLOOR
        '''


        self.checkPlotsButtonSF = Button(self.plotSF, text="Load Sea Floor Survey Plot File", command=(openfilesurveyplotSF))
        self.checkPlotsButtonSF.grid(row=1, column=1, sticky=E)

        self.entryInfoInputFilePlotSurveySF = Entry(self.plotSF, width=20, validate="key")
        self.entryInfoInputFilePlotSurveySF.grid(row=1, column=2, sticky=W)

        self.ExecutePlotsButtonSF = Button(self.plotSF, text="Execute", command=(plotMySurveySF))
        self.ExecutePlotsButtonSF.grid(row=1, column=5, sticky=E)
        
        mytextseparator='TAB'
        self.labeMySeparatorCSVPlotSF = Label(self.plotSF, text=mytextseparator)        
        self.labeMySeparatorCSVPlotSF['text'] = 'The separator is ' + str(mytextseparator)
        self.labeMySeparatorCSVPlotSF.grid(row=1, column=6, sticky=E)
               
        self.varPiePlotSF = IntVar()
        self.buttonvarPiePlotSF = Checkbutton(self.plotSF, text="Pie Plot", variable=self.varPiePlotSF)
        self.buttonvarPiePlotSF.grid(row=2, column=1,columnspan=10, sticky=W)
        
        self.varHBarPlotSF = IntVar()
        self.buttonvarHBarPlotSF = Checkbutton(self.plotSF, text="Horizontal Bar Plot", variable=self.varHBarPlotSF)
        self.buttonvarHBarPlotSF.grid(row=3, column=1,columnspan=10, sticky=W)
        
        self.varVBarPlotSF = IntVar()
        self.buttonvarVBarPlotSF = Checkbutton(self.plotSF, text="Vertical Bar Plot", variable=self.varVBarPlotSF)
        self.buttonvarVBarPlotSF.grid(row=4, column=1,columnspan=10, sticky=W)    

        self.varMapPlotSF = IntVar()
        self.buttonMapPlotSF = Checkbutton(self.plotSF, text="Show START coordinates of the survey (a web page will be opened on your browser)", variable=self.varMapPlotSF)
        self.buttonMapPlotSF.grid(row=5, column=1,columnspan=10, sticky=W)    

        self.varMapPlotEndSF = IntVar()
        self.buttonMapPlotEndSF = Checkbutton(self.plotSF, text="Show END coordinates of the survey (a web page will be opened on your browser)", variable=self.varMapPlotEndSF)
        self.buttonMapPlotEndSF.grid(row=6, column=1,columnspan=10, sticky=W)  

        self.mylistScrollListSurveySF = Listbox(self.plotSF,height=40, width=80,selectmode='multiple')

        
        '''
        END PLOTS SEA FLOOR
        '''







        '''
        START SCATTER PLOTS BEACH LITTER
        '''

        self.checkPlotsButtonParams = Button(self.scatterBL, text="Load Beach Litter Params Plot File", command=(openfileparamsplot))
        self.checkPlotsButtonParams.grid(row=1, column=1, sticky=W)

        self.entryInfoInputFilePlotParams = Entry(self.scatterBL, width=20, validate="key")
        self.entryInfoInputFilePlotParams.grid(row=1, column=2, sticky=W)

        self.ExecutePlotsButtonParams = Button(self.scatterBL, text="Execute", command=(plotMyParams))
        self.ExecutePlotsButtonParams.grid(row=1, column=5, sticky=E)

        self.varScatterPlot = IntVar()
        self.buttonvarScatterPlot = Checkbutton(self.scatterBL, text="Scatter 2D Plot", variable=self.varScatterPlot)
        self.buttonvarScatterPlot.grid(row=2, column=1,columnspan=10, sticky=W)
        
        self.varScatterDPlot = IntVar()
        self.buttonvarScatterDPlot = Checkbutton(self.scatterBL, text="Scatter 3D Plot", variable=self.varScatterDPlot)
        self.buttonvarScatterDPlot.grid(row=3, column=1,columnspan=10, sticky=W)
        
        self.varScatterLegendaPlot = IntVar()
        self.buttonScatterLegendaPlot = Checkbutton(self.scatterBL, text="Add legenda to plot", variable=self.varScatterLegendaPlot)
        self.buttonScatterLegendaPlot.grid(row=4, column=1,columnspan=10, sticky=W)  

        self.varScatterCoordPlot = IntVar()
        self.buttonScatterCoordPlot = Checkbutton(self.scatterBL, text="Create CSV file for Google Maps", variable=self.varScatterCoordPlot)
        self.buttonScatterCoordPlot.grid(row=5, column=1,columnspan=10, sticky=W) 

        self.mylistScrollListSurveyParams = Listbox(self.scatterBL,height=40, width=80)
        '''
        END SCATTER PLOTS BEACH LITTER
        '''



        '''
        START SCATTER PLOTS SEA FLOOR
        '''

        self.checkPlotsButtonParamsSF = Button(self.scatterSF, text="Load Sea Floor Params Plot File", command=(openfileparamsplotSF))
        self.checkPlotsButtonParamsSF.grid(row=1, column=1, sticky=W)

        self.entryInfoInputFilePlotParamsSF = Entry(self.scatterSF, width=20, validate="key")
        self.entryInfoInputFilePlotParamsSF.grid(row=1, column=2, sticky=W)

        self.ExecutePlotsButtonParamsSF = Button(self.scatterSF, text="Execute", command=(plotMyParamsSF))
        self.ExecutePlotsButtonParamsSF.grid(row=1, column=5, sticky=E)
        
        self.labeMySeparatorCSVscatterSF = Label(self.scatterSF, text=mytextseparator)
        self.labeMySeparatorCSVscatterSF['text'] = 'The separator is ' + str(mytextseparator)
        self.labeMySeparatorCSVscatterSF.grid(row=1, column=6, sticky=E)

        self.varScatterPlotSF = IntVar()
        self.buttonvarScatterPlotSF = Checkbutton(self.scatterSF, text="Scatter 2D Plot", variable=self.varScatterPlotSF)
        self.buttonvarScatterPlotSF.grid(row=2, column=1,columnspan=10, sticky=W)
        
        self.varScatterDPlotSF = IntVar()
        self.buttonvarScatterDPlotSF = Checkbutton(self.scatterSF, text="Scatter 3D Plot", variable=self.varScatterDPlotSF)
        self.buttonvarScatterDPlotSF.grid(row=3, column=1,columnspan=10, sticky=W)
        
        self.varScatterLegendaPlotSF = IntVar()
        self.buttonScatterLegendaPlotSF = Checkbutton(self.scatterSF, text="Add legenda to plot", variable=self.varScatterLegendaPlotSF)
        self.buttonScatterLegendaPlotSF.grid(row=4, column=1,columnspan=10, sticky=W)  

        self.varScatterCoordPlotSF = IntVar()
        self.buttonScatterCoordPlotSF = Checkbutton(self.scatterSF, text="Create CSV file for Google Maps", variable=self.varScatterCoordPlotSF)
        self.buttonScatterCoordPlotSF.grid(row=5, column=1,columnspan=10, sticky=W)  

        self.mylistScrollListSurveyParamsSF = Listbox(self.scatterSF,height=40, width=80)
        '''
        END SCATTER PLOTS BEACH LITTER
        '''

 

    def LinkButtonA(self):
        webbrowser.open('https://doi.org/10.6092/a75ba101-ebb9-4bad-9b7f-423a1327c76f')


    def LinkButtonB(self):
        webbrowser.open('http://dx.doi.org/10.6092/15c0d34c-a01a-4091-91ac-7c4f561ab508')


    def LinkButtonC(self):
        webbrowser.open('https://doi.org/10.6092/9593a449-37c1-4fd9-84bb-e91978ac8c40')


    def LinkButtonD(self):
        webbrowser.open('http://emodnet-chemistry.maris2.nl/v_cdi_v3/result.asp?formname=search&v0_30=LITT,UMLW,UMLS,BLIT,SLIT&v1_30=parameters_p02&v2_30=4')


    def LinkButtonE(self):
        webbrowser.open('http://ec.oceanbrowser.net/emodnet/?server=http://www.ifremer.fr/services/wms/emodnet_chemistry2')


    def LinkButtonF(self):
        webbrowser.open('https://www.emodnet-chemistry.eu/products/catalogue#/search?fast=index&_content_type=json&from=1&to=20&sortBy=changeDate&any=litter%20datasets')



    def checkGridXls(self, wichfield):

        mytext=self.FIELDBEACHES[int(wichfield)]
        tmpMyRow=int(self.entriesBeachesRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesBeachesCol[int(wichfield)].get()))-1


        stringtmpMyRow=str(self.entriesBeachesRow[int(wichfield)].get())
        stringtmpMyCol=str(self.LETTERS_ARRAY.index(self.entriesBeachesCol[int(wichfield)].get()))

        wichSheet=int(self.entryInfoBeaches.get())-1

        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':

                    self.input_beaches_work_sheet = self.book.sheet_by_index(wichSheet)
 
                    try:
                        actualvalue=self.input_beaches_work_sheet.cell_value(tmpMyRow, tmpMyCol)
                    except Exception as e:
                        print("WARNING!", e, "occurred.")
                        self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)

                    self.OnlyLabel['text'] = str(mytext)+'('+str(wichfield)+'): '+str(actualvalue)

        root.update()


    def checkGridXlsSurvey(self, wichfield):

        mytext=self.FIELDSURVEYS[int(wichfield)]
        tmpMyRow=int(self.entriesSurveysRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesSurveysCol[int(wichfield)].get()))-1

        stringtmpMyRow=str(self.entriesSurveysRow[int(wichfield)].get())
        stringtmpMyCol=str(self.LETTERS_ARRAY.index(self.entriesSurveysCol[int(wichfield)].get()))

        wichSheet=int(self.entryInfoSurveys.get())-1

        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':

                    self.input_surveys_work_sheet = self.book.sheet_by_index(wichSheet)
                    try:
                        actualvalue=self.input_surveys_work_sheet.cell_value(tmpMyRow, tmpMyCol)
                    except Exception as e:
                        print("WARNING!", e, "occurred.")
                        self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)
                        
                    self.OnlyLabelSurveys['text'] = str(mytext)+'('+str(wichfield)+'): '+str(actualvalue)

        root.update()


    def checkGridXlsSurveySF(self, wichfield):

        mytext=self.FIELDSURVEYSSEAFLOOR[int(wichfield)]
        tmpMyRow=int(self.entriesSurveysRowSF[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesSurveysColSF[int(wichfield)].get()))-1

        stringtmpMyRow=str(self.entriesSurveysRowSF[int(wichfield)].get())
        stringtmpMyCol=str(self.LETTERS_ARRAY.index(self.entriesSurveysColSF[int(wichfield)].get()))

        wichSheet=int(self.entryInfoSurveysSF.get())-1
        
        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':

                    self.input_surveys_work_sheetSF = self.bookSF.sheet_by_index(wichSheet)
                    try:
                        if wichfield in ('1','7','8','10','17','21'): # Need to manage the integer values (we have to cast them to delete the decimal)
                            actualvalue=math.floor(int(self.input_surveys_work_sheetSF.cell_value(tmpMyRow, tmpMyCol)))
                            
                        else:
                            actualvalue=self.input_surveys_work_sheetSF.cell_value(tmpMyRow, tmpMyCol)
                            
                    except Exception as e:
                        print("WARNING!", e, "occurred.")
                        self.infoSFarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)
                        
                    self.OnlyLabelSurveysSF['text'] = str(mytext)+'('+str(wichfield)+'): '+str(actualvalue)

        root.update()



    def checkGridXlsAnimals(self, wichfield):

        mytext=self.FIELDANIMALS[int(wichfield)]
        tmpMyRow=int(self.entriesAnimalsRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesAnimalsCol[int(wichfield)].get()))-1

        stringtmpMyRow=str(self.entriesAnimalsRow[int(wichfield)].get())
        stringtmpMyCol=str(self.LETTERS_ARRAY.index(self.entriesAnimalsCol[int(wichfield)].get()))

        wichSheet=int(self.entryInfoAnimals.get())-1

        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':

                    self.input_animals_work_sheet = self.book.sheet_by_index(wichSheet)
                    try:
                        actualvalue=self.input_animals_work_sheet.cell_value(tmpMyRow, tmpMyCol)
                    except Exception as e:
                        print("WARNING!", e, "occurred.")
                        self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)
                    self.OnlyLabelAnimals['text'] = str(mytext)+'('+str(wichfield)+'): '+str(actualvalue)

        root.update()


    def checkGridXlsLitter(self, wichfield):

        mytext=self.FIELDLITTER[int(wichfield)]
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1

        stringtmpMyRow=str(self.entriesLitterRow[int(wichfield)].get())
        stringtmpMyCol=str(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))

        wichSheet=int(self.entryInfoLitter.get())-1

        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':

                    tmpINTwichfield=int(wichfield)
                    if tmpINTwichfield == 2:

                        populateNameRow=str(self.entriesLitterRow[int(wichfield)].get())
                        populateOriginalNameRow=str(self.entriesLitterRow[int(wichfield)].get())
                        self.entriesLitterRow[3].delete(0,END)
                        self.entriesLitterRow[3].insert(0,populateNameRow)
                        self.entriesLitterRow[4].delete(0,END)
                        self.entriesLitterRow[4].insert(0,populateOriginalNameRow)

                        root.update()

                    self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
                    
                    try:
                        actualvalue=self.input_litter_work_sheet.cell_value(tmpMyRow, tmpMyCol)
                    except Exception as e:
                        print("WARNING!", e, "occurred.")
                        self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)
                        
                    self.OnlyLabelLitter['text'] = str(mytext)+'('+str(wichfield)+'): '+str(actualvalue)

        root.update()




    def checkGridXlsLitterSF(self, wichfield):

        mytext=self.FIELDLITTERSEAFLOOR[int(wichfield)]
        tmpMyRow=int(self.entriesLitterRowSF[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterColSF[int(wichfield)].get()))-1

        stringtmpMyRow=str(self.entriesLitterRowSF[int(wichfield)].get())
        stringtmpMyCol=str(self.LETTERS_ARRAY.index(self.entriesLitterColSF[int(wichfield)].get()))

        wichSheet=int(self.entryInfoLitterSF.get())-1

        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':

                    tmpINTwichfieldSF=int(wichfield)
                    if tmpINTwichfieldSF == 2:

                        populateNameRowSF=str(self.entriesLitterRowSF[int(wichfield)].get())
                        populateOriginalNameRowSF=str(self.entriesLitterRowSF[int(wichfield)].get())
                        self.entriesLitterRowSF[3].delete(0,END)
                        self.entriesLitterRowSF[3].insert(0,populateNameRowSF)
                        self.entriesLitterRowSF[4].delete(0,END)
                        self.entriesLitterRowSF[4].insert(0,populateOriginalNameRowSF)

                        root.update()

                    self.input_litter_work_sheetSF = self.bookSF.sheet_by_index(wichSheet)
                    try:
                        if wichfield in ('9','11'): # Need to manage the integer values (we have to cast them to delete the decimal)
                            actualvalueSF=int(self.input_litter_work_sheetSF.cell_value(tmpMyRow, tmpMyCol))
                        else:
                            actualvalueSF=self.input_litter_work_sheetSF.cell_value(tmpMyRow, tmpMyCol)
                        
                        
                    except Exception as e:
                        print("WARNING!", e, "occurred.")
                        self.infoSFarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)
                        
                    self.OnlyLabelLitterSF['text'] = str(mytext)+'('+str(wichfield)+'): '+str(actualvalueSF)

        root.update()




    def validatenumber(self, new_text):
        if not new_text: # the field is being cleared
            self.entered_number = 0
            return True

        try:
            self.entered_number = int(new_text)
            return True
        except ValueError:
            return False


    def ShowBeachLitterTabs(self):
        self.HideAllTabs()
        self.nb.tab(1, state="normal")
        self.nb.tab(2, state="normal")
        self.nb.tab(3, state="normal")
        self.nb.tab(4, state="normal")
        self.nb.tab(5, state="normal")
        self.nb.tab(6, state="normal")
        self.nb.tab(12, state="hidden")
        self.nb.tab(13, state="hidden")
        self.nb.tab(14, state="hidden")
        self.nb.tab(15, state="hidden")
        self.master.config(menu=self.menuBL)


    def ShowSeaFloorLitterTabs(self):
        self.HideAllTabs()
        self.nb.tab(7, state="normal")
        self.nb.tab(8, state="normal")
        self.nb.tab(9, state="normal")
        self.nb.tab(10, state="normal")
        self.nb.tab(11, state="normal")
        self.nb.tab(12, state="hidden")
        self.nb.tab(13, state="hidden")
        self.nb.tab(14, state="hidden")
        self.nb.tab(15, state="hidden")
        self.master.config(menu=self.menuSF)
        
        
        
    def ShowCoastalMacroLitterTabs(self):
        self.HideAllTabs()
        self.nb.tab(12, state="normal")
        self.nb.tab(13, state="hidden")
        self.nb.tab(14, state="hidden")
        self.nb.tab(15, state="hidden")
        self.master.config(menu=self.emptymenu)
        
    def ShowOpenSeaMacroLitterTabs(self):
        self.HideAllTabs()
        self.nb.tab(13, state="normal")
        self.nb.tab(14, state="hidden")
        self.nb.tab(15, state="hidden")
        self.master.config(menu=self.emptymenu)


    def ShowUtilitiesTabs(self):
        self.HideAllTabs()
        self.nb.tab(14, state="normal")
        self.nb.tab(15, state="normal")
        self.master.config(menu=self.emptymenu)


    def HideAllTabs(self):
        self.nb.tab(1, state="hidden")
        self.nb.tab(2, state="hidden")
        self.nb.tab(3, state="hidden")
        self.nb.tab(4, state="hidden")
        self.nb.tab(5, state="hidden")
        self.nb.tab(6, state="hidden")
        self.nb.tab(7, state="hidden")
        self.nb.tab(8, state="hidden")
        self.nb.tab(9, state="hidden")
        self.nb.tab(10, state="hidden")
        self.nb.tab(11, state="hidden")
        self.nb.tab(12, state="hidden")
        self.nb.tab(13, state="hidden")
        self.nb.tab(14, state="hidden")
        self.nb.tab(15, state="hidden")
        self.master.config(menu=self.emptymenu)


    def changeLabelSeparator(self):
        if self.varRadioOutputCSV.get() == 2:
            self.labeMySeparatorCSVPlotSF['text'] = 'The separator is COMMA'
            self.labeMySeparatorCSVscatterSF['text'] = 'The separator is COMMA'
        else:
            self.labeMySeparatorCSVPlotSF['text'] = 'The separator is TAB'
            self.labeMySeparatorCSVscatterSF['text'] = 'The separator is TAB'


    def executePlotParams(self):

        selectedScrollListParams=self.mylistScrollListSurveyParams.curselection()


        if selectedScrollListParams:

            for selectedParams in selectedScrollListParams:
                
                valueselectedScrollListParams = self.mylistScrollListSurveyParams.get(selectedParams)

                tempWhichNameOutputFilePlotParams = str(self.entryInfoInputFilePlotParams.get())
                if os.path.exists(tempWhichNameOutputFilePlotParams):

                    self.paramsPlot_book = xlrd.open_workbook(tempWhichNameOutputFilePlotParams)

                    self.all_ParamsSurveyPlot_work_sheet = self.paramsPlot_book.sheet_by_index(1)
                    self.all_ParamsValuePlot_work_sheet = self.paramsPlot_book.sheet_by_index(3)

                    allParamsValuePlot_current_row=1
                    allParamsValuePlot_num_rows = self.all_ParamsValuePlot_work_sheet.nrows

                    executeplotScatter=self.varScatterPlot.get()
                    executeplotScatterD=self.varScatterDPlot.get()
                    executeplotLegenda=self.varScatterLegendaPlot.get()
                    executeplotCSVGoogleMaps=self.varScatterCoordPlot.get()

                    paramLAT = []
                    paramLON = []
                    paramVALUE = []

                    del paramLAT [:]
                    del paramLON [:]
                    del paramVALUE [:]

                    pardescription=''

                    # When we need to create the Google Maps CSV file
                    if executeplotCSVGoogleMaps == 1:
                        GoogleMapsCSVName = str('GoogleMaps_'+time.strftime("%Y%m%d%H%M%S")+'_Param_labelled_'+valueselectedScrollListParams+'.csv')
                        GoogleMapsCSVOutputFile = open(GoogleMapsCSVName,"a")
                        GoogleMapsCSVOutputFile.write("ParamName,description,value,lat,lon")

                    while allParamsValuePlot_current_row < allParamsValuePlot_num_rows:

                        tmpParamsSurveyPlotName = str(self.all_ParamsValuePlot_work_sheet.cell_value(allParamsValuePlot_current_row, 2))
        
                        if str(tmpParamsSurveyPlotName) == str(valueselectedScrollListParams):
        
                            tmpParamsValuePlot=int(self.all_ParamsValuePlot_work_sheet.cell_value(allParamsValuePlot_current_row, 5))
                            if pardescription == '':
                                pardescription=str(self.all_ParamsValuePlot_work_sheet.cell_value(allParamsValuePlot_current_row, 3))

                            tmpParamsSurveyPlot=str(self.all_ParamsValuePlot_work_sheet.cell_value(allParamsValuePlot_current_row, 0))

                            allParamsSurveyPlot_current_row=1
                            allParamsSurveyPlot_num_rows = self.all_ParamsSurveyPlot_work_sheet.nrows

                            while allParamsSurveyPlot_current_row < allParamsSurveyPlot_num_rows:

                                tmpSurveyPlotName = str(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 1))
        
                                if str(tmpSurveyPlotName) == str(tmpParamsSurveyPlot):
                                    #We use only the starting coordinates (if they exist)
                                    tmpParamsLatString=str(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 8))
                                    tmpParamsLonString=str(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 9))

                                    if tmpParamsLatString != '' and tmpParamsLonString != '' and int(tmpParamsValuePlot) > 0:
                                        tmpParamsLat=float(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 8))
                                        tmpParamsLon=float(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 9))

                                        paramLAT.append(float(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 8))) 
                                        paramLON.append(float(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 9)))
                                        paramVALUE.append(int(tmpParamsValuePlot))

                                        # When we need to fill the Google Maps CSV file
                                        if executeplotCSVGoogleMaps == 1:
                                            GoogleMapsCSVOutputFile.write('\n'+str(valueselectedScrollListParams)+','+str(pardescription)+','+str(tmpParamsValuePlot)+','+str(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 8).replace(",","."))+','+str(self.all_ParamsSurveyPlot_work_sheet.cell_value(allParamsSurveyPlot_current_row, 9).replace(",",".")))


                                allParamsSurveyPlot_current_row += 1


                        allParamsValuePlot_current_row += 1

                    # When we need to close the Google Maps CSV file
                    if executeplotCSVGoogleMaps == 1:
                        GoogleMapsCSVOutputFile.close()
                        self.infoBLarea.insert(END, '\nThe Google Maps CSV '+GoogleMapsCSVName+' has been saved!') 



                    if executeplotScatter == 1:
                        self.infoBLarea.insert(END, "\n\nSCATTER PLOT 2D executed for param labelled "+str(valueselectedScrollListParams)+' '+pardescription)
                        fig1, scat1 = plt.subplots()
                    
                        myScatter = scat1.scatter(paramLON, paramLAT, s=300, c=paramVALUE, alpha=0.5)


                        # produce a legend with the unique colors from the scatter
                        legend1 = scat1.legend(*myScatter.legend_elements(), loc="lower left", title="Number of items")
                    
                        if executeplotLegenda == 1:
                            # produce a legend with a cross section of sizes from the scatter

                            scat1.grid(True)

                        scat1.set_title('SCATTER PLOT executed for param labelled '+str(valueselectedScrollListParams)+'\n'+pardescription+'\n X-axis:Longitude, Y-axis:Latitude')


                    if executeplotScatterD == 1:
                        self.infoBLarea.insert(END, "\n\nSCATTER PLOT 3D executed for param labelled "+str(valueselectedScrollListParams)+' '+pardescription)
                        fig2=plt.figure()
                        scat2=fig2.add_subplot(121,projection='3d')
                        scat2.scatter(paramLON, paramLAT,paramVALUE,c=paramVALUE,s=60)

                        if executeplotLegenda == 1:
                            scat2.legend(*myScatter.legend_elements(), loc="upper right", title="Number of items")
                            scat2.view_init(elev=20., azim=-35)

                        scat2.set_title('3D SCATTER PLOT executed for param labelled '+str(valueselectedScrollListParams)+' '+pardescription+'\n X-axis:Longitude, Y-axis:Latitude')

                    if executeplotScatter == 1 or executeplotScatterD == 1:
                        plt.show()

        else:

            self.infoBLarea.insert(END, '\nThis parameter does not exists!')




    def executePlot(self):

        selectedScrollListSurvey=self.mylistScrollListSurvey.curselection()
                
        executePiePlot=self.varPiePlot.get()
        executeHBarPlot=self.varHBarPlot.get()
        executeVBarPlot=self.varVBarPlot.get()
        
        
        if selectedScrollListSurvey:
            
            for selectedSurvey in selectedScrollListSurvey:
                
                valueselectedScrollListSurvey = self.mylistScrollListSurvey.get(selectedSurvey)
              
                labels = []
                labelsWithDesc = []
                labelsDesc = []
                sizes = []
                sizesInt = []
        
                ParamComplDescr=self.varMoreInfoPlot.get()

                ShowMapSurvey=self.varMapPlot.get()
                ShowMapSurveyEnd=self.varMapPlotEnd.get()
        
                tempWhichNameOutputFilePlot = str(self.entryInfoInputFilePlotSurvey.get())
                if os.path.exists(tempWhichNameOutputFilePlot):
        
                    self.survyesPlot_book = xlrd.open_workbook(tempWhichNameOutputFilePlot)

                    self.all_surveysPlot_work_sheet = self.survyesPlot_book.sheet_by_index(3)
        
                    allsurveysPlot_current_row=1
                    allsurveysPlot_num_rows = self.all_surveysPlot_work_sheet.nrows

                    self.infoBLarea.insert(END, "\nPARAMETERS LIST FOR SURVEY "+str(valueselectedScrollListSurvey)+"\nSTART LIST ------------------------------------------------------------------------------------------")
          
                    while allsurveysPlot_current_row < allsurveysPlot_num_rows:
        
                        tmpInputSurveyPlotValue = str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 0))
        
                        if tmpInputSurveyPlotValue == self.SurveyNamesList[int(selectedSurvey)]:
        
                            tmpFielNoItemValuePlot=int(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 5))
        
                            if ParamComplDescr == 1:
        
                                if tmpFielNoItemValuePlot > 0:
                                    labels.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 2))+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 1))+')'+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 3))+')')                   
                                    labelsWithDesc.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 2))+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 1))+')'+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 3))+')')   
                                    labelsDesc.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 3)))                                   
                                    sizes.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 5))) 
                                    sizesInt.append(int(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 5))) 

                                    self.infoBLarea.insert(END, '\n'+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 2))+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 1))+')'+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 3))+') N. of items: '+str(tmpFielNoItemValuePlot))
        
                            else :
        
                                if tmpFielNoItemValuePlot > 0:
                                    labels.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 2)))                   
                                    labelsWithDesc.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 2))+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 3))+')')   
                                    labelsDesc.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 3)))                                   
                                    sizes.append(str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 5))) 
                                    sizesInt.append(int(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 5))) 

                                    self.infoBLarea.insert(END, '\n'+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 2))+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 1))+')'+' ('+str(self.all_surveys_work_sheet.cell_value(allsurveysPlot_current_row, 3))+') N. of items: '+str(tmpFielNoItemValuePlot))
        
                          
                        allsurveysPlot_current_row += 1

                    self.infoBLarea.insert(END, "\nEND LIST ------------------------------------------------------------------------------------------")
        

                    # Only for the Maps links (START)
                    self.all_surveysMAPS_work_sheet = self.survyesPlot_book.sheet_by_index(1)

                    allsurveysMAPS_current_row=1
                    allsurveysMAPS_num_rows = self.all_surveysMAPS_work_sheet.nrows
  
                    while allsurveysMAPS_current_row < allsurveysMAPS_num_rows:

                        tmpInputSurveyMAPSValue = str(self.all_surveysMAPS_work_sheet.cell_value(allsurveysMAPS_current_row, 1))

                        if tmpInputSurveyMAPSValue == self.SurveyNamesList[int(selectedSurvey)]:

                            #We must rebuild the date for the output inside the plot
                            tmpFieldSurveyDate=self.all_surveysMAPS_work_sheet.cell_value(allsurveysMAPS_current_row, 4)

                            if isinstance(tmpFieldSurveyDate, datetime.datetime):
                                tmpRebuiltDate = datetime.datetime(*xlrd.xldate_as_tuple(tmpFieldSurveyDate, self.survyesPlot_book.datemode))
                            else:
                                tmpRebuiltDate = str(tmpFieldSurveyDate)

                            tmpFieldLatitudeStart=str(self.all_surveysMAPS_work_sheet.cell_value(allsurveysMAPS_current_row, 8))
                            tmpFieldLatitudeStop=str(self.all_surveysMAPS_work_sheet.cell_value(allsurveysMAPS_current_row, 10))
                            tmpFieldLongitudeStart=str(self.all_surveysMAPS_work_sheet.cell_value(allsurveysMAPS_current_row, 9))
                            tmpFieldLongitudeStop=str(self.all_surveysMAPS_work_sheet.cell_value(allsurveysMAPS_current_row, 11))

                            self.infoBLarea.insert(END, "\nSTART MAPS LINKS ------------------------------------------------------------------------------------")
                            self.infoBLarea.insert(END, "\nUse the following link to map the start of the survey: http://www.google.com/maps/place/"+tmpFieldLatitudeStart+','+tmpFieldLongitudeStart)
                            self.infoBLarea.insert(END, "\nUse the following link to map the end of the survey: http://www.google.com/maps/place/"+tmpFieldLatitudeStop+','+tmpFieldLongitudeStop)
                            self.infoBLarea.insert(END, "\nEND MAPS LINKS ------------------------------------------------------------------------------------")

                            if ShowMapSurvey == 1:

                                webbrowser.open('http://www.google.com/maps/place/'+tmpFieldLatitudeStart+','+tmpFieldLongitudeStart)

                            if ShowMapSurveyEnd == 1:

                                webbrowser.open('http://www.google.com/maps/place/'+tmpFieldLatitudeStop+','+tmpFieldLongitudeStop)
                  
                  
                        allsurveysMAPS_current_row += 1
                    # Only for the Maps links (END)




        
                    if executePiePlot == 1: 
                        self.infoBLarea.insert(END, "\n\nPIE PLOT executed for "+str(valueselectedScrollListSurvey))
                        fig1, ax1 = plt.subplots()
                        ax1.pie(sizes, labels=None, autopct='%1.1f%%', pctdistance=0.6, shadow=True, startangle=90, radius=1.5)
                        ax1.set_title("Pie plot for survey "+str(valueselectedScrollListSurvey)+"\ndate:"+str(tmpRebuiltDate)+"\nLAT start:"+str(tmpFieldLatitudeStart)+" LON start:"+str(tmpFieldLongitudeStart)+"\nLAT end:"+str(tmpFieldLatitudeStop)+" LON end:"+str(tmpFieldLongitudeStop))
                        legend = ax1.legend(labels=labelsWithDesc, loc='upper right', shadow=True, fontsize='x-small')
                        # Put a nicer background color on the legend.
                        legend.get_frame().set_facecolor('#ffffff')
                        ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        
                    if executeHBarPlot == 1:
                        self.infoBLarea.insert(END, "\n\nHORIZONTAL PLOT executed for "+str(valueselectedScrollListSurvey))
                        fig, ax2 = plt.subplots()
                        y_pos = np.arange(len(labels))
                        performance = sizesInt
                        error = 0
                        ax2.barh(y_pos, performance, xerr=error, align='center')
                        ax2.set_yticks(y_pos)
                        ax2.set_yticklabels(labels)
                        ax2.invert_yaxis()  # labels read top-to-bottom
                        ax2.set_xlabel('Number of items')
                        ax2.set_title("Horizontal bar chart for survey "+str(valueselectedScrollListSurvey)+"\ndate:"+str(tmpRebuiltDate)+"\nLAT start:"+str(tmpFieldLatitudeStart)+" LON start:"+str(tmpFieldLongitudeStart)+"\nLAT end:"+str(tmpFieldLatitudeStop)+" LON end:"+str(tmpFieldLongitudeStop))
        
                    if executeVBarPlot == 1:
                        self.infoBLarea.insert(END, "\n\nVERTICAL PLOT executed for "+str(valueselectedScrollListSurvey))
                        ind = np.arange(len(sizesInt))  # the x locations for the groups
                        width = 0.35  # the width of the bars
                        fig, axbar = plt.subplots()
                        rects1 = axbar.bar(ind - width/2, sizesInt, width, color='SkyBlue')
                        # Add some text for labels, title and custom x-axbaris tick labels, etc.
                        axbar.set_ylabel('Number of items')
                        axbar.set_title("Vertical bar chart for survey "+str(valueselectedScrollListSurvey)+"\ndate:"+str(tmpRebuiltDate)+"\nLAT start:"+str(tmpFieldLatitudeStart)+" LON start:"+str(tmpFieldLongitudeStart)+"\nLAT end:"+str(tmpFieldLatitudeStop)+" LON end:"+str(tmpFieldLongitudeStop))
                        axbar.set_xticks(ind)
                        axbar.set_xticklabels(labels)
                        axbar.legend()
            
            
            def autolabel(rects, xpos='center'):
                """
                Attach a text label above each bar in *rects*, displaying its height.
            
                *xpos* indicates which side to place the text w.r.t. the center of
                the bar. It can be one of the following {'center', 'right', 'left'}.
                """
           
                xpos = xpos.lower()  # normalize the case of the parameter
                ha = {'center': 'center', 'right': 'left', 'left': 'right'}
                offset = {'center': 0, 'right': 0, 'left': 0}  # x_txt = x + w*off
            
                for rect in rects:
                    height = rect.get_height()
                    axbar.text(rect.get_x() + rect.get_width()*offset[xpos], 1.01*height, '{}'.format(height), ha='center', va='bottom')
            
            if executePiePlot == 1 or executeHBarPlot == 1 or executeVBarPlot == 1: 
                if executeVBarPlot == 1:
                    autolabel(rects1, 'center')
                    
                plt.xticks(rotation=90)
                plt.show()
                
           
        else:

            self.infoBLarea.insert(END, '\nThis survey does not exists!')







    def executePlotSF(self):

        selectedScrollListSurveySF=self.mylistScrollListSurveySF.curselection()
                
        executePiePlotSF=self.varPiePlotSF.get()
        executeHBarPlotSF=self.varHBarPlotSF.get()
        executeVBarPlotSF=self.varVBarPlotSF.get()
        
        
        if selectedScrollListSurveySF:
            
            for selectedSurveySF in selectedScrollListSurveySF:
                
                valueselectedScrollListSurveySF = self.mylistScrollListSurveySF.get(selectedSurveySF)
        
                
                labels = []
                sizes = []
                sizesInt = []
                tmpRebuiltDate = ''
                tmpFieldLatitudeStart = ''
                tmpFieldLatitudeStop = ''
                tmpFieldLongitudeStart = ''
                tmpFieldLongitudeStop = ''


                ShowMapSurveySF=self.varMapPlotSF.get()
                ShowMapSurveyEndSF=self.varMapPlotEndSF.get()
        
                tempWhichNameOutputFilePlotSF = str(self.entryInfoInputFilePlotSurveySF.get())
                if os.path.exists(tempWhichNameOutputFilePlotSF):
        
        
                    self.infoSFarea.insert(END, "\nPARAMETERS LIST FOR SURVEY "+str(valueselectedScrollListSurveySF)+"\nSTART LIST ------------------------------------------------------------------------------------------")

                    countRow=0
                    checkMetaData=0
        
                    with open(tempWhichNameOutputFilePlotSF, 'r') as csvFile:
                        if self.varRadioOutputCSV.get() == 1:
                            reader = csv.reader(csvFile, dialect="excel-tab")
                            self.infoSFarea.insert(END, '\nCSV delimiter is: TAB')
                        else:
                            reader = csv.reader(csvFile, delimiter=',')
                            self.infoSFarea.insert(END, '\nCSV delimiter is: COMMA')
                        
                        for row in reader:
                            
                            tmpInputSurveyPlotValueSF = str(row[9])
        
                            if tmpInputSurveyPlotValueSF == self.SurveyNamesListSF[int(selectedSurveySF)] and countRow>=1:
                               
                                if row[31] != '':
                                    
                                    tmpFielNoItemValuePlot=float(row[31])
                
                                    if tmpFielNoItemValuePlot > 0:
                                        labels.append(str(row[23]))                   
                                        sizes.append(str(row[31])) 
                                        sizesInt.append(float(row[31])) 
                                        self.infoSFarea.insert(END, '\n'+str(row[23])+' N. of items: '+str(row[31]))



                                if countRow >= 1 and checkMetaData == 0:
                                    
                                    tmpFieldLatitudeStart=str(row[12])
                                    tmpFieldLatitudeStop=str(row[14])
                                    tmpFieldLongitudeStart=str(row[13])
                                    tmpFieldLongitudeStop=str(row[15])
                                    tmpRebuiltDate=str(row[3])

                                    self.infoSFarea.insert(END, "\nSTART MAPS LINKS ------------------------------------------------------------------------------------")
                                    self.infoSFarea.insert(END, "\nUse the following link to map the start of the survey: http://www.google.com/maps/place/"+tmpFieldLatitudeStart+','+tmpFieldLongitudeStart)
                                    self.infoSFarea.insert(END, "\nUse the following link to map the end of the survey: http://www.google.com/maps/place/"+tmpFieldLatitudeStop+','+tmpFieldLongitudeStop)
                                    self.infoSFarea.insert(END, "\nEND MAPS LINKS ------------------------------------------------------------------------------------")

                                    checkMetaData = 1

                                    if ShowMapSurveySF == 1:

                                        webbrowser.open('http://www.google.com/maps/place/'+tmpFieldLatitudeStart+','+tmpFieldLongitudeStart)

                                    if ShowMapSurveyEndSF == 1:

                                        webbrowser.open('http://www.google.com/maps/place/'+tmpFieldLatitudeStop+','+tmpFieldLongitudeStop)
        
        

                            countRow += 1
        
                    csvFile.close()
        
                    self.infoSFarea.insert(END, "\nEND LIST ------------------------------------------------------------------------------------------")

  
        
                    if executePiePlotSF == 1: 
                        self.infoBLarea.insert(END, "\n\nPIE PLOT executed for "+str(valueselectedScrollListSurveySF))
                        fig1, ax1 = plt.subplots()
                        ax1.pie(sizes, labels=None, autopct='%1.1f%%', pctdistance=0.6, shadow=True, startangle=90, radius=1.5)
                        ax1.set_title("Pie plot for survey "+str(valueselectedScrollListSurveySF)+"\ndate:"+str(tmpRebuiltDate)+"\nLAT start:"+str(tmpFieldLatitudeStart)+" LON start:"+str(tmpFieldLongitudeStart)+"\nLAT end:"+str(tmpFieldLatitudeStop)+" LON end:"+str(tmpFieldLongitudeStop))
                        legend = ax1.legend(labels=labels, loc='upper right', shadow=True, fontsize='x-small')
                        # Put a nicer background color on the legend.
                        legend.get_frame().set_facecolor('#ffffff')
                        ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        
                    if executeHBarPlotSF == 1:
                        self.infoBLarea.insert(END, "\n\nHORIZONTAL PLOT executed for "+str(valueselectedScrollListSurveySF))
                        fig, ax2 = plt.subplots()
                        y_pos = np.arange(len(labels))
                        performance = sizesInt
                        error = 0
                        ax2.barh(y_pos, performance, xerr=error, align='center')
                        ax2.set_yticks(y_pos)
                        ax2.set_yticklabels(labels)
                        ax2.invert_yaxis()  # labels read top-to-bottom
                        ax2.set_xlabel('Number of items')
                        ax2.set_title("Horizontal bar chart for survey "+str(valueselectedScrollListSurveySF)+"\ndate:"+str(tmpRebuiltDate)+"\nLAT start:"+str(tmpFieldLatitudeStart)+" LON start:"+str(tmpFieldLongitudeStart)+"\nLAT end:"+str(tmpFieldLatitudeStop)+" LON end:"+str(tmpFieldLongitudeStop))
        
                    if executeVBarPlotSF == 1:
                        self.infoBLarea.insert(END, "\n\nVERTICAL PLOT executed for "+str(valueselectedScrollListSurveySF))
                        ind = np.arange(len(sizesInt))  # the x locations for the groups
                        width = 0.35  # the width of the bars
                        fig, axbar = plt.subplots()
                        rects1 = axbar.bar(ind - width/2, sizesInt, width, color='SkyBlue')
                        # Add some text for labels, title and custom x-axbaris tick labels, etc.
                        axbar.set_ylabel('Number of items')
                        axbar.set_title("Vertical bar chart for survey "+str(valueselectedScrollListSurveySF)+"\ndate:"+str(tmpRebuiltDate)+"\nLAT start:"+str(tmpFieldLatitudeStart)+" LON start:"+str(tmpFieldLongitudeStart)+"\nLAT end:"+str(tmpFieldLatitudeStop)+" LON end:"+str(tmpFieldLongitudeStop))
                        axbar.set_xticks(ind)
                        axbar.set_xticklabels(labels)
                        axbar.legend()
            
            
            def autolabelSF(rects, xpos='center'):
                """
                Attach a text label above each bar in *rects*, displaying its height.
            
                *xpos* indicates which side to place the text w.r.t. the center of
                the bar. It can be one of the following {'center', 'right', 'left'}.
                """
           
                xpos = xpos.lower()  # normalize the case of the parameter
                ha = {'center': 'center', 'right': 'left', 'left': 'right'}
                offset = {'center': 0, 'right': 0, 'left': 0}  # x_txt = x + w*off
            
                for rect in rects:
                    height = rect.get_height()
                    axbar.text(rect.get_x() + rect.get_width()*offset[xpos], 1.01*height, '{}'.format(height), ha='center', va='bottom')
            
            if executePiePlotSF == 1 or executeHBarPlotSF == 1 or executeVBarPlotSF == 1: 
                if executeVBarPlotSF == 1:
                    autolabelSF(rects1, 'center')
                    
                plt.xticks(rotation=90)
                plt.show()
                
           
        else:

            self.infoSFarea.insert(END, '\nThis survey does not exists!')





    def executePlotParamsSF(self):

        selectedScrollListParamsSF=self.mylistScrollListSurveyParamsSF.curselection()


        if selectedScrollListParamsSF:

            for selectedParamsSF in selectedScrollListParamsSF:

                valueselectedScrollListParamsSF = self.mylistScrollListSurveyParamsSF.get(selectedParamsSF)

                tempWhichNameOutputFilePlotParamsSF = str(self.entryInfoInputFilePlotParamsSF.get())
                if os.path.exists(tempWhichNameOutputFilePlotParamsSF):

                    countRow=0
                    checkMetaData=0

                    executeplotScatter=self.varScatterPlotSF.get()
                    executeplotScatterD=self.varScatterDPlotSF.get()
                    executeplotLegenda=self.varScatterLegendaPlotSF.get()
                    executeplotCSVGoogleMaps=self.varScatterCoordPlotSF.get()

                    paramLAT = []
                    paramLON = []
                    paramVALUE = []

                    del paramLAT [:]
                    del paramLON [:]
                    del paramVALUE [:]

                    # When we need to create the Google Maps CSV file
                    if executeplotCSVGoogleMaps == 1:
                        GoogleMapsCSVName = str('GoogleMaps_'+time.strftime("%Y%m%d%H%M%S")+'_Param_labelled_'+valueselectedScrollListParamsSF+'.csv')
                        GoogleMapsCSVOutputFile = open(GoogleMapsCSVName,"a")
                        GoogleMapsCSVOutputFile.write("ParamName,value,lat,lon")

        
                    with open(tempWhichNameOutputFilePlotParamsSF, 'r') as csvFile:
                        if self.varRadioOutputCSV.get() == 1:
                            reader = csv.reader(csvFile, dialect="excel-tab")
                            self.infoSFarea.insert(END, '\nCSV delimiter is: TAB')
                        else:
                            reader = csv.reader(csvFile, delimiter=',')
                            self.infoSFarea.insert(END, '\nCSV delimiter is: COMMA')
                        
                        for row in reader:
                            
                            tmpParamsSurveyPlotNameSF = str(row[23])
        
                            if str(tmpParamsSurveyPlotNameSF) == str(valueselectedScrollListParamsSF) and countRow>=1:

                                if row[31] != '':
                                    tmpParamsValuePlot=float(row[31])
                                    tmpParamsLatString=str(row[12])
                                    tmpParamsLonString=str(row[13])
                                    if tmpParamsLatString != '' and tmpParamsLonString != '':
                                        tmpParamsLat=float(row[12])
                                        tmpParamsLon=float(row[13])
                                        paramLAT.append(float(row[12])) 
                                        paramLON.append(float(row[13]))
                                        paramVALUE.append(int(tmpParamsValuePlot))

                                        # When we need to fill the Google Maps CSV file
                                        if executeplotCSVGoogleMaps == 1:
                                            GoogleMapsCSVOutputFile.write('\n'+str(valueselectedScrollListParamsSF)+','+str(tmpParamsValuePlot)+','+str(row[12])+','+str(row[13]))


                            countRow += 1

                    # When we need to close the Google Maps CSV file
                    if executeplotCSVGoogleMaps == 1:
                        GoogleMapsCSVOutputFile.close()
                        self.infoSFarea.insert(END, '\nThe Google Maps CSV '+GoogleMapsCSVName+' has been saved!')        


                    csvFile.close()



                    if executeplotScatter == 1:
                        self.infoSFarea.insert(END, "\n\nSCATTER PLOT 2D executed for param labelled "+str(valueselectedScrollListParamsSF))
                        fig1, scat1 = plt.subplots()
                    
                        myScatter = scat1.scatter(paramLON, paramLAT, s=300, c=paramVALUE, alpha=0.5)

                        # produce a legend with the unique colors from the scatter
                        legend1 = scat1.legend(*myScatter.legend_elements(), loc="lower left", title="Number of items")
                        scat1.add_artist(legend1)
                    
                        if executeplotLegenda == 1:
                            # produce a legend with a cross section of sizes from the scatter
                            
                            scat1.grid(True)

                        scat1.set_title('SCATTER PLOT executed for param labelled '+str(valueselectedScrollListParamsSF)+'\n X-axis:Longitude, Y-axis:Latitude')




                    if executeplotScatterD == 1:
                        self.infoSFarea.insert(END, "\n\nSCATTER PLOT 3D executed for param labelled "+str(valueselectedScrollListParamsSF))
                        fig2=plt.figure()
                        scat2=fig2.add_subplot(121,projection='3d')
                        scat2.scatter(paramLON, paramLAT,paramVALUE,c=paramVALUE,s=60)

                        if executeplotLegenda == 1:
                            scat2.legend(*myScatter.legend_elements(), loc="upper right", title="Number of items")
                            scat2.view_init(elev=20., azim=-35)

                        scat2.set_title('3D SCATTER PLOT executed for param labelled '+str(valueselectedScrollListParamsSF)+'\n X-axis:Longitude, Y-axis:Latitude')


                    if executeplotScatter == 1 or executeplotScatterD == 1:
                        plt.show()



        else:

            self.infoSFarea.insert(END, '\nThis parameter does not exists!')


    #Define an input file for data.
    def OpenInputFile(self,):
        actualdirname = os.getcwd()
        self.name = askopenfilename(initialdir=actualdirname,
                               filetypes =(("Spreadsheets xls", "*.xls"),("Spreadsheets xlsx", "*.xlsx"),("All Files","*.*")),
                               title = "Choose a file."
                               )

        self.entryInfoInputFile.delete(0,END)
        self.entryInfoInputFile.insert(0,self.name)
        self.src = self.name
        self.book = xlrd.open_workbook(self.src)
        self.infoBLarea.insert(END, "\nBeach Litter Input File: "+str(self.name))



    #Define an input file for data.
    def OpenInputFileSF(self,):
        actualdirname = os.getcwd()
        self.name = askopenfilename(initialdir=actualdirname,
                               filetypes =(("Spreadsheets xls", "*.xls"),("Spreadsheets xlsx", "*.xlsx"),("All Files","*.*")),
                               title = "Choose a file."
                               )

        self.entryInfoInputFileSF.delete(0,END)
        self.entryInfoInputFileSF.insert(0,self.name)
        self.src = self.name
        self.bookSF = xlrd.open_workbook(self.src)
        self.infoSFarea.insert(END, "\nSea Floor Litter Input File: "+str(self.name))



    #Define an input file for data.
    def OpenInputFilePlotSurvey(self,):
        actualdirname = os.getcwd()
        self.name = askopenfilename(initialdir=actualdirname,
                               filetypes =(("Spreadsheets xls", "*.xls"),("Spreadsheets xlsx", "*.xlsx"),("All Files","*.*")),
                               title = "Choose a file."
                               )

        self.entryInfoInputFilePlotSurvey.delete(0,END)
        self.entryInfoInputFilePlotSurvey.insert(0,self.name)
        self.src = self.name
        self.infoBLarea.insert(END, "\nBeach Litter Survey Plot File: "+str(self.name))
        self.checkPlots()



    #Define an input file for data.
    def OpenInputFilePlotSurveySF(self,):
        actualdirname = os.getcwd()
        self.name = askopenfilename(initialdir=actualdirname,
                               filetypes =(("CSV document", "*.csv"),("All Files","*.*")),
                               title = "Choose a file."
                               )

        self.entryInfoInputFilePlotSurveySF.delete(0,END)
        self.entryInfoInputFilePlotSurveySF.insert(0,self.name)
        self.src = self.name
        self.infoSFarea.insert(END, "\nSea Floor Litter Survey Plot File: "+str(self.name))
        self.checkPlotsSF()



    #Define an input file for data.
    def OpenInputFilePlotParams(self,):
        actualdirname = os.getcwd()
        self.name = askopenfilename(initialdir=actualdirname,
                               filetypes =(("Spreadsheets xls", "*.xls"),("Spreadsheets xlsx", "*.xlsx"),("All Files","*.*")),
                               title = "Choose a file."
                               )

        self.entryInfoInputFilePlotParams.delete(0,END)
        self.entryInfoInputFilePlotParams.insert(0,self.name)
        self.src = self.name
        self.infoBLarea.insert(END, "\nBeach Litter Params Plot File: "+str(self.name))
        self.checkPlotsParams()




    #Define an input file for data.
    def OpenInputFilePlotParamsSF(self,):
        actualdirname = os.getcwd()
        self.name = askopenfilename(initialdir=actualdirname,
                               filetypes =(("CSV Document", "*.csv"),("All Files","*.*")),
                               title = "Choose a file."
                               )

        self.entryInfoInputFilePlotParamsSF.delete(0,END)
        self.entryInfoInputFilePlotParamsSF.insert(0,self.name)
        self.src = self.name
        self.infoSFarea.insert(END, "\nSea Floor Params Plot File: "+str(self.name))
        self.checkPlotsParamsSF()




    def SaveOutputFileXls(self,):
        files = [('Excel Document', '*.xls')] 
        file = asksaveasfile(filetypes = files, defaultextension = files)

        self.entryInfoOutputFile.config(state = NORMAL)
        self.entryInfoOutputFile.delete(0,END)
        self.entryInfoOutputFile.insert(0,file.name)
        self.entryInfoOutputFile.config(state = "readonly")

        self.createXlsOutput()




    def SaveOutputFileCsvSF(self,):
        files = [('CSV Document', '*.csv')] 
        file = asksaveasfile(filetypes = files, defaultextension = files)

        self.entryInfoOutputFileSF.config(state = NORMAL)
        self.entryInfoOutputFileSF.delete(0,END)
        self.entryInfoOutputFileSF.insert(0,file.name)
        self.entryInfoOutputFileSF.config(state = "readonly")

        self.createCSVOutputSF()



    def SaveOutputFileModel(self,):
        files = [('CSV Document', '*.csv')] 
        file = asksaveasfile(filetypes = files, defaultextension = files)

        self.entryInfoOutputModelFile.config(state = NORMAL)
        self.entryInfoOutputModelFile.delete(0,END)
        self.entryInfoOutputModelFile.insert(0,file.name)
        self.entryInfoOutputModelFile.config(state = "readonly")

        self.createXlsOutputModel()



    def SaveOutputFileModelSF(self,):
        files = [('CSV Document', '*.csv')] 
        file = asksaveasfile(filetypes = files, defaultextension = files)

        self.entryInfoOutputModelFileSF.config(state = NORMAL)
        self.entryInfoOutputModelFileSF.delete(0,END)
        self.entryInfoOutputModelFileSF.insert(0,file.name)
        self.entryInfoOutputModelFileSF.config(state = "readonly")

        self.createCSVOutputModelSF()


    #Define an input file for the model.
    def OpenModelInputFile(self,):
        actualdirname = os.getcwd()
        self.Modelname = askopenfilename(initialdir=actualdirname,
                               filetypes =(("CSV file", "*.csv"),("Text file", "*.txt"),("All Files","*.*")),
                               title = "Choose a file."
                               )
        print (self.Modelname)
        self.entryInfoModelInputFile.delete(0,END)
        self.entryInfoModelInputFile.insert(0,self.Modelname)

        self.loadModel()



    #Define an input file for the model.
    def OpenModelInputFileSF(self,):
        actualdirname = os.getcwd()
        self.Modelname = askopenfilename(initialdir=actualdirname,
                               filetypes =(("CSV file", "*.csv"),("Text file", "*.txt"),("All Files","*.*")),
                               title = "Choose a file."
                               )
        print (self.Modelname)
        self.entryInfoModelInputFileSF.delete(0,END)
        self.entryInfoModelInputFileSF.insert(0,self.Modelname)

        self.loadModelSF()



    #To make it work with PyInstaller (this is due to an update of PyInstaller)
    #this functions is used to define the right file path between two differents
    #enviroments: your PC and PInstaller wrapping
    def resource_path(self,relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
    
        return os.path.join(base_path, relative_path)


    #Define an input file for the LEGENDA.
    def SearchLegendaTermFile(self,):
        searchTerm=self.entrylegendaTerm.get()
        self.legendaarea.delete('1.0', END)
        

        executeLegendaSearch=self.varVocabBODClvLegenda.get()
        executeBODClvASearch=self.varVocabBODClvA.get()
        executeBODClvBSearch=self.varVocabBODClvB.get()
        executeBODClvCSearch=self.varVocabBODClvC.get()
        executeBODClvDSearch=self.varVocabBODClvD.get()
        executeBODClvESearch=self.varVocabBODClvE.get()
        executeBODClvFSearch=self.varVocabBODClvF.get()



        if executeLegendaSearch == 1:
            legendaTextUpper=''
            legendaCheckText=''
            inputLegenda = self.resource_path('legenda.txt')
            with open(inputLegenda, 'r') as le:
                while True:
                    legendaTextUpper = le.readline()
                    legendaCheckText = le.readline()
                    if str(searchTerm.upper()) in str(legendaTextUpper.upper()):

                        self.legendaarea.insert(END, str(legendaTextUpper.upper())+"\n")
                        self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
                    if legendaCheckText == '':
                        # We have reached the end of the file
                        break

        
        if executeBODClvASearch == 1:
            page_link = 'http://vocab.nerc.ac.uk/collection/H01/current/'+str(searchTerm.upper())
            self.legendaarea.insert(END, "---------------------SEARCH RESULT FOR BODC H01 VOCAB-------------------------------------\n")
            self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            page_response = requests.get(page_link, timeout=5)
            page_content = str(BeautifulSoup(page_response.content, "html.parser"))


            singlerow = page_content.split('</th>')

            dictLitterTxt = ''
            interestingText = []
            # this variable is necessary to extract the right portion
            pref='>Definition<'
            
            
            for i in range(0, len(singlerow)):
                if pref in str(singlerow[i]):
                    interestingText = singlerow[i+1].split('>')
                    preoutString=interestingText[1].split('<')
                    dictLitterTxt=str(preoutString[0])
                    self.legendaarea.insert(END, str(searchTerm.upper()+': '+dictLitterTxt.upper())+"\n")
                    self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            



        if executeBODClvBSearch == 1:
            page_link = 'http://vocab.nerc.ac.uk/collection/H02/current/'+str(searchTerm.upper())
            self.legendaarea.insert(END, "---------------------SEARCH RESULT FOR BODC H02 VOCAB-------------------------------------\n")
            self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            page_response = requests.get(page_link, timeout=5)
            page_content = str(BeautifulSoup(page_response.content, "html.parser"))
            #searchTerm='litter'
            singlerow = page_content.split('</th>')
            #dictLitter = []
            dictLitterTxt = ''
            interestingText = []
            # this variable is necessary to extract the right portion
            pref='>Definition<'
            
            
            for i in range(0, len(singlerow)):
                if pref in str(singlerow[i]):
                    interestingText = singlerow[i+1].split('>')
                    preoutString=interestingText[1].split('<')
                    dictLitterTxt=str(preoutString[0])
                    self.legendaarea.insert(END, str(searchTerm.upper()+': '+dictLitterTxt.upper())+"\n")
                    self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            



        if executeBODClvCSearch == 1:
            page_link = 'http://vocab.nerc.ac.uk/collection/H03/current/'+str(searchTerm.upper())
            self.legendaarea.insert(END, "---------------------SEARCH RESULT FOR BODC H03 VOCAB-------------------------------------\n")
            self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            page_response = requests.get(page_link, timeout=5)
            page_content = str(BeautifulSoup(page_response.content, "html.parser"))
            #searchTerm='litter'
            singlerow = page_content.split('</th>')
            #dictLitter = []
            dictLitterTxt = ''
            interestingText = []
            # this variable is necessary to extract the right portion
            pref='>Definition<'
            
            
            for i in range(0, len(singlerow)):
                if pref in str(singlerow[i]):
                    interestingText = singlerow[i+1].split('>')
                    preoutString=interestingText[1].split('<')
                    dictLitterTxt=str(preoutString[0])
                    self.legendaarea.insert(END, str(searchTerm.upper()+': '+dictLitterTxt.upper())+"\n")
                    self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            



        if executeBODClvDSearch == 1:
            page_link = 'http://vocab.nerc.ac.uk/collection/H04/current/'+str(searchTerm.upper())
            self.legendaarea.insert(END, "---------------------SEARCH RESULT FOR BODC H04 VOCAB-------------------------------------\n")
            self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            page_response = requests.get(page_link, timeout=5)
            page_content = str(BeautifulSoup(page_response.content, "html.parser"))
            #searchTerm='litter'
            singlerow = page_content.split('</th>')
            #dictLitter = []
            dictLitterTxt = ''
            interestingText = []
            # this variable is necessary to extract the right portion
            pref='>Definition<'
            
            
            for i in range(0, len(singlerow)):
                if pref in str(singlerow[i]):
                    interestingText = singlerow[i+1].split('>')
                    preoutString=interestingText[1].split('<')
                    dictLitterTxt=str(preoutString[0])
                    self.legendaarea.insert(END, str(searchTerm.upper()+': '+dictLitterTxt.upper())+"\n")
                    self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            



        if executeBODClvESearch == 1:
            page_link = 'http://vocab.nerc.ac.uk/collection/H05/current/'+str(searchTerm.upper())
            self.legendaarea.insert(END, "---------------------SEARCH RESULT FOR BODC H05 VOCAB-------------------------------------\n")
            self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            page_response = requests.get(page_link, timeout=5)
            page_content = str(BeautifulSoup(page_response.content, "html.parser"))
            
            singlerow = page_content.split('</th>')
            
            dictLitterTxt = ''
            interestingText = []
            # this variable is necessary to extract the right portion
            pref='>Definition<'
            
            
            for i in range(0, len(singlerow)):
                if pref in str(singlerow[i]):
                    interestingText = singlerow[i+1].split('>')
                    preoutString=interestingText[1].split('<')
                    dictLitterTxt=str(preoutString[0])
                    self.legendaarea.insert(END, str(searchTerm.upper()+': '+dictLitterTxt.upper())+"\n")
                    self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            





        if executeBODClvFSearch == 1:
            page_link = 'http://vocab.nerc.ac.uk/collection/P01/current/'+str(searchTerm.upper())
            self.legendaarea.insert(END, "---------------------SEARCH RESULT FOR BODC P01 VOCAB-------------------------------------\n")
            self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            page_response = requests.get(page_link, timeout=5)
            page_content = str(BeautifulSoup(page_response.content, "html.parser"))
            #searchTerm='litter'
            #print(page_content)
            singlerow = page_content.split('</th>')
            #dictLitter = []
            dictLitterTxt = ''
            interestingText = []
            # this variable is necessary to extract the right portion
            pref='>Definition<'
            
            for i in range(0, len(singlerow)):
                if pref in str(singlerow[i]):
                    interestingText = singlerow[i+1].split('>')
                    preoutString=interestingText[1].split('<')
                    dictLitterTxt=str(preoutString[0])
                    self.legendaarea.insert(END, str(searchTerm.upper()+': '+dictLitterTxt.upper())+"\n")
                    self.legendaarea.insert(END, "------------------------------------------------------------------------------------------\n")
            



    #Check if the output file already exist
    def checkPlots(self,):
        tempWhichNameOutputFile = str(self.entryInfoInputFilePlotSurvey.get())
        if os.path.exists(tempWhichNameOutputFile):

            self.survyes_book = xlrd.open_workbook(tempWhichNameOutputFile)
            self.all_surveys_work_sheet = self.survyes_book.sheet_by_index(3)
            del self.SurveyNamesList[:]
            allsurveys_current_row=1
            allsurveys_num_rows = self.all_surveys_work_sheet.nrows
  
            while allsurveys_current_row < allsurveys_num_rows:

                try:
                    tmpInputSurveyValue = str(self.all_surveys_work_sheet.cell_value(allsurveys_current_row, 0))
                except Exception as e:
                    print("Warning!", e, "occurred.")
                    self.infoBLarea.insert(END, '\nWarning! ', e, ' ocurred.')
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    print(exc_type, fname, exc_tb.tb_lineno)
                

                if tmpInputSurveyValue.index(tmpInputSurveyValue) == False:

                    self.SurveyNamesList.append(tmpInputSurveyValue)

                allsurveys_current_row += 1

            self.SurveyNamesList = list(set(self.SurveyNamesList))
            self.SurveyNamesList.sort()

            
            self.mylistScrollListSurvey.delete(0,'end')
            
            # Next we will loop through the list
            for c in range(len(self.SurveyNamesList)):

                self.mylistScrollListSurvey.insert(END, self.SurveyNamesList[c])
                self.mylistScrollListSurvey.grid(row=c, column=1,columnspan=10, sticky=W)

        else:

            self.infoBLarea.insert(END, '\nThis file does not exists!')





    #Check if the output file already exist
    def checkPlotsSF(self,):
        tempWhichNameOutputFileSF = str(self.entryInfoInputFilePlotSurveySF.get())
        if os.path.exists(tempWhichNameOutputFileSF):

            countRow=0
            del self.SurveyNamesListSF[:]
            with open(tempWhichNameOutputFileSF, 'r') as csvFile:
                if self.varRadioOutputCSV.get() == 1:
                    reader = csv.reader(csvFile, delimiter='\t')
                    self.infoSFarea.insert(END, '\nCSV delimiter is: TAB')
                else:
                    reader = csv.reader(csvFile, delimiter=',')
                    self.infoSFarea.insert(END, '\nCSV delimiter is: COMMA')
                
                for row in reader:
                                        
                    #We create a list with the uniques names of the surveys
                    try:
                        tmpInputSurveyValueSF = str(row[9])
                    except Exception as e:
                        print("Warning!", e, "occurred.")
                        self.infoBLarea.insert(END, '\nWarning! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)
                        

                    if tmpInputSurveyValueSF.index(tmpInputSurveyValueSF) == False and countRow>=1:

                        self.SurveyNamesListSF.append(tmpInputSurveyValueSF)

                    countRow += 1

            csvFile.close()

            self.SurveyNamesListSF = list(set(self.SurveyNamesListSF))
            self.SurveyNamesListSF.sort()


            self.mylistScrollListSurveySF.delete(0,'end')
            # Next we will loop through the list
            for c in range(len(self.SurveyNamesListSF)):

                self.mylistScrollListSurveySF.insert(END, self.SurveyNamesListSF[c])
                self.mylistScrollListSurveySF.grid(row=c+7, column=1,columnspan=10, sticky=W)
                

        else:

            self.infoSFarea.insert(END, '\nThis file does not exists!')


    #Check if the output file already exist
    def checkPlotsParams(self,):

        tempWhichNameOutputFileParams = str(self.entryInfoInputFilePlotParams.get())
        if os.path.exists(tempWhichNameOutputFileParams):

            self.survyes_bookParams = xlrd.open_workbook(tempWhichNameOutputFileParams)
            self.all_Params_work_sheet = self.survyes_bookParams.sheet_by_index(3)
            del self.ParamsNamesList[:]
            allParams_current_row=1
            allParams_num_rows = self.all_Params_work_sheet.nrows
  
            while allParams_current_row < allParams_num_rows:

                try:
                    tmpInputParamsValue = str(self.all_Params_work_sheet.cell_value(allParams_current_row, 2))
                    tmpParamsListValue = str(self.all_Params_work_sheet.cell_value(allParams_current_row, 5))
                except Exception as e:
                    print("Warning!", e, "occurred.")
                    self.infoBLarea.insert(END, '\nWarning! ', e, ' ocurred.')
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    print(exc_type, fname, exc_tb.tb_lineno)

                if tmpInputParamsValue.index(tmpInputParamsValue) == False and tmpParamsListValue != '0':

                    self.ParamsNamesList.append(tmpInputParamsValue)

                allParams_current_row += 1

            self.ParamsNamesList = list(set(self.ParamsNamesList))
            self.ParamsNamesList.sort()


            self.mylistScrollListSurveyParams.delete(0,'end')
            # Next we will loop through the list
            for c in range(len(self.ParamsNamesList)):

                self.mylistScrollListSurveyParams.insert(END, self.ParamsNamesList[c])
                self.mylistScrollListSurveyParams.grid(row=c, column=1,columnspan=10, sticky=W)

        else:

            self.infoBLarea.insert(END, '\nThis file does not exists!')



    #Check if the output file already exist
    def checkPlotsParamsSF(self,):

        tempWhichNameOutputFileParamsSF = str(self.entryInfoInputFilePlotParamsSF.get())
        if os.path.exists(tempWhichNameOutputFileParamsSF):

            countRow=0
            del self.ParamsNamesListSF[:]
            with open(tempWhichNameOutputFileParamsSF, 'r') as csvFile:
                if self.varRadioOutputCSV.get() == 1:
                    reader = csv.reader(csvFile, dialect="excel-tab")
                    self.infoSFarea.insert(END, '\nCSV delimiter is: TAB')
                else:
                    reader = csv.reader(csvFile, delimiter=',')
                    self.infoSFarea.insert(END, '\nCSV delimiter is: COMMA')
                
                for row in reader:
                    
                    try:
                        tmpInputParamsValueSF = str(row[23])
                        tmpParamsListValueSF = str(row[31])
                    except Exception as e:
                        print("Warning!", e, "occurred.")
                        self.infoBLarea.insert(END, '\nWarning! ', e, ' ocurred.')
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print(exc_type, fname, exc_tb.tb_lineno)

                    if tmpInputParamsValueSF.index(tmpInputParamsValueSF) == False and countRow>=1 and tmpParamsListValueSF != '0':

                        self.ParamsNamesListSF.append(tmpInputParamsValueSF)

                    countRow += 1

            csvFile.close()


            self.ParamsNamesListSF = list(set(self.ParamsNamesListSF))
            self.ParamsNamesListSF.sort()


            self.mylistScrollListSurveyParamsSF.delete(0,'end')
            # Next we will loop through the list
            for c in range(len(self.ParamsNamesListSF)):

                self.mylistScrollListSurveyParamsSF.insert(END, self.ParamsNamesListSF[c])
                self.mylistScrollListSurveyParamsSF.grid(row=c, column=1,columnspan=10, sticky=W)

        else:

            self.infoSFarea.insert(END, '\nThis file does not exists!')



    def loadModel(self,):

        # We import the Model
        tempInport=self.entryInfoModelInputFile.get()

        self.ParDescSheet.deselect()
        self.enWichSheetVars.set(int(0))
        self.enWichSheetRowVars.set(int(0))
        self.enWichSheetColVars.set(str(self.LETTERS_ARRAY[0]))
        self.enWichSheetNameColVars.set(str(self.LETTERS_ARRAY[0]))
        self.enWichSheetOriginalNameColVars.set(str(self.LETTERS_ARRAY[0]))


        with open(tempInport, 'r') as fp:
            while True:
                currentModelRow = fp.readline()
                # If the result is an empty string
                if currentModelRow == '':
                    # We have reached the end of the file
                    break

                ModelReadRow = currentModelRow.split(',')

                #SHEETS
                if ModelReadRow[0] == 'SHEETS':
                    if ModelReadRow[1] == '0':
                        self.entryInfoBeachesVars.set(int(ModelReadRow[2]))
                    if ModelReadRow[1] == '1':
                        self.entryInfoSurveysVars.set(int(ModelReadRow[2]))
                    if ModelReadRow[1] == '2':
                        self.entryInfoAnimalsVars.set(int(ModelReadRow[2]))
                    if ModelReadRow[1] == '3':
                        self.entryInfoLitterVars.set(int(ModelReadRow[2]))
                #BEACHES
                if ModelReadRow[0] == 'BEACHES':
                    self.entriesBeachesRowVars[int(ModelReadRow[2])].set(int(ModelReadRow[3])+1)
                    self.entriesBeachesColVars[int(ModelReadRow[2])].set(str(self.LETTERS_ARRAY[int(ModelReadRow[4])+1]))
                #SURVEYS
                if ModelReadRow[0] == 'SURVEYS':
                    self.entriesSurveysRowVars[int(ModelReadRow[2])].set(int(ModelReadRow[3])+1)
                    self.entriesSurveysColVars[int(ModelReadRow[2])].set(str(self.LETTERS_ARRAY[int(ModelReadRow[4])+1]))
                #ANIMALS
                if ModelReadRow[0] == 'ANIMALS':
                    self.entriesAnimalsRowVars[int(ModelReadRow[2])].set(int(ModelReadRow[3])+1)
                    self.entriesAnimalsColVars[int(ModelReadRow[2])].set(str(self.LETTERS_ARRAY[int(ModelReadRow[4])+1]))
                #LITTER
                if ModelReadRow[0] == 'LITTER':
                    self.entriesLitterRowVars[int(ModelReadRow[2])].set(int(ModelReadRow[3])+1)
                    self.entriesLitterColVars[int(ModelReadRow[2])].set(str(self.LETTERS_ARRAY[int(ModelReadRow[4])+1]))
                #ANOTHER_SHEET
                if ModelReadRow[0] == 'ANOTHER_SHEET':
                    self.ParDescSheet.select()
                    self.enWichSheetVars.set(int(ModelReadRow[2]))
                    self.enWichSheetRowVars.set(int(ModelReadRow[3]))
                    self.enWichSheetColVars.set(str(self.LETTERS_ARRAY[int(ModelReadRow[4])+1]))
                    self.enWichSheetNameColVars.set(str(self.LETTERS_ARRAY[int(ModelReadRow[5])+1]))
                    self.enWichSheetOriginalNameColVars.set(str(self.LETTERS_ARRAY[int(ModelReadRow[6])+1]))
                

        self.infoBLarea.insert(END, '\nThe model '+tempInport+' has been imported!')








    def loadModelSF(self,):

        # We import the Model
        tempInport=self.entryInfoModelInputFileSF.get()

        with open(tempInport, 'r') as fp:
            while True:
                currentModelRow = fp.readline()
                # If the result is an empty string
                if currentModelRow == '':
                    # We have reached the end of the file
                    break

                ModelReadRow = currentModelRow.split(',')

                #SHEETS
                if ModelReadRow[0] == 'SHEETS':
                    if ModelReadRow[1] == '0':
                        self.entryInfoSurveysVarsSF.set(int(ModelReadRow[2]))
                    if ModelReadRow[1] == '1':
                        self.entryInfoLitterVarsSF.set(int(ModelReadRow[2]))

                #SURVEYS
                if ModelReadRow[0] == 'SURVEYS':
                    self.entriesSurveysRowVarsSF[int(ModelReadRow[2])].set(int(ModelReadRow[3])+1)
                    self.entriesSurveysColVarsSF[int(ModelReadRow[2])].set(str(self.LETTERS_ARRAY[int(ModelReadRow[4])+1]))
                    
                #SURVEYS EXCEPTION
                if ModelReadRow[0] == 'SURVEYSEXCEPTION':
                    if int(ModelReadRow[1]) > 0:
                        self.R1Except.select()
                        print('Exception True '+str(ModelReadRow[1]))
                      
                #LITTER
                if ModelReadRow[0] == 'LITTER':
                    self.entriesLitterRowVarsSF[int(ModelReadRow[2])].set(int(ModelReadRow[3])+1)
                    self.entriesLitterColVarsSF[int(ModelReadRow[2])].set(str(self.LETTERS_ARRAY[int(ModelReadRow[4])+1]))

                

        self.infoSFarea.insert(END, '\nThe model '+tempInport+' has been imported!')





    def createXlsOutputModel(self,):
        
        #Check the pivot
        pivoting=''
        paramCodeRow=int(self.entriesLitterRow[2].get())
        paramValueRow=int(self.entriesLitterRow[5].get())

        if paramCodeRow < paramValueRow:
            pivoting=1
        else:
            pivoting=0

        # We start to save our preferences on a model
        tmpModelNameOut=str(self.entryInfoOutputModelFile.get())


        if os.path.exists(tmpModelNameOut):
            os.remove(tmpModelNameOut)


        if tmpModelNameOut != '':
                ModelName=str(self.entryInfoOutputModelFile.get())
        else:
                ModelName = str(time.strftime("%Y%m%d%H%M%S")+'.csv')

        ModelOutputFile = open(ModelName,"a")

        # We save inside the model
        ModelOutputFile.write("SHEETS,0,"+str(self.entryInfoBeaches.get()))
        ModelOutputFile.write("\nSHEETS,1,"+str(self.entryInfoSurveys.get()))
        ModelOutputFile.write("\nSHEETS,2,"+str(self.entryInfoAnimals.get()))
        ModelOutputFile.write("\nSHEETS,3,"+str(self.entryInfoLitter.get()))


        '''
        Read and write BEACHES
        '''
        for wichfield in range(42):

            mytext=self.FIELDBEACHES[int(wichfield)]
            tmpMyRow=''
            tmpMyCol=''

            tmpMyRow=int(self.entriesBeachesRow[int(wichfield)].get())-1
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesBeachesCol[int(wichfield)].get()))-1
            wichSheet=int(self.entryInfoBeaches.get())-1

            # We save inside the model
            ModelOutputFile.write("\nBEACHES,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        '''
        Read and write SURVEYS
        '''
        for wichfield in range(58):

            mytext=self.FIELDSURVEYS[int(wichfield)]
            tmpMyRow=''
            tmpMyCol=''

            tmpMyRow=int(self.entriesSurveysRow[int(wichfield)].get())-1
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesSurveysCol[int(wichfield)].get()))-1
            wichSheet=int(self.entryInfoSurveys.get())-1

            # We save inside the model
            ModelOutputFile.write("\nSURVEYS,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        '''
        Read and write ANIMALS
        '''
        for wichfield in range(7):

            mytext=self.FIELDANIMALS[int(wichfield)]
            tmpMyRow=''
            tmpMyCol=''

            tmpMyRow=int(self.entriesAnimalsRow[int(wichfield)].get())-1
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesAnimalsCol[int(wichfield)].get()))-1
            wichSheet=int(self.entryInfoAnimals.get())-1

            # We save inside the model
            ModelOutputFile.write("\nANIMALS,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        '''
        Read and write LITTER
        '''

        #THE SURVEYCODE
        wichfield=0
        mytext=self.FIELDLITTER[int(wichfield)]
        tmpMyRow=''
        tmpMyCol=''
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1
        wichSheet=int(self.entryInfoLitter.get())-1
        # We save inside the model
        ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        #THE REFERENCE LIST
        wichfield=1
        mytext=self.FIELDLITTER[int(wichfield)]
        tmpMyRow=''
        tmpMyCol=''
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1
        wichSheet=int(self.entryInfoLitter.get())-1
        # We save inside the model
        ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        #THE PARAMS
        wichfield=2
        mytext=self.FIELDLITTER[int(wichfield)]
        #The params description is in another sheet
        anotherSheet=self.var1ParDesc.get()
        tmpMyRow=''
        tmpMyCol=''
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1
        wichSheet=int(self.entryInfoLitter.get())-1
        # We save inside the model
        ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        #ITEM NAME
        wichfield=3
        mytext=self.FIELDLITTER[int(wichfield)]
        tmpMyRow=''
        tmpMyCol=''
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1
        wichSheet=int(self.entryInfoLitter.get())-1
        # We save inside the model
        ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        #ORIGINAL NAME
        wichfield=4
        mytext=self.FIELDLITTER[int(wichfield)]
        tmpMyRow=''
        tmpMyCol=''
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1
        wichSheet=int(self.entryInfoLitter.get())-1
        # We save inside the model
        ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        #THE VALUES
        wichfield=5
        mytext=self.FIELDLITTER[int(wichfield)]
        tmpMyRow=''
        tmpMyCol=''
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1
        wichSheet=int(self.entryInfoLitter.get())-1
        # We save inside the model
        ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        #THE NOTES
        wichfield=6
        mytext=self.FIELDLITTER[int(wichfield)]
        tmpMyRow=''
        tmpMyCol=''
        tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1
        wichSheet=int(self.entryInfoLitter.get())-1
        # We save inside the model
        ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))



        # If the params are in another sheet
        if anotherSheet == 1:

                ModelparamSheet=int(self.enWichSheet.get())
                ModeltmpParDescr_current_row=int(self.enWichSheetRow.get())
                ModeltmpParDescr_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetCol.get()))-1
                ModeltmpParDescrName_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetNameCol.get()))-1
                ModeltmpParDescrOriginalName_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetOriginalNameCol.get()))-1
                ModelOutputFile.write("\nANOTHER_SHEET,"+str(anotherSheet)+","+str(ModelparamSheet)+","+str(ModeltmpParDescr_current_row)+","+str(ModeltmpParDescr_current_col)+","+str(ModeltmpParDescrName_current_col)+","+str(ModeltmpParDescrOriginalName_current_col))

        # Finally we close the model
        ModelOutputFile.close()
        self.infoBLarea.insert(END, '\nThe model labelled '+ModelName+' has been saved!')





    def createCSVOutputModelSF(self,):
        

        # We start to save our preferences on a model
        tmpModelNameOut=str(self.entryInfoOutputModelFileSF.get())


        if tmpModelNameOut != '':
                ModelName=str(self.entryInfoOutputModelFileSF.get())
        else:
                ModelName = str(time.strftime("%Y%m%d%H%M%S")+'.csv')

        ModelOutputFile = open(ModelName,"a")

        # We save inside the model
        ModelOutputFile.write("SHEETS,0,"+str(self.entryInfoSurveysSF.get()))
        ModelOutputFile.write("\nSHEETS,1,"+str(self.entryInfoLitterSF.get()))



        '''
        Read and write SURVEYS
        '''
        for wichfield in range(24):

            mytext=self.FIELDSURVEYSSEAFLOOR[int(wichfield)]
            tmpMyRow=''
            tmpMyCol=''

            tmpMyRow=int(self.entriesSurveysRowSF[int(wichfield)].get())-1
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesSurveysColSF[int(wichfield)].get()))-1
            wichSheet=int(self.entryInfoSurveysSF.get())-1

            # We save inside the model
            ModelOutputFile.write("\nSURVEYS,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))


        '''
        Read and write SURVEYSEXCEPTION
        '''
        writeException=self.varCheckException.get()
        if writeException == 1 :
            ModelOutputFile.write("\nSURVEYSEXCEPTION,1") 
        else:
            ModelOutputFile.write("\nSURVEYSEXCEPTION,0")

        '''
        Read and write LITTER
        '''
        for wichfield in range(13):


            mytext=self.FIELDLITTERSEAFLOOR[int(wichfield)]
            tmpMyRow=''
            tmpMyCol=''
            tmpMyRow=int(self.entriesLitterRowSF[int(wichfield)].get())-1
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterColSF[int(wichfield)].get()))-1
            wichSheet=int(self.entryInfoLitterSF.get())-1
            # We save inside the model
            ModelOutputFile.write("\nLITTER,"+str(wichSheet)+","+str(wichfield)+","+str(tmpMyRow)+","+str(tmpMyCol))



        # Finally we close the model
        ModelOutputFile.close()
        self.infoBLarea.insert(END, '\nThe model labelled '+ModelName+' has been saved!')






    def createCSVOutputSF(self,):

        whichNameOutputFile = str(self.entryInfoOutputFileSF.get())

        if whichNameOutputFile != '':
                whichNameOutputFile=str(self.entryInfoOutputFileSF.get())
        else:
                whichNameOutputFile = str(time.strftime("%Y%m%d%H%M%S")+'.csv')

        inputFileSF=self.entryInfoInputFileSF.get()
        

        '''
        Read and write SEA FLOOR
        '''
        wichSheetLitter=int(self.entryInfoLitterSF.get())-1
        wichSheetSurvey=int(self.entryInfoSurveysSF.get())-1

        allArrayInString = ''

        allSurveysMetadataFilteredUp = []
        allSurveysMetadataFiltered = []
        allLitterDataFiltered = []


        self.input_litter_work_sheetSF = self.bookSF.sheet_by_index(wichSheetLitter)
        self.input_survey_work_sheetSF = self.bookSF.sheet_by_index(wichSheetSurvey)
        litterSF_num_rows = self.input_litter_work_sheetSF.nrows
        surveysSF_num_rows = self.input_survey_work_sheetSF.nrows
              
        
        allSurveysMetadata = [[0 for i in range(22)] for i in range(surveysSF_num_rows)]
        allLitterData = [[0 for i in range(13)] for i in range(litterSF_num_rows)]
        allSurveysException = [''] * surveysSF_num_rows
        print(str('surveysSF_num_rows: '+str(len(allSurveysException))))

        if inputFileSF != '':
            
            '''
            New section to manage  timestamp (shot_timestamp) and 
            haul duration (haul_dur) survey tab area
            '''
            checException=self.varCheckException.get()
            if checException == 1 :
                print('We must extract field 22 amd 23.... Shoot timestamp and Haul Duration from Survey sheet')
                
            
            for wichfield in range(22):

                tmpMyRow=int(self.entriesSurveysRowSF[int(wichfield)].get())-1
                tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesSurveysColSF[int(wichfield)].get()))-1
    
                stringtmpMyRow=str(self.entriesSurveysRowSF[int(wichfield)].get())
                stringtmpMyCol=str(self.LETTERS_ARRAY.index(self.entriesSurveysColSF[int(wichfield)].get()))

                if wichSheetSurvey != '':

                    surveysSF_current_row=tmpMyRow

                    while surveysSF_current_row < surveysSF_num_rows:


                        if tmpMyCol < 0:
                            actualvalue=''
                        else:
                            actualvalue=''
                            if wichfield == 3: # Need to manage the date
                                tmpValue_as_datetime=self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol)
                                if tmpValue_as_datetime != '':

                                    if isinstance(tmpValue_as_datetime, datetime.datetime):
                                        try:
                                            actualvalue = datetime.datetime(*xlrd.xldate_as_tuple(tmpValue_as_datetime, self.bookSF.datemode))
                                        except Exception as e:
                                            print("WARNING!", e, "occurred.")
                                            self.infoSFarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                        
                                    else:

                                        try:
                                            actualvalue=str(self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol))
                                        except Exception as e:
                                            print("WARNING!", e, "occurred.")
                                            self.infoSFarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                        
                                else:
                                    actualvalue=''
                            else:
                                try:
                                    if wichfield in (1,7,8,10,11,17,21) and self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol) != '': # Need to manage the integer values (we have to cast them to delete the decimal)
                                        actualvalue=str(int(self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol)))
                                    elif wichfield in (12,13,14,15,16,18,19,20) and self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol) != '': # Need to manage the float values
                                        actualvalue=str(float(self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol)))
                                    else:
                                        actualvalue=str(self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol))
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoSFarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                    exc_type, exc_obj, exc_tb = sys.exc_info()
                                    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                                    print(exc_type, fname, exc_tb.tb_lineno)
                                    print(str(self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol)))

                                    
                                    
                        '''
                        New section to manage  timestamp (shot_timestamp) and 
                        haul duration (haul_dur) survey tab area
                        '''            
                        if wichfield == 9:
                            if checException == 1 :
                                
                                #Shoot Timestamp 
                                tmpMyRowExcepST=int(self.entriesSurveysRowSF[int(22)].get())-1+surveysSF_current_row-1
                                tmpMyColExcepST=int(self.LETTERS_ARRAY.index(self.entriesSurveysColSF[int(22)].get()))-1
                                TMPShootTimestamp=str(self.input_survey_work_sheetSF.cell_value(tmpMyRowExcepST, tmpMyColExcepST))
                                #Haul Duration
                                tmpMyRowExcepHD=int(self.entriesSurveysRowSF[int(23)].get())-1+surveysSF_current_row-1
                                tmpMyColExcepHD=int(self.LETTERS_ARRAY.index(self.entriesSurveysColSF[int(23)].get()))-1
                                TMPHaulDuration=str(int(self.input_survey_work_sheetSF.cell_value(tmpMyRowExcepHD, tmpMyColExcepHD)))
                                #This list contains for each row: SurveyName,Shoot Timestamp,Haul Duration

                                if self.varRadioOutputCSV.get() == 1:
                                    allSurveysException[surveysSF_current_row]=str(self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol)+'\t'+TMPShootTimestamp+'\t'+TMPHaulDuration)
                                else:
                                    allSurveysException[surveysSF_current_row]=str(self.input_survey_work_sheetSF.cell_value(surveysSF_current_row, tmpMyCol)+','+TMPShootTimestamp+','+TMPHaulDuration)

                                
                                
                        allSurveysMetadata[surveysSF_current_row][wichfield]=actualvalue

                        surveysSF_current_row += 1



            for wichfield in range(13):

                tmpMyRowLitter=int(self.entriesLitterRowSF[int(wichfield)].get())-1
                tmpMyColLitter=int(self.LETTERS_ARRAY.index(self.entriesLitterColSF[int(wichfield)].get()))-1
    
                stringtmpMyRowLitter=str(self.entriesLitterRowSF[int(wichfield)].get())
                stringtmpMyColLitter=str(self.LETTERS_ARRAY.index(self.entriesLitterColSF[int(wichfield)].get()))

                if wichSheetLitter != '':

                    litterSF_current_row=tmpMyRowLitter

                    while litterSF_current_row < litterSF_num_rows:


                        if tmpMyColLitter < 0:
                            actualvalueLitter=''
                        else:
                                
                            actualvalueLitter=''

                            if wichfield in (9,11) and self.input_litter_work_sheetSF.cell_value(litterSF_current_row, tmpMyColLitter)!='': # Need to manage the integer values (we have to cast them to delete the decimal)
                                        
                                actualvalueLitter=str(int(self.input_litter_work_sheetSF.cell_value(litterSF_current_row, tmpMyColLitter)))
                            
                            else:
                               
                                actualvalueLitter=str(self.input_litter_work_sheetSF.cell_value(litterSF_current_row, tmpMyColLitter))
                        
                        allLitterData[litterSF_current_row][wichfield]=str(actualvalueLitter)

                        litterSF_current_row += 1





        # We have to change from a bi-dimensional list to a simple one dimension for Litter
        for tmpMyrowLitter in range (len(allLitterData)):
            if allLitterData[tmpMyrowLitter][0] != '' and allLitterData[tmpMyrowLitter][0] != 0: # We check the LTREF field (LitterReference for params)

                
                if self.varRadioOutputCSV.get() == 1:
                    allLitterArrayInString='\t'.join(map(str, allLitterData[tmpMyrowLitter]))
                else:
                    allLitterArrayInString=','.join(map(str, allLitterData[tmpMyrowLitter]))
                allLitterDataFiltered.append(allLitterArrayInString)
        
        allLitterDataFiltered = list(set(allLitterDataFiltered))

        
        if checException == 1 :
            tmpMyrowExceptions = 1
            while tmpMyrowExceptions < len(allSurveysException):
                # explode the exceptions values
                if self.varRadioOutputCSV.get() == 1:
                    mySurveyNameException = allSurveysException[tmpMyrowExceptions].split('\t')
                else:
                    mySurveyNameException = allSurveysException[tmpMyrowExceptions].split(',')
                

                for tmpMyrowLitterFilterExp in range (len(allLitterDataFiltered)):
                        
                        
                    # if 'IN' we explode this field of the list and then we substitute the values
                    if self.varRadioOutputCSV.get() == 1:
                        OpenedAllLitterDataFiltered = allLitterDataFiltered[tmpMyrowLitterFilterExp].split('\t')
                    else:
                        OpenedAllLitterDataFiltered = allLitterDataFiltered[tmpMyrowLitterFilterExp].split(',')

                    if mySurveyNameException[0] == OpenedAllLitterDataFiltered[12]:

                        OpenedAllLitterDataFiltered[10]=mySurveyNameException[1]

                        OpenedAllLitterDataFiltered[11]=mySurveyNameException[2]
                        # at last, we rebuild the correct shape of this field of the list

                        
                        if self.varRadioOutputCSV.get() == 1:
                            allLitterDataFiltered[tmpMyrowLitterFilterExp]='\t'.join(str(e) for e in OpenedAllLitterDataFiltered)
                        else:
                            allLitterDataFiltered[tmpMyrowLitterFilterExp]=','.join(str(e) for e in OpenedAllLitterDataFiltered)
                       

                    
                tmpMyrowExceptions = tmpMyrowExceptions+1
            
        

        # We have to change from a bi-dimensional list to a simple one dimension for Surveys
        for tmpMyrow in range (len(allSurveysMetadata)):
            if allSurveysMetadata[tmpMyrow][9] != '' and allSurveysMetadata[tmpMyrow][9] != 0: # We check the station name field (if exist then proceed)

                
                if self.varRadioOutputCSV.get() == 1:
                    allArrayInString='\t'.join(map(str, allSurveysMetadata[tmpMyrow]))
                else:
                    allArrayInString=','.join(map(str, allSurveysMetadata[tmpMyrow]))
                allSurveysMetadataFiltered.append(allArrayInString)




        # All duplicated surveys must be merged
        allSurveysMetadataFiltered = list(set(allSurveysMetadataFiltered))

        # We open the Sea Floor CSV outpu tfile
        CSVOutputFileSF = open(whichNameOutputFile,"a")
        # We save inside the Sea Floor CSV output file the labels
        # NEW FIELDS GroundSpeed,WingSpread,DoorSpread,WarpLength
        if self.varRadioOutputCSV.get() == 1:
            CSVOutputFileSF.write('SurveyName\tProjectCode\tDataPolicy\tDate\tShip\tGear\tCountry\tOriginator\tCollator\tStNo\tHaulNo\tCoordRefSys\tShootLat\tShootLong\tHaulLat\tHaulLong\tDepth\tDistance\tGroundSpeed\tWingSpread\tDoorSpread\tWarpLength\tLTREF\tPARAM\tLTSZC\tLTSRC\tTYPPL\tLTPRP\tUnitWgt\tLT_Weight\tUnitItem\tLT_Items\tShot_timestamp\tHaulDur')
        else:
            CSVOutputFileSF.write('SurveyName,ProjectCode,DataPolicy,Date,Ship,Gear,Country,Originator,Collator,StNo,HaulNo,CoordRefSys,ShootLat,ShootLong,HaulLat,HaulLong,Depth,Distance,GroundSpeed,WingSpread,DoorSpread,WarpLength,LTREF,PARAM,LTSZC,LTSRC,TYPPL,LTPRP,UnitWgt,LT_Weight,UnitItem,LT_Items,Shot_timestamp,HaulDur')

        print(str(len(allSurveysMetadataFiltered))+" Lunghezza allSurveysMetadataFiltered")
        print(str(len(allLitterDataFiltered))+" Lunghezza allLitterDataFiltered")

        for tmpMyrowFilter in range (len(allSurveysMetadataFiltered)):
            if self.varRadioOutputCSV.get() == 1:
                mySurveyName = allSurveysMetadataFiltered[tmpMyrowFilter].split('\t')  #mySurveyName[9] is the station name
            else:
                mySurveyName = allSurveysMetadataFiltered[tmpMyrowFilter].split(',')
            for tmpMyrowLitterFilter in range (len(allLitterDataFiltered)):
                # If the station name exist in the litter's row the add surveys metadata
                if self.varRadioOutputCSV.get() == 1:
                    checkMyStationmName = allLitterDataFiltered[tmpMyrowLitterFilter].split('\t')
                else:
                    checkMyStationmName = allLitterDataFiltered[tmpMyrowLitterFilter].split(',')
                
                if mySurveyName[9] == checkMyStationmName[12]:

                    if self.varRadioOutputCSV.get() == 1:
                        tmpReplaceStationName=str('\t'+mySurveyName[9])
                    else:
                        tmpReplaceStationName=str(','+mySurveyName[9])
                    # We save inside the Sea Floor CSV output file the data
                    # We have to delete the last field, used only to have a match with the station name
                    if self.varRadioOutputCSV.get() == 1:
                        outputrowCSV='\n'+str(allSurveysMetadataFiltered[tmpMyrowFilter])+'\t'+str(allLitterDataFiltered[tmpMyrowLitterFilter].replace(tmpReplaceStationName,''))
                    else:
                        outputrowCSV='\n'+str(allSurveysMetadataFiltered[tmpMyrowFilter])+','+str(allLitterDataFiltered[tmpMyrowLitterFilter].replace(tmpReplaceStationName,''))
                    CSVOutputFileSF.write(outputrowCSV)



        #file_uno.close()
        # Finally we close the model
        CSVOutputFileSF.close()
        self.infoSFarea.insert(END, '\nThe Sea Floor CSV file labelled '+whichNameOutputFile+' has been saved!')




    def createXlsOutput(self,):

        whichNameOutputFile = str(self.entryInfoOutputFile.get())

        if whichNameOutputFile != '':
                whichNameOutputFile=str(self.entryInfoOutputFile.get())
        else:
                whichNameOutputFile = str(time.strftime("%Y%m%d%H%M%S")+'.xls')


        output_workbook = xlsxwriter.Workbook(whichNameOutputFile)

        #We define a format for each field that contains a date
        format_date = output_workbook.add_format()
        format_date.set_num_format('mm/dd/yyyy')

        beaches_output_worksheet = output_workbook.add_worksheet('Beaches')
        surveys_output_worksheet = output_workbook.add_worksheet('Surveys')
        animals_output_worksheet = output_workbook.add_worksheet('Animals')
        litter_output_worksheet = output_workbook.add_worksheet('Litter')


        '''
        We fill all the labels of each sheet
        '''
        '''
        BEACHES
        '''
        for b in range(42):

            beaches_output_worksheet.write(0, b, str(self.FIELDBEACHES[b]))

        '''
        SURVEYS
        '''
        for s in range(58):
            surveys_output_worksheet.write(0, s, str(self.FIELDSURVEYS[s]))

        '''
        ANIMALS
        '''
        for a in range(8):
            animals_output_worksheet.write(0, a, str(self.FIELDANIMALS[a]))

        '''
        LITTER
        '''
        for l in range(7):
            litter_output_worksheet.write(0, l, str(self.FIELDLITTER[l]))


        self.create=1



        #Check the pivot
        pivoting=''
        paramCodeRow=int(self.entriesLitterRow[2].get())
        paramValueRow=int(self.entriesLitterRow[5].get())

        if paramCodeRow < paramValueRow:
            pivoting=1
        else:
            pivoting=0


        # We start to save our preferences on a model
        tmpModelNameOut=str(self.entryInfoOutputModelFile.get())

        if tmpModelNameOut != '':
                ModelName=str(self.entryInfoOutputModelFile.get())
        else:
                ModelName = str(time.strftime("%Y%m%d%H%M%S")+'.csv')


        '''
        Read and write BEACHES
        '''
        '''
        We write temporary the BEACHES
        '''
        for wichfield in range(42):

            mytext=self.FIELDBEACHES[int(wichfield)]
            tmpMyRow=''
            tmpMyCol=''

            stringtmpMyRow=str(self.entriesBeachesRow[int(wichfield)].get())
            stringtmpMyCol=str(self.entriesBeachesCol[int(wichfield)].get())

            if stringtmpMyRow != '':
                tmpMyRow=int(self.entriesBeachesRow[int(wichfield)].get())-1
            if stringtmpMyCol != '':
                tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesBeachesCol[int(wichfield)].get()))-1

            wichSheet=int(self.entryInfoBeaches.get())-1


            # We save inside the model

            if tmpMyRow != '':
                #if tmpMyCol != '':
                if tmpMyCol >= 0:
                    if wichSheet != '':


  
                        self.input_beaches_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        beaches_current_row=tmpMyRow
                        beaches_num_rows = self.input_beaches_work_sheet.nrows
  
                        while beaches_current_row < beaches_num_rows:

                            if wichfield == 4 or wichfield == 21:

                                
                                try:
                                    tmpInputValue = self.input_beaches_work_sheet.cell_value(beaches_current_row, tmpMyCol)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                    
                                    
                                if tmpInputValue != '' and isinstance(tmpInputValue, datetime.datetime):

                                    try:
                                        tmpInputValue_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(tmpInputValue, self.book.datemode))
                                    except Exception as e:
                                        print("WARNING!", e, "occurred.")
                                        self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                    
                                    beaches_output_worksheet.write(beaches_current_row, wichfield, tmpInputValue_as_datetime, format_date)
                                else:
                                    beaches_output_worksheet.write(beaches_current_row, wichfield, tmpInputValue)
                            else:

                                
                                try:
                                    tmpInputValue = self.input_beaches_work_sheet.cell_value(beaches_current_row, tmpMyCol)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                    
                                    
                                beaches_output_worksheet.write(beaches_current_row, wichfield, tmpInputValue)

                            beaches_current_row += 1






        '''
        Read and write SURVEYS
        '''
        for wichfield in range(58):

            mytext=self.FIELDSURVEYS[int(wichfield)]

            tmpMyRow=''
            tmpMyCol=''
            stringtmpMyRow=str(self.entriesSurveysRow[int(wichfield)].get())
            stringtmpMyCol=str(self.entriesSurveysCol[int(wichfield)].get())

            if stringtmpMyRow != '':
                tmpMyRow=int(self.entriesSurveysRow[int(wichfield)].get())-1
            if stringtmpMyCol != '':

                tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesSurveysCol[int(wichfield)].get()))-1

            wichSheet=int(self.entryInfoSurveys.get())-1


            if tmpMyRow != '':
                if tmpMyCol >= 0:
                    if wichSheet != '':
  
                        self.input_surveys_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        surveys_current_row=tmpMyRow
                        surveys_num_rows = self.input_surveys_work_sheet.nrows
  
                        while surveys_current_row < surveys_num_rows:

                            if wichfield == 4 or wichfield == 50:

                                
                                try:
                                    tmpInputValue = self.input_surveys_work_sheet.cell_value(surveys_current_row, tmpMyCol)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                if tmpInputValue != '' and isinstance(tmpInputValue, datetime.datetime):
                                    try:
                                        tmpInputValue_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(tmpInputValue, self.book.datemode))
                                    except Exception as e:
                                        print("WARNING!", e, "occurred.")
                                        self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                    
                                    surveys_output_worksheet.write(surveys_current_row, wichfield, tmpInputValue_as_datetime, format_date)
                                else:
                                    surveys_output_worksheet.write(surveys_current_row, wichfield, tmpInputValue)
                            else:
                                
                                try:
                                    tmpInputValue = self.input_surveys_work_sheet.cell_value(surveys_current_row, tmpMyCol)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                surveys_output_worksheet.write(surveys_current_row, wichfield, tmpInputValue)

                            surveys_current_row += 1




        '''
        Read and write ANIMALS
        '''
        for wichfield in range(7):

            mytext=self.FIELDANIMALS[int(wichfield)]

            tmpMyRow=''
            tmpMyCol=''
            stringtmpMyRow=str(self.entriesAnimalsRow[int(wichfield)].get())
            stringtmpMyCol=str(self.entriesAnimalsCol[int(wichfield)].get())

            if stringtmpMyRow != '':
                tmpMyRow=int(self.entriesAnimalsRow[int(wichfield)].get())-1
            if stringtmpMyCol != '':

                tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesAnimalsCol[int(wichfield)].get()))-1

            wichSheet=int(self.entryInfoAnimals.get())-1

            if tmpMyRow != '':

                if tmpMyCol >= 0:
                    if wichSheet != '':
  
                        self.input_animals_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        animals_current_row=tmpMyRow
                        animals_num_rows = self.input_animals_work_sheet.nrows
  
                        while animals_current_row < animals_num_rows:
                            
                            try:
                                tmpInputValue = self.input_animals_work_sheet.cell_value(animals_current_row, tmpMyCol)
                            except Exception as e:
                                print("WARNING!", e, "occurred.")
                                self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                            
                            animals_output_worksheet.write(animals_current_row, wichfield, tmpInputValue)
                            animals_current_row += 1







        '''
        Read and write LITTER
        '''

        #THE PARAMS
        wichfield=2
        mytext=self.FIELDLITTER[int(wichfield)]

        #The params description is in another sheet
        anotherSheet=self.var1ParDesc.get()

        tmpMyRow=''
        tmpMyCol=''
        stringtmpMyRow=str(self.entriesLitterRow[int(wichfield)].get())
        stringtmpMyCol=str(self.entriesLitterCol[int(wichfield)].get())

        if stringtmpMyRow != '':
            tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        if stringtmpMyCol != '':
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1

        wichSheet=int(self.entryInfoLitter.get())-1

        tmpColPivot=1
        tmpRowPivot=1
        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':
                    if pivoting == 0:
  
                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_current_col=tmpMyCol
                        litter_num_cols = self.input_litter_work_sheet.ncols
  
                        while litter_current_row < litter_num_rows:

                            
                            try:
                                tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                            except Exception as e:
                                print("WARNING!", e, "occurred.")
                                self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                            
                            litter_output_worksheet.write(litter_current_row, 2, tmpInputValue)

                            #Here we add the param descriptions
                            #if in another sheet....
                            if anotherSheet == 1:
                                tmpParDescr_current_row=int(self.enWichSheetRow.get())
                                tmpParDescr_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetCol.get()))-1

                                tmpParDescrName_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetNameCol.get()))-1
                                tmpParDescrOriginalName_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetOriginalNameCol.get()))-1

                                paramSheet=int(self.enWichSheet.get())
                                self.input_ParDescr_work_sheet = self.book.sheet_by_index(paramSheet)
                                tmpParDescr_num_rows = self.input_ParDescr_work_sheet.nrows

                                while tmpParDescr_current_row < tmpParDescr_num_rows:
                                    
                                    try:
                                        ParDefCheckID = self.input_ParDescr_work_sheet.cell_value(tmpParDescr_current_row, tmpParDescr_current_col)
                                    except Exception as e:
                                        print("WARNING!", e, "occurred.")
                                        self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                    

                                    if ParDefCheckID == tmpInputValue:

                                        try:
                                            ParDefCheckIDName = self.input_ParDescr_work_sheet.cell_value(tmpParDescr_current_row, tmpParDescrName_current_col)
                                            ParDefCheckIDOriginalName = self.input_ParDescr_work_sheet.cell_value(tmpParDescr_current_row, tmpParDescrOriginalName_current_col)
                                        except Exception as e:
                                            print("WARNING!", e, "occurred.")
                                            self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                        

                                        litter_output_worksheet.write(litter_current_row, 3, ParDefCheckIDName)
                                        litter_output_worksheet.write(litter_current_row, 4, ParDefCheckIDOriginalName)

                                    tmpParDescr_current_row += 1

                                
                            #if in the same sheet
                            else:
                                tmpParDescrName_current_col=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[3].get()))-1
                                tmpParDescrOriginalName_current_col=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[4].get()))-1

                                try:
                                    tmpInputValueName = self.input_litter_work_sheet.cell_value(litter_current_row, tmpParDescrName_current_col)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                litter_output_worksheet.write(litter_current_row, 3, tmpInputValueName)

                                try:
                                    tmpInputValueOriginalName = self.input_litter_work_sheet.cell_value(litter_current_row, tmpParDescrOriginalName_current_col)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                litter_output_worksheet.write(litter_current_row, 4, tmpInputValueOriginalName)

                            litter_current_row += 1
                    else:

                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow #So, we don't touch the first line
                        #less 1 row, so we will not create an empty survey on Item sheet
                        litter_num_rows = int(self.input_litter_work_sheet.nrows)-1
                        litter_num_cols = self.input_litter_work_sheet.ncols
                        while litter_current_row < litter_num_rows:

                            litter_current_col=tmpMyCol

                            while litter_current_col < litter_num_cols:

                                
                                try:
                                    tmpInputValue = self.input_litter_work_sheet.cell_value(tmpMyRow, litter_current_col)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                    
                                litter_output_worksheet.write(tmpRowPivot, 2, tmpInputValue)


                                #Here we add the param descriptions
                                #if in another sheet....
                                if anotherSheet == 1:
                                    tmpParDescr_current_row=int(self.enWichSheetRow.get())
                                    tmpParDescr_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetCol.get()))-1
    
                                    tmpParDescrName_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetNameCol.get()))-1
                                    tmpParDescrOriginalName_current_col=int(self.LETTERS_ARRAY.index(self.enWichSheetOriginalNameCol.get()))-1
    
                                    paramSheet=int(self.enWichSheet.get())
                                    self.input_ParDescr_work_sheet = self.book.sheet_by_index(paramSheet)
                                    tmpParDescr_num_rows = self.input_ParDescr_work_sheet.nrows
    
                                    while tmpParDescr_current_row < tmpParDescr_num_rows:
                                        ParDefCheckID = self.input_ParDescr_work_sheet.cell_value(tmpParDescr_current_row, tmpParDescr_current_col)
    
                                        if ParDefCheckID == tmpInputValue:
    
                                            try:
                                                ParDefCheckIDName = self.input_ParDescr_work_sheet.cell_value(tmpParDescr_current_row, tmpParDescrName_current_col)
                                                ParDefCheckIDOriginalName = self.input_ParDescr_work_sheet.cell_value(tmpParDescr_current_row, tmpParDescrOriginalName_current_col)
                                            except Exception as e:
                                                print("WARNING!", e, "occurred.")
                                                self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                            
    
                                            litter_output_worksheet.write(tmpRowPivot, 3, ParDefCheckIDName)
                                            litter_output_worksheet.write(tmpRowPivot, 4, ParDefCheckIDOriginalName)
    
    
                                        tmpParDescr_current_row += 1
    
                                    
                                #if in the same sheet
                                else:
                                    if tmpInputValue != '':
                                        tmpParDescrName_current_row=int(self.entriesLitterRow[3].get())
                                        tmpParDescrOriginalName_current_row=int(self.entriesLitterRow[4].get())
    
                                        try:
                                            tmpInputValueName = self.input_litter_work_sheet.cell_value(tmpParDescrName_current_row, litter_current_col)
                                        except Exception as e:
                                            print("WARNING!", e, "occurred.")
                                            self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                        
                                        litter_output_worksheet.write(tmpRowPivot, 3, tmpInputValueName)
    
                                        
                                        try:
                                            tmpInputValueOriginalName = self.input_litter_work_sheet.cell_value(tmpParDescrOriginalName_current_row, litter_current_col)
                                        except Exception as e:
                                            print("WARNING!", e, "occurred.")
                                            self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                        
                                        litter_output_worksheet.write(tmpRowPivot, 4, tmpInputValueOriginalName)
    
    
                                litter_current_col += 1
                                tmpRowPivot += 1

                            litter_current_row += 1




        #THE VALUES
        wichfield=5
        mytext=self.FIELDLITTER[int(wichfield)]

        tmpMyRow=''
        tmpMyCol=''
        stringtmpMyRow=str(self.entriesLitterRow[int(wichfield)].get())
        stringtmpMyCol=str(self.entriesLitterCol[int(wichfield)].get())

        if stringtmpMyRow != '':
            tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        if stringtmpMyCol != '':
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1


        wichSheet=int(self.entryInfoLitter.get())-1


        tmpColPivot=1
        tmpRowPivot=1
        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':
                    if pivoting == 0:
  
                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_current_col=tmpMyCol
                        litter_num_cols = self.input_litter_work_sheet.ncols
  
                        while litter_current_row < litter_num_rows:

                            try:
                                tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                            except Exception as e:
                                print("WARNING!", e, "occurred.")
                                self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                            
                            litter_output_worksheet.write(litter_current_row, 5, tmpInputValue)
                            litter_current_row += 1
                    else:

                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_num_cols = self.input_litter_work_sheet.ncols

                        while litter_current_row < litter_num_rows:

                            litter_current_col=tmpMyCol

                            while litter_current_col < litter_num_cols:

                                
                                try:
                                    tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, litter_current_col)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                litter_output_worksheet.write(tmpRowPivot, 5, tmpInputValue)
                                litter_current_col += 1
                                tmpRowPivot += 1

                            litter_current_row += 1





        #This value is common for next fields
        tmpMyColValues=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1




        #THE SURVEYCODE
        wichfield=0
        mytext=self.FIELDLITTER[int(wichfield)]

        tmpMyRow=''
        tmpMyCol=''
        stringtmpMyRow=str(self.entriesLitterRow[int(wichfield)].get())
        stringtmpMyCol=str(self.entriesLitterCol[int(wichfield)].get())

        if stringtmpMyRow != '':
            tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        if stringtmpMyCol != '':
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1

        wichSheet=int(self.entryInfoLitter.get())-1

        tmpColPivot=1
        tmpRowPivot=1
        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':
                    if pivoting == 0:
  
                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_current_col=tmpMyCol
                        litter_num_cols = self.input_litter_work_sheet.ncols
  
                        while litter_current_row < litter_num_rows:

                            try:
                                tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                            except Exception as e:
                                print("WARNING!", e, "occurred.")
                                self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                            
                            litter_output_worksheet.write(litter_current_row, 0, tmpInputValue)
                            litter_current_row += 1

                    else:

                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_num_cols = self.input_litter_work_sheet.ncols

                        while litter_current_row < litter_num_rows:

                            litter_current_col=tmpMyColValues

                            while litter_current_col < litter_num_cols:

                                try:
                                    tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                litter_output_worksheet.write(tmpRowPivot, 0, tmpInputValue)
                                litter_current_col += 1
                                tmpRowPivot += 1

                            litter_current_row += 1



        #THE NOTES
        wichfield=6
        mytext=self.FIELDLITTER[int(wichfield)]

        tmpMyRow=''
        tmpMyCol=''
        stringtmpMyRow=str(self.entriesLitterRow[int(wichfield)].get())
        stringtmpMyCol=str(self.entriesLitterCol[int(wichfield)].get())

        if stringtmpMyRow != '':
            tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        if stringtmpMyCol != '':
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1

        wichSheet=int(self.entryInfoLitter.get())-1

        tmpColPivot=1
        tmpRowPivot=1
        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':
                    if pivoting == 0:
  
                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_current_col=tmpMyCol
                        litter_num_cols = self.input_litter_work_sheet.ncols
  
                        while litter_current_row < litter_num_rows:

                            try:
                                tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                            except Exception as e:
                                print("WARNING!", e, "occurred.")
                                self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                            
                            litter_output_worksheet.write(litter_current_row, 6, tmpInputValue)
                            litter_current_row += 1
                    else:

                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_num_cols = self.input_litter_work_sheet.ncols

                        while litter_current_row < litter_num_rows:

                            litter_current_col=tmpMyColValues

                            while litter_current_col < litter_num_cols:

                                try:
                                    tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                litter_output_worksheet.write(tmpRowPivot, 6, tmpInputValue)

                                litter_current_col += 1
                                tmpRowPivot += 1

                            litter_current_row += 1





        #THE REFERENCE LIST
        wichfield=1
        mytext=self.FIELDLITTER[int(wichfield)]

        tmpMyRow=''
        tmpMyCol=''
        stringtmpMyRow=str(self.entriesLitterRow[int(wichfield)].get())
        stringtmpMyCol=str(self.entriesLitterCol[int(wichfield)].get())


        if stringtmpMyRow != '':
            tmpMyRow=int(self.entriesLitterRow[int(wichfield)].get())-1
        if stringtmpMyCol != '':
            tmpMyCol=int(self.LETTERS_ARRAY.index(self.entriesLitterCol[int(wichfield)].get()))-1

        wichSheet=int(self.entryInfoLitter.get())-1




        tmpColPivot=1
        tmpRowPivot=1
        if tmpMyRow != '':
            if tmpMyCol >= 0:
                if wichSheet != '':
                    if pivoting == 0:
  
                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_current_col=tmpMyCol
                        litter_num_cols = self.input_litter_work_sheet.ncols
  
                        while litter_current_row < litter_num_rows:

                            try:
                                tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                            except Exception as e:
                                print("WARNING!", e, "occurred.")
                                self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                            
                            litter_output_worksheet.write(litter_current_row, 1, tmpInputValue)
                            litter_current_row += 1
                    else:

                        self.input_litter_work_sheet = self.book.sheet_by_index(wichSheet)
  
                        litter_current_row=tmpMyRow
                        litter_num_rows = self.input_litter_work_sheet.nrows

                        litter_num_cols = self.input_litter_work_sheet.ncols

                        while litter_current_row < litter_num_rows:

                            litter_current_col=tmpMyColValues
                            
                            while litter_current_col < litter_num_cols:

                                try:
                                    tmpInputValue = self.input_litter_work_sheet.cell_value(litter_current_row, tmpMyCol)
                                except Exception as e:
                                    print("WARNING!", e, "occurred.")
                                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                                
                                litter_output_worksheet.write(tmpRowPivot, 1, tmpInputValue)

                                litter_current_col += 1
                                tmpRowPivot += 1

                            litter_current_row += 1



        output_workbook.close()


        self.bookbis = xlrd.open_workbook(whichNameOutputFile)
        self.beaches_output_worksheetTemp = self.bookbis.sheet_by_index(0)

        '''
        Start with a list for BEACHES uniques
        We populate the list with the uniq names of the beaches
        '''
        beaches_current_row=0
        beaches_num_rows = self.beaches_output_worksheetTemp.nrows
        beaches_current_col=0
        beaches_num_cols= self.beaches_output_worksheetTemp.ncols
        
        
        arr = []
        for i in range(beaches_num_rows):
            arr.append([])
        
        while beaches_current_row < beaches_num_rows:
            beaches_current_col=0
            flag=0
            while beaches_current_col < beaches_num_cols:
                try:
                    InputValue = self.beaches_output_worksheetTemp.cell_value(beaches_current_row, beaches_current_col)
                except Exception as e:
                    print("WARNING!", e, "occurred.")
                    self.infoBLarea.insert(END, '\nWARNING! ', e, ' ocurred.')
                

                if beaches_current_col == 0:
                    for x in range(beaches_current_row):
                        if InputValue in arr[x]:
                            flag=1
        
                if flag == 0:
                    arr[beaches_current_row].append(InputValue)
        
                beaches_current_col += 1
        
            beaches_current_row += 1

        

        '''
        Now we can write the list with the uniques values
        '''

        self.infoBLarea.insert(END, '\nThe output XLS file labelled '+whichNameOutputFile+' has been created!')

        return


#To make it work with PyInstaller (this is due to an update of PyInstaller)
#this functions is used to define the right file path between two differents
#enviroments: your PC and PInstaller wrapping
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)        





root = Tk()
root.option_add('*Font', 'Verdana 8')
my_gui = MarineLitterManager(root)
root.geometry('750x920')
cols = 0
while cols < 20:
    root.rowconfigure(cols, weight=1)
    root.columnconfigure(cols, weight=1)
    cols += 1

root.mainloop()
