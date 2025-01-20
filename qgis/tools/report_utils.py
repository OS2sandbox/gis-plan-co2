# -*- coding: utf-8 -*-
"""
report_utils
---------
"""
#from ctypes import util
#from http.client import FOUND
from ctypes import util
from qgis.PyQt.QtWidgets import QApplication, QFileDialog
from qgis.core import QgsProject, QgsReport, QgsLayoutExporter, QgsMasterLayoutInterface, QgsReadWriteContext, QgsPrintLayout
from qgis.PyQt.QtXml import QDomDocument, QDomElement

import os
from . import utils

# ---------------------------------------------------------------------------------------------------------
#
# ---------------------------------------------------------------------------------------------------------
def ExportReport(parent, iface, oDialog, dictTranslate, oLayerBuildings, oLayerOpenSurfacesLines, oLayerOpenSurfacesAreas, oLayerSubareas, dict_surface_width):

    try:
        oDialog.lblMessage.setStyleSheet('color: black')

        if oDialog.cmbStudyPeriod.currentText() == '' or int(oDialog.cmbStudyPeriod.currentText()) == 0:
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return

        #Det kontrolleres om der er ændringer i projektet 
        projectInstance = QgsProject.instance()
        projPath = projectInstance.fileName()
        if projPath == '':
            utils.Message('Der kan ikke indlægges en rapport i projektet, da projektet er ikke gemt i en fil.')
            return

        oDialog.btnExportReport.setEnabled(False)
        oDialog.lblMessage.setText('Danner rapport. Vent venligst...')
        QApplication.processEvents() 
        
        #Det kontrolleres om den seneste version af rapporten er oprettet i projektet
        #Navnet på den seneste rapport læses i xml-filen
        templateDir = os.path.join(os.path.abspath(os.path.dirname(__file__)), '../templates')
        if not os.path.exists(templateDir):
            utils.Message('Mappen eksisterer ikke: ' + templateDir)
            return
        filepathReport = os.path.join(templateDir, "UrbanDecarbReport.xml") 
        first_line = ""
        with open(filepathReport) as f:
            first_line = f.readline()
        latestReportVersion = ""
        pos1 = first_line.find("Urban Decarb Report")
        if pos1 > 0:
            pos2 = first_line[pos1:].find('"')
            if pos2 > 0:
                latestReportVersion = first_line[pos1:pos1+pos2]
        if latestReportVersion == "":
            utils.Message("Kan ikke finde rapport template")
            return
        layoutmanager = projectInstance.layoutManager()
        oReport = layoutmanager.layoutByName(latestReportVersion) #QgsReport
        if oReport == None:
            reportExists = False
            for layout in layoutmanager.layouts():
                if "Urban Decarb Report" in layout.name():
                    reportExists = True
            if reportExists:
                msg = "Den nyeste udgave af rapporten findes ikke i projektet.\n\nØnsker du at indlægge den nyeste version?"
            else:
                msg = "Rapporten findes ikke i projektet.\n\nØnsker du at indlægge rapporten?"
            if not utils.Confirm(msg):
                return
            #Advarsel hvis projekt skal gemmes
            if projectInstance.isDirty():
                if not utils.Confirm("Bemærk at der er lavet ændringer i QGIS projektet. Projektet vil blive gemt.\n\nØnsker du at fortsætte?"):
                    return
            #Tidligere udgaver fjernes
            for layout in layoutmanager.layouts():
                if "Urban Decarb Report" in layout.name():
                    layoutmanager.removeLayout(layout)
            #Projektet gemmes
            projectInstance.write()
            #Rapporten indlægges ved almindelig filhåndtering
            imageDir = os.path.join(os.path.abspath(os.path.dirname(__file__)), '../images')
            logoFilePath = os.path.join(imageDir, "hl.png") 
            InsertReportInProjectfile(filepathReport, projPath, logoFilePath)
            #Projektet genindlæses
            projectInstance.read(projPath)
            utils.Message("Rapporten er indlæst i projektet. Urban Decarb skal genstartes før der kan fortsættes.")
            return
            # parent.initGui()
            # oReport = layoutmanager.layoutByName(latestReportVersion) #QgsReport

        #JCN/2024-11-04: Overskrifter opdateres
        oSections = oReport.childSections()  #QgsAbstractReportSection
        for oSec in oSections:
            if oSec.body().itemById('HeaderTextTotal'):
                oLabel = oSec.body().itemById('HeaderTextTotal')
                headerText = 'Plan CO2 - Rapport'
                if oDialog.txtLokalplan.text() != "":
                    headerText += ' for ' + oDialog.txtLokalplan.text()
                if oDialog.txtKommune.text() != "":
                    headerText += ' i ' + oDialog.txtKommune.text()
                headerText = '<b>' + headerText + '</b>'
                labelText = headerText + '<br/><br/>Følgende rapport giver en oversigt over, hvordan klimapåvirkningen fordeler sig i lokalplanområdet. <br/><br/><b>Summerede resultater</b><br/><br/>Nedenfor ses de summerede resultater for lokalplanen fordelt på både overordnede urbane elementer samt deres underkategorier.<br/>'
                oLabel.setText(labelText)
            elif oSec.body().itemById('HeaderTextTypology'):
                oLabel = oSec.body().itemById('HeaderTextTypology')
                labelText = '<b>Klimapåvirkning fordelt på typologi</b><br/>'
                oLabel.setText(labelText)
            elif oSec.body().itemById('HeaderTextCategory'):
                oLabel = oSec.body().itemById('HeaderTextCategory')
                labelText = '<b>Delelementer</b><br/>Nedenfor ses hver af bygningernes klimapåvirkning samt andelen af denne udledning fordelt på bygningens underkategorier.'
                oLabel.setText(labelText)
            elif oSec.body().itemById('HeaderTextOpenSurfaces'):
                oLabel = oSec.body().itemById('HeaderTextOpenSurfaces')
                labelText = 'Nedenfor ses hver af overfladernes klimapåvirkning samt andelen af denne udledning fordelt på overfladens underkategorier.'
                oLabel.setText(labelText)
            #HeaderLabel1 ligger i samme section som HeaderTextTotal
            if oSec.body().itemById('HeaderLabel1'):
                oLabel = oSec.body().itemById('HeaderLabel1')
                if oLabel.text() == 'henninglarsen.com':
                    oLabel.setText('')

        #Indhold i tabeller opdateres
        UpdateTableTotalResults(oReport, oDialog, dictTranslate, oLayerBuildings, oLayerOpenSurfacesLines, oLayerSubareas, oLayerOpenSurfacesAreas)
        UpdateTableResultsCategories(oReport, oDialog, dictTranslate, oLayerBuildings, oLayerOpenSurfacesLines, oLayerOpenSurfacesAreas, dict_surface_width)
        UpdateTableResultsTypologies(oReport, oDialog, dictTranslate, oLayerBuildings, oLayerOpenSurfacesLines, oLayerOpenSurfacesAreas, dict_surface_width)
        UpdateTableResultsBuildings(oReport, oDialog, oLayerBuildings)
        UpdateTableResultsOpenSurfaces(oReport, oDialog, dictTranslate, oLayerOpenSurfacesLines, oLayerOpenSurfacesAreas, oLayerSubareas, dict_surface_width)

        #Der eksporteres til pdf.
        #Det foreslåede filnavn ændres til layoutnavn + filextension
        defaultPath = utils.ReadStringSetting("DefaultPath")
        if defaultPath == "" or not os.path.exists(defaultPath):
            defaultPath = "C:\\"
        defaultFilename = "PlanCO2_Rapport"
        if oDialog.txtLokalplan.text() != "":
            defaultFilename += "_" + oDialog.txtLokalplan.text()
        if oDialog.txtKommune.text() != "":
            defaultFilename += "_" + oDialog.txtKommune.text()
        filePath = defaultPath + "\\" + defaultFilename + ".pdf"
        filepath = QFileDialog.getSaveFileName(iface.mainWindow(), "Export til pdf", filePath, "*.pdf")
        if not all(filepath):
            oDialog.btnExportReport.setEnabled(True)
            return
        defaultPath = os.path.dirname(filepath[0])
        utils.SaveSetting("DefaultPath", defaultPath)

        settings = QgsLayoutExporter.PdfExportSettings() # default settings
        result, error = QgsLayoutExporter.exportToPdf(oReport, filepath[0], settings)
        if result == QgsLayoutExporter.Success:
            #Pdf-filen åbnes
            if utils.Confirm("Rapporten er nu dannet. Ønsker du at åbne filen?"):
                #Åben filen
                os.startfile(filepath[0])
        else:
            utils.Message('Fejl ved export af rapport til pdf. Fejlnummer: ' + str(result) + '\n' + error + '\nKontrollér at pdf-filen ikke er åben.')

        oDialog.btnExportReport.setEnabled(True)

    except Exception as e:
        utils.Error("Fejl i ExportReport: " + str(e))

    finally:
        oDialog.lblMessage.setText('')
        QApplication.processEvents() 

#---------------------------------------------------------------------------------------------------------    
# Opdaterer indhold i tabel Total results
#---------------------------------------------------------------------------------------------------------    
def UpdateTableTotalResults(oReport, oDialog, dictTranslate, oLayerBuildings, oLayerOpenSurfacesLines, oLayerSubareas, oLayerOpenSurfacesAreas):

    try:
        htmpTemplate = 'TableTotalResults.html'

        #html-filen indlæses
        templateDir = os.path.join(os.path.abspath(os.path.dirname(__file__)), '../templates')
        sHtmlPath = os.path.join(templateDir, htmpTemplate)
        if not os.path.exists(sHtmlPath):
            utils.Message('Tabel eksisterer ikke: ' + sHtmlPath)
            return

        oHtmlFile = open(sHtmlPath, 'rt', encoding='utf-8')
        oHtmlContent = oHtmlFile.read()
        oHtmlFile.close()

        iTotalNumberofPersons = 0
        if oDialog.txtTotalNumWorkplaces.text() != "":
            iTotalNumberofPersons = int(oDialog.txtTotalNumWorkplaces.text())
        if oDialog.txtTotalNumResidents.text() != "":
            iTotalNumberofPersons += int(oDialog.txtTotalNumResidents.text())
        if iTotalNumberofPersons == 0:
            utils.Message('Bemærk at der hverken er angivet totalt antal arbejdspladser eller totalt antal beboere under Indstillinger. Der kan ikke beregnes resultater pr. person.')
        if oDialog.cmbStudyPeriod.currentText() == '':
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        if study_period == 0:
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return

        #Bygninger
        buildingsEmissionsTotal = 0
        sumFloorArea = 0
        for feat in oLayerBuildings.getFeatures():
            buildingsEmissionsTotal += feat["EmissionsTotal"]
            #Etageareal
            area = feat["AreaCalc"]
            num_floors = feat["NumberOfFloors"]
            num_base_floors = feat["NumberOfBasementFloors"]
            floorarea = area * (num_floors + num_base_floors)
            sumFloorArea += floorarea
        if sumFloorArea == 0:
            buildingsEmissionsTotalM2Year = 0
        else:
            buildingsEmissionsTotalM2Year = buildingsEmissionsTotal / (sumFloorArea * study_period)
        if iTotalNumberofPersons == 0:
            buildingsEmissionsTotalPersonYear = 0
        else:
            buildingsEmissionsTotalPersonYear = buildingsEmissionsTotal /(iTotalNumberofPersons * study_period)
        #Veje, stier og befæstede ubebyggede arealer
        hardsurfacesEmissionsTotal = 0
        #Grønne ubebyggede arealer, træer og buske
        greensurfacesEmissionsTotal = 0
        #'roads and paths' tilhører kategori veje, stier og befæstede arealer
        for feat in oLayerOpenSurfacesLines.getFeatures():
            if feat["EmissionsType"] != None:
                hardsurfacesEmissionsTotal += feat["EmissionsType"]
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                greensurfacesEmissionsTotal += feat["EmissionsTreesShrubs"]
        #Arealer kan tilhøre befæstede og grønne arealer
        for feat in oLayerOpenSurfacesAreas.getFeatures():
            if feat["EmissionsType"] != None:
                type_landscape = dictTranslate.get(feat["type"].lower())
                category_landscape = 'hard surfaces'
                if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
                    category_landscape = 'green surfaces'
                if category_landscape == 'green surfaces':
                    greensurfacesEmissionsTotal += feat["EmissionsType"]
                else:
                    hardsurfacesEmissionsTotal += feat["EmissionsType"]
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                greensurfacesEmissionsTotal += feat["EmissionsTreesShrubs"]
        #JCN/2024-11-08: Delområder (belægninger) medtages
        subareasEmissionsTotal = 0
        for feat in oLayerSubareas.getFeatures():
            subareasEmissionsTotal += feat["EmissionsTotal"]
        #Total
        totalEmissionsTotal = buildingsEmissionsTotal + hardsurfacesEmissionsTotal + greensurfacesEmissionsTotal + subareasEmissionsTotal
        if sumFloorArea == 0:
            totalEmissionsTotalM2Year = 0
        else:
            totalEmissionsTotalM2Year = totalEmissionsTotal / (sumFloorArea * study_period)
        if iTotalNumberofPersons == 0:
            totalEmissionsTotalPersonYear = 0
        else:
            totalEmissionsTotalPersonYear = totalEmissionsTotal / (iTotalNumberofPersons * study_period)
        #Veje, stier og befæstede ubebyggede arealer
        if sumFloorArea == 0:
            hardsurfacesEmissionsTotalM2Year = 0
        else:
            hardsurfacesEmissionsTotalM2Year = hardsurfacesEmissionsTotal / (sumFloorArea * study_period)
        if iTotalNumberofPersons == 0:
            hardsurfacesEmissionsTotalPersonYear = 0
        else:
            hardsurfacesEmissionsTotalPersonYear = hardsurfacesEmissionsTotal / (iTotalNumberofPersons * study_period)
        #Grønne ubebyggede arealer, træer og buske
        if sumFloorArea == 0:
            greensurfacesEmissionsTotalM2Year = 0
        else:
            greensurfacesEmissionsTotalM2Year = greensurfacesEmissionsTotal / (sumFloorArea * study_period)
        if iTotalNumberofPersons == 0:
            greensurfacesEmissionsTotalPersonYear = 0
        else:
            greensurfacesEmissionsTotalPersonYear = greensurfacesEmissionsTotal / (iTotalNumberofPersons * study_period)
        #JCN/2024-11-08: Delområder (belægninger) medtages
        if sumFloorArea == 0:
            subareasEmissionsTotalM2Year = 0
        else:
            subareasEmissionsTotalM2Year = subareasEmissionsTotal / (sumFloorArea * study_period)
        if iTotalNumberofPersons == 0:
            subareasEmissionsTotalPersonYear = 0
        else:
            subareasEmissionsTotalPersonYear = subareasEmissionsTotal / (iTotalNumberofPersons * study_period)
        #I procent
        if totalEmissionsTotal == 0:
            buildingsAndel = 0.0
            hardsurfacesAndel = 0.0
            greensurfacesAndel = 0.0
            subareasAndel = 0.0
        else:
            buildingsAndel = int(100 * buildingsEmissionsTotal / totalEmissionsTotal)
            hardsurfacesAndel = int(100 * hardsurfacesEmissionsTotal / totalEmissionsTotal)
            greensurfacesAndel = int(100 * greensurfacesEmissionsTotal / totalEmissionsTotal)
            subareasAndel = int(100 * subareasEmissionsTotal / totalEmissionsTotal)
        totalAndel = 100

        #Replace indhold
        #Bygninger    
        oHtmlContent = oHtmlContent.replace('#BuildingsEmissionsTotal#', str(round(buildingsEmissionsTotal/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#BuildingsEmissionsTotalM2Year#', str(round(buildingsEmissionsTotalM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#BuildingsEmissionsTotalPersonYear#', str(round(buildingsEmissionsTotalPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#BuildingsAndel#', format(buildingsAndel, ',.1f').replace('.',','))
        #Veje, stier og befæstede ubebyggede arealer
        oHtmlContent = oHtmlContent.replace('#HardSurfacesEmissionsTotal#', str(round(hardsurfacesEmissionsTotal/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardSurfacesEmissionsTotalM2Year#', str(round(hardsurfacesEmissionsTotalM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardSurfacesEmissionsTotalPersonYear#', str(round(hardsurfacesEmissionsTotalPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardSurfacesAndel#', format(hardsurfacesAndel, ',.1f').replace('.',','))
        #Grønne ubebyggede arealer, træer og buske
        oHtmlContent = oHtmlContent.replace('#GreenSurfacesEmissionsTotal#', str(round(greensurfacesEmissionsTotal/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenSurfacesEmissionsTotalM2Year#', str(round(greensurfacesEmissionsTotalM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenSurfacesEmissionsTotalPersonYear#', str(round(greensurfacesEmissionsTotalPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenSurfacesAndel#', format(greensurfacesAndel, ',.1f').replace('.',','))
        #Delområder
        oHtmlContent = oHtmlContent.replace('#SubareasEmissionsTotal#', str(round(subareasEmissionsTotal/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#SubareasEmissionsTotalM2Year#', str(round(subareasEmissionsTotalM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#SubareasEmissionsTotalPersonYear#', str(round(subareasEmissionsTotalPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#SubareasAndel#', format(subareasAndel, ',.1f').replace('.',','))
        #Total
        oHtmlContent = oHtmlContent.replace('#TotalEmissionsTotal#', str(round(totalEmissionsTotal/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TotalEmissionsTotalM2Year#', str(round(totalEmissionsTotalM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TotalEmissionsTotalPersonYear#', str(round(totalEmissionsTotalPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TotalAndel#', format(totalAndel, ',.1f').replace('.',','))

        foundSecTot = False
        oSections = oReport.childSections()  #QgsAbstractReportSection
        for oSec in oSections:
            if oSec.body().itemById('TotalResults'):
                htmlFrame = oSec.body().itemById('TotalResults')
                htmlTable = htmlFrame.multiFrame()
                htmlTable.setHtml(oHtmlContent)
                htmlTable.setContentMode(1) #ManualHtml
                htmlTable.loadHtml()
                foundSecTot = True
                break

        if not foundSecTot:
            utils.Message("Tabel TotalResults ikke fundet")
            
    except Exception as e:
        utils.Error('Fejl i UpdateTableTotalResults: ' + str(e))

#---------------------------------------------------------------------------------------------------------    
# Opdaterer indhold i tabel Results Categories
#---------------------------------------------------------------------------------------------------------    
def UpdateTableResultsCategories(oReport, oDialog, dictTranslate, oLayerBuildings, oLayerOpenSurfacesLines, oLayerOpenSurfacesAreas, dict_surface_width):

    try:
        htmpTemplate = 'TableResultsCategories.html'

        #html-filen indlæses
        templateDir = os.path.join(os.path.abspath(os.path.dirname(__file__)), '../templates')
        sHtmlPath = os.path.join(templateDir, htmpTemplate)
        if not os.path.exists(sHtmlPath):
            utils.Message('Tabel eksisterer ikke: ' + sHtmlPath)
            return

        oHtmlFile = open(sHtmlPath, 'rt', encoding='utf-8')
        oHtmlContent = oHtmlFile.read()
        oHtmlFile.close()

        iTotalNumberofPersons = 0
        if oDialog.txtTotalNumWorkplaces.text() != "":
            iTotalNumberofPersons = int(oDialog.txtTotalNumWorkplaces.text())
        if oDialog.txtTotalNumResidents.text() != "":
            iTotalNumberofPersons += int(oDialog.txtTotalNumResidents.text())
        if oDialog.cmbStudyPeriod.currentText() == '':
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        if study_period == 0:
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return

        #Bygninger
        foundationTot = 0
        foundationTotM2Year = 0
        foundationTotPersonYear = 0
        structureTot = 0
        structureTotM2Year = 0
        structureTotPersonYear = 0
        envelopeTot = 0
        envelopeTotM2Year = 0
        envelopeTotPersonYear = 0
        windowTot = 0
        windowTotM2Year = 0
        windowTotPersonYear = 0
        interiorTot = 0
        interiorTotM2Year = 0
        interiorTotPersonYear = 0
        technicalTot = 0
        technicalTotM2Year = 0
        technicalTotPersonYear = 0
        constructionTot = 0
        constructionTotM2Year = 0
        constructionTotPersonYear = 0
        demolitionTot = 0
        demolitionTotM2Year = 0
        demolitionTotPersonYear = 0
        operationalTot = 0
        operationalTotM2Year = 0
        operationalTotPersonYear = 0
        sumFloorArea = 0
        for feat in oLayerBuildings.getFeatures():
            #numPersons = feat["numResidents"] + feat["numWorkplaces"]
            #Etageareal
            area = feat["AreaCalc"]
            num_floors = feat["NumberOfFloors"]
            num_base_floors = feat["NumberOfBasementFloors"]
            floorarea = area * (num_floors + num_base_floors)
            sumFloorArea += floorarea
            if feat["EmissionsFoundation"] != None:
                foundationTot += feat["EmissionsFoundation"]
            if feat["EmissionsFoundation"] != None:
                structureTot += feat["EmissionsStructure"]
            if feat["EmissionsEnvelope"] != None:
                envelopeTot += feat["EmissionsEnvelope"]
            if feat["EmissionsWindow"] != None:
                windowTot += feat["EmissionsWindow"]
            if feat["EmissionsInterior"] != None:
                interiorTot += feat["EmissionsInterior"]
            if feat["EmissionsTechnical"] != None:
                technicalTot += feat["EmissionsTechnical"]
            if feat["EmissionsConstruction"] != None:
                constructionTot += feat["EmissionsConstruction"]
            if feat["EmissionsDemolition"] != None:
                demolitionTot += feat["EmissionsDemolition"]
            if feat["EmissionsOperational"] != None:
                operationalTot += feat["EmissionsOperational"]
                
        if sumFloorArea > 0:
            foundationTotM2Year = foundationTot / (sumFloorArea * study_period)
            structureTotM2Year = structureTot / (sumFloorArea * study_period)
            envelopeTotM2Year = envelopeTot / (sumFloorArea * study_period)
            windowTotM2Year = windowTot / (sumFloorArea * study_period)
            interiorTotPersonYear = interiorTot / (sumFloorArea * study_period)
            technicalTotM2Year = technicalTot / (sumFloorArea * study_period)
            constructionTotM2Year = constructionTot / (sumFloorArea * study_period)
            demolitionTotM2Year = demolitionTot / (sumFloorArea * study_period)
            operationalTotM2Year = operationalTot / (sumFloorArea * study_period)
        if iTotalNumberofPersons > 0:
            foundationTotPersonYear = foundationTot / (iTotalNumberofPersons * study_period)
            structureTotPersonYear = structureTot / (iTotalNumberofPersons * study_period)
            envelopeTotPersonYear = envelopeTot / (iTotalNumberofPersons * study_period)
            windowTotPersonYear = windowTot / (iTotalNumberofPersons * study_period)
            interiorTotPersonYear = interiorTot / (iTotalNumberofPersons * study_period)
            technicalTotPersonYear = technicalTot / (iTotalNumberofPersons * study_period)
            constructionTotPersonYear = constructionTot / (iTotalNumberofPersons * study_period)
            demolitionTotPersonYear = demolitionTot / (iTotalNumberofPersons * study_period)
            operationalTotPersonYear = operationalTot / (iTotalNumberofPersons * study_period)

        #Veje, stier og befæstede ubebyggede arealer
        roadsTot = 0
        roadsTotM2Year = 0
        roadsTotPersonYear = 0
        hardTot = 0
        hardTotM2Year = 0
        hardTotPersonYear = 0
        #Grønne ubebyggede arealer, træer og buske
        greenTot = 0
        greenTotM2Year = 0
        greenTotPersonYear = 0
        treesTot = 0
        treesTotM2Year = 0
        treesTotPersonYear = 0
        sumAreaRoads = 0
        sumAreaHard = 0
        sumAreaGreen = 0
        for feat in oLayerOpenSurfacesLines.getFeatures():
            width = feat["width"]
            if width <= 0:
                width = float(dict_surface_width[feat["type"]])
            area = feat["LengthCalc"] * width
            sumAreaRoads += area
            if feat["EmissionsType"] != None:
                roadsTot += feat["EmissionsType"]
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                treesTot += feat["EmissionsTreesShrubs"]
        for feat in oLayerOpenSurfacesAreas.getFeatures():
            if feat["EmissionsType"] != None:
                type_landscape = dictTranslate.get(feat["type"].lower())
                category_landscape = 'hard surfaces'
                if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
                    category_landscape = 'green surfaces'
                if category_landscape == 'green surfaces':
                    greenTot += feat["EmissionsType"]
                    sumAreaGreen += feat["AreaCalc"]
                else:
                    hardTot += feat["EmissionsType"]
                    sumAreaHard += feat["AreaCalc"]
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                treesTot += feat["EmissionsTreesShrubs"]

        if sumAreaRoads > 0:
            roadsTotM2Year = roadsTot / (sumAreaRoads * study_period)
        if sumAreaHard > 0:
            hardTotM2Year = hardTot / (sumAreaHard * study_period)
        if sumAreaGreen > 0:
            greenTotM2Year = greenTot / (sumAreaGreen * study_period)
            treesTotM2Year = treesTot / (sumAreaGreen * study_period)
        if iTotalNumberofPersons > 0:
            roadsTotPersonYear = roadsTot / (iTotalNumberofPersons * study_period)
            hardTotPersonYear = hardTot / (iTotalNumberofPersons * study_period)
            greenTotPersonYear = greenTot / (iTotalNumberofPersons * study_period)
            treesTotPersonYear = treesTot / (iTotalNumberofPersons * study_period)

        #Total
        totalEmissionsTotal = foundationTot + structureTot + envelopeTot + windowTot + interiorTot + technicalTot + constructionTot + demolitionTot + operationalTot + roadsTot + hardTot + greenTot + treesTot

        if totalEmissionsTotal == 0:
            foundationAndel =  0.0
            structureAndel = 0.0
            envelopeAndel = 0.0
            windowAndel = 0.0
            interiorAndel = 0.0
            technicalAndel = 0.0
            constructionAndel = 0.0
            demolitionAndel = 0.0
            operationalAndel = 0.0
            roadsAndel = 0.0
            hardAndel = 0.0
            greenAndel = 0.0
            treesAndel = 0.0
        else:
            foundationAndel =  int(100 * foundationTot / totalEmissionsTotal)
            structureAndel = int(100 * structureTot / totalEmissionsTotal)
            envelopeAndel = int(100 * envelopeTot / totalEmissionsTotal)
            windowAndel = int(100 * windowTot / totalEmissionsTotal)
            interiorAndel = int(100 * interiorTot / totalEmissionsTotal)
            technicalAndel = int(100 * technicalTot / totalEmissionsTotal)
            constructionAndel = int(100 * constructionTot / totalEmissionsTotal)
            demolitionAndel = int(100 * demolitionTot / totalEmissionsTotal)
            operationalAndel = int(100 * operationalTot / totalEmissionsTotal)
            roadsAndel = int(100 * roadsTot / totalEmissionsTotal)
            hardAndel = int(100 * hardTot / totalEmissionsTotal)
            greenAndel = int(100 * greenTot / totalEmissionsTotal)
            treesAndel = int(100 * treesTot / totalEmissionsTotal)

        #Replace indhold
        #Fundament    
        oHtmlContent = oHtmlContent.replace('#FoundationTot#', str(round(foundationTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#FoundationTotM2Year#', str(round(foundationTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#FoundationTotPersonYear#', str(round(foundationTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#FoundationAndel#', format(foundationAndel, ',.1f').replace('.',','))
        #Bærende system
        oHtmlContent = oHtmlContent.replace('#StructureTot#', str(round(structureTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#StructureTotM2Year#', str(round(structureTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#StructureTotPersonYear#', str(round(structureTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#StructureAndel#', format(structureAndel, ',.1f').replace('.',','))
        #Klimaskærm
        oHtmlContent = oHtmlContent.replace('#EnvelopeTot#', str(round(envelopeTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#EnvelopeTotM2Year#', str(round(envelopeTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#EnvelopeTotPersonYear#', str(round(envelopeTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#EnvelopeAndel#', format(envelopeAndel, ',.1f').replace('.',','))
        #Vindue
        oHtmlContent = oHtmlContent.replace('#WindowTot#', str(round(windowTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#WindowTotM2Year#', str(round(windowTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#WindowTotPersonYear#', str(round(windowTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#WindowAndel#', format(windowAndel, ',.1f').replace('.',','))
        #Interiør
        oHtmlContent = oHtmlContent.replace('#InteriorTot#', str(round(interiorTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#InteriorTotM2Year#', str(round(interiorTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#InteriorTotPersonYear#', str(round(interiorTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#InteriorAndel#', format(interiorAndel, ',.1f').replace('.',','))
        #Tekniske systemer
        oHtmlContent = oHtmlContent.replace('#TechnicalTot#', str(round(technicalTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TechnicalTotM2Year#', str(round(technicalTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TechnicalTotPersonYear#', str(round(technicalTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TechnicalAndel#', format(technicalAndel, ',.1f').replace('.',','))
        #Opførelse
        oHtmlContent = oHtmlContent.replace('#ConstructionTot#', str(round(constructionTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ConstructionTotM2Year#', str(round(constructionTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ConstructionTotPersonYear#', str(round(constructionTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ConstructionAndel#', format(constructionAndel, ',.1f').replace('.',','))
        #Nedrivning
        oHtmlContent = oHtmlContent.replace('#DemolitionTot#', str(round(demolitionTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DemolitionTotM2Year#', str(round(demolitionTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DemolitionTotPersonYear#', str(round(demolitionTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DemolitionAndel#', format(demolitionAndel, ',.1f').replace('.',','))
        #Driftsenergi
        oHtmlContent = oHtmlContent.replace('#OperationalTot#', str(round(operationalTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#OperationalTotM2Year#', str(round(operationalTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#OperationalTotPersonYear#', str(round(operationalTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#OperationalAndel#', format(operationalAndel, ',.1f').replace('.',','))
        
        #Veje og stier
        oHtmlContent = oHtmlContent.replace('#RoadsTot#', str(round(roadsTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RoadsTotM2Year#', str(round(roadsTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RoadsTotPersonYear#', str(round(roadsTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RoadsAndel#', format(roadsAndel, ',.1f').replace('.',','))
        #Befæstede ubebyggede arealer
        oHtmlContent = oHtmlContent.replace('#HardTot#', str(round(hardTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardTotM2Year#', str(round(hardTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardTotPersonYear#', str(round(hardTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardAndel#', format(hardAndel, ',.1f').replace('.',','))
        #Grønne ubebyggede arealer
        oHtmlContent = oHtmlContent.replace('#GreenTot#', str(round(greenTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenTotM2Year#', str(round(greenTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenTotPersonYear#', str(round(greenTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenAndel#', format(greenAndel, ',.1f').replace('.',','))
        #Træer og buske
        oHtmlContent = oHtmlContent.replace('#TreesTot#', str(round(treesTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TreesTotM2Year#', str(round(treesTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TreesTotPersonYear#', str(round(treesTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TreesAndel#', format(treesAndel, ',.1f').replace('.',','))

        foundSec = False
        oSections = oReport.childSections()  #QgsAbstractReportSection
        for oSec in oSections:
            if oSec.body().itemById('ResultsCategories'):
                htmlFrame = oSec.body().itemById('ResultsCategories')
                htmlTable = htmlFrame.multiFrame()
                htmlTable.setHtml(oHtmlContent)
                htmlTable.setContentMode(1) #ManualHtml
                htmlTable.loadHtml()
                foundSec = True
                break

        if not foundSec:
            utils.Message("Tabel ResultsCategories ikke fundet")
            
    except Exception as e:
        utils.Error('Fejl i UpdateTableResultsCategories: ' + str(e))

#---------------------------------------------------------------------------------------------------------    
# Opdaterer indhold i tabel ResultsTypologies
#---------------------------------------------------------------------------------------------------------    
def UpdateTableResultsTypologies(oReport, oDialog, dictTranslate, oLayerBuildings, oLayerOpenSurfacesLines, oLayerOpenSurfacesAreas, dict_surface_width):

    try:
        htmpTemplate = 'TableResultsTypologies.html'

        #html-filen indlæses
        templateDir = os.path.join(os.path.abspath(os.path.dirname(__file__)), '../templates')
        sHtmlPath = os.path.join(templateDir, htmpTemplate)
        if not os.path.exists(sHtmlPath):
            utils.Message('Tabel eksisterer ikke: ' + sHtmlPath)
            return
        
        oHtmlFile = open(sHtmlPath, 'rt', encoding='utf-8')
        oHtmlContent = oHtmlFile.read()
        oHtmlFile.close()

        iTotalNumberofPersons = 0
        if oDialog.txtTotalNumWorkplaces.text() != "":
            iTotalNumberofPersons = int(oDialog.txtTotalNumWorkplaces.text())
        if oDialog.txtTotalNumResidents.text() != "":
            iTotalNumberofPersons += int(oDialog.txtTotalNumResidents.text())
        if oDialog.cmbStudyPeriod.currentText() == '':
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        if study_period == 0:
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return

        #Bygninger
        idx = oLayerBuildings.fields().indexOf("Building_usage")
        if idx < 0:
            utils.Message("Bygningslaget indeholder ikke feltet 'Building_usage'")
            return

        detachedHouseTot = 0
        detachedHouseTotM2Year = 0
        detachedHouseTotPersonYear = 0.0
        terracedHouseTot = 0
        terracedHouseTotM2Year = 0
        terracedHouseTotPersonYear = 0.0
        apartmentTot = 0
        apartmentTotM2Year = 0
        apartmentTotPersonYear = 0.0
        officeTot = 0
        officeTotM2Year = 0
        officeTotPersonYear = 0.0
        schoolTot = 0
        schoolTotM2Year = 0
        schoolTotPersonYear = 0.0
        institutionTot = 0
        institutionTotM2Year = 0
        institutionTotPersonYear = 0.0
        retailTot = 0
        retailTotM2Year = 0
        retailTotPersonYear = 0.0
        parkingTot = 0
        parkingTotM2Year = 0
        parkingTotPersonYear = 0.0
        industryTot = 0
        industryTotM2Year = 0
        industryTotPersonYear = 0.0
        transportTot = 0
        transportTotM2Year = 0
        transportTotPersonYear = 0.0
        hotelTot = 0
        hotelTotM2Year = 0
        hotelTotPersonYear = 0.0
        hospitalTot = 0
        hospitalTotM2Year = 0
        hospitalTotPersonYear = 0.0
        demolitionTot = 0
        demolitionTotM2Year = 0
        demolitionTotPersonYear = 0.0
        sumFloorArea = 0
        sumNumPersonsDetatchedHouse = 0
        sumNumPersonsTerracedHouse = 0
        sumNumPersonsApartment = 0
        sumNumWorkplacesOffice = 0
        for feat in oLayerBuildings.getFeatures():
            #Etageareal
            area = feat["AreaCalc"]
            num_floors = feat["NumberOfFloors"]
            num_base_floors = feat["NumberOfBasementFloors"]
            floorarea = area * (num_floors + num_base_floors)
            sumFloorArea += floorarea
            typology = dictTranslate.get(feat["Building_usage"].lower())
            if typology == "detached house":
                sumNumPersonsDetatchedHouse += feat["numResidents"]
                detachedHouseTot += feat["EmissionsTotal"]
            elif typology == "terraced house":
                sumNumPersonsTerracedHouse += feat["numResidents"]
                terracedHouseTot += feat["EmissionsTotal"]
            elif typology == "apartment building":
                sumNumPersonsApartment += feat["numResidents"]
                apartmentTot += feat["EmissionsTotal"]
            elif typology == "office":
                sumNumWorkplacesOffice += feat["numWorkplaces"]
                officeTot += feat["EmissionsTotal"]
            elif typology == "school":
                schoolTot += feat["EmissionsTotal"]
            elif typology == "institution":
                institutionTot += feat["EmissionsTotal"]
            elif typology == "retail":
                retailTot += feat["EmissionsTotal"]
            elif typology == "parking":
                parkingTot += feat["EmissionsTotal"]
            elif typology == "industry":
                industryTot += feat["EmissionsTotal"]
            elif typology == "transport":
                transportTot += feat["EmissionsTotal"]
            elif typology == "hotel":
                hotelTot += feat["EmissionsTotal"]
            elif typology == "hospital":
                hospitalTot += feat["EmissionsTotal"]
            elif typology == "demolition":
                demolitionTot += feat["EmissionsTotal"]

        if sumFloorArea > 0:
            detachedHouseTotM2Year = detachedHouseTot / (sumFloorArea * study_period)
            terracedHouseTotM2Year = terracedHouseTot / (sumFloorArea * study_period)
            apartmentTotM2Year = apartmentTot / (sumFloorArea * study_period)
            officeTotM2Year = officeTot / (sumFloorArea * study_period)
            schoolTotM2Year = schoolTot / (sumFloorArea * study_period)
            institutionTotM2Year = institutionTot / (sumFloorArea * study_period)
            retailTotM2Year = retailTot / (sumFloorArea * study_period)
            parkingTotM2Year = parkingTot / (sumFloorArea * study_period)
            industryTotM2Year = industryTot / (sumFloorArea * study_period)
            transportTotM2Year = transportTot / (sumFloorArea * study_period)
            hotelTotM2Year = hotelTot / (sumFloorArea * study_period)
            hospitalTotM2Year = hospitalTot / (sumFloorArea * study_period)
            demolitionTotM2Year = demolitionTot / (sumFloorArea * study_period)
        if sumNumPersonsDetatchedHouse > 0:
            detachedHouseTotPersonYear = detachedHouseTot / (sumNumPersonsDetatchedHouse * study_period)
        if sumNumPersonsTerracedHouse > 0:
            terracedHouseTotPersonYear = terracedHouseTot / (sumNumPersonsTerracedHouse * study_period)
        if sumNumPersonsApartment > 0:
            apartmentTotPersonYear = apartmentTot / (sumNumPersonsApartment * study_period)
        if sumNumWorkplacesOffice > 0:
            officeTotPersonYear = officeTot / (sumNumWorkplacesOffice * study_period)
        if iTotalNumberofPersons > 0:
            schoolTotPersonYear = schoolTot / (iTotalNumberofPersons * study_period)
            institutionTotPersonYear = institutionTot / (iTotalNumberofPersons * study_period)
            retailTotPersonYear = retailTot / (iTotalNumberofPersons * study_period)
            parkingTotPersonYear = parkingTot / (iTotalNumberofPersons * study_period)
            industryTotPersonYear = industryTot / (iTotalNumberofPersons * study_period)
            transportTotPersonYear = transportTot / (iTotalNumberofPersons * study_period)
            hotelTotPersonYear = hotelTot / (iTotalNumberofPersons * study_period)
            hospitalTotPersonYear = hospitalTot / (iTotalNumberofPersons * study_period)
            demolitionTotPersonYear = demolitionTot / (iTotalNumberofPersons * study_period)

        #Veje, stier og befæstede ubebyggede arealer
        roadsTot = 0
        roadsTotM2Year = 0
        roadsTotPersonYear = 0
        hardTot = 0
        hardTotM2Year = 0
        hardTotPersonYear = 0
        #Grønne ubebyggede arealer, træer og buske
        greenTot = 0
        greenTotM2Year = 0
        greenTotPersonYear = 0
        treesTot = 0
        treesTotM2Year = 0
        treesTotPersonYear = 0
        sumAreaRoads = 0
        sumAreaHard = 0
        sumAreaGreen = 0
        for feat in oLayerOpenSurfacesLines.getFeatures():
            width = feat["width"]
            if width <= 0:
                width = float(dict_surface_width[feat["type"]])
            area = feat["LengthCalc"] * width
            sumAreaRoads += area
            if feat["EmissionsType"] != None:
                roadsTot += feat["EmissionsType"]
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                treesTot += feat["EmissionsTreesShrubs"]
        for feat in oLayerOpenSurfacesAreas.getFeatures():
            if feat["EmissionsType"] != None:
                type_landscape = dictTranslate.get(feat["type"].lower())
                category_landscape = 'hard surfaces'
                if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
                    category_landscape = 'green surfaces'
                if category_landscape == 'green surfaces':
                    greenTot += feat["EmissionsType"]
                    sumAreaGreen += feat["AreaCalc"]
                else:
                    hardTot += feat["EmissionsType"]
                    sumAreaHard += feat["AreaCalc"]
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                treesTot += feat["EmissionsTreesShrubs"]

        if sumAreaRoads > 0:
            roadsTotM2Year = roadsTot / (sumAreaRoads * study_period)
        if sumAreaHard > 0:
            hardTotM2Year = hardTot / (sumAreaHard * study_period)
        if sumAreaGreen > 0:
            greenTotM2Year = greenTot / (sumAreaGreen * study_period)
            treesTotM2Year = treesTot / (sumAreaGreen * study_period)
        if iTotalNumberofPersons > 0:
            roadsTotPersonYear = roadsTot / (iTotalNumberofPersons * study_period)
            greenTotPersonYear = greenTot / (iTotalNumberofPersons * study_period)
            hardTotPersonYear = hardTot / (iTotalNumberofPersons * study_period)
            treesTotPersonYear = treesTot / (iTotalNumberofPersons * study_period)

        #Total
        totalEmissionsTotal = detachedHouseTot + terracedHouseTot + apartmentTot + officeTot + schoolTot + institutionTot + retailTot + parkingTot + industryTot + transportTot + hotelTot + hospitalTot + demolitionTot + roadsTot + hardTot + greenTot + treesTot

        if totalEmissionsTotal == 0:
            detachedHouseAndel =  0.0
            terracedHouseAndel = 0.0
            apartmentAndel = 0.0
            officeAndel = 0.0
            schoolAndel = 0.0
            institutionAndel = 0.0
            retailAndel = 0.0
            parkingAndel = 0.0
            industryAndel = 0.0
            transportAndel = 0.0
            hotelAndel = 0.0
            hospitalAndel = 0.0
            demolitionAndel = 0.0
            roadsAndel = 0.0
            hardAndel = 0.0
            greenAndel = 0.0
            treesAndel = 0.0
        else:
            detachedHouseAndel =  int(100 * detachedHouseTot / totalEmissionsTotal)
            terracedHouseAndel = int(100 * terracedHouseTot / totalEmissionsTotal)
            apartmentAndel = int(100 * apartmentTot / totalEmissionsTotal)
            officeAndel = int(100 * officeTot / totalEmissionsTotal)
            schoolAndel = int(100 * schoolTot / totalEmissionsTotal)
            institutionAndel = int(100 * institutionTot / totalEmissionsTotal)
            retailAndel = int(100 * retailTot / totalEmissionsTotal)
            parkingAndel = int(100 * parkingTot / totalEmissionsTotal)
            industryAndel = int(100 * industryTot / totalEmissionsTotal)
            transportAndel = int(100 * transportTot / totalEmissionsTotal)
            hotelAndel = int(100 * hotelTot / totalEmissionsTotal)
            hospitalAndel = int(100 * hospitalTot / totalEmissionsTotal)
            demolitionAndel = int(100 * demolitionTot / totalEmissionsTotal)
            roadsAndel = int(100 * roadsTot / totalEmissionsTotal)
            hardAndel = int(100 * hardTot / totalEmissionsTotal)
            greenAndel = int(100 * greenTot / totalEmissionsTotal)
            treesAndel = int(100 * treesTot / totalEmissionsTotal)

        # #Replace indhold
        #Enfamiliehus    
        oHtmlContent = oHtmlContent.replace('#DetachedHouseTot#', str(round(detachedHouseTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DetachedHouseTotM2Year#', str(round(detachedHouseTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DetachedHouseTotPersonYear#', str(round(detachedHouseTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DetachedHouseAndel#', format(detachedHouseAndel, ',.1f').replace('.',','))
        #Rækkehus    
        oHtmlContent = oHtmlContent.replace('#TerracedHouseTot#', str(round(terracedHouseTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TerracedHouseTotM2Year#', str(round(terracedHouseTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TerracedHouseTotPersonYear#', str(round(terracedHouseTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TerracedHouseAndel#', format(terracedHouseAndel, ',.1f').replace('.',','))
        #Etagebolig    
        oHtmlContent = oHtmlContent.replace('#ApartmentTot#', str(round(apartmentTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ApartmentTotM2Year#', str(round(apartmentTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ApartmentTotPersonYear#', str(round(apartmentTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ApartmentAndel#', format(apartmentAndel, ',.1f').replace('.',','))
        #Kontor    
        oHtmlContent = oHtmlContent.replace('#OfficeTot#', str(round(officeTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#OfficeTotM2Year#', str(round(officeTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#OfficeTotPersonYear#', str(round(officeTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#OfficeAndel#', format(officeAndel, ',.1f').replace('.',','))
        #Undervisning    
        oHtmlContent = oHtmlContent.replace('#SchoolTot#', str(round(schoolTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#SchoolTotM2Year#', str(round(schoolTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#SchoolTotPersonYear#', str(round(schoolTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#SchoolAndel#', format(schoolAndel, ',.1f').replace('.',','))
        #Daginstitution    
        oHtmlContent = oHtmlContent.replace('#InstitutionTot#', str(round(institutionTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#InstitutionTotM2Year#', str(round(institutionTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#InstitutionTotPersonYear#', str(round(institutionTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#InstitutionAndel#', format(institutionAndel, ',.1f').replace('.',','))
        #Butik    
        oHtmlContent = oHtmlContent.replace('#RetailTot#', str(round(retailTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RetailTotM2Year#', str(round(retailTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RetailTotPersonYear#', str(round(retailTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RetailAndel#', format(retailAndel, ',.1f').replace('.',','))
        #Parkering    
        oHtmlContent = oHtmlContent.replace('#ParkingTot#', str(round(parkingTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ParkingTotM2Year#', str(round(parkingTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ParkingTotPersonYear#', str(round(parkingTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#ParkingAndel#', format(parkingAndel, ',.1f').replace('.',','))
        #Industri    
        oHtmlContent = oHtmlContent.replace('#IndustryTot#', str(round(industryTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#IndustryTotM2Year#', str(round(industryTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#IndustryTotPersonYear#', str(round(industryTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#IndustryAndel#', format(industryAndel, ',.1f').replace('.',','))
        #Transport    
        oHtmlContent = oHtmlContent.replace('#TransportTot#', str(round(transportTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TransportTotM2Year#', str(round(transportTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TransportTotPersonYear#', str(round(transportTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TransportAndel#', format(transportAndel, ',.1f').replace('.',','))
        #Hotel    
        oHtmlContent = oHtmlContent.replace('#HotelTot#', str(round(hotelTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HotelTotM2Year#', str(round(hotelTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HotelTotPersonYear#', str(round(hotelTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HotelAndel#', format(hotelAndel, ',.1f').replace('.',','))
        #Sundhedshuse    
        oHtmlContent = oHtmlContent.replace('#HospitalTot#', str(round(hospitalTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HospitalTotM2Year#', str(round(hospitalTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HospitalTotPersonYear#', str(round(hospitalTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HospitalAndel#', format(hospitalAndel, ',.1f').replace('.',','))
        #Nedrivning    
        oHtmlContent = oHtmlContent.replace('#DemolitionTot#', str(round(demolitionTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DemolitionTotM2Year#', str(round(demolitionTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DemolitionTotPersonYear#', str(round(demolitionTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#DemolitionAndel#', format(demolitionAndel, ',.1f').replace('.',','))
       
        #Veje og stier
        oHtmlContent = oHtmlContent.replace('#RoadsTot#', str(round(roadsTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RoadsTotM2Year#', str(round(roadsTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RoadsTotPersonYear#', str(round(roadsTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#RoadsAndel#', format(roadsAndel, ',.1f').replace('.',','))
        #Befæstede ubebyggede arealer
        oHtmlContent = oHtmlContent.replace('#HardTot#', str(round(hardTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardTotM2Year#', str(round(hardTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardTotPersonYear#', str(round(hardTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#HardAndel#', format(hardAndel, ',.1f').replace('.',','))
        #Grønne ubebyggede arealer
        oHtmlContent = oHtmlContent.replace('#GreenTot#', str(round(greenTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenTotM2Year#', str(round(greenTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenTotPersonYear#', str(round(greenTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#GreenAndel#', format(greenAndel, ',.1f').replace('.',','))
        #Træer og buske
        oHtmlContent = oHtmlContent.replace('#TreesTot#', str(round(treesTot/1000, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TreesTotM2Year#', str(round(treesTotM2Year, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TreesTotPersonYear#', str(round(treesTotPersonYear, 1)).replace('.',','))
        oHtmlContent = oHtmlContent.replace('#TreesAndel#', format(treesAndel, ',.1f').replace('.',','))

        foundSec = False
        oSections = oReport.childSections()  #QgsAbstractReportSection
        for oSec in oSections:
            if oSec.body().itemById('ResultsTypologies'):
                htmlFrame = oSec.body().itemById('ResultsTypologies')
                htmlTable = htmlFrame.multiFrame()
                htmlTable.setHtml(oHtmlContent)
                htmlTable.setContentMode(1) #ManualHtml
                htmlTable.loadHtml()
                foundSec = True
                break

        if not foundSec:
            utils.Message("Tabel ResultsTypologies ikke fundet")
            
    except Exception as e:
        utils.Error('Fejl i UpdateTableResultsTypologies: ' + str(e))

#---------------------------------------------------------------------------------------------------------    
# Opdaterer indhold i tabel ResultsBuildings
#---------------------------------------------------------------------------------------------------------    
def UpdateTableResultsBuildings(oReport, oDialog, oLayerBuildings):

    try:
        htmpTemplate = 'TableResultsBuildings.html'

        #html-filen indlæses
        templateDir = os.path.join(os.path.abspath(os.path.dirname(__file__)), '../templates')
        sHtmlPath = os.path.join(templateDir, htmpTemplate)
        if not os.path.exists(sHtmlPath):
            utils.Message('Tabel eksisterer ikke: ' + sHtmlPath)

        oHtmlFile = open(sHtmlPath, 'rt', encoding='utf-8')
        oHtmlContent = oHtmlFile.read()
        oHtmlFile.close()

        sHtml = ""
        for feat in oLayerBuildings.getFeatures():
            foundationAndel =  0
            structureAndel = 0
            envelopeAndel = 0
            windowAndel = 0
            interiorAndel = 0
            technicalAndel = 0
            constructionAndel = 0
            demolitionAndel = 0
            operationalAndel = 0
            emissionsTotal = feat["EmissionsTotal"]
            if emissionsTotal > 0:
                if feat["EmissionsFoundation"] != None:
                    foundationAndel =  round((100 * feat["EmissionsFoundation"] / emissionsTotal), 1)
                if feat["EmissionsStructure"] != None:
                    structureAndel =  round((100 * feat["EmissionsStructure"] / emissionsTotal), 1)
                if feat["EmissionsEnvelope"] != None:
                    envelopeAndel =  round((100 * feat["EmissionsEnvelope"] / emissionsTotal), 1)
                if feat["EmissionsWindow"] != None:
                    windowAndel =  round((100 * feat["EmissionsWindow"] / emissionsTotal), 1)
                if feat["EmissionsInterior"] != None:
                    interiorAndel =  round((100 * feat["EmissionsInterior"] / emissionsTotal), 1)
                if feat["EmissionsTechnical"] != None:
                    technicalAndel =  round((100 * feat["EmissionsTechnical"] / emissionsTotal), 1)
                if feat["EmissionsConstruction"] != None:
                    constructionAndel =  round((100 * feat["EmissionsConstruction"] / emissionsTotal), 1)
                if feat["EmissionsDemolition"] != None:
                    demolitionAndel =  round((100 * feat["EmissionsDemolition"] / emissionsTotal), 1)
                if feat["EmissionsOperational"] != None:
                    operationalAndel =  round((100 * feat["EmissionsOperational"] / emissionsTotal), 1)
                
            sHtml += "<tr>"
            sHtml += "<td class='clsCell3a'><b>" + feat["id"] + "</b></td>"
            sHtml += "<td class='clsCell3'>" + str(round(feat["EmissionsTotal"]/1000, 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round(feat["EmissionsTotalM2Year"], 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round(feat["EmissionsTotalPersonYear"], 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(foundationAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(structureAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(envelopeAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(windowAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(interiorAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(technicalAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(constructionAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(demolitionAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + format(operationalAndel, ',.1f').replace('.',',') + "</td>"
            sHtml += "</tr>"

        oHtmlContent = oHtmlContent.replace('#HTMLTABLEROWS#', sHtml)

        foundSec = False
        oSections = oReport.childSections()  #QgsAbstractReportSection
        for oSec in oSections:
            if oSec.body().itemById('ResultsBuildings'):
                htmlFrame = oSec.body().itemById('ResultsBuildings')
                htmlTable = htmlFrame.multiFrame()
                htmlTable.setHtml(oHtmlContent)
                htmlTable.setContentMode(1) #ManualHtml
                htmlTable.loadHtml()
                foundSec = True
                break

        if not foundSec:
            utils.Message("Tabel med ResultsBuildings ikke fundet")
            
    except Exception as e:
        utils.Error('Fejl i UpdateTableResultsBuildings: ' + str(e))

#---------------------------------------------------------------------------------------------------------    
# Opdaterer indhold i tabel ResultsOpenSurfaces
#---------------------------------------------------------------------------------------------------------    
def UpdateTableResultsOpenSurfaces(oReport, oDialog, dictTranslate, oLayerOpenSurfacesLines, oLayerOpenSurfacesAreas, oLayerSubareas, dict_surface_width):

    try:
        htmpTemplate = 'TableResultsOpenSurfaces.html'

        #html-filen indlæses
        templateDir = os.path.join(os.path.abspath(os.path.dirname(__file__)), '../templates')
        sHtmlPath = os.path.join(templateDir, htmpTemplate)
        if not os.path.exists(sHtmlPath):
            utils.Message('Tabel eksisterer ikke: ' + sHtmlPath)

        oHtmlFile = open(sHtmlPath, 'rt', encoding='utf-8')
        oHtmlContent = oHtmlFile.read()
        oHtmlFile.close()

        if oDialog.cmbStudyPeriod.currentText() == '':
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        if study_period == 0:
            utils.Message('Vælg venligst beregningsperiode under Indstillinger')
            return

        sHtml = ""
        for feat in oLayerOpenSurfacesLines.getFeatures():
            roadsTot = 0.0
            roadsTotM2Year = 0.0
            treesTot = 0.0
            treesTotM2Year = 0.0
            width = feat["width"]
            if width <= 0:
                width = float(dict_surface_width[feat["type"]])
            area = feat["LengthCalc"] * width
            if feat["EmissionsType"] != None:
                roadsTot += feat["EmissionsType"]
                roadsTotM2Year += feat["EmissionsType"] / (area * study_period)
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                treesTot += feat["EmissionsTreesShrubs"]
                treesTotM2Year += feat["EmissionsTreesShrubs"] / (area * study_period)
            sHtml += "<tr>"
            sHtml += "<td class='clsCell3a'><b>" + feat["id"] + "</b></td>"
            sHtml += "<td class='clsCell3'>" + str(round((roadsTot+treesTot)/1000, 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round((roadsTotM2Year+treesTotM2Year), 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round((roadsTot), 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>0,0</td>"
            sHtml += "<td class='clsCell3'>0,0</td>"
            sHtml += "<td class='clsCell3'>" + str(round((treesTot), 1)).replace('.',',') + "</td>"
            sHtml += "</tr>"
            
        for feat in oLayerOpenSurfacesAreas.getFeatures():
            hardTot = 0.0
            hardTotM2Year = 0.0
            greenTot = 0.0
            greenTotM2Year = 0.0
            treesTot = 0.0
            treesTotM2Year = 0.0
            area = feat["AreaCalc"]
            if feat["EmissionsType"] != None:
                type_landscape = dictTranslate.get(feat["type"].lower())
                category_landscape = 'hard surfaces'
                if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
                    category_landscape = 'green surfaces'
                if category_landscape == 'green surfaces':
                    greenTot += feat["EmissionsType"]
                    greenTotM2Year += feat["EmissionsType"] / (area * study_period)
                else:
                    hardTot += feat["EmissionsType"]
                    hardTotM2Year += feat["EmissionsType"] / (area * study_period)
            #Træer og buske
            if feat["EmissionsTreesShrubs"] != None:    
                treesTot += feat["EmissionsTreesShrubs"]
                treesTotM2Year += feat["EmissionsTreesShrubs"] / (area * study_period)
            sHtml += "<tr>"
            sHtml += "<td class='clsCell3a'><b>" + feat["id"] + "</b></td>"
            sHtml += "<td class='clsCell3'>" + str(round((hardTot+greenTot+treesTot)/1000, 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round((hardTotM2Year+greenTotM2Year+treesTotM2Year), 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>0,0</td>"
            # sHtml += "<td class='clsCell3'>" + format(hardTot, ',.1f').replace('.',',')  + "</td>"
            # sHtml += "<td class='clsCell3'>" + format(greenTot, ',.1f').replace('.',',')  + "</td>"
            # sHtml += "<td class='clsCell3'>" + format(treesTot, ',.1f').replace('.',',')  + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round((hardTot), 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round((greenTot), 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round((treesTot), 1)).replace('.',',') + "</td>"
            sHtml += "</tr>"

        #JCN/2024-11-08: Delområder (belægninger) medtages
        for feat in oLayerSubareas.getFeatures():
            # hardTot = 0.0
            # hardTotM2Year = 0.0
            # greenTot = 0.0
            # greenTotM2Year = 0.0
            # treesTot = 0.0
            # treesTotM2Year = 0.0
            # area = feat["AreaCalc"]
            # if feat["EmissionsType"] != None:
            #     type_landscape = dictTranslate.get(feat["type"].lower())
            #     category_landscape = 'hard surfaces'
            #     if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
            #         category_landscape = 'green surfaces'
            #     if category_landscape == 'green surfaces':
            #         greenTot += feat["EmissionsType"]
            #         greenTotM2Year += feat["EmissionsType"] / (area * study_period)
            #     else:
            #         hardTot += feat["EmissionsType"]
            #         hardTotM2Year += feat["EmissionsType"] / (area * study_period)
            # #Træer og buske
            # if feat["EmissionsTreesShrubs"] != None:    
            #     treesTot += feat["EmissionsTreesShrubs"]
            #     treesTotM2Year += feat["EmissionsTreesShrubs"] / (area * study_period)
            sHtml += "<tr>"
            sHtml += "<td class='clsCell3a'><b>" + feat["Name"] + "</b></td>"
            sHtml += "<td class='clsCell3'>" + str(round(feat["EmissionsTotal"]/1000, 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>" + str(round(feat["EmissionsTotalM2Year"], 1)).replace('.',',') + "</td>"
            sHtml += "<td class='clsCell3'>0,0</td>"
            sHtml += "<td class='clsCell3'>0,0</td>"
            sHtml += "<td class='clsCell3'>0,0</td>"
            sHtml += "<td class='clsCell3'>0,0</td>"
            sHtml += "</tr>"

        oHtmlContent = oHtmlContent.replace('#HTMLTABLEROWS#', sHtml)

        foundSec = False
        oSections = oReport.childSections()  #QgsAbstractReportSection
        for oSec in oSections:
            if oSec.body().itemById('ResultsOpenSurfaces'):
                htmlFrame = oSec.body().itemById('ResultsOpenSurfaces')
                htmlTable = htmlFrame.multiFrame()
                htmlTable.setHtml(oHtmlContent)
                htmlTable.setContentMode(1) #ManualHtml
                htmlTable.loadHtml()
                foundSec = True
                break

        if not foundSec:
            utils.Message("Tabel ResultsOpenSurfaces ikke fundet")
            
    except Exception as e:
        utils.Error('Fejl i UpdateTableResultsOpenSurfaces: ' + str(e))

#---------------------------------------------------------------------------------------------------------    
# Indsætter report fra template
#---------------------------------------------------------------------------------------------------------    
def InsertReportInProjectfile(filepathReport, filepathProject, logoFilePath):

    try:
        if not os.path.exists(filepathReport):
            utils.Message('Filen eksisterer ikke: ' + filepathReport)
            return False
        elif not os.path.exists(filepathProject):
            utils.Message('Projektet eksisterer ikke: ' + filepathProject)
            return False

        oReportFile = open(filepathReport, 'rt', encoding='utf-8')
        oReportContent = oReportFile.read()
        oReportFile.close()
        oReportContent = oReportContent.replace("#LogoFilePath#", logoFilePath)
            
        newContent = ""
        oProjectFile = open(filepathProject, 'rt', encoding='utf-8')
        for line in oProjectFile:
            if line == '  <Layouts/>\n':
                newContent += '  <Layouts>\n' 
                newContent += oReportContent
                newContent += '  </Layouts>\n'
            elif line == '  <Layouts>\n':
                newContent += line
                newContent += oReportContent
            else:
                newContent += line
        oProjectFile.close()

        dirname = os.path.dirname(filepathProject)
        basefilename = os.path.splitext(filepathProject)[0]
        newname = os.path.join(dirname, basefilename + '-backup.qgs')
        if os.path.exists(newname):
            os.remove(newname)
        if not os.path.exists(newname):
            os.rename(filepathProject, newname)

        oProjectFileNew = open(filepathProject, "wt")
        oProjectFileNew.write(newContent)
        oProjectFileNew.close()

        return True

    except Exception as e:
        utils.Error('Fejl i InsertReportInProjectfile. Fejlbesked: ' + str(e))
        return False


