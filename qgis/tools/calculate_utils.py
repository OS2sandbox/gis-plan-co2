# -*- coding: utf-8 -*-
"""
utils
---------
"""
from ctypes import util
from qgis.PyQt.QtCore import Qt
from qgis.PyQt.QtWidgets import QApplication, QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox
from qgis.core import QgsFeatureRequest

from . import utils
from . import plotly_utils

#---------------------------------------------------------------------------------------------------------    
#
#---------------------------------------------------------------------------------------------------------    
# def ValidatePip():

#     try:
#         import pip
#         return True

#     except Exception as e:
#         utils.Error('En eller flere python-packages skal installeres. Disse kan ikke installeres, da der ikke er installeret "pip".')
#         return False

#---------------------------------------------------------------------------------------------------------    
#
#---------------------------------------------------------------------------------------------------------    
# def ValidatePackages():

#     try:
#         from qgis.PyQt.QtWidgets import QDockWidget
#         from qgis.utils import iface #change made to original code
#         pythonConsole = iface.mainWindow().findChild(QDockWidget, 'PythonConsole')
    
#         try:
#             from tabulate import tabulate
#         except Exception as e:
#             if not ValidatePip():
#                 return False
#             if utils.Confirm('Package "tabulate" skal installeres. Dette kan gøres ved at afvikle denne kommando i Python Console: pip.main(["install","tabulate"])\n\nHvis du åbner Python Console, kan vi gøre det nu. Ønsker du at forsøge at installere tabulate nu?'):
#                 try:
#                     if not pythonConsole or not pythonConsole.isVisible():
#                         iface.actionShowPythonDialog().trigger()
#                     import pip
#                     pip.main(["install","tabulate"])
#                 except Exception as e:
#                     return False             
#             else:
#                 return False

#         try:
#             import joblib
#         except Exception as e:
#             if not ValidatePip():
#                 return False
#             if utils.Confirm('Package "joblib" skal installeres. Dette kan gøres ved at afvikle denne kommando i Python Console: pip.main(["install","joblib"])\n\nHvis du åbner Python Console, kan vi gøre det nu. Ønsker du at forsøge at installere joblib nu?'):
#                 try:
#                     if not pythonConsole or not pythonConsole.isVisible():
#                         iface.actionShowPythonDialog().trigger()
#                     import pip
#                     pip.main(["install","joblib"])
#                 except Exception as e:
#                     return False             
#             else:
#                 return False

#         try:
#             import sklearn
#         except Exception as e:
#             if not ValidatePip():
#                 return False
#             if utils.Confirm('Package "sklearn" skal installeres. Dette kan gøres ved at afvikle denne kommando i Python Console: pip.main(["install","scikit-learn"])\n\nHvis du åbner Python Console, kan vi gøre det nu. Ønsker du at forsøge at installere sklearn nu?'):
#                 try:
#                     if not pythonConsole or not pythonConsole.isVisible():
#                         iface.actionShowPythonDialog().trigger()
#                     import pip
#                     pip.main(["install","scikit-learn"])
#                 except Exception as e:
#                     return False             
#             else:
#                 return False

#         try:
#             import tensorflow as tf
#         except Exception as e:
#             if not ValidatePip():
#                 return False
#             if utils.Confirm('Package "tensorflow" skal installeres. Dette kan gøres ved at afvikle denne kommando i Python Console: pip.main(["install","tensorflow"])\n\nHvis du åbner Python Console, kan vi gøre det nu. Ønsker du at forsøge at installere tensorflow nu?'):
#                 try:
#                     if not pythonConsole or not pythonConsole.isVisible():
#                         iface.actionShowPythonDialog().trigger()
#                     import pip
#                     pip.main(["install","tensorflow"])
#                 except Exception as e:
#                     return False             
#             else:
#                 return False

#         return True

#     except Exception as e:
#         pass
            
#---------------------------------------------------------------------------------------------------------    
#
#---------------------------------------------------------------------------------------------------------    
def CalculateBuilding(oDialog, oLayerBuildings, dictTranslate):

    try:

        # if not ValidatePackages():
        #     return -1
        
        oDialog.lblMessage.setStyleSheet('color: black')

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.building import Building

        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -1
        if oDialog.txtConstructionYear.text() == "":
            utils.Error("Der er ikke indtastet årstal for opførelse under indstillinger")
            return -1
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -1
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        construction_year = int(oDialog.txtConstructionYear.text())
        emission_factor = oDialog.cmbEmissionFactor.currentText()
        phases = oDialog.cmbPhases.currentText()
        if oDialog.cmbIncludeDemolition.currentText() == "Ja" or oDialog.cmbIncludeDemolition.currentText() == "Ja/Yes" or oDialog.cmbIncludeDemolition.currentText() == "Yes" or oDialog.cmbIncludeDemolition.currentText() == "yes":
            include_demolition = True
        else:
            include_demolition = False
        sequestration = oDialog.cmbSekvestrering.currentText()
        num_residents = None
        num_workplaces = None
        calculate_people = "Nej/No"
        if oDialog.numResidents.value() > 0:
            num_residents = oDialog.numResidents.value()
            calculate_people = "Ja/Yes"
        if oDialog.numWorkplaces.value() > 0:
            num_workplaces = oDialog.numWorkplaces.value()
            calculate_people = "Ja/Yes"
        
        #Building parameters
        area = 0 #490
        width = 0 #20
        perimeter = 0 #89
        oTable = oDialog.tblBuildings
        idx = oTable.currentRow()
        oArea = oTable.cellWidget(idx, 4)
        if isinstance(oArea, QLineEdit) and oArea.text() != "":
            area = float(oArea.text())
        oPerimeter = oTable.cellWidget(idx, 5)
        if isinstance(oPerimeter, QLineEdit) and oPerimeter.text() != "":
            perimeter = float(oPerimeter.text())
        oWidth = oTable.cellWidget(idx, 6)
        if isinstance(oWidth, QDoubleSpinBox):
            width = float(oWidth.value())
        if area == 0 or width == 0 or perimeter == 0:
            utils.Error('Kan ikke fastlægge grundareal, perimeter eller dybde')
            return -1
        height = oDialog.buildingHeight.value()
        num_floors = oDialog.numberOfFloors.value()
        num_base_floors = oDialog.numberOfBasementFloors.value()
        basement_depth = None
        if oDialog.basementDepth.value() > 0 and num_base_floors > 0:
            basement_depth = oDialog.basementDepth.value()
        parking_basement = "no"
        if oDialog.parkingBasement.isChecked():
            parking_basement = "yes"
        condition = dictTranslate.get(oDialog.cb_Building_condition.currentText().lower())
        if condition == None:
            condition = ""
        typology = dictTranslate.get(oDialog.cb_Building_usage.currentText().lower())
        if typology == None:
            typology = ""
        ground_typology = None
        ground_ftf = None
        if num_floors > 1:
            #If there is only one floor the ground typology and ground floor to floor height should be None
            ground_typology = dictTranslate.get(oDialog.cb_Ground_usage.currentText().lower())
            if ground_typology != None:
                ground_ftf = oDialog.groundFTFHeight.value()
        # num_residents = 0
        # num_workplaces = None
        structure = dictTranslate.get(oDialog.cb_Building_bearing_sys.currentText().lower())
        if structure == None:
            structure = ""
        # Here we consider that the user didn't input anything if the value is None
        wwr = None
        if oDialog.windowPercent.value() > 0:
            wwr = oDialog.windowPercent.value()
        prim_facade = dictTranslate.get(oDialog.cb_Building_facade_prime.currentText().lower())
        if prim_facade == None:
            prim_facade = ""
        sec_facade = dictTranslate.get(oDialog.cb_Building_facade_sek.currentText().lower())
        if sec_facade == None:
            sec_facade = ""
        prim_facade_proportion = oDialog.primFacedePercent.value()
        roof_type = dictTranslate.get(oDialog.cb_Building_roof_type.currentText().lower())
        if roof_type == None:
            roof_type = ""
        roof_angle = None
        if oDialog.roofAngle.value() > 0:
            roof_angle = oDialog.roofAngle.value()
        roof_material = ""
        roof_material = dictTranslate.get(oDialog.cb_Building_roof_material.currentText().lower())
        if roof_material == None:
            roof_material = ""
        heating = dictTranslate.get(oDialog.cb_Building_heating_type.currentText().lower())
        if heating == None:
            heating = ""

        #Error handling    
        if height == 0:
            utils.Error("Bygningshøjde kan ikke være 0")
            return -1
        if condition == "":
            utils.Error("Der er ikke angivet bygningstilstand")
            return -1
        if typology == "":
            utils.Error("Der er ikke angivet anvendelse")
            return -1
        if structure == "":
            utils.Error("Der er ikke angivet bærende system")
            return -1
        if prim_facade == "":
            utils.Error("Der er ikke angivet primær facade")
            return -1
        if sec_facade == "" and prim_facade_proportion < 100:
            utils.Error("Der skal angives sekundær facade, når der er angivet en procent større end 0")
            return -1
        if roof_material == "":
            utils.Error("Der er ikke angivet tagmateriale")
            return -1
        if roof_type == "":
            utils.Error("Der er ikke angivet tagtype")
            return -1

        try:  
            building_1 = Building(study_period, construction_year, emission_factor, phases, calculate_people, include_demolition, num_residents,
                                    num_workplaces,
                                    area, width, perimeter, height, num_floors, num_base_floors, basement_depth, parking_basement,
                                    condition, typology,
                                    ground_typology, ground_ftf,
                                    structure, wwr,
                                    prim_facade, sec_facade, prim_facade_proportion, roof_type, roof_material, roof_angle, heating)
        except Exception as e:
            utils.Message("Uforudset fejl ved beregning af bygning: " + str(e))
            return -1
            
        if building_1.ERROR_MESSAGES != []:
            utils.Error("Fejl ved beregning: " + ''.join(building_1.ERROR_MESSAGES))
            return -1
            
        # Så opdateres laget
        if not oLayerBuildings.isEditable():
            oLayerBuildings.startEditing()
        sId = oTable.item(idx, 0).text()
        if sId != "":
            if not utils.isInt(sId):
                utils.Error("Kan ikke opdatere bygning.")
                return -1
            id = int(sId)
            idx = oLayerBuildings.fields().indexOf("EmissionsTotal")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.TOTAL_EMISSIONS, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsTotalM2")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.TOTAL_EMISSIONS_M2, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsTotalM2Year")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.TOTAL_EMISSIONS_M2_YEAR,2 )))
            idx = oLayerBuildings.fields().indexOf("EmissionsTotalPerson")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.TOTAL_EMISSIONS_PERSON, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsTotalPersonYear")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.TOTAL_EMISSIONS_PERSON_YEAR, 2)))
            #JCN/2024-06-03: Emissions på kategorier
            idx = oLayerBuildings.fields().indexOf("EmissionsFoundation")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_FOUNDATION_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsStructure")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_STRUCTURE_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsEnvelope")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_ENVELOPE_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsWindow")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_WINDOW_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsInterior")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_INTERIOR_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsTechnical")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_TECHNICAL_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsConstruction")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_CONSTRUCTION_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsDemolition")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_DEMOLITION_GWP_TOTAL, 2)))
            idx = oLayerBuildings.fields().indexOf("EmissionsOperational")
            if idx >= 0:
                oLayerBuildings.changeAttributeValue(id, idx, float(round(building_1.EMISSIONS_OPERATIONAL_GWP_TOTAL, 2)))

        QApplication.restoreOverrideCursor()
        oDialog.lblMessage.setText("Beregning gik fint. Totale emissioner: " + format(building_1.TOTAL_EMISSIONS, ",.0f") + " kg CO\u2082e, Emissioner pr. areal: "+ format(building_1.TOTAL_EMISSIONS_M2_YEAR, ",.2f")  + " kg CO\u2082e/m\u00B2/år")
            
    except Exception as e:
        utils.Error("Fejl i CalculateBuilding: " + str(e))
        return -1

    finally:
        QApplication.restoreOverrideCursor()

#---------------------------------------------------------------------------------------------------------   
# 
#---------------------------------------------------------------------------------------------------------   
# region
def CalculateAllBuildings(oDialog, oLayerBuildings, dictTranslate):
        
    try:
        # if not ValidatePackages():
        #     return -1, -1

        oDialog.lblMessage.setStyleSheet('color: black')

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.building import Building
            
        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -1, -1
        if oDialog.txtConstructionYear.text() == "":
            utils.Error("Der er ikke indtastet årstal for opførelse under indstillinger")
            return -1, -1
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -1, -1
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        construction_year = int(oDialog.txtConstructionYear.text())
        emission_factor = oDialog.cmbEmissionFactor.currentText()
        phases = oDialog.cmbPhases.currentText()
        if oDialog.cmbIncludeDemolition.currentText() == "Ja" or oDialog.cmbIncludeDemolition.currentText() == "Ja/Yes" or oDialog.cmbIncludeDemolition.currentText() == "Yes" or oDialog.cmbIncludeDemolition.currentText() == "yes":
            include_demolition = True
        else:
            include_demolition = False
        sequestration = oDialog.cmbSekvestrering.currentText()
        calculate_people = "Ja/Yes"
        # num_residents = 70
        # num_workplaces = None

        idxName = oLayerBuildings.fields().lookupField("id")
        idxAreaCalc = oLayerBuildings.fields().lookupField("AreaCalc")
        idxPerimeterCalc = oLayerBuildings.fields().lookupField("PerimeterCalc")
        idxDepth = oLayerBuildings.fields().lookupField("depth")
        if idxName < 0 or idxAreaCalc < 0 or idxPerimeterCalc < 0 or idxDepth < 0:
            utils.Error('Et af disse felter mangler i laget med bygninger: id, AreaCalc, PerimeterCalc, depth')
            return -1, -1

        sumTotalEmissions = 0
        #JCN/2024-07-30: Man kan ikke lægge gennemsnitsdata sammen på denne måde
        # sumTotalEmissionsM2Year = 0
        # sumTotalEmissionsPersonYear = 0
        sumFloorArea = 0
        for feat in oLayerBuildings.getFeatures():
            oDialog.lblMessage.setText("Beregner bygning " + feat[idxName])
            QApplication.processEvents() 

            if feat[idxDepth] != None and feat[idxDepth] > 0:
                width = feat[idxDepth]
            else:
                width = 0
                if feat[idxAreaCalc] != None and feat[idxPerimeterCalc] != None:
                    omb = feat.geometry().orientedMinimumBoundingBox()
                    width = round(omb[-2], 2)
                    length = round(omb[-1], 2)
                    if width > length:
                        width, length = length, width

            #Building parameters
            area = feat[idxAreaCalc]
            perimeter = feat[idxPerimeterCalc]
            height = feat["Building_height"]
            num_floors = feat["NumberOfFloors"]
            num_base_floors = feat["NumberOfBasementFloors"]
            #Etageareal
            floorarea = area * (num_floors + num_base_floors)
            sumFloorArea += floorarea
            basement_depth = None
            if feat["Basement_depth"] != None and feat["Basement_depth"] > 0 and num_base_floors > 0:
                basement_depth = feat["Basement_depth"]
            parking_basement = "no"
            if feat["parkingBasement"] != None and feat["parkingBasement"] == True:
                parking_basement = "yes"
            condition = ""
            if feat["Building_condition"] != None and feat["Building_condition"] != "":
                condition = dictTranslate.get(feat["Building_condition"].lower())
                if condition == None:
                    condition = ""
            typology = ""
            if feat["Building_usage"] != None and feat["Building_usage"] != "":
                typology = dictTranslate.get(feat["Building_usage"].lower())
                if typology == None:
                    typology = ""
            ground_typology = None
            ground_ftf = None
            if num_floors > 1:
                #If there is only one floor the ground typology and ground floor to floor height should be None
                if feat["Ground_usage"] != None and feat["Ground_usage"] != "":
                    ground_typology = dictTranslate.get(feat["Ground_usage"].lower())
                    if ground_typology != None:
                        ground_ftf = feat["Ground_ftf_height"]
            num_residents = None
            num_workplaces = None
            calculate_people = "Nej/No"
            if feat["numResidents"] != None and feat["numResidents"] > 0:
                num_residents = feat["numResidents"]
                calculate_people = "Ja/Yes"
            if feat["numWorkplaces"] != None and feat["numWorkplaces"] > 0:
                num_workplaces = feat["numWorkplaces"]
                calculate_people = "Ja/Yes"
            structure = ""
            if feat["Building_bearing_sys"] != None and feat["Building_bearing_sys"] != "":
                structure = dictTranslate.get(feat["Building_bearing_sys"].lower())
                if structure == None:
                    structure = ""
            # Here we consider that the user didn't input anything if the value is None
            wwr = None
            if feat["windowPercent"] != 0:
                wwr = feat["windowPercent"]
            prim_facade = ""
            if feat["Building_facade_prime"] != None and feat["Building_facade_prime"] != "":
                prim_facade = dictTranslate.get(feat["Building_facade_prime"].lower())
                if prim_facade == None:
                    prim_facade = ""
            sec_facade = ""
            if feat["Building_facade_sek"] != None and feat["Building_facade_sek"] != "":
                sec_facade = dictTranslate.get(feat["Building_facade_sek"].lower())
                if sec_facade == None:
                    sec_facade = ""
            prim_facade_proportion = feat["primFacedePercent"]
            if feat["Building_roof_type"] != None and feat["Building_roof_type"] != "":
                roof_type = dictTranslate.get(feat["Building_roof_type"].lower())
                if roof_type == None:
                    roof_type = ""
            roof_angle = None
            if feat["Roof_angle"] != None and feat["Roof_angle"] > 0:
                roof_angle = feat["Roof_angle"]
            roof_material = ""
            if feat["Building_roof_material"] != None and feat["Building_roof_material"] != "":
                roof_material = dictTranslate.get(feat["Building_roof_material"].lower())
                if roof_material == None:
                    roof_material = ""
            heating = ""
            if feat["Building_heating_type"] != None and feat["Building_heating_type"] != "":
                heating = dictTranslate.get(feat["Building_heating_type"].lower())
                if heating == None:
                    heating = ""

            if height == 0 or condition == "" or typology == "" or ground_typology == "" or structure == "" or\
                prim_facade == "" or (sec_facade == "" and prim_facade_proportion < 100) or roof_type == "" or roof_material == "":
                pass
            else:
                try:
                    #Error handling
                    if roof_material == "":
                        utils.Error("Fejl ved beregning af bygning '" + feat[idxName] + ": Der er ikke angivet tagmateriale")
                        return -1, -1
                    if roof_type == "":
                        utils.Error("Fejl ved beregning af bygning '" + feat[idxName] + ": Der er ikke angivet tagtype")
                        return -1, -1
                    building_1 = Building(study_period, construction_year, emission_factor, phases, calculate_people, include_demolition, num_residents,
                                            num_workplaces,
                                            area, width, perimeter, height, num_floors, num_base_floors, basement_depth, parking_basement,
                                            condition, typology,
                                            ground_typology, ground_ftf,
                                            structure, wwr,
                                            prim_facade, sec_facade, prim_facade_proportion, roof_type, roof_material, roof_angle, heating)
                except Exception as e:
                    # print("study_period: " + str(study_period))
                    # print("construction_year: " + str(construction_year))
                    # print("emission_factor: " + str(emission_factor))
                    # print("phases: " + str(phases))
                    # print("calculate_people: " + str(calculate_people))
                    # print("include_demolition: " + str(include_demolition))
                    # print("num_residents: " + str(num_residents))
                    # print("num_workplaces: " + str(num_workplaces))
                    # print("area: " + str(area))
                    # print("width: " + str(width))
                    # print("perimeter: " + str(perimeter))
                    # print("height: " + str(height))
                    # print("num_floors: " + str(num_floors))
                    # print("num_base_floors: " + str(num_base_floors))
                    # print("basement_depth: " + str(basement_depth))
                    # print("parking_basement: " + str(parking_basement))
                    # print("condition: " + str(condition))
                    # print("typology: " + str(typology))
                    # print("ground_typology: " + str(ground_typology))
                    # print("ground_ftf: " + str(ground_ftf))
                    # print("structure: " + str(structure))
                    # print("wwr: " + str(wwr))
                    # print("prim_facade: " + str(prim_facade))
                    # print("sec_facade: " + str(sec_facade))
                    # print("prim_facade_proportion: " + str(prim_facade_proportion))
                    # print("roof_type: " + str(roof_type))
                    # print("roof_material: " + str(roof_material))
                    # print("roof_angle: " + str(roof_angle))
                    # print("heating: " + str(heating))
                    utils.Message("Uforudset fejl ved beregning af bygning '" + feat[idxName] + "'. Fejlbesked: " + str(e))
                    return -1, -1
                    
                if building_1.ERROR_MESSAGES != []:
                    utils.Error("Fejl ved beregning af bygning '" + feat[idxName] + "'. Fejlbesked: " + ''.join(building_1.ERROR_MESSAGES))
                    return -1, -1

                sumTotalEmissions += building_1.TOTAL_EMISSIONS
                # sumTotalEmissionsM2Year += building_1.TOTAL_EMISSIONS_M2_YEAR
                # if calculate_people == "Ja/Yes":
                #     sumTotalEmissionsPersonYear += building_1.TOTAL_EMISSIONS_PERSON_YEAR
 
                # oDialog.lblMessage.setText("TOTAL EMISSIONS: " + str(building_1.TOTAL_EMISSIONS))
                # QApplication.processEvents() 
                if not oLayerBuildings.isEditable():
                    oLayerBuildings.startEditing()
                idx = oLayerBuildings.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.TOTAL_EMISSIONS, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.TOTAL_EMISSIONS_M2, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.TOTAL_EMISSIONS_M2_YEAR, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsTotalPerson")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.TOTAL_EMISSIONS_PERSON, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsTotalPersonYear")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.TOTAL_EMISSIONS_PERSON_YEAR, 2)))
                #JCN/2024-06-03: Emissions på kategorier
                idx = oLayerBuildings.fields().indexOf("EmissionsFoundation")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_FOUNDATION_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsStructure")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_STRUCTURE_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsEnvelope")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_ENVELOPE_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsWindow")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_WINDOW_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsInterior")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_INTERIOR_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsTechnical")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_TECHNICAL_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsConstruction")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_CONSTRUCTION_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsDemolition")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_DEMOLITION_GWP_TOTAL, 2)))
                idx = oLayerBuildings.fields().indexOf("EmissionsOperational")
                if idx >= 0:
                    oLayerBuildings.changeAttributeValue(feat.id(), idx, float(round(building_1.EMISSIONS_OPERATIONAL_GWP_TOTAL, 2)))
                        
        QApplication.restoreOverrideCursor()
        
        return sumTotalEmissions, sumFloorArea

    except Exception as e:
        utils.Error("Fejl i CalculateAllBuildings: " + str(e))
        return -1, -1

    finally:
        QApplication.restoreOverrideCursor()
        

# endregion
# region
#---------------------------------------------------------------------------------------------------------    
#
#---------------------------------------------------------------------------------------------------------    
def CalculateOpenSurfacesLine(oDialog, oLayer, dictTranslate):

    try:
        oDialog.lblMessage.setStyleSheet('color: black')

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.landscape import Landscape

        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -1
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -1
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        phases = oDialog.cmbPhases.currentText()
        
        #category_landscape = 'hard surfaces'
        #type_landscape = 'parking space (asphalt)'
        #condition_landscape = 'new'
        #length = None
        #width_landscape = None
        #area = 250

        category_landscape = 'roads and paths'
        oTable = oDialog.tblOpenSurfacesLines
        idx = oTable.currentRow()

        type_landscape = None
        oType = oTable.cellWidget(idx, 4)
        if isinstance(oType, QComboBox):
            type_landscape = dictTranslate.get(oType.currentText().lower())
        if type_landscape == None:
            utils.Message("Kan ikke fastlægge type")
            return -1
        condition_landscape = None
        oCondition = oTable.cellWidget(idx, 5)
        if isinstance(oCondition, QComboBox):
            condition_landscape = dictTranslate.get(oCondition.currentText().lower())
        if condition_landscape == None:
            utils.Message("Kan ikke fastlægge tilstand")
            return -1

        length = None
        width_landscape = None
        area = None
        num_trees = 0
        type_trees = None
        size_trees = "small (< 10m)"
        num_shrubs = 0
        type_shrubs = None
        size_shrubs = "small (< 1m)"
        oLength = oTable.cellWidget(idx, 6)
        if isinstance(oLength, QLineEdit):
            if oLength.text() == "":
                utils.Error("Kan ikke fastlægge længde - vælg eventuelt 'Gem og opdater'")
                return -1
            length = float(oLength.text())
        oWidth = oTable.cellWidget(idx, 7)
        if isinstance(oWidth, QDoubleSpinBox):
            if oWidth.value() == 0:
                utils.Error("Kan ikke fastlægge bredde")
                return -1
            width_landscape = float(oWidth.value())
        if length == None or width_landscape == None:
            utils.Error("Kan ikke fastlægge længde og/eller bredde")
            return -1
        area = length * width_landscape
        oText = oTable.cellWidget(idx, 8)
        if isinstance(oText, QSpinBox) and oText.text() != "":
            num_trees = int(oText.text())
        oCombo = oTable.cellWidget(idx, 9)
        if isinstance(oCombo, QComboBox):
            type_trees = dictTranslate.get(oCombo.currentText().lower())
        oCombo = oTable.cellWidget(idx, 10)
        if isinstance(oCombo, QComboBox) and oCombo.currentText() != "":
            size_trees = dictTranslate.get(oCombo.currentText().lower())
        oText = oTable.cellWidget(idx, 11)
        if isinstance(oText, QSpinBox) and oText.text() != "":
            num_shrubs = int(oText.text())
        oCombo = oTable.cellWidget(idx, 12)
        if isinstance(oCombo, QComboBox):
            type_shrubs = dictTranslate.get(oCombo.currentText().lower())
        oCombo = oTable.cellWidget(idx, 13)
        if isinstance(oCombo, QComboBox) and oCombo.currentText() != "":
            size_shrubs = dictTranslate.get(oCombo.currentText().lower())

        # print("category_landscape: " + category_landscape)
        # print("type_landscape: " + type_landscape)
        # print("condition_landscape: " + condition_landscape)
        # print("length: " + str(length))
        # print("width_landscape: " + str(width_landscape))
        # print("area: " + str(area))
        # print("num_trees: " + str(num_trees))
        # print("type_trees: " + type_trees)
        # print("size_trees: " + size_trees)
        # print("num_shrubs: " + str(num_shrubs))
        # print("type_shrubs: " + type_shrubs)
        # print("size_shrubs: " + size_shrubs)
        try:  
            landscape = Landscape(category_landscape, type_landscape, condition_landscape, length, width_landscape, area,
                                  num_trees, type_trees, size_trees,
                                  num_shrubs, type_shrubs, size_shrubs,
                                  study_period, phases)
        except Exception as e:
            utils.Message("Uforudset fejl ved beregning af landscape: " + str(e))
            return -1

        # Så opdateres laget
        if not oLayer.isEditable():
            oLayer.startEditing()
        sId = oTable.item(idx, 0).text()
        if sId != "":
            if not utils.isInt(sId):
                utils.Error("Kan ikke opdatere vejen.")
                return -1
            id = int(sId)
            idx = oLayer.fields().indexOf("EmissionsType")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, float(round(landscape.EMISSIONS_TYPE, 2)))
            idx = oLayer.fields().indexOf("EmissionsTreesShrubs")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, float(round(landscape.EMISSIONS_TREES_SHRUBS, 2)))
            #JCN/2024-08-05: Nye felter
            EmissionsTotal = landscape.EMISSIONS_TYPE + landscape.EMISSIONS_TREES_SHRUBS
            EmissionsTotalM2 = EmissionsTotal / area
            EmissionsTotalM2Year = EmissionsTotalM2 / study_period
            idx = oLayer.fields().indexOf("EmissionsTotal")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotal, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2Year")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2Year, 2))

        QApplication.restoreOverrideCursor()
        oDialog.lblMessage.setText("Beregning gik fint. Totale emissioner: " + format(EmissionsTotal, ",.0f") + " kg CO\u2082e, Emissioner pr. areal: "+ format(EmissionsTotalM2Year, ",.2f")  + " kg CO\u2082e/m\u00B2/år")

    except Exception as e:
        utils.Error("Fejl i CalculateOpenSurfacesLine: " + str(e))
        return -1

    finally:
        QApplication.restoreOverrideCursor()

#---------------------------------------------------------------------------------------------------------   
# 
#---------------------------------------------------------------------------------------------------------   
def CalculateAllOpenSurfacesLines(oDialog, oLayer, dictTranslate, dict_surface_width):
        
    try:
        oDialog.lblMessage.setStyleSheet('color: black')

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.landscape import Landscape
            
        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -1, -1
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -1, -1
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        phases = oDialog.cmbPhases.currentText()

        idxName = oLayer.fields().lookupField("id")
        idxType = oLayer.fields().lookupField("type")
        idxCondition = oLayer.fields().lookupField("condition")
        idxWidth = oLayer.fields().lookupField("width")
        idxLengthCalc = oLayer.fields().lookupField("LengthCalc")
        if idxName < 0 or idxType < 0 or idxCondition < 0 or idxLengthCalc < 0:
            utils.Error('Et af disse felter mangler i laget med veje: id, type, condition LengthCalc')
            return -1, -1

        if not oLayer.isEditable():
            oLayer.startEditing()

        sumTotalEmissionsType = 0
        sumTotalEmissionsTrees = 0

        category_landscape = 'roads and paths'

        for feat in oLayer.getFeatures():
            length = None
            width_landscape = None
            area = None
            num_trees = 0
            type_trees = None
            size_trees = "small (< 10m)"
            num_shrubs = 0
            type_shrubs = None
            size_shrubs = "small (< 1m)"

            oDialog.lblMessage.setText("Beregner landscape " + feat[idxName])
            QApplication.processEvents() 

            if feat[idxType] == None or feat[idxType] == "" or feat[idxCondition] == None or feat[idxCondition] == "":
                # utils.Message("Fejl ved beregning af bygning '" + feat[idxName] + ": Kan ikke fastlægge tilstand")
                # return
                pass
            else:
                type_landscapeDK = feat[idxType]
                type_landscape = dictTranslate.get(feat[idxType].lower())
                condition_landscape = dictTranslate.get(feat[idxCondition].lower())
                length = feat[idxLengthCalc]
                width_landscape = feat[idxWidth]
                if width_landscape <= 0:
                    width_landscape = float(dict_surface_width[type_landscapeDK])
                area = length * width_landscape
                if length == 0 or width_landscape == 0 or area == 0:
                    utils.Message("Fejl ved beregning af landscape '" + feat[idxName] + "'. Length, width og area kan ikke være 0")
                    return -1, -1

                if feat["num_trees"] != None and feat["num_trees"] != "" and feat["num_trees"] > 0:
                    num_trees = int(feat["num_trees"])
                if feat["type_trees"] != None and feat["type_trees"] != "":
                    type_trees = dictTranslate.get(feat["type_trees"].lower())
                if feat["size_trees"] != None and feat["size_trees"] != "":
                    size_trees = dictTranslate.get(feat["size_trees"].lower())
                if feat["num_shrubs"] != None and feat["num_shrubs"] != "" and feat["num_shrubs"] > 0:
                    num_shrubs = int(feat["num_shrubs"])
                if feat["type_shrubs"] != None and feat["type_shrubs"] != "":
                    type_shrubs = dictTranslate.get(feat["type_shrubs"].lower())
                if feat["size_shrubs"] != None and feat["size_shrubs"] != "":
                    size_shrubs = dictTranslate.get(feat["size_shrubs"].lower())

                try:
                    landscape = Landscape(category_landscape, type_landscape, condition_landscape, length, width_landscape, area,
                                          num_trees, type_trees, size_trees,
                                          num_shrubs, type_shrubs, size_shrubs,
                                          study_period, phases)
                except Exception as e:
                    # print("category_landscape: " + category_landscape)
                    # print("type_landscape: " + type_landscape)
                    # print("condition_landscape: " + condition_landscape)
                    # print("length: " + str(length))
                    # print("width_landscape: " + str(width_landscape))
                    # print("area: " + str(area))
                    # print("num_trees: " + str(num_trees))
                    # print("type_trees: " + type_trees)
                    # print("size_trees: " + size_trees)
                    # print("num_shrubs: " + str(num_shrubs))
                    # print("type_shrubs: " + type_shrubs)
                    # print("size_shrubs: " + size_shrubs)
                    utils.Message("Uforudset fejl ved beregning af landscape '" + feat[idxName] + "'. Fejlbesked: " + str(e))
                    return -1, -1
                    
                sumTotalEmissionsType += landscape.EMISSIONS_TYPE
                sumTotalEmissionsTrees += landscape.EMISSIONS_TREES_SHRUBS
                #print("Emissions based on landscape type: " + str(landscape.EMISSIONS_TYPE))
                #print("Emissions based on trees and shrubs: " + str(landscape.EMISSIONS_TREES_SHRUBS))

                idx = oLayer.fields().indexOf("EmissionsType")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, float(round(landscape.EMISSIONS_TYPE, 2)))
                idx = oLayer.fields().indexOf("EmissionsTreesShrubs")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, float(round(landscape.EMISSIONS_TREES_SHRUBS, 2)))
                #JCN/2024-08-05: Nye felter
                EmissionsTotal = landscape.EMISSIONS_TYPE + landscape.EMISSIONS_TREES_SHRUBS
                EmissionsTotalM2 = EmissionsTotal / area
                EmissionsTotalM2Year = EmissionsTotalM2 / study_period
                idx = oLayer.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, round(EmissionsTotal, 2))
                idx = oLayer.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, round(EmissionsTotalM2, 2))
                idx = oLayer.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, round(EmissionsTotalM2Year, 2))
                        
        QApplication.restoreOverrideCursor()
        
        return sumTotalEmissionsType, sumTotalEmissionsTrees

    except Exception as e:
        utils.Error("Fejl i CalculateAllOpenSurfacesLines: " + str(e))
        return -1, -1

    finally:
        QApplication.restoreOverrideCursor()
        
#---------------------------------------------------------------------------------------------------------    
#
#---------------------------------------------------------------------------------------------------------    
def CalculateOpenSurfacesArea(oDialog, oLayer, dictTranslate):

    try:

        oDialog.lblMessage.setStyleSheet('color: black')

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.landscape import Landscape

        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -1
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -1
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        phases = oDialog.cmbPhases.currentText()

        oTable = oDialog.tblOpenSurfacesAreas
        idx = oTable.currentRow()

        type_landscape = None
        oType = oTable.cellWidget(idx, 4)
        if isinstance(oType, QComboBox):
            type_landscape = dictTranslate.get(oType.currentText().lower())
        if type_landscape == None:
            utils.Message("Kan ikke fastlægge type")
            return -1

        category_landscape = 'hard surfaces'
        if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
            category_landscape = 'green surfaces'
        condition_landscape = None
        oCondition = oTable.cellWidget(idx, 5)
        if isinstance(oCondition, QComboBox):
            condition_landscape = dictTranslate.get(oCondition.currentText().lower())
        if condition_landscape == None:
            utils.Message("Kan ikke fastlægge tilstand")
            return -1

        length = None
        width_landscape = None
        area = None
        num_trees = 0
        type_trees = None
        size_trees = "small (< 10m)"
        num_shrubs = 0
        type_shrubs = None
        size_shrubs = "small (< 1m)"
        oArea = oTable.cellWidget(idx, 6)
        if isinstance(oArea, QLineEdit):
            if oArea.text() == "":
                utils.Error("Kan ikke fastlægge areal - vælg eventuelt 'Gem og opdater'")
                return -1
            area = float(oArea.text())
        if area == None:
            utils.Error("Kan ikke fastlægge areal")
            return -1
        oText = oTable.cellWidget(idx, 7)
        if isinstance(oText, QSpinBox) and oText.text() != "":
            num_trees = int(oText.text())
        oCombo = oTable.cellWidget(idx, 8)
        if isinstance(oCombo, QComboBox):
            type_trees = dictTranslate.get(oCombo.currentText().lower())
        oCombo = oTable.cellWidget(idx, 9)
        if isinstance(oCombo, QComboBox) and oCombo.currentText() != "":
            size_trees = dictTranslate.get(oCombo.currentText().lower())
        oText = oTable.cellWidget(idx, 10)
        if isinstance(oText, QSpinBox) and oText.text() != "":
            num_shrubs = int(oText.text())
        oCombo = oTable.cellWidget(idx, 11)
        if isinstance(oCombo, QComboBox):
            type_shrubs = dictTranslate.get(oCombo.currentText().lower())
        oCombo = oTable.cellWidget(idx, 12)
        if isinstance(oCombo, QComboBox) and oCombo.currentText() != "":
            size_shrubs = dictTranslate.get(oCombo.currentText().lower())

        try:  
            landscape = Landscape(category_landscape, type_landscape, condition_landscape, length, width_landscape, area,
                                  num_trees, type_trees, size_trees,
                                  num_shrubs, type_shrubs, size_shrubs,
                                  study_period, phases)
        except Exception as e:
            utils.Message("Uforudset fejl ved beregning af landscape: " + str(e))
            return -1

        # Så opdateres laget
        if not oLayer.isEditable():
            oLayer.startEditing()
        sId = oTable.item(idx, 0).text()
        if sId != "":
            if not utils.isInt(sId):
                utils.Error("Kan ikke opdatere den åbne overflade.")
                return -1
            id = int(sId)
            idx = oLayer.fields().indexOf("EmissionsType")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, float(round(landscape.EMISSIONS_TYPE, 2)))
            idx = oLayer.fields().indexOf("EmissionsTreesShrubs")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, float(round(landscape.EMISSIONS_TREES_SHRUBS, 2)))
            #JCN/2024-08-05: Nye felter
            EmissionsTotal = landscape.EMISSIONS_TYPE + landscape.EMISSIONS_TREES_SHRUBS
            EmissionsTotalM2 = EmissionsTotal / area
            EmissionsTotalM2Year = EmissionsTotalM2 / study_period
            idx = oLayer.fields().indexOf("EmissionsTotal")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotal, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2Year")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2Year, 2))

        QApplication.restoreOverrideCursor()
        oDialog.lblMessage.setText("Beregning gik fint. Totale emissioner: " + format(EmissionsTotal, ",.0f") + " kg CO\u2082e, Emissioner pr. areal: "+ format(EmissionsTotalM2Year, ",.2f")  + " kg CO\u2082e/m\u00B2/år")

    except Exception as e:
        utils.Error("Fejl i CalculateOpenSurfacesArea: " + str(e))

    finally:
        QApplication.restoreOverrideCursor()

#---------------------------------------------------------------------------------------------------------   
# 
#---------------------------------------------------------------------------------------------------------   
def CalculateAllOpenSurfacesAreas(oDialog, oLayer, dictTranslate):
        
    try:
        oDialog.lblMessage.setStyleSheet('color: black')

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.landscape import Landscape
            
        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -1, -1, -1
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -1, -1, -1
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        phases = oDialog.cmbPhases.currentText()

        idxName = oLayer.fields().lookupField("id")
        idxType = oLayer.fields().lookupField("type")
        idxCondition = oLayer.fields().lookupField("condition")
        idxAreaCalc = oLayer.fields().lookupField("AreaCalc")
        if idxName < 0 or idxType < 0 or idxCondition < 0 or idxAreaCalc < 0:
            utils.Error("Et af disse felter mangler i laget med veje: id, type, condition, AreaCalc")
            return -1, -1, -1

        if not oLayer.isEditable():
            oLayer.startEditing()

        sumHardSurfaces = 0
        sumGreenSurfaces = 0
        sumTrees = 0

        for feat in oLayer.getFeatures():
            length = None
            width_landscape = None
            area = None
            num_trees = 0
            type_trees = None
            size_trees = "small (< 10m)"
            num_shrubs = 0
            type_shrubs = None
            size_shrubs = "small (< 1m)"

            oDialog.lblMessage.setText("Beregner landscape " + feat[idxName])
            QApplication.processEvents() 

            if feat[idxType] == None or feat[idxType] == "" or feat[idxCondition] == None or feat[idxCondition] == "":
                # utils.Message("Fejl ved beregning af bygning '" + feat[idxName] + ": Kan ikke fastlægge tilstand")
                # return
                pass
            else:
                type_landscapeDK = feat[idxType]
                type_landscape = dictTranslate.get(feat[idxType].lower())
                condition_landscape = dictTranslate.get(feat[idxCondition].lower())
                area = feat[idxAreaCalc]
                if area == 0:
                    utils.Message("Fejl ved beregning af landscape '" + feat[idxName] + "'. Area kan ikke være 0")
                    return -1, -1, -1

                category_landscape = 'hard surfaces'
                if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
                    category_landscape = 'green surfaces'

                if feat["num_trees"] != None and feat["num_trees"] != "" and feat["num_trees"] > 0:
                    num_trees = int(feat["num_trees"])
                if feat["type_trees"] != None and feat["type_trees"] != "":
                    type_trees = dictTranslate.get(feat["type_trees"].lower())
                if feat["size_trees"] != None and feat["size_trees"] != "":
                    size_trees = dictTranslate.get(feat["size_trees"].lower())
                if feat["num_shrubs"] != None and feat["num_shrubs"] != "" and feat["num_shrubs"] > 0:
                    num_shrubs = int(feat["num_shrubs"])
                if feat["type_shrubs"] != None and feat["type_shrubs"] != "":
                    type_shrubs = dictTranslate.get(feat["type_shrubs"].lower())
                if feat["size_shrubs"] != None and feat["size_shrubs"] != "":
                    size_shrubs = dictTranslate.get(feat["size_shrubs"].lower())

                try:
                    landscape = Landscape(category_landscape, type_landscape, condition_landscape, length, width_landscape, area,
                                          num_trees, type_trees, size_trees,
                                          num_shrubs, type_shrubs, size_shrubs,
                                          study_period, phases)
                except Exception as e:
                    # print("category_landscape: " + category_landscape)
                    # print("type_landscape: " + type_landscape)
                    # print("condition_landscape: " + condition_landscape)
                    # print("length: " + str(length))
                    # print("width_landscape: " + str(width_landscape))
                    # print("area: " + str(area))
                    # print("num_trees: " + str(num_trees))
                    # print("type_trees: " + type_trees)
                    # print("size_trees: " + size_trees)
                    # print("num_shrubs: " + str(num_shrubs))
                    # print("type_shrubs: " + type_shrubs)
                    # print("size_shrubs: " + size_shrubs)
                    utils.Message("Uforudset fejl ved beregning af landscape '" + feat[idxName] + "'. Fejlbesked: " + str(e))
                    return -1, -1, -1
                    
                if category_landscape == 'hard surfaces':
                    sumHardSurfaces += landscape.EMISSIONS_TYPE
                else:
                    sumGreenSurfaces += landscape.EMISSIONS_TYPE
                sumTrees += landscape.EMISSIONS_TREES_SHRUBS
                #print("Emissions based on landscape type: " + str(landscape.EMISSIONS_TYPE))
                #print("Emissions based on trees and shrubs: " + str(landscape.EMISSIONS_TREES_SHRUBS))

                idx = oLayer.fields().indexOf("EmissionsType")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, float(round(landscape.EMISSIONS_TYPE, 2)))
                idx = oLayer.fields().indexOf("EmissionsTreesShrubs")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, float(round(landscape.EMISSIONS_TREES_SHRUBS, 2)))
                #JCN/2024-08-05: Nye felter
                EmissionsTotal = landscape.EMISSIONS_TYPE + landscape.EMISSIONS_TREES_SHRUBS
                EmissionsTotalM2 = EmissionsTotal / area
                EmissionsTotalM2Year = EmissionsTotalM2 / study_period
                idx = oLayer.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, round(EmissionsTotal, 2))
                idx = oLayer.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, round(EmissionsTotalM2, 2))
                idx = oLayer.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    oLayer.changeAttributeValue(feat.id(), idx, round(EmissionsTotalM2Year, 2))
                        
        QApplication.restoreOverrideCursor()
        
        return sumHardSurfaces, sumGreenSurfaces, sumTrees
    
    except Exception as e:
        utils.Error("Fejl i CalculateAllOpenSurfacesAreas: " + str(e))
        return -1, -1, -1

    finally:
        QApplication.restoreOverrideCursor()
        
#---------------------------------------------------------------------------------------------------------    
# JCN/2024-11-07: Beregner for en belægning (en række i tblSurfaces). 
#                 Bruges både til areas and roads (da arealet er givet)
#---------------------------------------------------------------------------------------------------------    
def CalculateSurface(oDialog, oLayer, row, dictTranslate):

    try:

        oDialog.lblMessage.setStyleSheet('color: black')

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.landscape import Landscape

        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -999
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -999
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        phases = oDialog.cmbPhases.currentText()

        oTable = oDialog.tblSurfaces
        #idx = oTable.currentRow()

        type_landscape = None
        oType = oTable.cellWidget(row, 1)
        if isinstance(oType, QComboBox):
            type_landscape = dictTranslate.get(oType.currentText().lower())
        if type_landscape == None:
            utils.Message("Kan ikke fastlægge belægningstype")
            return -999

        category_landscape = 'hard surfaces'
        if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
            category_landscape = 'green surfaces'
        condition_landscape = None
        oCondition = oTable.cellWidget(row, 2)
        if isinstance(oCondition, QComboBox):
            condition_landscape = dictTranslate.get(oCondition.currentText().lower())
        if condition_landscape == None:
            utils.Message("Kan ikke fastlægge tilstand")
            return -999

        length = None
        width_landscape = None
        area = None
        num_trees = None
        type_trees = None
        size_trees = None #"small (< 10m)"
        num_shrubs = None
        type_shrubs = None
        size_shrubs = None #"small (< 1m)"
        oArea = oTable.cellWidget(row, 4)
        if isinstance(oArea, QLineEdit):
            if oArea.text() == "":
                utils.Error("Kan ikke fastlægge areal - vælg eventuelt 'Gem og opdater'")
                return -999
            area = float(oArea.text())
        if area == None:
            utils.Error("Kan ikke fastlægge areal")
            return -999
        # oText = oTable.cellWidget(row, 7)
        # if isinstance(oText, QSpinBox) and oText.text() != "":
        #     num_trees = int(oText.text())
        # oCombo = oTable.cellWidget(row, 8)
        # if isinstance(oCombo, QComboBox):
        #     type_trees = dictTranslate.get(oCombo.currentText().lower())
        # oCombo = oTable.cellWidget(row, 9)
        # if isinstance(oCombo, QComboBox) and oCombo.currentText() != "":
        #     size_trees = dictTranslate.get(oCombo.currentText().lower())
        # oText = oTable.cellWidget(row, 10)
        # if isinstance(oText, QSpinBox) and oText.text() != "":
        #     num_shrubs = int(oText.text())
        # oCombo = oTable.cellWidget(row, 11)
        # if isinstance(oCombo, QComboBox):
        #     type_shrubs = dictTranslate.get(oCombo.currentText().lower())
        # oCombo = oTable.cellWidget(row, 12)
        # if isinstance(oCombo, QComboBox) and oCombo.currentText() != "":
        #     size_shrubs = dictTranslate.get(oCombo.currentText().lower())

        try:  
            landscape = Landscape(category_landscape, type_landscape, condition_landscape, length, width_landscape, area,
                                  num_trees, type_trees, size_trees,
                                  num_shrubs, type_shrubs, size_shrubs,
                                  study_period, phases)
        except Exception as e:
            utils.Message("Uforudset fejl ved beregning af landscape: " + str(e))
            return -999

        # Så opdateres laget
        if not oLayer.isEditable():
            oLayer.startEditing()
        sId = oTable.item(row, 0).text()
        if sId != "":
            if not utils.isInt(sId):
                utils.Error("Kan ikke opdatere belægningen.")
                return -999
            id = int(sId)
            # idx = oLayer.fields().indexOf("EmissionsType")
            # if idx >= 0:
            #     oLayer.changeAttributeValue(id, idx, float(round(landscape.EMISSIONS_TYPE, 2)))
            # idx = oLayer.fields().indexOf("EmissionsTreesShrubs")
            # if idx >= 0:
            #     oLayer.changeAttributeValue(id, idx, float(round(landscape.EMISSIONS_TREES_SHRUBS, 2)))

            EmissionsTotal = landscape.EMISSIONS_TYPE + landscape.EMISSIONS_TREES_SHRUBS
            EmissionsTotalM2 = EmissionsTotal / area
            EmissionsTotalM2Year = EmissionsTotalM2 / study_period
            idx = oLayer.fields().indexOf("EmissionsTotal")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotal, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2Year")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2Year, 2))

        QApplication.restoreOverrideCursor()
        #oDialog.lblMessage.setText("Beregning gik fint. Totale emissioner: " + format(EmissionsTotal, ",.0f") + " kg CO\u2082e, Emissioner pr. areal: "+ format(EmissionsTotalM2Year, ",.2f")  + " kg CO\u2082e/m\u00B2/år")

        return EmissionsTotal

    except Exception as e:
        utils.Error("Fejl i CalculateSurface: " + str(e))
        return -999

    finally:
        QApplication.restoreOverrideCursor()

#---------------------------------------------------------------------------------------------------------    
# JCN/2024-11-08: Beregner for alle belægninger i alle Subareas. 
#                 Bruges både til areas and roads (da arealet er givet)
#---------------------------------------------------------------------------------------------------------    
def CalculateAllSurfaces(oDialog, layerSubareas, layerSurfaces, layerBuildings, layerOpenSurfacesLines, layerOpenSurfacesAreas, dictTranslate, dict_surface_width):

    try:

        oDialog.lblMessage.setStyleSheet('color: black')

        if layerSubareas == None:
            utils.Message('CalculateAllSurfaces: Der er ikke angivet et lag med delområder.')
            return -999, -999
        if layerSurfaces == None:
            utils.Message('CalculateAllSurfaces: Der er ikke angivet et lag med belægninger.')
            return -999, -999
        if layerBuildings == None:
            utils.Message('CalculateAllSurfaces: Der er ikke angivet et lag med bygninger.')
            return -999, -999
        if layerOpenSurfacesLines == None:
            utils.Message('CalculateAllSurfaces: Der er ikke angivet et lag med åbne overflader (veje).')
            return -999, -999
        if layerOpenSurfacesAreas == None:
            utils.Message('CalculateAllSurfaces: Der er ikke angivet et lag med åbne overflader (andre).')
            return -999, -999

        oDialog.lblMessage.setText("Starter beregning...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents() 

        from ..calculations.landscape import Landscape

        #Settings
        if oDialog.cmbStudyPeriod.currentText() == "":
            utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
            return -999, -999
        if oDialog.cmbPhases.currentText() == "":
            utils.Error("Der er ikke valgt inkluderede livscyklusstadier under indstillinger")
            return -999, -999
        study_period = int(oDialog.cmbStudyPeriod.currentText())
        phases = oDialog.cmbPhases.currentText()

        #Check af felter i Surfaces
        idxAreaName = layerSurfaces.fields().lookupField("AreaName")
        idxType = layerSurfaces.fields().lookupField("SurfaceName")
        idxCondition = layerSurfaces.fields().lookupField("condition")
        idxAreaPercent = layerSurfaces.fields().lookupField("AreaPercent")
        if idxAreaName < 0 or idxType < 0 or idxCondition < 0 or idxAreaPercent < 0:
            utils.Error("Et af disse felter mangler i laget med belægninger: AreaName, SurfaceName, condition, AreaPercent")
            return -999, -999
        #Findes nye felter i Surfaces?
        idxEmTot = layerSurfaces.fields().indexOf("EmissionsTotal")
        idxEmTotM2 = layerSurfaces.fields().indexOf("EmissionsTotalM2")
        idxEmTotM2Year = layerSurfaces.fields().indexOf("EmissionsTotalM2Year")
        if idxEmTot < 0 or idxEmTotM2 < 0 or idxEmTotM2Year < 0:
            utils.Error('Et af disse felter mangler i laget med belægninger: EmissionsTotal, EmissionsTotalM2, EmissionsTotalM2Year')
            return -999, -999
        #Check af felter i Subareas
        idxName = layerSubareas.fields().indexOf("Name")
        idxAreaCalc = layerSubareas.fields().indexOf("AreaCalc")
        if idxName < 0 or idxAreaCalc < 0:
            utils.Error("Et af disse felter mangler i laget med delområder: Name, AreaCalc")
            return -999, -999
        #Findes nye felter i Subareas?
        idxEmTotSubarea = layerSubareas.fields().indexOf("EmissionsTotal")
        idxEmTotM2Subarea = layerSubareas.fields().indexOf("EmissionsTotalM2")
        idxEmTotM2YearSubarea = layerSubareas.fields().indexOf("EmissionsTotalM2Year")
        if idxEmTotSubarea < 0 or idxEmTotM2Subarea < 0 or idxEmTotM2YearSubarea < 0:
            utils.Error('Et af disse felter mangler i laget med delområder: EmissionsTotal, EmissionsTotalM2, EmissionsTotalM2Year')
            return -999, -999

        if not layerSubareas.isEditable():
            layerSubareas.startEditing()
        if not layerSurfaces.isEditable():
            layerSurfaces.startEditing()

        sumHardSurfaces = 0
        sumGreenSurfaces = 0

        #Alle subareas
        for featSubarea in layerSubareas.getFeatures():
            EmissionsTotalSubarea = 0

            AreaName = featSubarea[idxName]
            #Frit areal af delområde
            areaFree = featSubarea[idxAreaCalc]
            #Find alle bygninger som tilhører det valgte delområde
            request = QgsFeatureRequest().setFilterExpression("AreaName = '" + AreaName + "'")
            request.setFlags(QgsFeatureRequest.NoGeometry)
            for featB in layerBuildings.getFeatures(request):
                areaFree -= featB.geometry().area()
            #Find alle OpenSurfacesLines som tilhører det valgte delområde
            request = QgsFeatureRequest().setFilterExpression("AreaName = '" + AreaName + "'")
            request.setFlags(QgsFeatureRequest.NoGeometry)
            for featB in layerOpenSurfacesLines.getFeatures(request):
                width = dict_surface_width[featB["type"]]
                if width > 0:
                    areaFree -= featB.geometry().length() * width
            #Find alle OpenSurfacesAreas som tilhører det valgte delområde
            request = QgsFeatureRequest().setFilterExpression("AreaName = '" + AreaName + "'")
            request.setFlags(QgsFeatureRequest.NoGeometry)
            for featB in layerOpenSurfacesAreas.getFeatures(request):
                areaFree -= featB.geometry().area()

            if areaFree > 0:
                #Alle belægninger som tilhører det valgte delområde gennemløbes og beregnes
                request = QgsFeatureRequest().setFilterExpression("AreaName = '" + AreaName + "'")
                request.setFlags(QgsFeatureRequest.NoGeometry)
                for feat in layerSurfaces.getFeatures(request):
                    area = areaFree * feat[idxAreaPercent]/100
                    length = None
                    width_landscape = None
                    num_trees = None
                    type_trees = None
                    size_trees = None
                    num_shrubs = None
                    type_shrubs = None
                    size_shrubs = None

                    oDialog.lblMessage.setText("Beregner landscape " + feat[idxAreaName] + "/" + feat[idxType])
                    QApplication.processEvents() 

                    if feat[idxType] == None or feat[idxType] == "" or feat[idxCondition] == None or feat[idxCondition] == "":
                        pass
                    else:
                        type_landscapeDK = feat[idxType]
                        type_landscape = dictTranslate.get(feat[idxType].lower())
                        condition_landscape = dictTranslate.get(feat[idxCondition].lower())
                        if area == 0:
                            utils.Message("Fejl ved beregning af landscape '" + feat[idxAreaName] + "'/'" + feat[idxType] + "'. Area kan ikke være 0")
                            return -999, -999

                        category_landscape = 'hard surfaces'
                        if type_landscape == "lawn (maintained)" or type_landscape == "lawn (wild)" or type_landscape == "wetlands" or type_landscape == "forest":
                            category_landscape = 'green surfaces'

                        try:
                            landscape = Landscape(category_landscape, type_landscape, condition_landscape, length, width_landscape, area,
                                                  num_trees, type_trees, size_trees,
                                                  num_shrubs, type_shrubs, size_shrubs,
                                                  study_period, phases)
                        except Exception as e:
                            utils.Message("Uforudset fejl ved beregning af landscape '" + feat[idxName] + "'. Fejlbesked: " + str(e))
                            return -999, -999
                    
                        if category_landscape == 'hard surfaces':
                            sumHardSurfaces += landscape.EMISSIONS_TYPE
                        else:
                            sumGreenSurfaces += landscape.EMISSIONS_TYPE

                        EmissionsTotal = landscape.EMISSIONS_TYPE + landscape.EMISSIONS_TREES_SHRUBS
                        EmissionsTotalM2 = EmissionsTotal / area
                        EmissionsTotalM2Year = EmissionsTotalM2 / study_period
                        idx = layerSurfaces.fields().indexOf("EmissionsTotal")
                        if idx >= 0:
                            layerSurfaces.changeAttributeValue(feat.id(), idx, round(EmissionsTotal, 2))
                        idx = layerSurfaces.fields().indexOf("EmissionsTotalM2")
                        if idx >= 0:
                            layerSurfaces.changeAttributeValue(feat.id(), idx, round(EmissionsTotalM2, 2))
                        idx = layerSurfaces.fields().indexOf("EmissionsTotalM2Year")
                        if idx >= 0:
                            layerSurfaces.changeAttributeValue(feat.id(), idx, round(EmissionsTotalM2Year, 2))
                        #Samlet Emissions for delområde
                        EmissionsTotalSubarea += EmissionsTotal
                        
                #Delområde færdigberegnet - laget opdateres
                EmissionsTotalM2Subarea = EmissionsTotalSubarea / areaFree
                EmissionsTotalM2YearSubarea = EmissionsTotalM2Subarea / study_period
                idx = layerSubareas.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    layerSubareas.changeAttributeValue(featSubarea.id(), idx, round(EmissionsTotalSubarea, 2))
                idx = layerSubareas.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    layerSubareas.changeAttributeValue(featSubarea.id(), idx, round(EmissionsTotalM2Subarea, 2))
                idx = layerSubareas.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    layerSubareas.changeAttributeValue(featSubarea.id(), idx, round(EmissionsTotalM2YearSubarea, 2))

        QApplication.restoreOverrideCursor()
        
        return sumHardSurfaces, sumGreenSurfaces

    except Exception as e:
        utils.Error("Fejl i CalculateAllSurfaces: " + str(e))
        return -999, -999

    finally:
        QApplication.restoreOverrideCursor()

# endregion
#---------------------------------------------------------------------------------------------------------   
# 
#---------------------------------------------------------------------------------------------------------   
