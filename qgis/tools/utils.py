# -*- coding: utf-8 -*-
"""
utils
---------
"""

from qgis.PyQt.QtCore import QCoreApplication, QSettings, Qt
from qgis.PyQt.QtWidgets import QMessageBox
from qgis.core import QgsMessageLog, QgsPointXY, QgsDistanceArea, QgsWkbTypes

#---------------------------------------------------------------------------------------------------------    
# Messages
#---------------------------------------------------------------------------------------------------------    
def debug(msg):
    QgsMessageLog.logMessage(msg, "PlanCO2")

def Error(msg):
    QMessageBox.critical(None, "PlanCO2",  unicode(msg))

def Message(msg):
    QMessageBox.information(None, "PlanCO2",  unicode(msg))

def tr(message):
    return QCoreApplication.translate('PlanCO2', message)

def Confirm(msg = None):
    title = "Er du sikker?"
    reply = QMessageBox.question(None, title, msg, QMessageBox.Yes | QMessageBox.No)
    if reply == QMessageBox.Yes:
        return True
    else:
        return False

#---------------------------------------------------------------------------------------------------------    
# Kontrollerer om en string er et tal. Returns True is string is a number. (isnumeric() returnerer true hvis alle tegn er 0-0).
#---------------------------------------------------------------------------------------------------------  
def isNumber(s):
    try:
        float(s)
        return True
    except ValueError:
        return False
def isInt(s):
    try:
        int(s)
        return True
    except ValueError:
        return False
#---------------------------------------------------------------------------------------------------------    
# Settings i registry
#---------------------------------------------------------------------------------------------------------    
def SaveSetting(sname, value):
    s = QSettings()
    s.setValue("UrbanDecarb/" + sname, value)

def ReadStringSetting(sname):
    s = QSettings()
    sValue = s.value("UrbanDecarb/" + sname, "")
    return sValue

def ReadIntSetting(sname):
    s = QSettings()
    val = s.value("UrbanDecarb/" + sname, -1)
    if val == None:
        return -1
    else:
        sValue = int(val)
        return sValue

def ReadFloatSetting(sname):
    s = QSettings()
    val = s.value("UrbanDecarb/" + sname, -1)
    if val == None:
        return -1
    else:
        sValue = float(val)
        return sValue

#---------------------------------------------------------------------------------------------------------    
# Finder et knudepunkt i en polygon
# polyline og knudepunktnummer returneres
#---------------------------------------------------------------------------------------------------------    
def selectVertex(SearchPoint, feat, minDist):

    try:
        retPointIndex = -1

        distance = QgsDistanceArea()
        minDistActive = minDist
        polyline = getPolyLineXYFromFeature(feat)
        #Find den knude, der ligger nærmest mouseclick - og indenfor en radius af 1 meter
        nodeIdx = -1
        i = 0
        for node in polyline:
            if distance.measureLine(SearchPoint, QgsPointXY(node)) < minDistActive:
                minDistActive = distance.measureLine(SearchPoint, QgsPointXY(node))
                nodeIdx = i
            i += 1
        if nodeIdx >= 0:
            retPointIndex = nodeIdx

        return polyline, retPointIndex

    except Exception as e:
        Error("Fejl i selectVertex: " + str(e))

#---------------------------------------------------------------------------------------------------------    
# Returnerer list med QgsPointXY for linier/polygoner
#---------------------------------------------------------------------------------------------------------    
def getPolyLineXYFromFeature(feat):

    try:
        polyLineXY = []

        if feat.geometry().isNull():
            Message('Tom/slettet geometri valgt')
        elif feat.geometry().wkbType() == QgsWkbTypes.LineString or feat.geometry().wkbType() == QgsWkbTypes.LineStringZ or feat.geometry().wkbType() == QgsWkbTypes.LineStringZM or feat.geometry().wkbType() == QgsWkbTypes.LineString25D:
            polyLineXY = feat.geometry().asPolyline() #returnerer QgsPolylineXY
        elif feat.geometry().wkbType() == QgsWkbTypes.MultiLineString or feat.geometry().wkbType() == QgsWkbTypes.MultiLineStringZ or feat.geometry().wkbType() == QgsWkbTypes.MultiLineStringZM or feat.geometry().wkbType() == QgsWkbTypes.MultiLineString25D:
            multiPolyLine = feat.geometry().asMultiPolyline() #returnerer QgsMultiPolylineXY
            for pLine in multiPolyLine:
                for pointXY in pLine:
                    polyLineXY.append(pointXY)
        elif feat.geometry().wkbType() == QgsWkbTypes.Polygon or feat.geometry().wkbType() == QgsWkbTypes.PolygonZ or feat.geometry().wkbType() == QgsWkbTypes.PolygonZM:
            polygon = feat.geometry().asPolygon()
            for pLine in polygon:
                for pointXY in pLine:
                    polyLineXY.append(pointXY)
        elif feat.geometry().wkbType() == QgsWkbTypes.MultiPolygon or feat.geometry().wkbType() == QgsWkbTypes.MultiPolygonZ or feat.geometry().wkbType() == QgsWkbTypes.MultiPolygonZM:
            multiPolygon = feat.geometry().asMultiPolygon()
            for polygon in multiPolygon:
                for pLine in polygon:
                    for pointXY in pLine:
                        polyLineXY.append(pointXY)
        else:
            Message('Ugyldig geometri valgt: ' + QgsWkbTypes.displayString(feat.geometry().wkbType()))

        return polyLineXY

    except Exception as e:
        Error("Fejl i getPolyLineXYFromFeature: " + str(e))
        return []

#---------------------------------------------------------------------------------------------------------   
# 
#---------------------------------------------------------------------------------------------------------   
