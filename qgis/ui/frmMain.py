# -*- coding: utf-8 -*-
"""
/***************************************************************************
 Main
 ***************************************************************************/
"""

from ctypes import util
import os
import math
from qgis.PyQt import uic
from qgis.PyQt.QtCore import Qt, pyqtSignal, QVariant
from qgis.PyQt.QtWidgets import (
    QDialog,
    QFrame,
    QHeaderView,
    QLineEdit,
    QToolButton,
    QComboBox,
    QTableWidgetItem,
    QAbstractItemView,
    QApplication,
    QSpinBox,
    QDoubleSpinBox,
    QFileDialog,
)
from qgis.PyQt.QtGui import QColor, QIcon, QPixmap
from qgis.gui import QgsMapTool, QgsRubberBand
from qgis.core import (
    QgsProject,
    QgsMapLayerProxyModel,
    QgsRectangle,
    QgsFeature,
    QgsGeometry,
    QgsFeatureSink,
    QgsPointXY,
    QgsWkbTypes,
    QgsVectorLayer,
    QgsProcessingFeatureSourceDefinition,
    QgsFeatureRequest,
    QgsField,
)

import uuid
from ..tools import utils
from ..tools import calculate_utils
from ..tools import plotly_utils

FORM_CLASS, _ = uic.loadUiType(os.path.join(os.path.dirname(__file__), "frmMain.ui"))


# ---------------------------------------------------------------------------------------------------------
#
# ---------------------------------------------------------------------------------------------------------
class MainDialog(QDialog, FORM_CLASS):

    hidden = pyqtSignal()

    def __init__(self, parent):
        super(MainDialog, self).__init__(parent)
        self.setupUi(self)

    def hideEvent(self, ev):
        self.hidden.emit()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.hidden.emit()


# ---------------------------------------------------------------------------------------------------------
#
# ---------------------------------------------------------------------------------------------------------
class MainDialogTool(QgsMapTool, MainDialog):

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def __init__(self, iface):
        self.iface = iface
        QgsMapTool.__init__(self, iface.mapCanvas())
        self.dlgMain = MainDialog(self.iface.mainWindow())
        # area surfaces
        self.dict_open_surfaces = {
            "": "",
            "Cykelsti": "line",
            "Tung vej": "line",
            "Motorvej": "line",
            #"Græsplæne (intensiv håndtering)": "area",
            #"Græsplæne (minimal håndtering)": "area",
            "Græsplæne (vedligeholdt)": "area",
            "Græsplæne (vild)": "area",
            "Mellem vej": "line",
            "Parkeringsplads (asfalt)": "area",
            "Parkeringsplads (fliser)": "area",
            "Parkeringsplads (grus)": "area",
            "Parkeringsplads (permeabel)": "area",
            "Sti": "line",
            "Fortov": "area",
            # "Stauder": "area",
            "Let vej": "line",
            "Vådområder": "area",
            "Skov": "area",
            "Opholdsareal (fliser)": "area",
            "Opholdsareal (asfalt)": "area",
        }
        self.dict_surface_condition = [
            "",
            "Ny",
            "Istandsat",
            "Beholdt",
        ]
        self.dict_surface_treetype = [
            "",
            "Løvfældende",
            "Stedsegrønne",
        ]
        self.dict_surface_treesize = [
            "",
            "Lille (< 10m)",
            "Medium (10 - 15m)",
            "Stor (> 15m)",
        ]
        self.dict_surface_shrubsize = [
            "",
            "Lille (< 1m)",
            "Medium (1 - 2m)",
            "Stor (> 2m)",
        ]
        self.dict_surface_width = {
            "": 0,
            "Cykelsti": 1.8,
            "Tung vej": 7.5,
            "Motorvej": 23,
            "Mellem vej": 6.5,
            "Sti": 2.0,
            "Let vej": 6.5,
        }
        self.WWR_DICT = {
            "detached house": 0.22,
            "terraced house": 0.25,
            "apartment building": 0.30,
            "office": 0.40,
            "school": 0.25,
            "institution": 0.25,
            "retail": 0.60,
            "parking": 0.00,
            "industry": 0.20,
            "transport": 0.20,
            "hotel": 0.30,
            "hospital": 0.30
        }

        self.initGui()
        self.iface.mapCanvas().setMapTool(self)
        self.dlgMain.setVisible(True)

        self.dict_translate = {
            "ny": "new",
            "nybyg": "new",
            "transformation": "transformation",
            "renovering": "renovation",
            "istandsat": "refurbishment",
            "beholdt": "keep",
            "bevaret": "keep",
            "nedrivning": "demolition",
            "enfamiliehus": "detached house",
            "rækkehus": "terraced house",
            "etagebolig": "apartment building",
            "kontor": "office",
            "undervisning": "school",
            "dagsinstitution": "institution",
            "butik": "retail",
            "parkering": "parking",
            "industri": "industry",
            "transport": "transport",
            "hotel": "hotel",
            "sundhedshuse": "hospital",
            "beton (bærende vægge)": "concrete (load bearing)",
            "beton (søjle/bjælke)": "concrete (framed)",
            "hybrid (søjle/bjælke)": "hybrid timber (framed)",
            "træ (bærende vægge)": "timber (load bearing)",
            "træ (søjle/bjælke)": "timber (framed)",
            "ikke bestemt": "not decided",
            "mursten (tung facade)": "brick (heavy facade)",
            "aluminium (let facade)": "aluminium (light facade)",
            "zink (let facade)": "zinc (light facade)",
            "skærmtegl (let facade)": "panel brick (light facade)",
            "fibercement (let facade)": "fibercement (light facade)",
            "træ (let facade)": "timber (light facade)",
            "fladt": "flat",
            "skråt": "steep",
            "tagpap": "bitumen",
            "bitumen": "bitumen",
            "teglsten": "tiles",
            "stål": "steel",
            "fibercement": "fibercement",
            "skifer": "slate",
            "zink": "zinc",
            "fjernvarme": "district heating",
            "el": "electricity",
            "cykelsti": "bike path",
            "tung vej": "heavy road",
            "motorvej": "highway",
            "græsplæne (vedligeholdt)": "lawn (maintained)",
            "græsplæne (vild)": "lawn (wild)",
            "mellem vej": "medium road",
            "parkeringsplads (asfalt)": "parking space (asphalt)",
            "parkeringsplads (fliser)": "parking space (pavement)",
            "parkeringsplads (grus)": "parking space (gravel)",
            "parkeringsplads (permeabel)": "parking space (permeable)",
            "sti": "path",
            "fortov": "pavement",
            # "stauder": "perennials",
            "let vej": "light road",
            "vådområder": "wetlands",
            "skov": "forest",
            "opholdsareal (fliser)": "square (pavement)",
            "opholdsareal (asfalt)": "square (asphalt)",
            "løvfældende": "deciduous", 
            "stedsegrønne": "evergreen", 
            "lille (< 10m)": "small (< 10m)", 
            "medium (10 - 15m)": "medium (10 - 15m)", 
            "stor (> 15m)": "large (> 15m)", 
            "lille (< 1m)": "small (< 1m)", 
            "medium (1 - 2m)": "medium (1 - 2m)", 
            "stor (> 2m)": "large (> 2m)", 
        }

    # ---------------------------------------------------------------------------------------------------------
    # Initialization
    # ---------------------------------------------------------------------------------------------------------
    def initGui(self):

        self.ActiveTool = ""
        self.polyline = []
        self.vertex = -1
        self.dirty = False
        self.isEditing = False
        self.initScrollHeight = self.dlgMain.scrollArea_2.height()

        self.rubberband = QgsRubberBand(
            self.iface.mapCanvas(), QgsWkbTypes.PolygonGeometry
        )
        self.rubberband.setColor(QColor(127, 127, 255, 127))
        self.rubberbandLine = QgsRubberBand(
            self.iface.mapCanvas(), QgsWkbTypes.LineGeometry
        )
        self.rubberbandLine.setColor(QColor(0, 0, 255, 255))

        # ---------------------------------------------------------------------------------------------------------
        # Menu
        # ---------------------------------------------------------------------------------------------------------
        # Set icons
        icon_path = os.path.join(
            os.path.abspath(os.path.join(os.path.dirname(__file__), "..")), "icons"
        )
        self.dlgMain.listMenu.item(0).setIcon(
            QIcon(os.path.join(icon_path, "plan_co2.png"))
        )
        self.dlgMain.listMenu.item(1).setIcon(
            QIcon(os.path.join(icon_path, "settings.png"))
        )
        self.dlgMain.listMenu.item(2).setIcon(
            QIcon(os.path.join(icon_path, "area_parts.png"))
        )
        self.dlgMain.listMenu.item(3).setIcon(
            QIcon(os.path.join(icon_path, "buildings.png"))
        )
        self.dlgMain.listMenu.item(4).setIcon(
            QIcon(os.path.join(icon_path, "open_surfaces.png"))
        )
        self.dlgMain.listMenu.item(5).setIcon(
            QIcon(os.path.join(icon_path, "results.png"))
        )
        self.dlgMain.listMenu.item(6).setIcon(
            QIcon(os.path.join(icon_path, "info.png"))
        )
        # Menu size
        self.dlgMain.listMenu.setMinimumWidth(
            self.dlgMain.listMenu.sizeHintForColumn(0) + 10
        )
        # Initial selection
        self.dlgMain.listMenu.setCurrentRow(0)
        self.dlgMain.stackMain.setCurrentIndex(0)
        # signals
        self.dlgMain.listMenu.currentRowChanged.connect(self.listMenuChanged)

        # ---------------------------------------------------------------------------------------------------------
        # Plan CO2
        # ---------------------------------------------------------------------------------------------------------
        # Add logos
        pixmap_kk = QPixmap(
            os.path.join(os.path.dirname(__file__), "../images/kk.png")
        ).scaledToHeight(98)
        self.dlgMain.lblLogoCopenhagen.setPixmap(pixmap_kk)
        self.dlgMain.lblLogoCopenhagen.setFrameStyle(QFrame.NoFrame)
        pixmap_mk = QPixmap(
            os.path.join(os.path.dirname(__file__), "../images/mk.png")
        ).scaledToHeight(98)
        self.dlgMain.lblLogoMiddelfart.setPixmap(pixmap_mk)
        self.dlgMain.lblLogoMiddelfart.setFrameStyle(QFrame.NoFrame)
        pixmap_hl = QPixmap(
            os.path.join(os.path.dirname(__file__), "../images/hl.png")
        ).scaledToHeight(49)
        self.dlgMain.lblLogoHenningLarsen.setPixmap(pixmap_hl)
        self.dlgMain.lblLogoHenningLarsen.setFrameStyle(QFrame.NoFrame)
        pixmap_ram = QPixmap(
            os.path.join(os.path.dirname(__file__), "../images/ramboll.png")
        ).scaledToWidth(153)
        self.dlgMain.lblLogoRamboll.setPixmap(pixmap_ram)
        self.dlgMain.lblLogoRamboll.setFrameStyle(QFrame.NoFrame)
        pixmap_build = QPixmap(
            os.path.join(os.path.dirname(__file__), "../images/AAU_CIRCLE_blue_rgb.png")
        ).scaledToHeight(98)
        self.dlgMain.lblLogoBuild.setPixmap(pixmap_build)
        self.dlgMain.lblLogoBuild.setFrameStyle(QFrame.NoFrame)
        pixmap_plan = QPixmap(
            os.path.join(os.path.dirname(__file__), "../images/plan_co2.png")
        ).scaledToHeight(70)
        self.dlgMain.lblLogoPlan.setPixmap(pixmap_plan)
        self.dlgMain.lblLogoPlan.setFrameStyle(QFrame.NoFrame)


        # ---------------------------------------------------------------------------------------------------------
        # Indstillinger
        # ---------------------------------------------------------------------------------------------------------
        self.layerSubareas = None
        self.layerSurfaces = None
        self.layerBuildings = None
        self.layerOpenSurfacesLines = None
        self.layerOpenSurfacesAreas = None
        self.dlgMain.wInputLayerSubareas.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        self.dlgMain.wInputLayerSurfaces.setFilters(QgsMapLayerProxyModel.NoGeometry)
        self.dlgMain.wInputLayerBuildings.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        self.dlgMain.wInputLayerOpenSurfacesLines.setFilters(
            QgsMapLayerProxyModel.LineLayer
        )
        self.dlgMain.wInputLayerOpenSurfacesAreas.setFilters(
            QgsMapLayerProxyModel.PolygonLayer
        )

        layerNameSubareas = utils.ReadStringSetting("UrbanDecarbLayerSubareas")
        if layerNameSubareas == "":
            self.dlgMain.wInputLayerSubareas.setCurrentIndex(0)
        else:
            index = self.dlgMain.wInputLayerSubareas.findText(
                layerNameSubareas, Qt.MatchFixedString
            )
            if index >= 0:
                self.dlgMain.wInputLayerSubareas.setCurrentIndex(index)
                self.layerSubareas = self.dlgMain.wInputLayerSubareas.currentLayer()
        layerNameSurfaces = utils.ReadStringSetting("UrbanDecarbLayerSurfaces")
        if layerNameSurfaces == "":
            self.dlgMain.wInputLayerSurfaces.setCurrentIndex(0)
        else:
            index = self.dlgMain.wInputLayerSurfaces.findText(
                layerNameSurfaces, Qt.MatchFixedString
            )
            if index >= 0:
                self.dlgMain.wInputLayerSurfaces.setCurrentIndex(index)
                self.layerSurfaces = self.dlgMain.wInputLayerSurfaces.currentLayer()
        layerNameBuildings = utils.ReadStringSetting("UrbanDecarbLayerBuildings")
        if layerNameBuildings == "":
            self.dlgMain.wInputLayerBuildings.setCurrentIndex(0)
        else:
            index = self.dlgMain.wInputLayerBuildings.findText(
                layerNameBuildings, Qt.MatchFixedString
            )
            if index >= 0:
                self.dlgMain.wInputLayerBuildings.setCurrentIndex(index)
                self.layerBuildings = self.dlgMain.wInputLayerBuildings.currentLayer()
        layerNameOpenSurfacesLines = utils.ReadStringSetting(
            "UrbanDecarbLayerOpenSurfacesLines"
        )
        if layerNameOpenSurfacesLines == "":
            self.dlgMain.wInputLayerOpenSurfacesLines.setCurrentIndex(0)
        else:
            index = self.dlgMain.wInputLayerOpenSurfacesLines.findText(
                layerNameOpenSurfacesLines, Qt.MatchFixedString
            )
            if index >= 0:
                self.dlgMain.wInputLayerOpenSurfacesLines.setCurrentIndex(index)
                self.layerOpenSurfacesLines = (
                    self.dlgMain.wInputLayerOpenSurfacesLines.currentLayer()
                )
        layerNameOpenSurfacesAreas = utils.ReadStringSetting("UrbanDecarbLayerOpenSurfacesAreas")
        if layerNameOpenSurfacesAreas == "":
            self.dlgMain.wInputLayerOpenSurfacesAreas.setCurrentIndex(0)
        else:
            index = self.dlgMain.wInputLayerOpenSurfacesAreas.findText(
                layerNameOpenSurfacesAreas, Qt.MatchFixedString
            )
            if index >= 0:
                self.dlgMain.wInputLayerOpenSurfacesAreas.setCurrentIndex(index)
                self.layerOpenSurfacesAreas = (
                    self.dlgMain.wInputLayerOpenSurfacesAreas.currentLayer()
                )

        #Initial values from QGIS settings
        sLokalplan = utils.ReadStringSetting("Lokalplan")
        if sLokalplan != '':
            self.dlgMain.txtLokalplan.setText(sLokalplan)
        sKommune = utils.ReadStringSetting("Kommune")
        if sKommune != '':
            self.dlgMain.txtKommune.setText(sKommune)
        iBeregningsperiode = utils.ReadIntSetting("Beregningsperiode")
        for i in range(self.dlgMain.cmbStudyPeriod.count()):
            if self.dlgMain.cmbStudyPeriod.itemText(i) == str(iBeregningsperiode):
                self.dlgMain.cmbStudyPeriod.setCurrentIndex(i)
                break
        iConstructionYear = utils.ReadIntSetting("ConstructionYear")
        if iConstructionYear > 0:
            self.dlgMain.txtConstructionYear.setText(str(iConstructionYear))
        iTotalNumWorkplaces = utils.ReadIntSetting("TotalNumWorkplaces")
        if iTotalNumWorkplaces > 0:
            self.dlgMain.txtTotalNumWorkplaces.setText(str(iTotalNumWorkplaces))
        iTotalNumResidents = utils.ReadIntSetting("TotalNumResidents")
        if iTotalNumResidents > 0:
            self.dlgMain.txtTotalNumResidents.setText(str(iTotalNumResidents))
        sIncludeDemolition = utils.ReadStringSetting("IncludeDemolition")
        for i in range(self.dlgMain.cmbIncludeDemolition.count()):
            if self.dlgMain.cmbIncludeDemolition.itemText(i) == sIncludeDemolition:
                self.dlgMain.cmbIncludeDemolition.setCurrentIndex(i)
                break
        sEmissionFactor = utils.ReadStringSetting("EmissionFactor")
        for i in range(self.dlgMain.cmbEmissionFactor.count()):
            if self.dlgMain.cmbEmissionFactor.itemText(i) == sEmissionFactor:
                self.dlgMain.cmbEmissionFactor.setCurrentIndex(i)
                break
        sSekvestrering = utils.ReadStringSetting("Sekvestrering")
        for i in range(self.dlgMain.cmbSekvestrering.count()):
            if self.dlgMain.cmbSekvestrering.itemText(i) == sSekvestrering:
                self.dlgMain.cmbSekvestrering.setCurrentIndex(i)
                break
        sPhases = utils.ReadStringSetting("Phases")
        for i in range(self.dlgMain.cmbPhases.count()):
            if self.dlgMain.cmbPhases.itemText(i) == sPhases:
                self.dlgMain.cmbPhases.setCurrentIndex(i)
                break

        # Signals
        self.dlgMain.txtLokalplan.editingFinished.connect(self.lokalplanChanged)
        self.dlgMain.txtKommune.editingFinished.connect(self.kommuneChanged)
        self.dlgMain.cmbStudyPeriod.activated.connect(self.cmbStudyPeriodChanged)
        self.dlgMain.txtConstructionYear.editingFinished.connect(self.constructionYearChanged)
        self.dlgMain.txtTotalNumWorkplaces.editingFinished.connect(self.totalNumWorkplacesChanged)
        self.dlgMain.txtTotalNumResidents.editingFinished.connect(self.totalNumResidentsChanged)
        self.dlgMain.cmbIncludeDemolition.activated.connect(self.cmbIncludeDemolitionChanged)
        self.dlgMain.cmbEmissionFactor.activated.connect(self.cmbEmissionFactorChanged)
        self.dlgMain.cmbSekvestrering.activated.connect(self.cmbSekvestreringChanged)
        self.dlgMain.cmbPhases.activated.connect(self.cmbPhasesChanged)

        self.dlgMain.wInputLayerSubareas.activated.connect(
            self.wInputLayerSubareasChanged
        )
        self.dlgMain.wInputLayerSurfaces.activated.connect(
            self.wInputLayerSurfacesChanged
        )
        self.dlgMain.wInputLayerBuildings.activated.connect(
            self.wInputLayerBuildingsChanged
        )
        self.dlgMain.wInputLayerOpenSurfacesLines.activated.connect(
            self.wInputLayerOpenSurfacesLinesChanged
        )
        self.dlgMain.wInputLayerOpenSurfacesAreas.activated.connect(
            self.wInputLayerOpenSurfacesAreasChanged
        )

        # ---------------------------------------------------------------------------------------------------------
        # Delområder
        # ---------------------------------------------------------------------------------------------------------
        self.sortColSubareas = 2
        self.sortOrderSubareas = 0
        if self.layerSubareas != None:
            if self.layerSubareas.fields().indexOf("AreaCalc") < 0:
                self.layerSubareas.addExpressionField(
                    " area($geometry) ",
                    QgsField("AreaCalc", QVariant.Double, "", 12, 2),
                )

        self.sortColSurfaces = 1
        self.sortOrderSurfaces = 0

        # Signals
        self.dlgMain.tblSubareas.keyPressEvent = self.keyPressEventSubareas
        self.dlgMain.tblSubareas.verticalHeader().sectionClicked.connect(
            self.rowClickedSubareas
        )
        self.dlgMain.tblSubareas.selectionModel().currentRowChanged.connect(
            self.rowChangedSubareas
        )
        self.dlgMain.btnSaveSubareas.clicked.connect(self.SaveSubareas)
        self.dlgMain.btnRollbackSubareas.clicked.connect(self.RollbackSubareas)
        # Knapperne vises ikke
        #self.dlgMain.btnSaveSubareas.setVisible(False)
        self.dlgMain.btnRollbackSubareas.setVisible(False)
        # Surfaces
        self.dlgMain.tblSurfaces.keyPressEvent = self.keyPressEventSurfaces
        self.dlgMain.tblSurfaces.verticalHeader().sectionClicked.connect(
            self.rowClickedSurfaces
        )
        self.dlgMain.tblSurfaces.selectionModel().currentRowChanged.connect(
            self.rowChangedSurfaces
        )

        # ---------------------------------------------------------------------------------------------------------
        # Bygninger
        # ---------------------------------------------------------------------------------------------------------
        # self.dlgMain.label_13.setVisible(False)
        # self.dlgMain.floorHeight.setVisible(False)
        self.sortColBuildings = 2
        self.sortOrderBuildings = 0
        if self.layerBuildings != None:
            if self.layerBuildings.fields().indexOf("AreaCalc") < 0:
                self.layerBuildings.addExpressionField(
                    " area($geometry) ",
                    QgsField("AreaCalc", QVariant.Double, "", 12, 2),
                )
            if self.layerBuildings.fields().indexOf("PerimeterCalc") < 0:
                self.layerBuildings.addExpressionField(
                    " perimeter($geometry) ",
                    QgsField("PerimeterCalc", QVariant.Double, "", 12, 2),
                )

        # Signals
        self.dlgMain.tblBuildings.keyPressEvent = self.keyPressEventBuildings
        self.dlgMain.tblBuildings.verticalHeader().sectionClicked.connect(
            self.rowClickedBuildings
        )
        self.dlgMain.tblBuildings.selectionModel().currentRowChanged.connect(
            self.rowChangedBuildings
        )
        self.dlgMain.btnCalculateBuilding.clicked.connect(self.CalculateBuilding)
        self.dlgMain.btnSaveBuildings.clicked.connect(self.SaveBuildings)
        self.dlgMain.btnRollbackBuildings.clicked.connect(self.RollbackBuildings)
        self.dlgMain.numberOfFloors.editingFinished.connect(self.calcFloorHeight)
        self.dlgMain.buildingHeight.editingFinished.connect(self.calcFloorHeight)
        self.dlgMain.groundFTFHeight.editingFinished.connect(self.calcFloorHeight)
        self.dlgMain.cb_Building_usage.activated.connect(self.Building_usageChanged) #aktiveres ikke når combobox ændres via script
        self.dlgMain.cb_Ground_usage.activated.connect(self.Ground_usageChanged) #aktiveres ikke når combobox ændres via script
        self.dlgMain.cb_Building_roof_type.activated.connect(self.Building_roof_typeChanged) #aktiveres ikke når combobox ændres via script
        self.dlgMain.numberOfBasementFloors.editingFinished.connect(self.numberOfBasementFloorsChanged)
        self.dlgMain.basementDepth.editingFinished.connect(self.basementDepthChanged)
        self.dlgMain.parkingBasement.stateChanged.connect(self.parkingBasementChanged)
        #self.dlgMain.primFacedePercent.valueChanged.connect(self.primFacedePercentChanged)
        self.dlgMain.sekFacedePercent.valueChanged.connect(self.sekFacedePercentChanged)
        # Fortryd-knappen vises ikke
        self.dlgMain.btnRollbackBuildings.setVisible(False)
        #self.dlgMain.sekFacedePercent.setEnabled(False)

        # ComboBoxes content
        building_conditions_lst = [
            "",
            "Nybyg",
            "Transformation",
            "Renovering",
            "Istandsat",
            "Bevaret",
            "Nedrivning"
        ]
        self.dlgMain.cb_Building_condition.addItems(building_conditions_lst)

        building_usage_lst = [
            "",
            "Enfamiliehus",
            "Rækkehus",
            "Etagebolig",
            "Kontor",
            "Undervisning",
            "Dagsinstitution",
            "Butik",
            "Parkering",
            "Industri",
            "Transport",
            "Hotel",
            "Sundhedshuse",
        ]
        self.dlgMain.cb_Building_usage.addItems(building_usage_lst)
        self.dlgMain.cb_Ground_usage.addItems(building_usage_lst)

        building_bearing_sys_lst = [
            "",
            "Beton (bærende vægge)",
            "Beton (søjle/bjælke)",
            "Hybrid (søjle/bjælke)",
            "Træ (bærende vægge)",
            "Træ (søjle/bjælke)",
            "Ikke bestemt",
        ]
        self.dlgMain.cb_Building_bearing_sys.addItems(building_bearing_sys_lst)

        building_facade_lst = [
            "",
            "Mursten (tung facade)",
            "Aluminium (let facade)",
            "Zink (let facade)", 
            "Skærmtegl (let facade)", 
            "Fibercement (let facade)",
            "Træ (let facade)",
        ]
        self.dlgMain.cb_Building_facade_prime.addItems(building_facade_lst)
        self.dlgMain.cb_Building_facade_sek.addItems(building_facade_lst)

        building_roof_type_lst = ["", "Fladt", "Skråt"]
        self.dlgMain.cb_Building_roof_type.addItems(building_roof_type_lst)

        building_roof_material_lst = ["", "Tagpap", "Teglsten", "Stål", "Fibercement", "Skifer", "Zink"]
        self.dlgMain.cb_Building_roof_material.addItems(building_roof_material_lst)

        building_heating_type_lst = ["", "Fjernvarme", "El", "Ikke bestemt"]
        self.dlgMain.cb_Building_heating_type.addItems(building_heating_type_lst)

        # ---------------------------------------------------------------------------------------------------------
        # Åbne overflader
        # ---------------------------------------------------------------------------------------------------------
        self.sortColOpenSurfacesLines = 2
        self.sortOrderOpenSurfacesLines = 0
        if self.layerOpenSurfacesLines != None:
            if self.layerOpenSurfacesLines.fields().indexOf("LengthCalc") < 0:
                self.layerOpenSurfacesLines.addExpressionField(
                    " length($geometry) ",
                    QgsField("LengthCalc", QVariant.Double, "", 12, 2),
                )
        # Signals
        self.dlgMain.btnCalculateOpenSurfacesLine.clicked.connect(self.CalculateOpenSurfacesLine)
        self.dlgMain.tblOpenSurfacesLines.keyPressEvent = self.keyPressEventOpenSurfacesLines
        self.dlgMain.tblOpenSurfacesLines.verticalHeader().sectionClicked.connect(
            self.rowClickedOpenSurfacesLines
        )
        self.dlgMain.tblOpenSurfacesLines.selectionModel().currentRowChanged.connect(
            self.rowChangedOpenSurfacesLines
        )
        self.dlgMain.btnSaveOpenSurfacesLines.clicked.connect(self.SaveOpenSurfacesLines)

        self.sortColOpenSurfacesAreas = 2
        self.sortOrderOpenSurfacesAreas = 0
        if self.layerOpenSurfacesAreas != None:
            if self.layerOpenSurfacesAreas.fields().indexOf("AreaCalc") < 0:
                self.layerOpenSurfacesAreas.addExpressionField(
                    " area($geometry) ",
                    QgsField("AreaCalc", QVariant.Double, "", 12, 2),
                )
        # Signals
        self.dlgMain.btnCalculateOpenSurfacesArea.clicked.connect(self.CalculateOpenSurfacesArea)
        self.dlgMain.tblOpenSurfacesAreas.keyPressEvent = self.keyPressEventOpenSurfacesAreas
        self.dlgMain.tblOpenSurfacesAreas.verticalHeader().sectionClicked.connect(
            self.rowClickedOpenSurfacesAreas
        )
        self.dlgMain.tblOpenSurfacesAreas.selectionModel().currentRowChanged.connect(
            self.rowChangedOpenSurfacesAreas
        )
        self.dlgMain.btnSaveOpenSurfacesAreas.clicked.connect(self.SaveOpenSurfacesAreas)

        # ---------------------------------------------------------------------------------------------------------
        # Resultater
        # ---------------------------------------------------------------------------------------------------------
        self.sortColBuildingsResult = 1
        self.sortOrderBuildingsResult = 0
        self.sortColRoadsResult = 1
        self.sortOrderRoadsResult = 0
        self.sortColSurfacesResult = 1
        self.sortOrderSurfacesResult = 0
        #Values from QGIS settings
        if utils.ReadIntSetting("SumTotalEmissions") > 0:
            self.dlgMain.result_total_co2.setText(str(int(utils.ReadIntSetting("SumTotalEmissions")/1000)))
        if utils.ReadFloatSetting("SumTotalEmissionsM2Year") > 0:
            self.dlgMain.result_sqm_co2.setText(format(utils.ReadFloatSetting("SumTotalEmissionsM2Year"), ",.2f"))
        if utils.ReadFloatSetting("SumTotalEmissionsPersonYear") > 0:
            self.dlgMain.result_person_co2.setText(format(utils.ReadFloatSetting("SumTotalEmissionsPersonYear"), ",.2f"))

        #Grafer udfyldes
        plotly_utils.initPlotly(self.dlgMain)

        # Signals
        self.dlgMain.tblBuildingsResult.verticalHeader().sectionClicked.connect(self.rowClickedBuildingsResult)
        self.dlgMain.tblBuildingsResult.selectionModel().currentRowChanged.connect(self.rowChangedBuildingsResult)
        self.dlgMain.tblRoadsResult.verticalHeader().sectionClicked.connect(self.rowClickedRoadsResult)
        self.dlgMain.tblRoadsResult.selectionModel().currentRowChanged.connect(self.rowChangedRoadsResult)
        self.dlgMain.tblSurfacesResult.verticalHeader().sectionClicked.connect(self.rowClickedSurfacesResult)
        self.dlgMain.tblSurfacesResult.selectionModel().currentRowChanged.connect(self.rowChangedSurfacesResult)

        self.dlgMain.btnCalculate.clicked.connect(self.Calculate)
        self.dlgMain.btnGogglesM2.clicked.connect(self.GogglesM2)
        self.dlgMain.btnGogglesTot.clicked.connect(self.GogglesTot)
        self.dlgMain.btnExportReport.clicked.connect(self.ExportReport)

        #Carbon Goggles knapper vises ikke
        self.dlgMain.btnGogglesTot.setVisible(False)
        self.dlgMain.btnGogglesM2.setVisible(False)

        # Signals on main form
        self.dlgMain.hidden.connect(self.__onDialogHidden)
        self.deactivated.connect(self.__cleanup)

    # ---------------------------------------------------------------------------------------------------------
    # End Initialization
    # ---------------------------------------------------------------------------------------------------------

    # ---------------------------------------------------------------------------------------------------------
    # Beregning af floor-to-floor height (fra building.py)
    # ---------------------------------------------------------------------------------------------------------
    def calcFloorHeight(self):

        try:

            self.dlgMain.lblMessage.setStyleSheet('color: black')
            self.dlgMain.lblMessage.setText('')

            from ..calculations.const import constants as c
            # Define the path to the JSON file
            import json
            PATH = os.path.join(os.path.dirname(__file__), '../calculations')
            json_file_path = PATH + '/const/build/opbyg_dict.json'

            # Read and parse the JSON file
            with open(json_file_path, "r") as json_file:
                opbygDict = json.load(json_file)

            roof_type = self.dict_translate.get(self.dlgMain.cb_Building_roof_type.currentText().lower())
            roof_material = self.dict_translate.get(self.dlgMain.cb_Building_roof_material.currentText().lower())
            typology = self.dict_translate.get(self.dlgMain.cb_Building_usage.currentText().lower())
            if typology == None:
                return
            ground_typology = self.dict_translate.get(self.dlgMain.cb_Ground_usage.currentText().lower())
            structure = self.dict_translate.get(self.dlgMain.cb_Building_bearing_sys.currentText().lower())
            if structure == None or structure == "not decided":
                structure = c.STRUCTURE_DICT[typology]
            if structure == None:
                return
            width = 0
            oTable = self.dlgMain.tblBuildings
            idx = oTable.currentRow()
            oWidth = oTable.cellWidget(idx, 6)
            if isinstance(oWidth, QDoubleSpinBox):
                width = float(oWidth.value())
            height = self.dlgMain.buildingHeight.value()
            num_floors = self.dlgMain.numberOfFloors.value()
            roof_angle = self.dlgMain.roofAngle.value()
            ground_ftf = self.dlgMain.groundFTFHeight.value()
            ground_ftf = None
            if num_floors > 1:
                #If there is only one floor the ground typology and ground floor to floor height should be None
                ground_typology = self.dict_translate.get(self.dlgMain.cb_Ground_usage.currentText().lower())
                if ground_typology != None:
                    ground_ftf = self.dlgMain.groundFTFHeight.value()

            # region lookup to calculate roof thickness
            struct = structure.split()[0]
            item_roof_thickness = None
            if roof_type != None and roof_material != None:
                if struct in ["concrete", "not"] and typology in ["detached house","terraced house",]:
                    item_roof_thickness = next(
                        (
                            item
                            for item in opbygDict
                            if item["assembly_version"]
                            == "low settlement concrete building "
                            + roof_type
                            + " "
                            + roof_material
                        ),
                        None,
                    )
                elif struct in ["timber", "hybrid"] and typology in ["detached house","terraced house",]:
                    item_roof_thickness = next(
                        (
                            item
                            for item in opbygDict
                            if item["assembly_version"]
                            == "low settlement timber building "
                            + roof_type
                            + " "
                            + roof_material
                        ),
                        None,
                    )
                else:
                    item_roof_thickness = next(
                        (
                            item
                            for item in opbygDict
                            if item["assembly_version"] == roof_type + " " + roof_material
                        ),
                        None,
                    )
            # endregion
            roof_angle = 0 if roof_type == "flat" else roof_angle or c.ROOF_ANGLE[typology]
            if item_roof_thickness == None:
                ROOF_THICKNESS = 0
            else:
                ROOF_THICKNESS = item_roof_thickness["thickness [m]"]
            ROOF_HEIGHT = math.tan(math.radians(roof_angle)) * (width / 2) + ROOF_THICKNESS

            if ground_typology is None:
                ground_typology = typology

            if ground_ftf is not None:
                try:
                    ftf = ((height - ROOF_HEIGHT) - ground_ftf) / (num_floors - 1)
                except ZeroDivisionError:
                    self.dlgMain.lblMessage.setText('The values from floor-to-floor height for ground floor and basement do not match the number of floors.')
                    self.dlgMain.lblMessage.setStyleSheet('color: red')
            else:
                ground_ftf = (height - ROOF_HEIGHT) / num_floors
                ftf = ground_ftf

            # region calculate slab thickness
            slab_thickness = 0

            if typology in [
                "detached house",
                "terraced house",
            ]:
                struct = structure.split()[0]
                if struct in ["concrete", "not"]:
                    slab_item = next(
                        (item for item in opbygDict if item["assembly_id"] == "#o0411"),
                        None,
                    )
                    slab_thickness = slab_item["thickness [m]"]
                else:
                    slab_item = next(
                        (item for item in opbygDict if item["assembly_id"] == "#o0412"),
                        None,
                    )
                    slab_thickness = slab_item["thickness [m]"]
            else:
                slab_thickness = c.SLAB_THICKNESS[structure]

            # endregion
            ftc = ftf - slab_thickness
            ftc_ground = ground_ftf - slab_thickness
            if ftc < c.MIN_FTC[typology] or ftc_ground < c.MIN_FTC[ground_typology]:
                self.dlgMain.lblMessage.setText('For lav loftshøjde.')
                self.dlgMain.lblMessage.setStyleSheet('color: red')
                #utils.Message("For lav loftshøjde.")

            self.dlgMain.groundFTFHeight.setValue(ground_ftf)
            self.dlgMain.floorHeight.setValue(ftf)

            # if self.dlgMain.numberOfFloors.value() <= 1:
            #     self.dlgMain.cb_Ground_usage.setEnabled(False)
            #     self.dlgMain.groundFTFHeight.setEnabled(False)
            #     #self.dlgMain.groundFTFHeight.setValue(0)
            # else:
            #     self.dlgMain.cb_Ground_usage.setEnabled(True)
            #     if self.dlgMain.cb_Ground_usage.currentText() == '':
            #         self.dlgMain.groundFTFHeight.setEnabled(False)
            #     else:
            #         self.dlgMain.groundFTFHeight.setEnabled(True)
            # if (self.dlgMain.numberOfFloors.value() == 0 or self.dlgMain.buildingHeight.value() == 0):
            #     self.dlgMain.floorHeight.setValue(0)
            # elif (self.dlgMain.numberOfFloors.value() == 1):
            #     self.dlgMain.floorHeight.setValue(self.dlgMain.buildingHeight.value())
            # else:
            #     if self.dlgMain.groundFTFHeight.value() == 0:
            #         self.dlgMain.floorHeight.setValue(self.dlgMain.buildingHeight.value() / self.dlgMain.numberOfFloors.value())
            #         # height = self.dlgMain.buildingHeight.value() / self.dlgMain.numberOfFloors.value()
            #         # self.dlgMain.floorHeight.setValue(height)
            #         # self.dlgMain.groundFTFHeight.setValue(height)
            #     else:
            #         self.dlgMain.floorHeight.setValue((self.dlgMain.buildingHeight.value() - self.dlgMain.groundFTFHeight.value()) / (self.dlgMain.numberOfFloors.value() - 1))
            
        except Exception as e:
            utils.Error("Fejl i calcFloorHeight: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    # MENU
    # ---------------------------------------------------------------------------------------------------------
    def listMenuChanged(self, row):
        try:

            self.dlgMain.lblMessage.setStyleSheet('color: black')
            self.dlgMain.lblMessage.setText('')

            #JCN/2024-11-06: Hvis den tidligere side var bygninger gemmes ændringer
            if self.dlgMain.stackMain.currentWidget().objectName() == "page4_buildings" and self.dlgMain.tblBuildings.currentRow() > -1:
                self.saveRowBuildings(self.dlgMain.tblBuildings.currentRow())

            self.dlgMain.stackMain.setCurrentIndex(row)
            if self.dlgMain.stackMain.currentWidget().objectName() == "page3_subareas":
                if self.layerSubareas == None:
                    utils.Message('Der skal vælges et lag med delområder under Indstillinger.')
                    return
                if self.layerSurfaces == None:
                    utils.Message('Der skal vælges et lag med belægninger under Indstillinger.')
                    return
                #Tabel med delområder udfyldes altid da arealer fra bygninger kan være ændret
                if True or self.dlgMain.tblSubareas.rowCount() == 0:
                    self.FillTableSubareas()
                    self.dlgMain.tblSubareas.horizontalHeader().sortIndicatorChanged.connect(
                        self.sortingChangedSubareas
                    )
                    self.dlgMain.tblSurfaces.horizontalHeader().sortIndicatorChanged.connect(
                        self.sortingChangedSurfaces
                    )
                # Udvælgelse i andre lag fjernes og første række vælges
                if self.layerBuildings != None:
                    self.layerBuildings.removeSelection()
                if self.layerOpenSurfacesLines != None:
                    self.layerOpenSurfacesLines.removeSelection()
                if self.layerOpenSurfacesAreas != None:
                    self.layerOpenSurfacesAreas.removeSelection()
                self.dlgMain.tblSubareas.selectRow(0)
                self.selectRowSubareas(0)
            elif self.dlgMain.stackMain.currentWidget().objectName() == "page4_buildings":
                if self.layerBuildings == None:
                    utils.Message('Der skal vælges et lag med bygninger under Indstillinger.')
                    return
                if self.dlgMain.tblBuildings.rowCount() == 0:
                    self.FillTableBuildings()
                    self.dlgMain.tblBuildings.horizontalHeader().sortIndicatorChanged.connect(
                        self.sortingChangedBuildings
                    )
                # Udvælgelse i andre lag fjernes og første række vælges
                if self.layerSubareas != None:
                    self.layerSubareas.removeSelection()
                if self.layerOpenSurfacesLines != None:
                    self.layerOpenSurfacesLines.removeSelection()
                if self.layerOpenSurfacesAreas != None:
                    self.layerOpenSurfacesAreas.removeSelection()
                self.dlgMain.tblBuildings.selectRow(0)
                self.selectRowBuildings(0)  # ellers highlightes bygningen ikke
            elif self.dlgMain.stackMain.currentWidget().objectName() == "page5_opensurfaces":
                if self.layerOpenSurfacesLines == None or self.layerOpenSurfacesAreas == None:
                    utils.Message('Der skal vælges lag med åbne overflader under Indstillinger.')
                    return
                if self.dlgMain.tblOpenSurfacesLines.rowCount() == 0:
                    self.FillTableOpenSurfacesLines()
                    self.dlgMain.tblOpenSurfacesLines.horizontalHeader().sortIndicatorChanged.connect(
                        self.sortingChangedOpenSurfacesLines
                    )
                if self.dlgMain.tblOpenSurfacesAreas.rowCount() == 0:
                    self.FillTableOpenSurfacesAreas()
                    self.dlgMain.tblOpenSurfacesAreas.horizontalHeader().sortIndicatorChanged.connect(
                        self.sortingChangedOpenSurfacesAreas
                    )
                # Udvælgelse i andre lag fjernes og første række vælges
                if self.layerSubareas != None:
                    self.layerSubareas.removeSelection()
                if self.layerBuildings != None:
                    self.layerBuildings.removeSelection()
                self.dlgMain.tblOpenSurfacesLines.selectRow(0)
                self.selectRowOpenSurfacesLines(0)  # ellers highlightes linien ikke
            elif self.dlgMain.stackMain.currentWidget().objectName() == "page6_results":
                if self.dlgMain.tblBuildingsResult.rowCount() == 0:
                    self.FillTableBuildingsResult()
                    self.dlgMain.tblBuildingsResult.horizontalHeader().sortIndicatorChanged.connect(self.sortingChangedBuildingsResult)
                if self.dlgMain.tblRoadsResult.rowCount() == 0:
                    self.FillTableRoadsResult()
                    self.dlgMain.tblRoadsResult.horizontalHeader().sortIndicatorChanged.connect(self.sortingChangedRoadsResult)
                if self.dlgMain.tblSurfacesResult.rowCount() == 0:
                    self.FillTableSurfacesResult()
                    self.dlgMain.tblSurfacesResult.horizontalHeader().sortIndicatorChanged.connect(self.sortingChangedSurfacesResult)

        except Exception as e:
            utils.Error("Fejl i listMenuChanged. Fejlbesked: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    # INDSTILLINGER
    # ---------------------------------------------------------------------------------------------------------
    # region
    def lokalplanChanged(self):
        try:
            utils.SaveSetting("Lokalplan", self.dlgMain.txtLokalplan.text())
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def kommuneChanged(self):
        try:
            utils.SaveSetting("Kommune", self.dlgMain.txtKommune.text())
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def cmbStudyPeriodChanged(self):
        try:
            if self.dlgMain.cmbStudyPeriod.currentText() == "":
                utils.SaveSetting("Beregningsperiode", str(0))
            else:
                utils.SaveSetting("Beregningsperiode", str(self.dlgMain.cmbStudyPeriod.currentText()))
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def constructionYearChanged(self):
        try:
            if self.dlgMain.txtConstructionYear.text().isnumeric():
                utils.SaveSetting("ConstructionYear", self.dlgMain.txtConstructionYear.text())
            else:
                utils.SaveSetting("ConstructionYear", 0)
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def totalNumWorkplacesChanged(self):
        try:
            if self.dlgMain.txtTotalNumWorkplaces.text().isnumeric():
                utils.SaveSetting("TotalNumWorkplaces", self.dlgMain.txtTotalNumWorkplaces.text())
            else:
                utils.SaveSetting("TotalNumWorkplaces", 0)
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def totalNumResidentsChanged(self):
        try:
            if self.dlgMain.txtTotalNumResidents.text().isnumeric():
                utils.SaveSetting("TotalNumResidents", self.dlgMain.txtTotalNumResidents.text())
            else:
                utils.SaveSetting("TotalNumResidents", 0)
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def cmbIncludeDemolitionChanged(self):
        try:
            utils.SaveSetting("IncludeDemolition", str(self.dlgMain.cmbIncludeDemolition.currentText()))
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def cmbEmissionFactorChanged(self):
        try:
            utils.SaveSetting("EmissionFactor", str(self.dlgMain.cmbEmissionFactor.currentText()))
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def cmbSekvestreringChanged(self):
        try:
            utils.SaveSetting("Sekvestrering", str(self.dlgMain.cmbSekvestrering.currentText()))
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def cmbPhasesChanged(self):
        try:
            utils.SaveSetting("Phases", str(self.dlgMain.cmbPhases.currentText()))
        except Exception as e:
            utils.Error("Fejl: " + str(e))

    def wInputLayerSubareasChanged(self, newindex):
        try:
            if self.layerSubareas != self.dlgMain.wInputLayerSubareas.currentLayer():
                self.layerSubareas = self.dlgMain.wInputLayerSubareas.currentLayer()
                utils.SaveSetting(
                    "UrbanDecarbLayerSubareas",
                    self.dlgMain.wInputLayerSubareas.currentText(),
                )

        except Exception as e:
            utils.Error("Fejl i wInputLayerSubareasChanged. Fejlbesked: " + str(e))

    def wInputLayerSurfacesChanged(self, newindex):
        try:
            if (
                self.layerSurfaces == None
                or self.layerSurfaces.name()
                != self.dlgMain.wInputLayerSurfaces.currentText()
            ):
                self.layerSurfaces = self.dlgMain.wInputLayerSurfaces.currentLayer()
                utils.SaveSetting(
                    "UrbanDecarbLayerSurfaces",
                    self.dlgMain.wInputLayerSurfaces.currentText(),
                )

        except Exception as e:
            utils.Error("Fejl i wInputLayerSurfacesChanged. Fejlbesked: " + str(e))

    def wInputLayerBuildingsChanged(self, newindex):
        try:
            if self.layerBuildings != self.dlgMain.wInputLayerBuildings.currentLayer():
                self.layerBuildings = self.dlgMain.wInputLayerBuildings.currentLayer()
                utils.SaveSetting(
                    "UrbanDecarbLayerBuildings",
                    self.dlgMain.wInputLayerBuildings.currentText(),
                )
                self.FillTableBuildings()

        except Exception as e:
            utils.Error("Fejl i wInputLayerBuildingsChanged. Fejlbesked: " + str(e))

    def wInputLayerOpenSurfacesLinesChanged(self, newindex):
        try:
            if (
                self.layerOpenSurfacesLines
                != self.dlgMain.wInputLayerOpenSurfacesLines.currentLayer()
            ):
                self.layerOpenSurfacesLines = (
                    self.dlgMain.wInputLayerOpenSurfacesLines.currentLayer()
                )
                utils.SaveSetting(
                    "UrbanDecarbLayerOpenSurfacesLines",
                    self.dlgMain.wInputLayerOpenSurfacesLines.currentText(),
                )
                self.FillTableOpenSurfacesLines()

        except Exception as e:
            utils.Error(
                "Fejl i wInputLayerOpenSurfacesLinesChanged. Fejlbesked: " + str(e)
            )

    def wInputLayerOpenSurfacesAreasChanged(self, newindex):
        try:
            if (
                self.layerOpenSurfacesAreas
                != self.dlgMain.wInputLayerOpenSurfacesAreas.currentLayer()
            ):
                self.layerOpenSurfacesAreas = (
                    self.dlgMain.wInputLayerOpenSurfacesAreas.currentLayer()
                )
                utils.SaveSetting(
                    "UrbanDecarbLayerOpenSurfacesAreas",
                    self.dlgMain.wInputLayerOpenSurfacesAreas.currentText(),
                )
                self.FillTableOpenSurfacesAreas()

        except Exception as e:
            utils.Error("Fejl i wInputLayerOpenSurfacesAreasChanged. Fejlbesked: " + str(e))

    # endregion
    # ---------------------------------------------------------------------------------------------------------
    # DELOMRÅDER
    # ---------------------------------------------------------------------------------------------------------
    # region

    def FillTableSubareas(self):

        try:
            if self.layerSubareas == None:
                return

            if self.sortColSubareas <= 0:
                return

            oTable = self.dlgMain.tblSubareas  # QTableWidget
            oTable.setRowCount(0)
            oTable.setColumnCount(5)
            oTable.setSortingEnabled(True)
            oTable.sortByColumn(self.sortColSubareas, self.sortOrderSubareas)

            oTable.setHorizontalHeaderLabels(
                ["", "", "Navn", "Samlet areal [m\u00B2]", "Frit areal [m\u00B2]"]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.setColumnHidden(0, True)  # Id skjules
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(1, 50)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            oTable.setWordWrap(True)

            idxName = self.layerSubareas.fields().lookupField("Name")
            idxAreaCalc = self.layerSubareas.fields().lookupField("AreaCalc")
            if idxName < 0 or idxAreaCalc < 0:
                utils.Error('Et af disse felter mangler i laget med delområder: Name, AreaCalc')
                return

            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            i = 0
            featList = []
            colId = 0
            colName = 1
            colAreaCalc = 2
            colAreaFree = 3
            for feat in self.layerSubareas.getFeatures():
                #Beregning af frit areal (samlet areal minus areal af bygninger, veje og åbne overflader i delområdet)
                sAreaFree = ""
                areaFree = feat[idxAreaCalc]
                if feat[idxName] != None:
                    if self.layerBuildings == None:
                        utils.Message('Der skal vælges et lag med bygninger under Indstillinger.')
                    else:
                        #Find alle bygninger som tilhører det valgte delområde
                        request = QgsFeatureRequest().setFilterExpression("AreaName = '" + feat[idxName] + "'")
                        request.setFlags(QgsFeatureRequest.NoGeometry)
                        for featB in self.layerBuildings.getFeatures(request):
                            areaFree -= featB.geometry().area()
                    if self.layerOpenSurfacesLines == None:
                        utils.Message('Der skal vælges lag med åbne overflader (veje) under Indstillinger.')
                    else:
                        #Find alle OpenSurfacesLines som tilhører det valgte delområde
                        request = QgsFeatureRequest().setFilterExpression("AreaName = '" + feat[idxName] + "'")
                        request.setFlags(QgsFeatureRequest.NoGeometry)
                        for featB in self.layerOpenSurfacesLines.getFeatures(request):
                            width = self.dict_surface_width[featB["type"]]
                            if width > 0:
                                areaFree -= featB.geometry().length() * width
                    if self.layerOpenSurfacesAreas == None:
                        utils.Message('Der skal vælges lag med åbne overflader (andre) under Indstillinger.')
                    else:
                        #Find alle OpenSurfacesAreas som tilhører det valgte delområde
                        request = QgsFeatureRequest().setFilterExpression("AreaName = '" + feat[idxName] + "'")
                        request.setFlags(QgsFeatureRequest.NoGeometry)
                        for featB in self.layerOpenSurfacesAreas.getFeatures(request):
                            areaFree -= featB.geometry().area()
                sAreaFree = str(round(areaFree, 2))
                featList.append([feat.id(), feat[idxName], feat[idxAreaCalc], sAreaFree])
            if self.sortColSubareas == colAreaCalc+1 or self.sortColSubareas == colAreaFree+1:
                if self.sortOrderSubareas == 0:
                    featList.sort(key=lambda x:float(x[self.sortColSubareas - 1]), reverse=False)
                else:
                    featList.sort(key=lambda x:float(x[self.sortColSubareas - 1]), reverse=True)
            else:
                if self.sortOrderSubareas == 0:
                    featList.sort(key=lambda x: str(x[self.sortColSubareas - 1]).replace("NULL", "").lower(), reverse=False)
                else:
                    featList.sort(key=lambda x: str(x[self.sortColSubareas - 1]).replace("NULL", "").lower(), reverse=True)

            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Rediger geometri
                oEdit = QToolButton()
                oEdit.setIcon(oIcon)
                oEdit.setToolTip("Rediger geometri")
                oEdit.clicked.connect(self.editSubareaClicked)
                oTable.setCellWidget(i, 1, oEdit)
                # Navn
                oName = QLineEdit()
                if feat[colName] == None:
                    oName.setText("")
                else:
                    oName.setText(feat[colName])
                # oName.textChanged.connect(self.setDirty)
                oTable.setCellWidget(i, 2, oName)
                # Samlet areal
                oAreaCalc = QLineEdit()
                if feat[colAreaCalc] == None:
                    oAreaCalc.setText("")
                else:
                    oAreaCalc.setText(str(feat[colAreaCalc]))
                oAreaCalc.setAlignment(Qt.AlignRight)
                oAreaCalc.setEnabled(False)
                oTable.setCellWidget(i, 3, oAreaCalc)
                # Frit areal
                oAreaFree = QLineEdit()
                if feat[colAreaFree] == None:
                    oAreaFree.setText("")
                else:
                    oAreaFree.setText(feat[colAreaFree])
                oAreaFree.setAlignment(Qt.AlignRight)
                oAreaFree.setEnabled(False)
                oTable.setCellWidget(i, 4, oAreaFree)

                i += 1

            # Der indsættes en tom række
            self.InsertBlankRowSubareas(oTable)

        except Exception as e:
            utils.Error("Fejl i FillTableSubareas. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def on_combobox_changed(self, index, row):
        print(f"Row {row}, Combo box new index: {index}")

    # ---------------------------------------------------------------------------------------------------------
    # Procent ændret - resulterende areal genberegnes
    # ---------------------------------------------------------------------------------------------------------
    def recalcArea(self):
        try:
            oTable = self.dlgMain.tblSurfaces
            # ch = self.sender()
            # ix = oTable.indexAt(ch.parent().parent().pos())
            # idxRow = ix.row()
            idxRow = oTable.currentRow()
            self.dlgMain.lblMessage.setText(str(idxRow))
            #Frit areal af delområde
            freeArea = 0
            oAreaM2 = oTable.cellWidget(idxRow, 4)
            if isinstance(oAreaM2, QLineEdit):
                areaM2 = ""
                if self.dlgMain.tblSubareas.currentRow() >= 0:
                    oAreaFree = self.dlgMain.tblSubareas.cellWidget(self.dlgMain.tblSubareas.currentRow(), 4)
                    if isinstance(oAreaFree, QLineEdit) and oAreaFree.text() != "":
                        freeArea = float(oAreaFree.text())
                    if freeArea > 0:
                        oAreaPercent = oTable.cellWidget(idxRow, 3)
                        if isinstance(oAreaPercent, QLineEdit) and oAreaPercent.text() != "":
                            areaPercent = float(oAreaPercent.text())
                            if areaPercent > 0:
                                areaM2 = str(round((freeArea * areaPercent)/100, 2))
                oAreaM2.setText(areaM2)
    
        except Exception as e:
            utils.Error("Fejl i recalcArea. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def FillTableSurfaces(self, AreaName):

        try:
            if self.layerSurfaces == None:
                return

            if self.sortColSurfaces <= 0:
                return

            oTable = self.dlgMain.tblSurfaces  # QTableWidget
            oTable.setRowCount(0)
            oTable.setColumnCount(5)
            oTable.setSortingEnabled(True)
            oTable.sortByColumn(self.sortColSurfaces, self.sortOrderSurfaces)

            oTable.setHorizontalHeaderLabels(
                ["", "Belægning", "Tilstand", "Procent af området [%]", "Resulterende areal [m\u00B2]"]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setMinimumSectionSize(0)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.horizontalHeader().resizeSection(0, 1)
            oTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(3, 150)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(4, 150)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            # oTable.setColumnHidden(0, True)  # Id skjules
            oTable.setWordWrap(True)
            oTable.setStyleSheet("item {border: 0px; padding: 0px; margin: 0px;}")

            if AreaName == "":
                return

            idxSurfaceName = self.layerSurfaces.fields().lookupField("SurfaceName")
            idxCondition = self.layerSurfaces.fields().lookupField("condition")
            idxAreaPercent = self.layerSurfaces.fields().lookupField("AreaPercent")
            if idxSurfaceName < 0 or idxCondition < 0 or idxAreaPercent < 0:
                utils.Error('Et af disse felter mangler i laget med belægninger: SurfaceName, condition, AreaPercent')
                return
            
            dictSurfaceNames = self.dict_open_surfaces.keys()

            #Frit areal af delområde
            freeArea = 0
            if self.dlgMain.tblSubareas.currentRow() >= 0:
                oAreaFree = self.dlgMain.tblSubareas.cellWidget(self.dlgMain.tblSubareas.currentRow(), 4)
                if isinstance(oAreaFree, QLineEdit) and oAreaFree.text() != "":
                    freeArea = float(oAreaFree.text())
            i = 0
            lstRowHeader = []
            featList = []
            colId = 0
            colSurfaceName = 1
            colCondition = 2
            colAreaPercent = 3
            colAreaM2 = 4
            # Kun rækker som tilhører det valgte delområde
            request = QgsFeatureRequest().setFilterExpression("AreaName = '" + AreaName + "'")
            request.setFlags(QgsFeatureRequest.NoGeometry)
            for feat in self.layerSurfaces.getFeatures(request):
                if freeArea < 0 or feat[idxAreaPercent] == None:
                    sAreaM2 = ""
                else:
                    sAreaM2 = str(round(freeArea * feat[idxAreaPercent]/100, 2))
                featList.append([feat.id(), feat[idxSurfaceName], feat[idxCondition], feat[idxAreaPercent], sAreaM2])
            #Sort
            if self.sortColSurfaces == colAreaPercent or self.sortColSurfaces == colAreaM2:
                if self.sortOrderSurfaces == 0:
                    featList.sort(key=lambda x:float(x[self.sortColSurfaces]), reverse=False)
                else:
                    featList.sort(key=lambda x:float(x[self.sortColSurfaces]), reverse=True)
            else:
                if self.sortOrderSurfaces == 0:
                    featList.sort(key=lambda x: str(x[self.sortColSurfaces]).replace("NULL", "").lower(), reverse=False)
                else:
                    featList.sort(key=lambda x: str(x[self.sortColSurfaces]).replace("NULL", "").lower(), reverse=True)
            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Belægning (SurfaceName)
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                j = 0
                for key in dictSurfaceNames:
                    qCombo.addItem(key, key)
                    if key == feat[colSurfaceName]:
                        qCombo.setCurrentIndex(j)
                    j += 1
                oTable.setCellWidget(i, 1, qCombo)
                # Tilstand
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_condition:
                        qCombo.addItem(key, key)
                        if key == feat[colCondition]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 2, qCombo)
                # Procent af området
                oAreaPercent = QLineEdit()
                if feat[colAreaPercent] == None:
                    oAreaPercent.setText("")
                else:
                    oAreaPercent.setText(str(feat[colAreaPercent]))
                oAreaPercent.setAlignment(Qt.AlignRight)
                oAreaPercent.textChanged.connect(self.recalcArea)
                oTable.setCellWidget(i, 3, oAreaPercent)
                # Resulterende areal
                oAreaM2 = QLineEdit()
                if freeArea < 0 or feat[colAreaPercent] == None:
                    oAreaM2.setText("")
                else:
                    restArea = round(freeArea * feat[colAreaPercent]/100, 2)
                    oAreaM2.setText(str(restArea))
                oAreaM2.setAlignment(Qt.AlignRight)
                oAreaM2.setEnabled(False)
                oTable.setCellWidget(i, 4, oAreaM2)

                i += 1
                lstRowHeader.append("")

            self.checkSurfacePercent(oTable)
            # Der indsættes en tom række
            self.InsertBlankRowSurfaces(oTable)

        except Exception as e:
            utils.Error("Fejl i FillTableSurfaces. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def InsertBlankRowSubareas(self, oTable):

        try:
            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            # Der indsættes en tom række
            i = oTable.verticalHeader().count()
            oTable.insertRow(i)
            # Id
            oId = QTableWidgetItem(-1)
            oId.setFlags(Qt.NoItemFlags)  # not enabled
            oTable.setItem(i, 0, oId)
            # Rediger geometri
            qMapLookup = QToolButton()
            qMapLookup.setIcon(oIcon)
            qMapLookup.setToolTip("Opret geometri")
            qMapLookup.clicked.connect(self.editSubareaClicked)
            oTable.setCellWidget(i, 1, qMapLookup)
            # Navn
            oName = QLineEdit()
            oName.setText("")
            # oName.textChanged.connect(self.setDirty)
            oTable.setCellWidget(i, 2, oName)
            # Samlet areal
            oAreaCalc = QLineEdit()
            oAreaCalc.setText("")
            oAreaCalc.setEnabled(False)
            oTable.setCellWidget(i, 3, oAreaCalc)
            # Frit areal
            oAreaFree = QLineEdit()
            oAreaFree.setText("")
            oAreaFree.setEnabled(False)
            oTable.setCellWidget(i, 4, oAreaFree)

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count() - 1):
                lstRowHeader.append("")
            lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i InsertBlankRowSubareas. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def InsertBlankRowSurfaces(self, oTable):

        try:
            # Der indsættes en tom række
            i = oTable.verticalHeader().count()
            oTable.insertRow(i)
            # Id
            oId = QTableWidgetItem(-1)
            oId.setFlags(Qt.NoItemFlags)  # not enabled
            oTable.setItem(i, 0, oId)
            # Belægning
            qCombo = QComboBox()
            qCombo.addItems(list(sorted(self.dict_open_surfaces.keys())))
            # Connect the signal to a slot if necessary
            # qCombo.currentIndexChanged.connect(lambda index, r=i: self.on_combobox_changed(index, r))
            oTable.setCellWidget(i, 1, qCombo)
            # Tilstand
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_condition:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 2, qCombo)
            # Procent af området
            oAreaPercent = QLineEdit()
            oAreaPercent.setText("")
            oAreaPercent.setAlignment(Qt.AlignRight)
            oTable.setCellWidget(i, 3, oAreaPercent)
            # Resulterende areal
            oAreaM2 = QLineEdit()
            oAreaM2.setText("")
            oAreaM2.setAlignment(Qt.AlignRight)
            oAreaM2.setEnabled(False)
            oTable.setCellWidget(i, 4, oAreaM2)

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count() - 1):
                lstRowHeader.append("")
            lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i InsertBlankRowSubareas. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    # Rediger geometri for delområde
    # ---------------------------------------------------------------------------------------------------------
    def editSubareaClicked(self):

        try:
            self.resetGeometryEdit()

            oTable = self.dlgMain.tblSubareas
            ch = self.sender()
            ix = oTable.indexAt(ch.pos())
            oTable.selectRow(ix.row())

            sId = oTable.item(ix.row(), 0).text()
            if sId == "":
                # Opret ny polygon
                self.ActiveTool = "CreateSubarea"
                self.dlgMain.lblMessage.setText("Tegn nyt delområde i kortet.")
            else:
                # Rediger eksisterende polygon
                if not utils.isInt(sId):
                    utils.Error("Ugyldigt delområde valgt.")
                    return
                id = int(sId)
                self.layerSubareas.selectByIds([id])
                if self.layerSubareas.selectedFeatureCount() == 0:
                    utils.Error("Kunne ikke finde delområde i kortet.")
                elif self.layerSubareas.selectedFeatureCount() == 1:
                    # Hvis polygonen ikke er synlig flyttes kortudsnit til contrum af polygonen
                    polygonGeometry = self.layerSubareas.selectedFeatures()[
                        0
                    ].geometry()
                    if not polygonGeometry.within(
                        QgsGeometry.fromRect(self.iface.mapCanvas().extent())
                    ):
                        self.iface.mapCanvas().setCenter(
                            polygonGeometry.centroid().asPoint()
                        )
                        # self.iface.mapCanvas().setExtent(self.layerSubareas.boundingBoxOfSelected())
                        # self.iface.mapCanvas().refresh()
                    if not self.layerSubareas.isEditable():
                        self.layerSubareas.startEditing()
                    self.ActiveTool = "EditSubarea"
                    self.dlgMain.lblMessage.setText("Vælg vertex og flyt det.")
                else:
                    utils.Error("Flere delområder valgt i kortet.")

        except Exception as e:
            utils.Error("Fejl i editSubareaClicked. Fejlbesked: " + str(e))
            return -1

    # ---------------------------------------------------------------------------------------------------------
    # Der skiftes sortering efter klik i header
    # ---------------------------------------------------------------------------------------------------------
    def sortingChangedSubareas(self, col, sortorder):
        if col <= 0:
            self.dlgMain.tblSubareas.horizontalHeader().setSortIndicatorShown(False)
        else:
            self.sortColSubareas = col
            self.sortOrderSubareas = sortorder
            self.FillTableSubareas()
            self.selectRowSubareas(0)

    def sortingChangedSurfaces(self, col, sortorder):
        if col <= 0:
            self.dlgMain.tblSurfaces.horizontalHeader().setSortIndicatorShown(False)
        else:
            self.sortColSurfaces = col
            self.sortOrderSurfaces = sortorder
            oTable = self.dlgMain.tblSubareas
            if oTable.currentRow() >= 0:
                sId = oTable.item(oTable.currentRow(), 0).text()
                if sId == "":
                    self.FillTableSurfaces("")
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldigt delområde valgt.")
                        return
                    #id = int(sId)
                    AreaName = ""
                    oAreaName = self.dlgMain.tblSubareas.cellWidget(oTable.currentRow(), 2)
                    if isinstance(oAreaName, QLineEdit):
                        AreaName = oAreaName.text()
                    self.FillTableSurfaces(AreaName)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def keyPressEventSubareas(self, e):

        try:
            if e.key() == Qt.Key_Delete:
                oTable = self.dlgMain.tblSubareas
                sId = oTable.item(oTable.currentRow(), 0).text()
                if sId == "":
                    utils.Error("Kan ikke fastlægge rækkens id.")
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldigt delområde valgt.")
                        return
                    id = int(sId)
                    if utils.Confirm("Er du sikker på at du ønsker at slette det valgte delområde og tilhørende belægninger?"):
                        AreaName = ""
                        oAreaName = self.dlgMain.tblSubareas.cellWidget(oTable.currentRow(), 2)
                        if isinstance(oAreaName, QLineEdit):
                            AreaName = oAreaName.text()
                        oTable.removeRow(oTable.currentRow())
                        if not self.layerSubareas.isEditable():
                            self.layerSubareas.startEditing()
                        self.layerSubareas.deleteFeature(id)
                        #Tilhørende belægninger slettes
                        if AreaName != "":
                            if not self.layerSurfaces.isEditable():
                                self.layerSurfaces.startEditing()
                            request = QgsFeatureRequest().setFilterExpression("AreaName = '" + AreaName + "'")
                            request.setFlags(QgsFeatureRequest.NoGeometry)
                            for feat in self.layerSurfaces.getFeatures(request):
                                self.layerSurfaces.deleteFeature(feat.id())

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der klikkes i 'verticalHeader' vælges polygon med delområde og rækken gemmes
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedSubareas(self, index):

        try:
            self.saveRowSubareas(index)
            self.selectRowSubareas(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedSubareas(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            oldrow = indexOld.row()
            if oldrow >= 0:
                self.saveRowSubareas(oldrow)
            self.selectRowSubareas(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Række og polygon highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowSubareas(self, index):

        try:
            oTable = self.dlgMain.tblSubareas
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)
                # Polygon vælges i kortet
                sId = oTable.item(index, 0).text()
                if sId == "":
                    #Tom række
                    self.layerSubareas.removeSelection()
                    self.FillTableSurfaces("")
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldigt delområde valgt.")
                        return
                    id = int(sId)
                    self.layerSubareas.selectByIds([id])
                    # Refresh tblSurfaces
                    #nulstil sortering af belægninger
                    self.sortColSurfaces = 1
                    self.sortOrderSurfaces = 0
                    AreaName = ""
                    oAreaName = self.dlgMain.tblSubareas.cellWidget(index, 2)
                    if isinstance(oAreaName, QLineEdit):
                        AreaName = oAreaName.text()
                    self.FillTableSurfaces(AreaName)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Data gemmes
    # ---------------------------------------------------------------------------------------------------------
    def saveRowSubareas(self, index):

        try:
            # Data fra tabel gemmes
            oTable = self.dlgMain.tblSubareas
            sId = oTable.item(index, 0).text()
            if sId == "":
                return False
            if not utils.isInt(sId):
                utils.Error("Ugyldigt delområde valgt.")
                return False

            id = int(sId)
            if not self.layerSubareas.isEditable():
                self.layerSubareas.startEditing()
            Name = ""
            oName = oTable.cellWidget(index, 2)
            if isinstance(oName, QLineEdit):
                Name = oName.text()
                idx = self.layerSubareas.fields().indexOf("Name")
                if idx >= 0:
                    self.layerSubareas.changeAttributeValue(id, idx, Name)
            if self.layerSubareas.featureCount() > 0:
                #Emission værdier sættes til 0 aht. tematisk kort
                feat = self.layerSubareas.getFeature(id)
                idx = self.layerSubareas.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    if feat["EmissionsTotal"] == None:
                        self.layerSubareas.changeAttributeValue(id, idx, 0)
                idx = self.layerSubareas.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    if feat["EmissionsTotalM2"] == None:
                        self.layerSubareas.changeAttributeValue(id, idx, 0)
                idx = self.layerSubareas.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    if feat["EmissionsTotalM2Year"] == None:
                        self.layerSubareas.changeAttributeValue(id, idx, 0)
            # self.layerSubareas.commitChanges()
            #Data i tblSurfaces gemmes også
            if self.dlgMain.tblSurfaces.currentRow() >= 0:
                self.saveRowSurfaces(self.dlgMain.tblSurfaces.currentRow(), index)

            return True

        except Exception as e:
            utils.Error(e)
            return False

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def SaveSubareas(self, silentMode = False):

        try:
            if self.layerSubareas == None:
                return
            activeRow = -1
            oTable = self.dlgMain.tblSubareas
            if oTable.currentRow() >= 0:
                activeRow = oTable.currentRow()
                if not self.saveRowSubareas(oTable.currentRow()):
                    utils.Message("Den aktive række (delområde) indeholder ingen geometri. Der skal tegnes et område før rækken kan gemmes.")
                    return

            if not silentMode:
                if self.layerSubareas.undoStack().count() == 0:
                    utils.Message("Der er ingen ændringer at gemme.")
                    return
                if not utils.Confirm("Er du sikker på at du vil gemme?"):
                    return
            if self.layerSubareas.isEditable():
                if self.layerSubareas.isModified():
                    saveOkKomp = self.layerSubareas.commitChanges()
                    if not saveOkKomp:
                        if utils.Confirm("Ændringerne i delområder kunne ikke gemmes. Vis fejl?"):
                            utils.Message("\n".join(self.layerSubareas.commitErrors()))

            if self.layerSubareas.selectedFeatureCount() > 0:
                self.layerSubareas.removeSelection()
            self.FillTableSubareas()
            if not silentMode:
                if activeRow > -1:
                    self.dlgMain.tblSubareas.selectRow(activeRow)
                    self.selectRowSubareas(activeRow)
                #JCN/2024-11-07: beregning skal ikke kaldes her - midlertidig løsning
                self.CalculateSubareasFOR_TESTING_PURPOSE()

        except Exception as e:
            utils.Error("Fejl i SaveSubareas: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    # TODO: JCN/2024-11-07: MIDLERTIDIG LØSNING - bruges til hurtig test af beregning
    # ---------------------------------------------------------------------------------------------------------
    def CalculateSubareasFOR_TESTING_PURPOSE(self):

        try:
            oTable = self.dlgMain.tblSubareas
            if oTable.currentRow() < 0:
                utils.Message("Vælg en række som skal beregnes")
                return
            #Der beregnes for alle belægninger tilhørende delområdet
            idxEmTot = self.layerSubareas.fields().indexOf("EmissionsTotal")
            idxEmTotM2 = self.layerSubareas.fields().indexOf("EmissionsTotalM2")
            idxEmTotM2Year = self.layerSubareas.fields().indexOf("EmissionsTotalM2Year")
            if idxEmTot < 0 or idxEmTotM2 < 0 or idxEmTotM2Year < 0:
                utils.Error('Et af disse felter mangler i laget med delområder: EmissionsTotal, EmissionsTotalM2, EmissionsTotalM2Year')
                return
            idxEmTot = self.layerSurfaces.fields().indexOf("EmissionsTotal")
            idxEmTotM2 = self.layerSurfaces.fields().indexOf("EmissionsTotalM2")
            idxEmTotM2Year = self.layerSurfaces.fields().indexOf("EmissionsTotalM2Year")
            if idxEmTot < 0 or idxEmTotM2 < 0 or idxEmTotM2Year < 0:
                utils.Error('Et af disse felter mangler i laget med belægninger: EmissionsTotal, EmissionsTotalM2, EmissionsTotalM2Year')
                return

            sumEmissionsTotal = 0
            for row in range(0, self.dlgMain.tblSurfaces.verticalHeader().count() - 1):
                EmissionsTotal = calculate_utils.CalculateSurface(self.dlgMain, self.layerSurfaces, row, self.dict_translate)
                if EmissionsTotal == -999:
                    return
                sumEmissionsTotal += EmissionsTotal
            #Til sidst beregnes for hele delområdet
            study_period = int(self.dlgMain.cmbStudyPeriod.currentText())
            areafree = None
            sId = oTable.item(oTable.currentRow(), 0).text()
            if sId != "":
                if not utils.isInt(sId):
                    utils.Error("Kan ikke opdatere belægningen.")
                    return
            id = int(sId)
            oAreaFree = oTable.cellWidget(oTable.currentRow(), 4)
            if isinstance(oAreaFree, QLineEdit):
                if oAreaFree.text() == "":
                    utils.Error("Kan ikke fastlægge frit areal.")
                    return -1
                areafree = float(oAreaFree.text())
            if areafree == None:
                utils.Error("Kan ikke fastlægge frit areal")

            oLayer = self.layerSubareas
            if not oLayer.isEditable():
                oLayer.startEditing()
            idx = oLayer.fields().indexOf("EmissionsTotal")
            if idx >= 0:
                oLayer.changeAttributeValue(id, idx, round(sumEmissionsTotal, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2")
            if idx >= 0:
                EmissionsTotalM2 = sumEmissionsTotal / areafree
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2, 2))
            idx = oLayer.fields().indexOf("EmissionsTotalM2Year")
            if idx >= 0:
                EmissionsTotalM2Year = EmissionsTotalM2 / study_period
                oLayer.changeAttributeValue(id, idx, round(EmissionsTotalM2Year, 2))

            self.dlgMain.lblMessage.setText("færdig")

            
        except Exception as e:
            utils.Error("Fejl i CalculateSubareas: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def RollbackSubareas(self):

        try:
            if self.layerSubareas == None:
                return

            if self.layerSubareas.undoStack().count() == 0:
                if self.layerSubareas.isEditable():
                    self.layerSubareas.commitChanges()
            else:
                doContinue = utils.Confirm(
                    "Der er lavet ændringer - er du sikker på at disse ikke skal gemmes?"
                )
                if doContinue:
                    if self.layerSubareas.isEditable():
                        self.layerSubareas.rollBack()

            if self.layerSubareas.selectedFeatureCount() > 0:
                self.layerSubareas.removeSelection()
            self.FillTableSubareas()

        except Exception as e:
            utils.Error("Fejl i RollbackSubareas: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    # Surfaces
    # ---------------------------------------------------------------------------------------------------------
    def keyPressEventSurfaces(self, e):

        try:
            if e.key() == Qt.Key_Delete:
                oTable = self.dlgMain.tblSurfaces
                sId = oTable.item(oTable.currentRow(), 0).text()
                if sId == "":
                    utils.Error("Kan ikke fastlægge rækkens id.")
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig række med belægning valgt.")
                        return
                    id = int(sId)
                    if utils.Confirm(
                        "Er du sikker på at du ønsker at slette den valgte belægning?"
                    ):
                        if not self.layerSurfaces.isEditable():
                            self.layerSurfaces.startEditing()
                        self.layerSurfaces.deleteFeature(id)
                        oTable.removeRow(oTable.currentRow())

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der klikkes i 'verticalHeader' gemmes data
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedSurfaces(self, index):

        try:
            self.selectRowSurfaces(index)
            self.saveRowSurfaces(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedSurfaces(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            oldrow = indexOld.row()
            if oldrow >= 0:
                self.saveRowSurfaces(oldrow)
            self.selectRowSurfaces(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Data gemmes
    # ---------------------------------------------------------------------------------------------------------
    def saveRowSurfaces(self, index, idxSubarea = -1):

        try:
            # Data fra tabel gemmes
            oTable = self.dlgMain.tblSurfaces
            # AreaId, AreaName hentes fra self.dlgMain.tblSubareas
            AreaId = 0
            AreaName = ""
            if idxSubarea < 0:
                idxSubarea = self.dlgMain.tblSubareas.currentRow()
            if idxSubarea >= 0:
                sAreaId = self.dlgMain.tblSubareas.item(
                    idxSubarea, 0
                ).text()
                if sAreaId != "":
                    if utils.isInt(sAreaId):
                        AreaId = int(sAreaId)
                oAreaName = self.dlgMain.tblSubareas.cellWidget(idxSubarea, 2)
                if isinstance(oAreaName, QLineEdit):
                    AreaName = oAreaName.text()
            if AreaId == 0:
                utils.Error("Kunne ikke finde informationer om delområde - der kan ikke gemmes.")
                return
            if AreaName == "":
                utils.Error("Der er ikke angivet navn på delområde - der kan ikke gemmes.")
                return
            # Værdier hentes fra tabel
            SurfaceName = ""
            Condition = ""
            AreaPercent = ""
            AreaFree = ""
            oSurfaceName = oTable.cellWidget(index, 1)
            if isinstance(oSurfaceName, QComboBox):
                SurfaceName = oSurfaceName.currentText()
            oCondition = oTable.cellWidget(index, 2)
            if isinstance(oCondition, QComboBox):
                Condition = oCondition.currentText()
            oAreaPercent = oTable.cellWidget(index, 3)
            if isinstance(oAreaPercent, QLineEdit):
                AreaPercent = oAreaPercent.text()
            oAreaFree = oTable.cellWidget(index, 4)
            if isinstance(oAreaFree, QLineEdit):
                AreaFree = oAreaFree.text()

            # Så opdateres laget
            if not self.layerSurfaces.isEditable():
                self.layerSurfaces.startEditing()
            sId = oTable.item(index, 0).text()
            if sId == "":
                # Ny række
                if SurfaceName == "":
                    return
                feat = QgsFeature(self.layerSurfaces.fields())
                # if feat.fieldNameIndex("AreaId") >= 0:
                #     feat.setAttribute("AreaId", AreaId)
                if feat.fieldNameIndex("AreaName") >= 0:
                    feat.setAttribute("AreaName", AreaName)
                if feat.fieldNameIndex("SurfaceName") >= 0:
                    feat.setAttribute("SurfaceName", SurfaceName)
                if feat.fieldNameIndex("condition") >= 0:
                    feat.setAttribute("condition", Condition)
                if feat.fieldNameIndex("AreaPercent") >= 0 and AreaPercent != "":
                    feat.setAttribute("AreaPercent", AreaPercent)
                if feat.fieldNameIndex("AreaFree") >= 0 and AreaFree != "":
                    feat.setAttribute("AreaFree", AreaFree)
                self.layerSurfaces.addFeatures([feat], QgsFeatureSink.FastInsert)
            else:
                # Eksisterende række
                if not utils.isInt(sId):
                    utils.Error("Ugyldig række med belægninger valgt.")
                    return
                id = int(sId)
                # idx = self.layerSurfaces.fields().indexOf("AreaId")
                # if idx >= 0:
                #     self.layerSurfaces.changeAttributeValue(id, idx, AreaId)
                idx = self.layerSurfaces.fields().indexOf("AreaName")
                if idx >= 0:
                    self.layerSurfaces.changeAttributeValue(id, idx, AreaName)
                idx = self.layerSurfaces.fields().indexOf("SurfaceName")
                if idx >= 0:
                    self.layerSurfaces.changeAttributeValue(id, idx, SurfaceName)
                idx = self.layerSurfaces.fields().indexOf("condition")
                if idx >= 0:
                    self.layerSurfaces.changeAttributeValue(id, idx, Condition)
                idx = self.layerSurfaces.fields().indexOf("AreaPercent")
                if idx >= 0:
                    self.layerSurfaces.changeAttributeValue(id, idx, AreaPercent)
                idx = self.layerSurfaces.fields().indexOf("AreaFree")
                if idx >= 0:
                    self.layerSurfaces.changeAttributeValue(id, idx, AreaFree)

            self.layerSurfaces.commitChanges()
            if sId == "":
                self.FillTableSurfaces(AreaName)
                
            self.checkSurfacePercent(oTable)
                        
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Calculate sum of percent
    # ---------------------------------------------------------------------------------------------------------
    def checkSurfacePercent(self, oTable):

        try:
            sumPercent = 0
            for i in range(0, oTable.verticalHeader().count()):
                oAreaPercent = oTable.cellWidget(i, 3)
                if isinstance(oAreaPercent, QLineEdit) and oAreaPercent.text() != "":
                    sumPercent += float(oAreaPercent.text())
            for i in range(0, oTable.verticalHeader().count()):
                oAreaPercent = oTable.cellWidget(i, 3)
                if isinstance(oAreaPercent, QLineEdit):
                    if sumPercent > 100: 
                        oAreaPercent.setStyleSheet("QLineEdit {background-color: yellow;}")
                    else:
                        oAreaPercent.setStyleSheet("QLineEdit {background-color: none;}")
            if sumPercent > 100: 
                self.dlgMain.lblMessage.setStyleSheet('color: red')
                self.dlgMain.lblMessage.setText("Bemærk at der samlet er angivet belægninger for mere end 100% i området")
            elif sumPercent < 100: 
                self.dlgMain.lblMessage.setStyleSheet('color: red')
                self.dlgMain.lblMessage.setText("Bemærk at der ikke er angivet belægninger for hele området")
            else:
                self.dlgMain.lblMessage.setText("")

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Rækken highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowSurfaces(self, index):

        try:
            oTable = self.dlgMain.tblSurfaces
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def getSubareaNames(self):
        """
        Retrieves a sorted list of unique 'Name' attributes from the 'layerSubareas' layer.

        This function queries the 'layerSubareas' layer to collect all distinct values in the
        'Name' field, sorts the list of names, and returns it. If 'layerSubareas' is not set,
        the function returns None. On error, logs the error message and returns None.

        Returns:
            list: A sorted list of unique subarea names, or None if the layer is not set or an error occurs.
        """

        try:
            if self.layerSubareas == None:
                return

            idxName = self.layerSubareas.fields().lookupField("Name")
            if idxName <= 0:
                return

            lstNames = [""]

            for feat in self.layerSubareas.getFeatures():
                z = feat[idxName]
                lstNames.append(str(feat[idxName]))

            lstNames = sorted(list(set(lstNames)))

            return lstNames

        except Exception as e:
            utils.Error(
                "Fejl ved forsøg på at læse delområdenavne. Fejlbesked: " + str(e)
            )
            return

    # endregion
    # region
    # ---------------------------------------------------------------------------------------------------------
    # BYGNINGER
    # ---------------------------------------------------------------------------------------------------------
    # ---------------------------------------------------------------------------------------------------------
    # Vindues procent genberegnes
    # ---------------------------------------------------------------------------------------------------------
    def Building_usageChanged(self):
        
        try:
            typology = self.dict_translate.get(self.dlgMain.cb_Building_usage.currentText().lower())
            if typology == None:
                self.dlgMain.lblStandardWindowPercent.setText("")
            else:
                wwr = int(self.WWR_DICT[typology] * 100)
                self.dlgMain.lblStandardWindowPercent.setText(str(wwr))
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Enable/disable etagehøjde i stueetage
    # ---------------------------------------------------------------------------------------------------------
    def Ground_usageChanged(self):
        
        try:
            if self.dlgMain.cb_Ground_usage.currentText() == '':
                self.dlgMain.groundFTFHeight.setEnabled(False)
                #self.dlgMain.groundFTFHeight.setValue(0)
            else:
                self.dlgMain.groundFTFHeight.setEnabled(True)

            self.calcFloorHeight() 
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Taghældning disables for fladt tag
    # ---------------------------------------------------------------------------------------------------------
    def Building_roof_typeChanged(self):
        
        try:
            if self.dlgMain.cb_Building_roof_type.currentText() == 'Fladt':
                self.dlgMain.roofAngle.setEnabled(False)
            else:
                self.dlgMain.roofAngle.setEnabled(True)
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Kælderhøjde disables hvis ingen kælder
    # ---------------------------------------------------------------------------------------------------------
    def numberOfBasementFloorsChanged(self):
        
        try:
            if (self.dlgMain.numberOfBasementFloors.value() == 0):
                self.dlgMain.basementDepth.setEnabled(False)
                self.dlgMain.basementDepth.setValue(0)
            else:
                self.dlgMain.basementDepth.setEnabled(True)
                standardBasementDepth = round(self.getStandardBasementDepth(), 2)
                self.dlgMain.basementDepth.setValue(standardBasementDepth)
            
            self.checkBasementDepth()

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Check af minimumskælderhøjde
    # ---------------------------------------------------------------------------------------------------------
    def basementDepthChanged(self):

        try:
            self.checkBasementDepth()

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # 
    # ---------------------------------------------------------------------------------------------------------
    def parkingBasementChanged(self):
        
        try:
            self.checkBasementDepth()
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Check af minimumskælderhøjde
    # ---------------------------------------------------------------------------------------------------------
    def checkBasementDepth(self):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            from ..calculations.const import constants as c
            index_basement = c.STRUCTURE_OVERVIEW["Assembly version"].index("basement")
            num_base_floors = self.dlgMain.numberOfBasementFloors.value()
            if num_base_floors > 0:
                standardBasementDepth = round(self.getStandardBasementDepth(), 2)
                if self.dlgMain.basementDepth.value() < standardBasementDepth:
                    self.dlgMain.lblMessage.setText('For lav gulv-til-gulv kælderhøjde.')
                    self.dlgMain.lblMessage.setStyleSheet('color: red')

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def getStandardBasementDepth(self):
        
        try:
            #JCN/2024-10-31: Beregning af default kælder dybde
            from ..calculations.const import constants as c

            num_base_floors = self.dlgMain.numberOfBasementFloors.value()
            if self.dlgMain.parkingBasement.isChecked():
                TYPOLOGY_BASEMENT = "parking"
            else:
                typology = self.dict_translate.get(self.dlgMain.cb_Building_usage.currentText().lower())
                if typology == None:
                    #self.dlgMain.basementDepth.setValue(0)
                    return 0
                TYPOLOGY_BASEMENT = typology

            index_basement = c.STRUCTURE_OVERVIEW["Assembly version"].index("basement")
            index_hollow_core_slab = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "hollow-core slab"
            )
            index_concrete_finish = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "concrete finish"
            )
            basement_depth = (
                c.MIN_FTC_BASEMENT[TYPOLOGY_BASEMENT]
                + c.STRUCTURE_OVERVIEW["Thickness [m]"][index_hollow_core_slab]
                + c.STRUCTURE_OVERVIEW["Thickness [m]"][index_concrete_finish]
            ) * num_base_floors + c.STRUCTURE_OVERVIEW["Thickness [m]"][index_basement]

            return basement_depth
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def sekFacedePercentChanged(self):
        
        try:
            # if (self.dlgMain.primFacedePercent.value() > 0):
            #     self.dlgMain.sekFacedePercent.setValue(100-self.dlgMain.primFacedePercent.value())
            self.dlgMain.primFacedePercent.setValue(100-self.dlgMain.sekFacedePercent.value())
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def FillTableBuildings(self):
        try:
            if self.layerBuildings == None:
                return

            if self.sortColBuildings <= 0:
                return

            oTable = self.dlgMain.tblBuildings  # QTableWidget
            oTable.setRowCount(0)
            oTable.setColumnCount(8)
            oTable.setSortingEnabled(True)
            oTable.sortByColumn(self.sortColBuildings, self.sortOrderBuildings)

            oTable.setHorizontalHeaderLabels(
                [
                    "",
                    "",
                    "ID",
                    "Delområde",
                    "Grundareal\n[m\u00B2]",
                    "Omkreds [m]",
                    "Dybde [m]",
                    "Etageareal\n[m\u00B2]",
                ]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(1, 50)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(4, 100)
            oTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(5, 100)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            oTable.setColumnHidden(0, True)  # Id skjules
            oTable.setWordWrap(True)

            idxName = self.layerBuildings.fields().lookupField("id")
            idxAreaName = self.layerBuildings.fields().lookupField("AreaName")
            idxAreaCalc = self.layerBuildings.fields().lookupField("AreaCalc")
            idxPerimeterCalc = self.layerBuildings.fields().lookupField("PerimeterCalc")
            idxDepth = self.layerBuildings.fields().lookupField("depth")
            idxNumberOfFloors = self.layerBuildings.fields().lookupField("NumberOfFloors")
            idxNumberOfBasementFloors = self.layerBuildings.fields().lookupField("NumberOfBasementFloors")
            if idxName < 0 or idxAreaName < 0 or idxAreaCalc < 0 or idxPerimeterCalc < 0 or idxDepth < 0 or idxNumberOfFloors < 0 or idxNumberOfBasementFloors < 0:
                utils.Error('Et af disse felter mangler i laget med bygninger: id, AreaName, AreaCalc, PerimeterCalc, Depth, NumberOfFloors, NumberOfBasementFloors')
                return

            lstSubareaNames = self.getSubareaNames()

            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            i = 0
            featList = []
            colId = 0
            colName = 1
            colAreaName = 2
            colAreaCalc = 3
            colPerimeterCalc = 4
            colWidth = 5
            colFloorarea = 6

            for feat in self.layerBuildings.getFeatures():
                if feat[idxDepth] != None and feat[idxDepth] > 0:
                    width = feat[idxDepth]
                else:
                    if feat[idxAreaCalc] != None and feat[idxPerimeterCalc] != None:
                        omb = feat.geometry().orientedMinimumBoundingBox()
                        width = round(omb[-2], 2)
                        length = round(omb[-1], 2)
                        if width > length:
                            width, length = length, width
                floorarea = round(feat[idxAreaCalc] * (feat[idxNumberOfFloors] + feat[idxNumberOfBasementFloors]), 2)
                featList.append(
                    [
                        feat.id(),
                        feat[idxName],
                        feat[idxAreaName],
                        feat[idxAreaCalc],
                        feat[idxPerimeterCalc],
                        width,
                        floorarea,
                    ]
                )
            if self.sortColBuildings == colAreaCalc+1 or self.sortColBuildings == colPerimeterCalc+1 or self.sortColBuildings == colWidth+1:
                if self.sortOrderBuildings == 0:
                    featList.sort(key=lambda x:float(x[self.sortColBuildings - 1]), reverse=False)
                else:
                    featList.sort(key=lambda x:float(x[self.sortColBuildings - 1]), reverse=True)
            else:
                if self.sortOrderBuildings == 0:
                    featList.sort(key=lambda x: (str(x[self.sortColBuildings - 1]).replace("NULL", "").lower(), x[0]), reverse=False)
                else:
                    featList.sort(key=lambda x: (str(x[self.sortColBuildings - 1]).replace("NULL", "").lower(), x[0]), reverse=True)

            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Rediger geometri
                oEdit = QToolButton()
                oEdit.setIcon(oIcon)
                oEdit.setToolTip("Rediger geometri")
                oEdit.clicked.connect(self.editBuildingClicked)
                oTable.setCellWidget(i, 1, oEdit)
                # Navn
                oName = QLineEdit()
                if feat[colName] == None:
                    oName.setText("")
                else:
                    oName.setText(feat[colName])
                oName.setAlignment(Qt.AlignLeft)
                # oName.setEnabled(False)
                oTable.setCellWidget(i, 2, oName)
                # Delområde
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if lstSubareaNames != None:
                    j = 0
                    for key in lstSubareaNames:
                        qCombo.addItem(key, key)
                        if key == feat[colAreaName]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 3, qCombo)
                # Samlet areal
                oAreaCalc = QLineEdit()
                if feat[colAreaCalc] == None:
                    oAreaCalc.setText("")
                else:
                    oAreaCalc.setText(str(feat[colAreaCalc]))
                oAreaCalc.setAlignment(Qt.AlignRight)
                oAreaCalc.setEnabled(False)
                oTable.setCellWidget(i, 4, oAreaCalc)
                # Omkreds
                oPerimeterCalc = QLineEdit()
                if feat[colPerimeterCalc] == None:
                    oPerimeterCalc.setText("")
                else:
                    oPerimeterCalc.setText(str(feat[colPerimeterCalc]))
                oPerimeterCalc.setAlignment(Qt.AlignRight)
                oPerimeterCalc.setEnabled(False)
                oTable.setCellWidget(i, 5, oPerimeterCalc)
                # Dybde (Width)
                oWidth = QDoubleSpinBox()
                if feat[colWidth] != None:
                    oWidth.setValue(feat[colWidth])
                #oWidth.setAlignment(Qt.AlignRight)
                #oWidth.setEnabled(False)
                oTable.setCellWidget(i, 6, oWidth)
                # Etageareal
                oFloorarea = QLineEdit()
                if feat[colFloorarea] == None:
                    oFloorarea.setText("")
                else:
                    oFloorarea.setText(str(feat[colFloorarea]))
                oFloorarea.setAlignment(Qt.AlignRight)
                oFloorarea.setEnabled(False)
                oTable.setCellWidget(i, 7, oFloorarea)

                i += 1

            # Der indsættes en tom række
            self.InsertBlankRowBuildings(oTable)

        except Exception as e:
            utils.Error("Fejl i FillTableBuildings. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def InsertBlankRowBuildings(self, oTable):
        """
        Inserts an empty row into the building table at the specified index and initializes cell widgets
        for ID, geometry editing, and building attributes.

        Parameters:
            oTable (QTableWidget): The table widget into which the row will be inserted.
            i (int): The index at which the row should be inserted.

        Raises:
            Exception: If an error occurs during the row insertion process.
        """

        try:
            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            # Der indsættes en tom række
            i = oTable.verticalHeader().count()
            oTable.insertRow(i)
            # Id
            oId = QTableWidgetItem(-1)
            oId.setFlags(Qt.NoItemFlags)  # not enabled
            oTable.setItem(i, 0, oId)
            # Rediger geometri
            qMapLookup = QToolButton()
            qMapLookup.setIcon(oIcon)
            qMapLookup.setToolTip("Opret geometri")
            qMapLookup.clicked.connect(self.editBuildingClicked)
            oTable.setCellWidget(i, 1, qMapLookup)
            # Navn
            oName = QLineEdit()
            oName.setText("")
            oTable.setCellWidget(i, 2, oName)
            # Delområde
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.getSubareaNames() != None:
                for key in self.getSubareaNames():
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 3, qCombo)
            # Grundareal
            oAreaCalc = QLineEdit()
            oAreaCalc.setText("")
            oAreaCalc.setEnabled(False)
            oTable.setCellWidget(i, 4, oAreaCalc)
            # Omkreds
            oLenghtPerimeter = QLineEdit()
            oLenghtPerimeter.setText("")
            oLenghtPerimeter.setEnabled(False)
            oTable.setCellWidget(i, 5, oLenghtPerimeter)
            # Dybde (Width)
            oWidth = QLineEdit()
            oWidth.setText("")
            oWidth.setEnabled(False)
            oTable.setCellWidget(i, 6, oWidth)
            # Etageareal
            oFloorarea = QLineEdit()
            oFloorarea.setText("")
            oFloorarea.setEnabled(False)
            oTable.setCellWidget(i, 7, oFloorarea)

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count() - 1):
                lstRowHeader.append("")
            lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i InsertBlankRowBuildings. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    # Rediger geometri for bygning
    # ---------------------------------------------------------------------------------------------------------
    def editBuildingClicked(self):
        """
        Handles the action triggered when the 'Edit Building' button is clicked.

        This method performs several operations based on the user selection in the application interface:
        - If a new building is to be created, it sets the tool to 'CreateBuilding' and prompts user to draw a polygon on the map.
        - If an existing building is selected, it retrieves the building's GUID, selects the corresponding feature in the layer,
        and if necessary, moves the map view to center on the building geometry for editing.
        - Displays error messages in the UI in cases of failure to find or select the building's feature in the map.

        The method participates in managing the state of map tools and layer editing status for handling the creation
        or modification of building geometries within the application.

        Returns:
            -1 if an exception is caught during the operation, otherwise returns None.

        Raises:
            Exception: Captures any exceptions that occur during execution and shows an error message in the interface.
        """

        try:
            self.resetGeometryEdit()

            oTable = self.dlgMain.tblBuildings
            ch = self.sender()
            ix = oTable.indexAt(ch.pos())
            oTable.selectRow(ix.row())

            sId = oTable.item(ix.row(), 0).text()
            if sId == "":
                # Opret ny polygon
                self.ActiveTool = "CreateBuilding"
                self.dlgMain.lblMessage.setText("Tegn ny bygning i kortet.")
            else:
                # Rediger eksisterende bygning
                if not utils.isInt(sId):
                    utils.Error("Ugyldig bygning valgt.")
                    return
                id = int(sId)
                self.layerBuildings.selectByIds([id])
                if self.layerBuildings.selectedFeatureCount() == 0:
                    utils.Error("Kunne ikke finde bygning i kortet.")
                elif self.layerBuildings.selectedFeatureCount() == 1:
                    # Hvis polygonen ikke er synlig flyttes kortudsnit til contrum af polygonen
                    polygonGeometry = self.layerBuildings.selectedFeatures()[
                        0
                    ].geometry()
                    if not polygonGeometry.within(
                        QgsGeometry.fromRect(self.iface.mapCanvas().extent())
                    ):
                        self.iface.mapCanvas().setCenter(
                            polygonGeometry.centroid().asPoint()
                        )
                        # self.iface.mapCanvas().setExtent(self.layerSubareas.boundingBoxOfSelected())
                        # self.iface.mapCanvas().refresh()
                    if not self.layerBuildings.isEditable():
                        self.layerBuildings.startEditing()
                    self.ActiveTool = "EditBuilding"
                    self.dlgMain.lblMessage.setText("Vælg vertex og flyt det.")
                else:
                    utils.Error("Flere bygninger valgt i kortet.")

        except Exception as e:
            utils.Error("Fejl i editBuildingClicked. Fejlbesked: " + str(e))
            return -1

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def sortingChangedBuildings(self, col, sortorder):
        if col <= 0:
            self.dlgMain.tblBuildings.horizontalHeader().setSortIndicatorShown(False)
        else:
            self.sortColBuildings = col
            self.sortOrderBuildings = sortorder
            self.FillTableBuildings()

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def keyPressEventBuildings(self, e):

        try:
            if e.key() == Qt.Key_Delete:
                oTable = self.dlgMain.tblBuildings
                sId = oTable.item(oTable.currentRow(), 0).text()
                if sId == "":
                    utils.Error("Kan ikke fastlægge rækkens id.")
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig bygning valgt.")
                        return
                    id = int(sId)
                    if utils.Confirm(
                        "Er du sikker på at du ønsker at slette den valgte bygning?"
                    ):
                        oTable.removeRow(oTable.currentRow())
                        if not self.layerBuildings.isEditable():
                            self.layerBuildings.startEditing()
                        self.layerBuildings.deleteFeature(id)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der klikkes i 'verticalHeader' vælges polygon med bygning og rækken gemmes
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedBuildings(self, index):
        """
        Handles the event when a row in the building table is clicked.

        Selects the corresponding polygon on the map and attempts to save the data row
        associated with the provided index and building GUID.

        Args:
            index (int): The index of the clicked row in the buildings table.

        Returns:
            bool: True if the operation was successful, False otherwise.

        Raises:
            Exception: If an error occurs, logs the exception message.
        """

        try:
            self.saveRowBuildings(index)
            self.selectRowBuildings(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedBuildings(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            oldrow = indexOld.row()
            if oldrow >= 0:
                self.saveRowBuildings(oldrow)
            self.selectRowBuildings(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Række og polygon highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowBuildings(self, index):

        try:
            oTable = self.dlgMain.tblBuildings
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)
                # Polygon vælges i kortet
                sId = oTable.item(index, 0).text()
                if sId == "":
                    #Tom række
                    self.layerBuildings.removeSelection()
                    #Nulstil indhold
                    self.dlgMain.numberOfFloors.setValue(0)
                    self.dlgMain.buildingHeight.setValue(0)
                    self.dlgMain.floorHeight.setValue(0)
                    self.dlgMain.groundFTFHeight.setValue(0)
                    self.dlgMain.numberOfBasementFloors.setValue(0)
                    self.dlgMain.basementDepth.setValue(0)
                    self.dlgMain.cb_Building_condition.setCurrentIndex(0)
                    self.dlgMain.cb_Building_usage.setCurrentIndex(0)
                    self.dlgMain.cb_Ground_usage.setCurrentIndex(0)
                    self.dlgMain.cb_Building_bearing_sys.setCurrentIndex(0)
                    self.dlgMain.windowPercent.setValue(0)
                    self.dlgMain.numResidents.setValue(0)
                    self.dlgMain.numWorkplaces.setValue(0)
                    self.dlgMain.cb_Building_facade_prime.setCurrentIndex(0)
                    self.dlgMain.primFacedePercent.setValue(100)
                    self.dlgMain.cb_Building_facade_sek.setCurrentIndex(0)
                    self.dlgMain.sekFacedePercent.setValue(0)
                    self.dlgMain.cb_Building_roof_type.setCurrentIndex(0)
                    self.dlgMain.roofAngle.setValue(0)
                    self.dlgMain.cb_Building_roof_material.setCurrentIndex(0)
                    self.dlgMain.cb_Building_heating_type.setCurrentIndex(0)
                    #Disable editing
                    self.EnableDisableBuildingParameters(False)
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig bygning valgt.")
                        return
                    id = int(sId)
                    self.layerBuildings.selectByIds([id])
                    # Udfyld bygningsparametre - udenfor tabel
                    if self.layerBuildings.selectedFeatureCount() == 1:
                        #Enable editing
                        self.EnableDisableBuildingParameters(True)
                        feat = self.layerBuildings.selectedFeatures()[0]
                        if feat.fieldNameIndex("NumberOfFloors") >= 0:
                            if feat["NumberOfFloors"] == None:
                                self.dlgMain.numberOfFloors.setValue(0)
                            else:
                                self.dlgMain.numberOfFloors.setValue(
                                    feat["NumberOfFloors"]
                                )
                        if feat.fieldNameIndex("Building_height") >= 0:
                            if feat["Building_height"] == None:
                                self.dlgMain.buildingHeight.setValue(0)
                            else:
                                self.dlgMain.buildingHeight.setValue(
                                    feat["Building_height"]
                                )
                        if (
                            feat.fieldNameIndex("NumberOfFloors") >= 0
                            and feat.fieldNameIndex("Building_height") >= 0
                        ):
                            if (
                                feat["NumberOfFloors"] == None
                                or feat["NumberOfFloors"] == 0
                                or feat["Building_height"] == None
                            ):
                                self.dlgMain.floorHeight.setValue(0)
                            else:
                                self.dlgMain.floorHeight.setValue(
                                    feat["Building_height"] / feat["NumberOfFloors"]
                                )
                        if feat.fieldNameIndex("Ground_ftf_height") >= 0:
                            if feat["Ground_ftf_height"] == None:
                                self.dlgMain.groundFTFHeight.setValue(0)
                            else:
                                self.dlgMain.groundFTFHeight.setValue(
                                    feat["Ground_ftf_height"]
                                )
                        if feat.fieldNameIndex("NumberOfBasementFloors") >= 0:
                            if feat["NumberOfBasementFloors"] == None or feat["NumberOfBasementFloors"] == 0:
                                self.dlgMain.basementDepth.setEnabled(False)
                                self.dlgMain.numberOfBasementFloors.setValue(0)
                            else:
                                self.dlgMain.basementDepth.setEnabled(True)
                                self.dlgMain.numberOfBasementFloors.setValue(
                                    feat["NumberOfBasementFloors"]
                                )
                        if feat.fieldNameIndex("Basement_depth") >= 0:
                            if feat["Basement_depth"] == None:
                                self.dlgMain.basementDepth.setValue(0)
                            else:
                                self.dlgMain.basementDepth.setValue(
                                    feat["Basement_depth"]
                                )
                        if feat.fieldNameIndex("ParkingBasement") >= 0:
                            if feat["ParkingBasement"] == None or feat["ParkingBasement"] == False:
                                self.dlgMain.parkingBasement.setChecked(False)
                            else:
                                self.dlgMain.parkingBasement.setChecked(True)
                        if feat.fieldNameIndex("Building_condition") >= 0:
                            self.dlgMain.cb_Building_condition.setCurrentIndex(0)
                            for i in range(self.dlgMain.cb_Building_condition.count()):
                                if (self.dlgMain.cb_Building_condition.itemText(i).lower() == feat["Building_condition"].lower().replace('beholdt','bevaret')):
                                    self.dlgMain.cb_Building_condition.setCurrentIndex(i)
                                    break
                        if feat.fieldNameIndex("Building_usage") >= 0:
                            self.dlgMain.cb_Building_usage.setCurrentIndex(0)
                            for i in range(self.dlgMain.cb_Building_usage.count()):
                                if (
                                    self.dlgMain.cb_Building_usage.itemText(i)
                                    == feat["Building_usage"]
                                ):
                                    self.dlgMain.cb_Building_usage.setCurrentIndex(i)
                                    break
                        if feat.fieldNameIndex("Ground_usage") >= 0:
                            self.dlgMain.cb_Ground_usage.setCurrentIndex(0)
                            for i in range(self.dlgMain.cb_Ground_usage.count()):
                                if (
                                    self.dlgMain.cb_Ground_usage.itemText(i)
                                    == feat["Ground_usage"]
                                ):
                                    self.dlgMain.cb_Ground_usage.setCurrentIndex(i)
                                    break
                        if feat.fieldNameIndex("Building_bearing_sys") >= 0:
                            self.dlgMain.cb_Building_bearing_sys.setCurrentIndex(0)
                            for i in range(
                                self.dlgMain.cb_Building_bearing_sys.count()
                            ):
                                if (
                                    self.dlgMain.cb_Building_bearing_sys.itemText(i)
                                    == feat["Building_bearing_sys"]
                                ):
                                    self.dlgMain.cb_Building_bearing_sys.setCurrentIndex(
                                        i
                                    )
                                    break
                        if feat.fieldNameIndex("windowPercent") >= 0:
                            if feat["windowPercent"] == None:
                                self.dlgMain.windowPercent.setValue(0)
                            else:
                                self.dlgMain.windowPercent.setValue(
                                    feat["windowPercent"]
                                )
                        if feat.fieldNameIndex("numResidents") >= 0:
                            if feat["numResidents"] == None:
                                self.dlgMain.numResidents.setValue(0)
                            else:
                                self.dlgMain.numResidents.setValue(
                                    feat["numResidents"]
                                )
                        if feat.fieldNameIndex("numWorkplaces") >= 0:
                            if feat["numWorkplaces"] == None:
                                self.dlgMain.numWorkplaces.setValue(0)
                            else:
                                self.dlgMain.numWorkplaces.setValue(
                                    feat["numWorkplaces"]
                                )
                        if feat.fieldNameIndex("Building_facade_prime") >= 0:
                            self.dlgMain.cb_Building_facade_prime.setCurrentIndex(0)
                            for i in range(
                                self.dlgMain.cb_Building_facade_prime.count()
                            ):
                                if (
                                    self.dlgMain.cb_Building_facade_prime.itemText(i)
                                    == feat["Building_facade_prime"]
                                ):
                                    self.dlgMain.cb_Building_facade_prime.setCurrentIndex(
                                        i
                                    )
                                    break
                        if feat.fieldNameIndex("primFacedePercent") >= 0:
                            if feat["primFacedePercent"] == None:
                                self.dlgMain.primFacedePercent.setValue(0)
                            else:
                                self.dlgMain.primFacedePercent.setValue(
                                    feat["primFacedePercent"]
                                )
                        if feat.fieldNameIndex("Building_facade_sek") >= 0:
                            self.dlgMain.cb_Building_facade_sek.setCurrentIndex(0)
                            for i in range(self.dlgMain.cb_Building_facade_sek.count()):
                                if (
                                    self.dlgMain.cb_Building_facade_sek.itemText(i)
                                    == feat["Building_facade_sek"]
                                ):
                                    self.dlgMain.cb_Building_facade_sek.setCurrentIndex(
                                        i
                                    )
                                    break
                        if feat.fieldNameIndex("sekFacedePercent") >= 0:
                            if feat["sekFacedePercent"] == None:
                                self.dlgMain.sekFacedePercent.setValue(0)
                            else:
                                self.dlgMain.sekFacedePercent.setValue(
                                    feat["sekFacedePercent"]
                                )
                        if feat.fieldNameIndex("Building_roof_type") >= 0:
                            self.dlgMain.cb_Building_roof_type.setCurrentIndex(0)
                            for i in range(self.dlgMain.cb_Building_roof_type.count()):
                                if (
                                    self.dlgMain.cb_Building_roof_type.itemText(i)
                                    == feat["Building_roof_type"]
                                ):
                                    self.dlgMain.cb_Building_roof_type.setCurrentIndex(
                                        i
                                    )
                                    break
                        if feat.fieldNameIndex("Roof_angle") >= 0:
                            if feat["Roof_angle"] == None:
                                self.dlgMain.roofAngle.setValue(0)
                            else:
                                self.dlgMain.roofAngle.setValue(
                                    feat["Roof_angle"]
                                )
                        if feat.fieldNameIndex("Building_roof_material") >= 0:
                            self.dlgMain.cb_Building_roof_material.setCurrentIndex(0)
                            for i in range(self.dlgMain.cb_Building_roof_material.count()):
                                if (self.dlgMain.cb_Building_roof_material.itemText(i).lower() == feat["Building_roof_material"].lower().replace('bitumen','tagpap')):
                                    self.dlgMain.cb_Building_roof_material.setCurrentIndex(i)
                                    break
                        if feat.fieldNameIndex("Building_heating_type") >= 0:
                            self.dlgMain.cb_Building_heating_type.setCurrentIndex(0)
                            for i in range(
                                self.dlgMain.cb_Building_heating_type.count()
                            ):
                                if (
                                    self.dlgMain.cb_Building_heating_type.itemText(i)
                                    == feat["Building_heating_type"]
                                ):
                                    self.dlgMain.cb_Building_heating_type.setCurrentIndex(
                                        i
                                    )
                                    break
            #Events
            self.Building_usageChanged()
            self.Ground_usageChanged()
            self.sekFacedePercentChanged()
            self.Building_roof_typeChanged()
            #self.numberOfBasementFloorsChanged()
            self.checkBasementDepth()
            self.calcFloorHeight()
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def EnableDisableBuildingParameters(self, bStatus):
        
        try:    
            #Enable/disable editing
            self.dlgMain.numberOfFloors.setEnabled(bStatus)
            self.dlgMain.buildingHeight.setEnabled(bStatus)
            #self.dlgMain.floorHeight.setEnabled(bStatus)
            if self.dlgMain.cb_Ground_usage.currentText() == '':
                self.dlgMain.groundFTFHeight.setEnabled(False)
            else:
                self.dlgMain.groundFTFHeight.setEnabled(bStatus)
            self.dlgMain.numberOfBasementFloors.setEnabled(bStatus)
            self.dlgMain.cb_Building_condition.setEnabled(bStatus)
            self.dlgMain.cb_Building_usage.setEnabled(bStatus)
            self.dlgMain.cb_Ground_usage.setEnabled(bStatus)
            self.dlgMain.cb_Building_bearing_sys.setEnabled(bStatus)
            self.dlgMain.windowPercent.setEnabled(bStatus)
            self.dlgMain.cb_Building_facade_prime.setEnabled(bStatus)
            #self.dlgMain.primFacedePercent.setEnabled(bStatus)
            self.dlgMain.cb_Building_facade_sek.setEnabled(bStatus)
            self.dlgMain.sekFacedePercent.setEnabled(bStatus)
            self.dlgMain.cb_Building_roof_type.setEnabled(bStatus)
            self.dlgMain.roofAngle.setEnabled(bStatus)
            self.dlgMain.cb_Building_roof_material.setEnabled(bStatus)
            self.dlgMain.cb_Building_heating_type.setEnabled(bStatus)
        
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Data gemmes
    # ---------------------------------------------------------------------------------------------------------
    def saveRowBuildings(self, index):

        try:
            # Data fra tabel gemmes
            oTable = self.dlgMain.tblBuildings
            sId = oTable.item(index, 0).text()
            if sId == "":
                return False
            if not utils.isInt(sId):
                utils.Error("Ugyldig bygning valgt.")
                return False

            id = int(sId)
            # Så opdateres laget
            if not self.layerBuildings.isEditable():
                self.layerBuildings.startEditing()
            oName = oTable.cellWidget(index, 2)
            if isinstance(oName, QLineEdit):
                idx = self.layerBuildings.fields().indexOf("id")
                if idx >= 0:
                    self.layerBuildings.changeAttributeValue(id, idx, oName.text())
            oAreaName = oTable.cellWidget(index, 3)
            if isinstance(oAreaName, QComboBox):
                idx = self.layerBuildings.fields().indexOf("AreaName")
                if idx >= 0:
                    self.layerBuildings.changeAttributeValue(id, idx, oAreaName.currentText())
            oWidth = oTable.cellWidget(index, 6)
            if isinstance(oWidth, QDoubleSpinBox):
                idx = self.layerBuildings.fields().indexOf("depth")
                if idx >= 0:
                    self.layerBuildings.changeAttributeValue(id, idx, oWidth.value())
            # Øvrige data fra skærmbillede
            idx = self.layerBuildings.fields().indexOf("NumberOfFloors")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.numberOfFloors.value()
                )
            idx = self.layerBuildings.fields().indexOf("Building_height")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.buildingHeight.value()
                )
            idx = self.layerBuildings.fields().indexOf("Ground_ftf_height")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.groundFTFHeight.value()
                )
            idx = self.layerBuildings.fields().indexOf("NumberOfBasementFloors")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.numberOfBasementFloors.value()
                )
            idx = self.layerBuildings.fields().indexOf("Basement_depth")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.basementDepth.value()
                )
            idx = self.layerBuildings.fields().indexOf("ParkingBasement")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.parkingBasement.isChecked()
                )
            idx = self.layerBuildings.fields().indexOf("Building_condition")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_condition.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("Building_usage")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_usage.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("Ground_usage")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Ground_usage.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("Building_bearing_sys")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_bearing_sys.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("windowPercent")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.windowPercent.value()
                )
            idx = self.layerBuildings.fields().indexOf("numResidents")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.numResidents.value()
                )
            idx = self.layerBuildings.fields().indexOf("numWorkplaces")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.numWorkplaces.value()
                )
            idx = self.layerBuildings.fields().indexOf("Building_facade_prime")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_facade_prime.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("primFacedePercent")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.primFacedePercent.value()
                )
            idx = self.layerBuildings.fields().indexOf("Building_facade_sek")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_facade_sek.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("sekFacedePercent")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.sekFacedePercent.value()
                )
            idx = self.layerBuildings.fields().indexOf("Building_roof_type")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_roof_type.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("Roof_angle")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.roofAngle.value()
                )
            idx = self.layerBuildings.fields().indexOf("Building_roof_material")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_roof_material.currentText()
                )
            idx = self.layerBuildings.fields().indexOf("Building_heating_type")
            if idx >= 0:
                self.layerBuildings.changeAttributeValue(
                    id, idx, self.dlgMain.cb_Building_heating_type.currentText()
                )
            if self.layerBuildings.featureCount() > 0:
                #Emission værdier sættes til 0 aht. tematisk kort
                feat = self.layerBuildings.getFeature(id)
                idx = self.layerBuildings.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    if feat["EmissionsTotal"] == None:
                        self.layerBuildings.changeAttributeValue(id, idx, 0)
                idx = self.layerBuildings.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    if feat["EmissionsTotalM2"] == None:
                        self.layerBuildings.changeAttributeValue(id, idx, 0)
                idx = self.layerBuildings.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    if feat["EmissionsTotalM2Year"] == None:
                        self.layerBuildings.changeAttributeValue(id, idx, 0)
                idx = self.layerBuildings.fields().indexOf("EmissionsTotalPerson")
                if idx >= 0:
                    if feat["EmissionsTotalPerson"] == None:
                        self.layerBuildings.changeAttributeValue(id, idx, 0)
                idx = self.layerBuildings.fields().indexOf("EmissionsTotalPersonYear")
                if idx >= 0:
                    if feat["EmissionsTotalPersonYear"] == None:
                        self.layerBuildings.changeAttributeValue(id, idx, 0)

            return True

        except Exception as e:
            utils.Error("Fejl i saveRowBuildings. Fejlbesked: " + str(e))
            return False

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def CalculateBuilding(self):

        try:
            self.dlgMain.lblMessage.setText("")
            if self.dlgMain.tblBuildings.currentRow() < 0:
                utils.Message("Vælg en række som skal beregnes")
                return
            sId = self.dlgMain.tblBuildings.item(self.dlgMain.tblBuildings.currentRow(), 0).text()
            if sId != "":
                if utils.isInt(sId):
                    self.saveRowBuildings(self.dlgMain.tblBuildings.currentRow())
                    calculate_utils.CalculateBuilding(self.dlgMain, self.layerBuildings, self.dict_translate)
            
        except Exception as e:
            utils.Error("Fejl i CalculateBuilding: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def SaveBuildings(self):

        try:
            if self.layerBuildings == None:
                return
            oTable = self.dlgMain.tblBuildings
            if oTable.currentRow() >= 0:
                if not self.saveRowBuildings(oTable.currentRow()):
                    utils.Message("Den aktive række indeholder ingen geometri. Der skal tegnes et område før rækken kan gemmes.")
                    return

            if self.layerBuildings.undoStack().count() == 0:
                utils.Message("Der er ingen ændringer at gemme.")
                return
            if not utils.Confirm("Er du sikker på at du vil gemme?"):
                return
            if self.layerBuildings.isEditable():
                if self.layerBuildings.isModified():
                    saveOkKomp = self.layerBuildings.commitChanges()
                    if not saveOkKomp:
                        if utils.Confirm("Ændringerne kunne ikke gemmes. Vis fejl?"):
                            utils.Message("\n".join(self.layerBuildings.commitErrors()))

            if self.layerBuildings.selectedFeatureCount() > 0:
                self.layerBuildings.removeSelection()
            self.FillTableBuildings()
            self.SaveSubareas(True)
            
        except Exception as e:
            utils.Error("Fejl i SaveBuildings: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def RollbackBuildings(self):

        try:
            if self.layerBuildings == None:
                return

            if self.layerBuildings.undoStack().count() == 0:
                if self.layerBuildings.isEditable():
                    self.layerBuildings.commitChanges()
            else:
                doContinue = utils.Confirm(
                    "Der er lavet ændringer - er du sikker på at disse ikke skal gemmes?"
                )
                if doContinue:
                    if self.layerBuildings.isEditable():
                        self.layerBuildings.rollBack()

            if self.layerBuildings.selectedFeatureCount() > 0:
                self.layerBuildings.removeSelection()
            self.FillTableBuildings()

        except Exception as e:
            utils.Error("Fejl i RollbackBuildings: " + str(e))

    # endregion
    # ---------------------------------------------------------------------------------------------------------
    # ÅBNE OVERFLADER
    # ---------------------------------------------------------------------------------------------------------
    # region
    # ---------------------------------------------------------------------------------------------------------
    # Lines
    # ---------------------------------------------------------------------------------------------------------
    def FillTableOpenSurfacesLines(self):
        # Define the update_width function here
        def update_width(index, row):
            # Get the current text of the combo box
            current_text = oTable.cellWidget(row, 4).itemText(index)
            # Get the oWidth widget
            oWidth = oTable.cellWidget(row, 7)
            # Update the text of oWidth based on the current selection in combo_box
            oWidth.setValue(self.dict_surface_width[current_text])

        try:
            if self.layerOpenSurfacesLines == None:
                return

            if self.sortColOpenSurfacesLines <= 0:
                return

            oTable = self.dlgMain.tblOpenSurfacesLines
            oTable.setRowCount(0)
            oTable.setColumnCount(14)
            oTable.setSortingEnabled(True)
            oTable.sortByColumn(self.sortColOpenSurfacesLines, self.sortOrderOpenSurfacesLines)
            oTable.setHorizontalHeaderLabels(
                ["", "", "ID", "Delområde", "Type", "Tilstand", "Længde\n[m]", "Bredde\n[m]", "Antal\ntræer", "Type af\ntræer", "Størrelse af\ntræer", "Antal\nbuske", "Type a\nbuske", "Størrelse af\nbuske"]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.setColumnHidden(0, True)  # Id skjules
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(1, 50)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
            oTable.horizontalHeader().resizeSection(2, 50)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(6, 50)
            oTable.horizontalHeader().setSectionResizeMode(7, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(7, 50)
            oTable.horizontalHeader().setSectionResizeMode(8, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(8, 50)
            oTable.horizontalHeader().setSectionResizeMode(9, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(10, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(11, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(11, 50)
            oTable.horizontalHeader().setSectionResizeMode(12, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(13, QHeaderView.Stretch)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            oTable.setWordWrap(True)

            idxName = self.layerOpenSurfacesLines.fields().lookupField("id")
            idxAreaName = self.layerOpenSurfacesLines.fields().lookupField("AreaName")
            idxType = self.layerOpenSurfacesLines.fields().lookupField("type")
            idxCondition = self.layerOpenSurfacesLines.fields().lookupField("condition")
            idxWidth = self.layerOpenSurfacesLines.fields().lookupField("width")
            idxLengthCalc = self.layerOpenSurfacesLines.fields().lookupField("LengthCalc")
            if idxName < 0 or idxType < 0 or idxCondition < 0 or idxLengthCalc < 0:
                utils.Error('Et af disse felter mangler i laget med veje: id, type, condition, LengthCalc')
                return
            idxNum_trees = self.layerOpenSurfacesLines.fields().lookupField("num_trees")
            idxType_trees = self.layerOpenSurfacesLines.fields().lookupField("type_trees")
            idxSize_trees = self.layerOpenSurfacesLines.fields().lookupField("size_trees")
            idxNum_shrubs = self.layerOpenSurfacesLines.fields().lookupField("num_shrubs")
            idxType_shrubs = self.layerOpenSurfacesLines.fields().lookupField("type_shrubs")
            idxSize_shrubs = self.layerOpenSurfacesLines.fields().lookupField("size_shrubs")
            if idxNum_trees < 0 or idxType_trees < 0 or idxSize_trees < 0 or idxNum_shrubs < 0 or idxType_shrubs < 0 or idxSize_shrubs < 0:
                utils.Error('Et af disse felter mangler i laget med veje: num_trees, type_trees, size_trees, num_shrubs, type_shrubs, size_shrubs')
                return

            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            i = 0
            featList = []
            colId = 0
            colName = 1
            colAreaName = 2
            colType = 3
            colCondition = 4
            colLengthCalc = 5
            colWidth = 6
            colNum_trees = 7
            colType_trees = 8
            colSize_trees = 9
            colNum_shrubs = 10
            colType_shrubs =11
            colSize_shrubs = 12
            for feat in self.layerOpenSurfacesLines.getFeatures():
                #Er bredden gemt i tabellen bruges denne
                if feat[idxWidth] > 0:
                    width = feat[idxWidth]
                else:
                    if feat[idxType] == None or feat[idxType] == "":
                        width = 0
                    else:
                        width = self.dict_surface_width[feat[idxType]]
                featList.append(
                    [
                        feat.id(),
                        feat[idxName],
                        feat[idxAreaName],
                        feat[idxType],
                        feat[idxCondition],
                        feat[idxLengthCalc],
                        width,
                        feat[idxNum_trees],
                        feat[idxType_trees],
                        feat[idxSize_trees],
                        feat[idxNum_shrubs],
                        feat[idxType_shrubs],
                        feat[idxSize_shrubs],
                    ]
                )

            if self.sortColOpenSurfacesLines == colLengthCalc+1 or self.sortColOpenSurfacesLines == colWidth+1:
                #Lidt langsomt - overvej at indsætte 0 i stedet for NULL i listen
                if self.sortOrderOpenSurfacesLines == 0:
                    featList.sort(key=lambda x:float(str(x[self.sortColOpenSurfacesLines - 1]).replace("", "0")), reverse=False)
                else:
                    featList.sort(key=lambda x:float(str(x[self.sortColOpenSurfacesLines - 1]).replace("", "0")), reverse=True)
            else:
                if self.sortOrderOpenSurfacesLines == 0:
                    featList.sort(key=lambda x: str(x[self.sortColOpenSurfacesLines - 1]).replace("NULL", "").lower(), reverse=False)
                else:
                    featList.sort(key=lambda x: str(x[self.sortColOpenSurfacesLines - 1]).replace("NULL", "").lower(), reverse=True)

            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Rediger geometri
                oEdit = QToolButton()
                oEdit.setIcon(oIcon)
                oEdit.setToolTip("Rediger geometri")
                oEdit.clicked.connect(self.editOpenSurfacesLinesClicked)
                oTable.setCellWidget(i, 1, oEdit)
                # Navn / id
                oName = QLineEdit()
                if feat[colName] == None:
                    oName.setText("")
                else:
                    oName.setText(feat[colName])
                oName.setAlignment(Qt.AlignLeft)
                oTable.setCellWidget(i, 2, oName)
                # Delområde
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.getSubareaNames() != None:
                    j = 0
                    for key in self.getSubareaNames():
                        qCombo.addItem(key, key)
                        if key == feat[colAreaName]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 3, qCombo)
                # Type
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_open_surfaces != None:
                    j = 0
                    for key, val in self.dict_open_surfaces.items():
                        if val == "line" or val == "":
                            qCombo.addItem(key, key)
                            if key == feat[colType]:
                                qCombo.setCurrentIndex(j)
                            j += 1
                qCombo.currentIndexChanged.connect(lambda index, r=i: update_width(index, r))
                oTable.setCellWidget(i, 4, qCombo)
                # Tilstand
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_condition:
                        qCombo.addItem(key, key)
                        if key == feat[colCondition]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 5, qCombo)
                # Length
                oLengthCalc = QLineEdit()
                if feat[colLengthCalc] == None:
                    oLengthCalc.setText("")
                else:
                    oLengthCalc.setText(str(feat[colLengthCalc]))
                oLengthCalc.setAlignment(Qt.AlignRight)
                oLengthCalc.setEnabled(False)
                oTable.setCellWidget(i, 6, oLengthCalc)
                # Width
                oWidth = QDoubleSpinBox()
                if feat[colWidth] != None:
                    oWidth.setValue(feat[colWidth])
                #oWidth.setAlignment(Qt.AlignRight)
                #oWidth.setEnabled(False)
                oTable.setCellWidget(i, 7, oWidth)
                # Num_trees
                oBox = QSpinBox()
                oBox.setMaximum(100000)
                if feat[colNum_trees] != None:
                    oBox.setValue(feat[colNum_trees])
                oBox.setAlignment(Qt.AlignRight)
                oTable.setCellWidget(i, 8, oBox)
                # Type_trees
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_treetype:
                        qCombo.addItem(key, key)
                        if key == feat[colType_trees]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 9, qCombo)
                # Size_trees
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_treesize:
                        qCombo.addItem(key, key)
                        if key == feat[colSize_trees]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 10, qCombo)
                # Num_shrubs
                oBox = QSpinBox()
                oBox.setMaximum(100000)
                if feat[colNum_shrubs] != None:
                    oBox.setValue(feat[colNum_shrubs])
                oBox.setAlignment(Qt.AlignRight)
                oTable.setCellWidget(i, 11, oBox)
                # Type_shrubs
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_treetype:
                        qCombo.addItem(key, key)
                        if key == feat[colType_shrubs]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 12, qCombo)
                # Size_shrubs
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_shrubsize:
                        qCombo.addItem(key, key)
                        if key == feat[colSize_shrubs]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 13, qCombo)

                i += 1

            # Der indsættes en tom række
            self.InsertBlankRowOpenSurfacesLines(oTable)

        except Exception as e:
            utils.Error("Fejl i FillTableOpenSurfacesLines. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def InsertBlankRowOpenSurfacesLines(self, oTable):
        # Define the update_width function here
        def update_width(index, row):
            # Get the current text of the combo box
            current_text = oTable.cellWidget(row, 4).itemText(index)
            # Get the oWidth widget
            oWidth = oTable.cellWidget(row, 7)
            # Update the text of oWidth based on the current selection in combo_box
            oWidth.setValue(self.dict_surface_width[current_text])

        try:
            lstSubareaNames = self.getSubareaNames()
            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            i = oTable.verticalHeader().count()
            combo_box = QComboBox()
            combo_box.addItems(
                list(
                    sorted(
                        [
                            k
                            for k, v in self.dict_open_surfaces.items()
                            if v == "line" or v == ""
                        ]
                    )
                )
            )
            # Connect the signal to a slot if necessary
            combo_box.currentIndexChanged.connect(
                lambda index, r=i: update_width(index, r)
            )

            # Der indsættes en tom række
            oTable.insertRow(i)
            # Id
            oId = QTableWidgetItem(-1)
            oId.setFlags(Qt.NoItemFlags)  # not enabled
            oTable.setItem(i, 0, oId)
            # Rediger geometri
            qMapLookup = QToolButton()
            qMapLookup.setIcon(oIcon)
            qMapLookup.setToolTip("Opret geometri")
            qMapLookup.clicked.connect(self.editOpenSurfacesLinesClicked)
            oTable.setCellWidget(i, 1, qMapLookup)
            # Navn / id
            oName = QLineEdit()
            oName.setText("")
            oName.setAlignment(Qt.AlignLeft)
            oTable.setCellWidget(i, 2, oName)
            # Delområde
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.getSubareaNames() != None:
                for key in self.getSubareaNames():
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 3, qCombo)
            # Type
            oTable.setCellWidget(i, 4, combo_box)
            # Tilstand
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_condition:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 5, qCombo)
            # Length
            oLenght = QLineEdit()
            oLenght.setText("")
            oLenght.setEnabled(False)
            oTable.setCellWidget(i, 6, oLenght)
            # Width
            oWidth = QDoubleSpinBox()
            #oWidth.setAlignment(Qt.AlignRight)
            #oWidth.setEnabled(False)
            oTable.setCellWidget(i, 7, oWidth)
            # Num_trees
            oBox = QSpinBox()
            oBox.setAlignment(Qt.AlignRight)
            oBox.setMaximum(100000)
            oTable.setCellWidget(i, 8, oBox)
            # Type_trees
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_treetype:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 9, qCombo)
            # Size_trees
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_treesize:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 10, qCombo)
            # Num_shrubs
            oBox = QSpinBox()
            oBox.setAlignment(Qt.AlignRight)
            oBox.setMaximum(100000)
            oTable.setCellWidget(i, 11, oBox)
            # Type_shrubs
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_treetype:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 12, qCombo)
            # Size_shrubs
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_shrubsize:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 13, qCombo)

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count() - 1):
                lstRowHeader.append("")
            lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i insertBlankRowOpenSurfacesLines. Fejlbesked: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    # Rediger geometri for veje
    # ---------------------------------------------------------------------------------------------------------
    def editOpenSurfacesLinesClicked(self):
        
        try:
            self.resetGeometryEdit()

            oTable = self.dlgMain.tblOpenSurfacesLines
            ch = self.sender()
            ix = oTable.indexAt(ch.pos())
            oTable.selectRow(ix.row())

            sId = oTable.item(ix.row(), 0).text()
            if sId == "":
                # Opret ny line
                self.ActiveTool = "CreateSurfaceLine"
                self.dlgMain.lblMessage.setText("Tegn linje i kortet.")
            else:
                # Rediger eksisterende linie
                if not utils.isInt(sId):
                    utils.Error("Ugyldig linie valgt.")
                    return
                id = int(sId)
                self.layerOpenSurfacesLines.selectByIds([id])
                if self.layerOpenSurfacesLines.selectedFeatureCount() == 0:
                    utils.Error("Kunne ikke finde linie i kortet.")
                elif self.layerOpenSurfacesLines.selectedFeatureCount() == 1:
                    # Hvis linje ikke er synlig flyttes kortudsnit til midpunkt af linjen
                    lineGeometry = self.layerOpenSurfacesLines.selectedFeatures()[
                        0
                    ].geometry()
                    if not lineGeometry.within(
                        QgsGeometry.fromRect(self.iface.mapCanvas().extent())
                    ):
                        self.iface.mapCanvas().setCenter(
                            lineGeometry.centroid().asPoint()
                        )
                    if not self.layerOpenSurfacesLines.isEditable():
                        self.layerOpenSurfacesLines.startEditing()
                    self.ActiveTool = "EditSurfaceLine"
                    self.dlgMain.lblMessage.setText("Vælg vertex og flyt det.")
                else:
                    utils.Error("Flere linjer valgt i kortet.")

        except Exception as e:
            utils.Error("Fejl i editOpenSurfacesLinesClicked. Fejlbesked: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    # Der skiftes sortering efter klik i header
    # ---------------------------------------------------------------------------------------------------------
    def sortingChangedOpenSurfacesLines(self, col, sortorder):
        if col <= 0:
            self.dlgMain.tblOpenSurfacesLines.horizontalHeader().setSortIndicatorShown(False)
        else:
            self.sortColOpenSurfacesLines = col
            self.sortOrderOpenSurfacesLines = sortorder
            self.FillTableOpenSurfacesLines()

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def keyPressEventOpenSurfacesLines(self, e):

        try:
            if e.key() == Qt.Key_Delete:
                oTable = self.dlgMain.tblOpenSurfacesLines
                sId = oTable.item(oTable.currentRow(), 0).text()
                if sId == "":
                    utils.Error("Kan ikke fastlægge rækkens id.")
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig vej valgt.")
                        return
                    id = int(sId)
                    if utils.Confirm(
                        "Er du sikker på at du ønsker at slette den valgte vej?"
                    ):
                        oTable.removeRow(oTable.currentRow())
                        if not self.layerOpenSurfacesLines.isEditable():
                            self.layerOpenSurfacesLines.startEditing()
                        self.layerOpenSurfacesLines.deleteFeature(id)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der klikkes i 'verticalHeader' vælges linien med vejen og rækken gemmes
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedOpenSurfacesLines(self, index):

        try:
            self.saveRowOpenSurfacesLines(index)
            self.selectRowOpenSurfacesLines(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedOpenSurfacesLines(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            oldrow = indexOld.row()
            if oldrow >= 0:
                self.saveRowOpenSurfacesLines(oldrow)
            self.selectRowOpenSurfacesLines(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Række og linie highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowOpenSurfacesLines(self, index):

        try:
            oTable = self.dlgMain.tblOpenSurfacesLines
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)
                # Linie vælges i kortet
                sId = oTable.item(index, 0).text()
                if sId == "":
                    #Tom række
                    self.layerOpenSurfacesLines.removeSelection()
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig vej valgt.")
                        return
                    id = int(sId)
                    self.layerOpenSurfacesLines.selectByIds([id])

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Data gemmes
    # ---------------------------------------------------------------------------------------------------------
    def saveRowOpenSurfacesLines(self, index):

        try:
            # Data fra tabel gemmes
            oTable = self.dlgMain.tblOpenSurfacesLines
            sId = oTable.item(index, 0).text()
            if sId == "":
                return False
            if not utils.isInt(sId):
                utils.Error("Ugyldig vej/linie valgt.")
                return False
            
            id = int(sId)
            if not self.layerOpenSurfacesLines.isEditable():
                self.layerOpenSurfacesLines.startEditing()
            oName = oTable.cellWidget(index, 2)
            if isinstance(oName, QLineEdit):
                idx = self.layerOpenSurfacesLines.fields().indexOf("id")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oName.text())
            oAreaName = oTable.cellWidget(index, 3)
            if isinstance(oAreaName, QComboBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("AreaName")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oAreaName.currentText())
            oType = oTable.cellWidget(index, 4)
            if isinstance(oType, QComboBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("type")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oType.currentText())
            oCondition = oTable.cellWidget(index, 5)
            if isinstance(oCondition, QComboBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("condition")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oCondition.currentText())
            oWidth = oTable.cellWidget(index, 7)
            if isinstance(oWidth, QDoubleSpinBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("width")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oWidth.value())
            oNum_trees = oTable.cellWidget(index, 8)
            if isinstance(oNum_trees, QSpinBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("num_trees")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oNum_trees.value())
            oType_trees = oTable.cellWidget(index, 9)
            if isinstance(oType_trees, QComboBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("type_trees")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oType_trees.currentText())
            oSize_trees = oTable.cellWidget(index, 10)
            if isinstance(oSize_trees, QComboBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("size_trees")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oSize_trees.currentText())
            oNum_shrubs = oTable.cellWidget(index, 11)
            if isinstance(oNum_shrubs, QSpinBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("num_shrubs")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oNum_shrubs.value())
            oType_shrubs = oTable.cellWidget(index, 12)
            if isinstance(oType_shrubs, QComboBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("type_shrubs")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oType_shrubs.currentText())
            oSize_shrubs = oTable.cellWidget(index, 13)
            if isinstance(oSize_shrubs, QComboBox):
                idx = self.layerOpenSurfacesLines.fields().indexOf("size_shrubs")
                if idx >= 0:
                    self.layerOpenSurfacesLines.changeAttributeValue(id, idx, oSize_shrubs.currentText())
            if self.layerOpenSurfacesLines.featureCount() > 0:
                #Emission værdier sættes til 0 aht. tematisk kort
                feat = self.layerOpenSurfacesLines.getFeature(id)
                idx = self.layerOpenSurfacesLines.fields().indexOf("EmissionsType")
                if idx >= 0:
                    if feat["EmissionsType"] == None:
                        self.layerOpenSurfacesLines.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesLines.fields().indexOf("EmissionsTreesShrubs")
                if idx >= 0:
                    if feat["EmissionsTreesShrubs"] == None:
                        self.layerOpenSurfacesLines.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesLines.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    if feat["EmissionsTotal"] == None:
                        self.layerOpenSurfacesLines.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesLines.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    if feat["EmissionsTotalM2"] == None:
                        self.layerOpenSurfacesLines.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesLines.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    if feat["EmissionsTotalM2Year"] == None:
                        self.layerOpenSurfacesLines.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesLines.fields().indexOf("EmissionsTotalPerson")
                if idx >= 0:
                    if feat["EmissionsTotalPerson"] == None:
                        self.layerOpenSurfacesLines.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesLines.fields().indexOf("EmissionsTotalPersonYear")
                if idx >= 0:
                    if feat["EmissionsTotalPersonYear"] == None:
                        self.layerOpenSurfacesLines.changeAttributeValue(id, idx, 0)

            return True

        except Exception as e:
            utils.Error(e)
            return False

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def SaveOpenSurfacesLines(self):

        try:
            if self.layerOpenSurfacesLines == None:
                return
            oTable = self.dlgMain.tblOpenSurfacesLines
            if oTable.currentRow() >= 0:
                if not self.saveRowOpenSurfacesLines(oTable.currentRow()):
                    utils.Message("Den aktive række indeholder ingen geometri. Der skal tegnes et område før rækken kan gemmes.")
                    return

            if self.layerOpenSurfacesLines.undoStack().count() == 0:
                utils.Message("Der er ingen ændringer at gemme.")
                return
            if not utils.Confirm("Er du sikker på at du vil gemme?"):
                return
            if self.layerOpenSurfacesLines.isEditable():
                if self.layerOpenSurfacesLines.isModified():
                    saveOkKomp = self.layerOpenSurfacesLines.commitChanges()
                    if not saveOkKomp:
                        if utils.Confirm("Ændringerne kunne ikke gemmes. Vis fejl?"):
                            utils.Message("\n".join(self.layerSubareas.commitErrors()))

            if self.layerOpenSurfacesLines.selectedFeatureCount() > 0:
                self.layerOpenSurfacesLines.removeSelection()
            self.FillTableOpenSurfacesLines()

        except Exception as e:
            utils.Error("Fejl i SaveOpenSurfacesLines: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def CalculateOpenSurfacesLine(self):

        try:
            self.dlgMain.lblMessage.setText("")
            if self.dlgMain.tblOpenSurfacesLines.currentRow() < 0:
                utils.Message("Vælg en række som skal beregnes")
                return
            sId = self.dlgMain.tblOpenSurfacesLines.item(self.dlgMain.tblOpenSurfacesLines.currentRow(), 0).text()
            if sId != "":
                if utils.isInt(sId):
                    self.saveRowOpenSurfacesLines(self.dlgMain.tblOpenSurfacesLines.currentRow())
                    calculate_utils.CalculateOpenSurfacesLine(self.dlgMain, self.layerOpenSurfacesLines, self.dict_translate)
            
        except Exception as e:
            utils.Error("Fejl i CalculateOpenSurfacesLine: " + str(e))

    # endregion
    # ---------------------------------------------------------------------------------------------------------
    # Areas
    # ---------------------------------------------------------------------------------------------------------
    # region
    def FillTableOpenSurfacesAreas(self):
        try:
            if self.layerOpenSurfacesAreas == None:
                return

            if self.sortColOpenSurfacesAreas <= 0:
                return

            oTable = self.dlgMain.tblOpenSurfacesAreas
            oTable.setRowCount(0)
            oTable.setColumnCount(13)
            oTable.setSortingEnabled(False)
            oTable.sortByColumn(self.sortColOpenSurfacesAreas, self.sortOrderOpenSurfacesAreas)

            oTable.setHorizontalHeaderLabels(
                ["", "", "ID", "Delområde", "Type", "Tilstand", "Areal [m\u00B2]", "Antal\ntræer", "Type af\ntræer", "Størrelse af\ntræer", "Antal\nbuske", "Type a\nbuske", "Størrelse af\nbuske"]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.setColumnHidden(0, True)  # Id skjules
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(1, 50)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
            oTable.horizontalHeader().resizeSection(2, 50)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
            oTable.horizontalHeader().resizeSection(6, 50)
            oTable.horizontalHeader().setSectionResizeMode(7, QHeaderView.Fixed) #Num_trees
            oTable.horizontalHeader().resizeSection(7, 50)
            oTable.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(9, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(10, QHeaderView.Fixed) #Num_shrubs
            oTable.horizontalHeader().resizeSection(10, 50)
            oTable.horizontalHeader().setSectionResizeMode(11, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(12, QHeaderView.Stretch)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            oTable.setWordWrap(True)

            idxName = self.layerOpenSurfacesAreas.fields().lookupField("id")
            idxAreaName = self.layerOpenSurfacesAreas.fields().lookupField("AreaName")
            idxType = self.layerOpenSurfacesAreas.fields().lookupField("type")
            idxCondition= self.layerOpenSurfacesAreas.fields().lookupField("condition")
            idxAreaCalc = self.layerOpenSurfacesAreas.fields().lookupField("AreaCalc")
            if idxName < 0 or idxType < 0 or idxCondition < 0 or idxAreaCalc < 0:
                utils.Error('Et af disse felter mangler i laget med overflader: id, type, condition, AreaCalc')
                return
            idxNum_trees = self.layerOpenSurfacesAreas.fields().lookupField("num_trees")
            idxType_trees = self.layerOpenSurfacesAreas.fields().lookupField("type_trees")
            idxSize_trees = self.layerOpenSurfacesAreas.fields().lookupField("size_trees")
            idxNum_shrubs = self.layerOpenSurfacesAreas.fields().lookupField("num_shrubs")
            idxType_shrubs = self.layerOpenSurfacesAreas.fields().lookupField("type_shrubs")
            idxSize_shrubs = self.layerOpenSurfacesAreas.fields().lookupField("size_shrubs")
            if idxNum_trees < 0 or idxType_trees < 0 or idxSize_trees < 0 or idxNum_shrubs < 0 or idxType_shrubs < 0 or idxSize_shrubs < 0:
                utils.Error('Et af disse felter mangler i laget med overflader: num_trees, type_trees, size_trees, num_shrubs, type_shrubs, size_shrubs')
                return

            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            i = 0
            featList = []
            colId = 0
            colName = 1
            colAreaName = 2
            colType = 3
            colCondition = 4
            colAreaCalc = 5
            colNum_trees = 6
            colType_trees = 7
            colSize_trees = 8
            colNum_shrubs = 9
            colType_shrubs =10
            colSize_shrubs = 11
            for feat in self.layerOpenSurfacesAreas.getFeatures():
                featList.append(
                    [
                        feat.id(),
                        feat[idxName],
                        feat[idxAreaName],
                        feat[idxType],
                        feat[idxCondition],
                        feat[idxAreaCalc],
                        feat[idxNum_trees],
                        feat[idxType_trees],
                        feat[idxSize_trees],
                        feat[idxNum_shrubs],
                        feat[idxType_shrubs],
                        feat[idxSize_shrubs],
                    ]
                )

            if self.sortColOpenSurfacesAreas == colAreaCalc+1:
                if self.sortOrderOpenSurfacesAreas == 0:
                    featList.sort(key=lambda x:float(x[self.sortColOpenSurfacesAreas - 1]), reverse=False)
                else:
                    featList.sort(key=lambda x:float(x[self.sortColOpenSurfacesAreas - 1]), reverse=True)
            else:
                if self.sortOrderOpenSurfacesAreas == 0:
                    featList.sort(key=lambda x: str(x[self.sortColOpenSurfacesAreas - 1]).replace("NULL", "").lower(), reverse=False)
                else:
                    featList.sort(key=lambda x: str(x[self.sortColOpenSurfacesAreas - 1]).replace("NULL", "").lower(), reverse=True)

            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Rediger geometri
                oEdit = QToolButton()
                oEdit.setIcon(oIcon)
                oEdit.setToolTip("Rediger geometri")
                oEdit.clicked.connect(self.editOpenSurfacesAreasClicked)
                oTable.setCellWidget(i, 1, oEdit)
                # Navn / id
                oName = QLineEdit()
                if feat[colName] == None:
                    oName.setText("")
                else:
                    oName.setText(feat[colName])
                oName.setAlignment(Qt.AlignLeft)
                oTable.setCellWidget(i, 2, oName)
                # Delområde
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.getSubareaNames() != None:
                    j = 0
                    for key in self.getSubareaNames():
                        qCombo.addItem(key, key)
                        if key == feat[colAreaName]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 3, qCombo)
                # Type
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                qCombo.addItem("", "")
                if self.dict_open_surfaces != None:
                    j = 1
                    for key, val in self.dict_open_surfaces.items():
                        if val == "area":
                            qCombo.addItem(key, key)
                            if key == feat[colType]:
                                qCombo.setCurrentIndex(j)
                            j += 1
                oTable.setCellWidget(i, 4, qCombo)
                # Tilstand
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_condition:
                        qCombo.addItem(key, key)
                        if key == feat[colCondition]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 5, qCombo)
                # Samlet areal
                oAreaCalc = QLineEdit()
                if feat[colAreaCalc] == None:
                    oAreaCalc.setText("")
                else:
                    oAreaCalc.setText(str(feat[colAreaCalc]))
                oAreaCalc.setAlignment(Qt.AlignRight)
                oAreaCalc.setEnabled(False)
                oTable.setCellWidget(i, 6, oAreaCalc)
                # Num_trees
                oBox = QSpinBox()
                oBox.setMaximum(100000)
                if feat[colNum_trees] != None:
                    oBox.setValue(feat[colNum_trees])
                oBox.setAlignment(Qt.AlignRight)
                oTable.setCellWidget(i, 7, oBox)
                # Type_trees
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_treetype:
                        qCombo.addItem(key, key)
                        if key == feat[colType_trees]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 8, qCombo)
                # Size_trees
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_treesize:
                        qCombo.addItem(key, key)
                        if key == feat[colSize_trees]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 9, qCombo)
                # Num_shrubs
                oBox = QSpinBox()
                oBox.setMaximum(100000)
                if feat[colNum_shrubs] != None:
                    oBox.setValue(feat[colNum_shrubs])
                oBox.setAlignment(Qt.AlignRight)
                oTable.setCellWidget(i, 10, oBox)
                # Type_shrubs
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_treetype:
                        qCombo.addItem(key, key)
                        if key == feat[colType_shrubs]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 11, qCombo)
                # Size_shrubs
                qCombo = QComboBox()
                qCombo.wheelEvent = lambda event: None
                if self.dict_surface_condition != None:
                    j = 0
                    for key in self.dict_surface_shrubsize:
                        qCombo.addItem(key, key)
                        if key == feat[colSize_shrubs]:
                            qCombo.setCurrentIndex(j)
                        j += 1
                oTable.setCellWidget(i, 12, qCombo)

                i += 1

            # Der indsættes en tom række
            self.InsertBlankRowOpenSurfacesAreas(oTable)

        except Exception as e:
            utils.Error("Fejl i FillTableOpenSurfacesAreas. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def InsertBlankRowOpenSurfacesAreas(self, oTable):
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "../icons/edit.png")
            oIcon = QIcon(icon_path)

            i = oTable.verticalHeader().count()
            # Der indsættes en tom række
            oTable.insertRow(i)
            # Id
            oId = QTableWidgetItem(-1)
            oId.setFlags(Qt.NoItemFlags)  # not enabled
            oTable.setItem(i, 0, oId)
            # Rediger geometri
            qMapLookup = QToolButton()
            qMapLookup.setIcon(oIcon)
            qMapLookup.setToolTip("Opret geometri")
            qMapLookup.clicked.connect(self.editOpenSurfacesAreasClicked)
            oTable.setCellWidget(i, 1, qMapLookup)
            # Navn / id
            oName = QLineEdit()
            oName.setText("")
            oName.setAlignment(Qt.AlignLeft)
            oTable.setCellWidget(i, 2, oName)
            # Delområde
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.getSubareaNames() != None:
                for key in self.getSubareaNames():
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 3, qCombo)
            # Type
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            qCombo.addItem("", "")
            if self.dict_open_surfaces != None:
                for key, val in self.dict_open_surfaces.items():
                    if val == "area":
                        qCombo.addItem(key, key)
            oTable.setCellWidget(i, 4, qCombo)
            # Tilstand
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_condition:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 5, qCombo)
            # Grundareal
            oAreaCalc = QLineEdit()
            oAreaCalc.setText("")
            oAreaCalc.setEnabled(False)
            oTable.setCellWidget(i, 6, oAreaCalc)
            # Num_trees
            oBox = QSpinBox()
            oBox.setAlignment(Qt.AlignRight)
            oBox.setMaximum(100000)
            oTable.setCellWidget(i, 7, oBox)
            # Type_trees
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_treetype:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 8, qCombo)
            # Size_trees
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_treesize:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 9, qCombo)
            # Num_shrubs
            oBox = QSpinBox()
            oBox.setAlignment(Qt.AlignRight)
            oBox.setMaximum(100000)
            oTable.setCellWidget(i, 10, oBox)
            # Type_shrubs
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_treetype:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 11, qCombo)
            # Size_shrubs
            qCombo = QComboBox()
            qCombo.wheelEvent = lambda event: None
            if self.dict_surface_condition != None:
                for key in self.dict_surface_shrubsize:
                    qCombo.addItem(key, key)
            oTable.setCellWidget(i, 12, qCombo)

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count() - 1):
                lstRowHeader.append("")
            lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i InsertBlankRowOpenSurfacesAreas. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def editOpenSurfacesAreasClicked(self):
        try:
            self.resetGeometryEdit()

            oTable = self.dlgMain.tblOpenSurfacesAreas
            ch = self.sender()
            ix = oTable.indexAt(ch.pos())
            oTable.selectRow(ix.row())
            
            sId = oTable.item(ix.row(), 0).text()
            if sId == "":
                # Opret ny line
                self.ActiveTool = "CreateSurfaceArea"
                self.dlgMain.lblMessage.setText("Tegn polygon i kortet.")
            else:
                # Rediger eksisterende linie
                if not utils.isInt(sId):
                    utils.Error("Ugyldig polygon valgt.")
                    return
                id = int(sId)
                self.layerOpenSurfacesAreas.selectByIds([id])
                if self.layerOpenSurfacesAreas.selectedFeatureCount() == 0:
                    utils.Error("Kunne ikke finde polygon i kortet.")
                elif self.layerOpenSurfacesAreas.selectedFeatureCount() == 1:
                    # Hvis polygonen ikke er synlig flyttes kortudsnit til contrum af polygonen
                    polygonGeometry = self.layerOpenSurfacesAreas.selectedFeatures()[
                        0
                    ].geometry()
                    if not polygonGeometry.within(
                        QgsGeometry.fromRect(self.iface.mapCanvas().extent())
                    ):
                        self.iface.mapCanvas().setCenter(
                            polygonGeometry.centroid().asPoint()
                        )
                    if not self.layerOpenSurfacesAreas.isEditable():
                        self.layerOpenSurfacesAreas.startEditing()
                    self.ActiveTool = "EditSurfaceArea"
                    self.dlgMain.lblMessage.setText("Vælg vertex og flyt det.")
                else:
                    utils.Error("Flere polygoner valgt i kortet.")

        except Exception as e:
            utils.Error("Fejl i editOpenSurfacesAreasClicked. Fejlbesked: " + str(e))
            return -1

    # ---------------------------------------------------------------------------------------------------------
    # Der skiftes sortering efter klik i header
    # ---------------------------------------------------------------------------------------------------------
    def sortingChangedOpenSurfacesAreas(self, col, sortorder):
        if col <= 0:
            self.dlgMain.tblOpenSurfacesAreas.horizontalHeader().setSortIndicatorShown(False)
        else:
            self.sortColOpenSurfacesAreas = col
            self.sortOrderOpenSurfacesAreas = sortorder
            self.FillTableOpenSurfacesAreas()

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def keyPressEventOpenSurfacesAreas(self, e):

        try:
            if e.key() == Qt.Key_Delete:
                oTable = self.dlgMain.tblOpenSurfacesAreas
                sId = oTable.item(oTable.currentRow(), 0).text()
                if sId == "":
                    utils.Error("Kan ikke fastlægge rækkens id.")
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig række valgt.")
                        return
                    id = int(sId)
                    if utils.Confirm(
                        "Er du sikker på at du ønsker at slette den valgte række?"
                    ):
                        oTable.removeRow(oTable.currentRow())
                        if not self.layerOpenSurfacesAreas.isEditable():
                            self.layerOpenSurfacesAreas.startEditing()
                        self.layerOpenSurfacesAreas.deleteFeature(id)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedOpenSurfacesAreas(self, index):

        try:
            self.saveRowOpenSurfacesAreas(index)
            self.selectRowOpenSurfacesAreas(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedOpenSurfacesAreas(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            oldrow = indexOld.row()
            if oldrow >= 0:
                self.saveRowOpenSurfacesAreas(oldrow)
            self.selectRowOpenSurfacesAreas(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Række og polygon highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowOpenSurfacesAreas(self, index):

        try:
            oTable = self.dlgMain.tblOpenSurfacesAreas
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)
                # Polygon vælges i kortet
                sId = oTable.item(index, 0).text()
                if sId == "":
                    #Tom række
                    self.layerOpenSurfacesAreas.removeSelection()
                else:
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig polygon valgt.")
                        return
                    id = int(sId)
                    self.layerOpenSurfacesAreas.selectByIds([id])

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Data gemmes
    # ---------------------------------------------------------------------------------------------------------
    def saveRowOpenSurfacesAreas(self, index):

        try:
            # Data fra tabel gemmes
            oTable = self.dlgMain.tblOpenSurfacesAreas
            sId = oTable.item(index, 0).text()
            if sId == "":
                return False
            if not utils.isInt(sId):
                utils.Error("Ugyldig række valgt.")
                return False

            id = int(sId)
            if not self.layerOpenSurfacesAreas.isEditable():
                self.layerOpenSurfacesAreas.startEditing()
            oName = oTable.cellWidget(index, 2)
            if isinstance(oName, QLineEdit):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("id")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oName.text())
            oAreaName = oTable.cellWidget(index, 3)
            if isinstance(oAreaName, QComboBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("AreaName")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oAreaName.currentText())
            oType = oTable.cellWidget(index, 4)
            if isinstance(oType, QComboBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("type")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oType.currentText())
            oCondition = oTable.cellWidget(index, 5)
            if isinstance(oCondition, QComboBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("condition")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oCondition.currentText())
            oNum_trees = oTable.cellWidget(index, 7)
            if isinstance(oNum_trees, QSpinBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("num_trees")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oNum_trees.value())
            oType_trees = oTable.cellWidget(index, 8)
            if isinstance(oType_trees, QComboBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("type_trees")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oType_trees.currentText())
            oSize_trees = oTable.cellWidget(index, 9)
            if isinstance(oSize_trees, QComboBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("size_trees")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oSize_trees.currentText())
            oNum_shrubs = oTable.cellWidget(index, 10)
            if isinstance(oNum_shrubs, QSpinBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("num_shrubs")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oNum_shrubs.value())
            oType_shrubs = oTable.cellWidget(index, 11)
            if isinstance(oType_shrubs, QComboBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("type_shrubs")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oType_shrubs.currentText())
            oSize_shrubs = oTable.cellWidget(index, 12)
            if isinstance(oSize_shrubs, QComboBox):
                idx = self.layerOpenSurfacesAreas.fields().indexOf("size_shrubs")
                if idx >= 0:
                    self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, oSize_shrubs.currentText())
            if self.layerOpenSurfacesAreas.featureCount() > 0:
                #Emission værdier sættes til 0 aht. tematisk kort
                feat = self.layerOpenSurfacesAreas.getFeature(id)
                idx = self.layerOpenSurfacesAreas.fields().indexOf("EmissionsType")
                if idx >= 0:
                    if feat["EmissionsType"] == None:
                        self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesAreas.fields().indexOf("EmissionsTreesShrubs")
                if idx >= 0:
                    if feat["EmissionsTreesShrubs"] == None:
                        self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesAreas.fields().indexOf("EmissionsTotal")
                if idx >= 0:
                    if feat["EmissionsTotal"] == None:
                        self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesAreas.fields().indexOf("EmissionsTotalM2")
                if idx >= 0:
                    if feat["EmissionsTotalM2"] == None:
                        self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesAreas.fields().indexOf("EmissionsTotalM2Year")
                if idx >= 0:
                    if feat["EmissionsTotalM2Year"] == None:
                        self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesAreas.fields().indexOf("EmissionsTotalPerson")
                if idx >= 0:
                    if feat["EmissionsTotalPerson"] == None:
                        self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, 0)
                idx = self.layerOpenSurfacesAreas.fields().indexOf("EmissionsTotalPersonYear")
                if idx >= 0:
                    if feat["EmissionsTotalPersonYear"] == None:
                        self.layerOpenSurfacesAreas.changeAttributeValue(id, idx, 0)

            return True
        
        except Exception as e:
            utils.Error("Fejl i saveRowOpenSurfacesAreas. Fejlbesked: " + str(e))
            return False

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def SaveOpenSurfacesAreas(self):

        try:
            if self.layerOpenSurfacesAreas == None:
                return
            oTable = self.dlgMain.tblOpenSurfacesAreas
            if oTable.currentRow() >= 0:
                if not self.saveRowOpenSurfacesAreas(oTable.currentRow()):
                    utils.Message("Den aktive række indeholder ingen geometri. Der skal tegnes et område før rækken kan gemmes.")
                    return

            # if oTable.item(oTable.verticalHeader().count()-1, 0).text() == "":
            #     if not utils.Confirm("Den seneste række indeholder ingen geometri. Denne række vil ikke blive gemt. Ønsker du at fortsætte?"):
            #         return
                    
            #self.setStatusSaveButtons()
            if self.layerOpenSurfacesAreas.undoStack().count() == 0:
                utils.Message("Der er ingen ændringer at gemme.")
                return
            if not utils.Confirm("Er du sikker på at du vil gemme?"):
                return
            if self.layerOpenSurfacesAreas.isEditable():
                if self.layerOpenSurfacesAreas.isModified():
                    saveOkKomp = self.layerOpenSurfacesAreas.commitChanges()
                    if not saveOkKomp:
                        if utils.Confirm("Ændringerne kunne ikke gemmes. Vis fejl?"):
                            utils.Message("\n".join(self.layerSubareas.commitErrors()))

            if self.layerOpenSurfacesAreas.selectedFeatureCount() > 0:
                self.layerOpenSurfacesAreas.removeSelection()
            self.FillTableOpenSurfacesAreas()

        except Exception as e:
            utils.Error("Fejl i SaveOpenSurfacesAreas: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def CalculateOpenSurfacesArea(self):

        try:
            self.dlgMain.lblMessage.setText("")
            if self.dlgMain.tblOpenSurfacesAreas.currentRow() < 0:
                utils.Message("Vælg en række som skal beregnes")
                return
            sId = self.dlgMain.tblOpenSurfacesAreas.item(self.dlgMain.tblOpenSurfacesAreas.currentRow(), 0).text()
            if sId != "":
                if utils.isInt(sId):
                    self.saveRowOpenSurfacesAreas(self.dlgMain.tblOpenSurfacesAreas.currentRow())
                    calculate_utils.CalculateOpenSurfacesArea(self.dlgMain, self.layerOpenSurfacesAreas, self.dict_translate)
            
        except Exception as e:
            utils.Error("Fejl i CalculateOpenSurfacesArea: " + str(e))

    # endregion
    # ---------------------------------------------------------------------------------------------------------
    # RESULTATER
    # ---------------------------------------------------------------------------------------------------------
    def FillTableBuildingsResult(self):
        try:
            if self.layerBuildings == None:
                return

            if self.sortColBuildingsResult <= 0:
                return

            oTable = self.dlgMain.tblBuildingsResult
            oTable.setRowCount(0)
            oTable.setColumnCount(5)
            oTable.setSortingEnabled(True)
            oTable.sortByColumn(self.sortColBuildingsResult, self.sortOrderBuildingsResult)

            oTable.setHorizontalHeaderLabels(
                [
                    "",
                    "ID",
                    "Anvendelse",
                    "Total [kg CO\u2082e]",
                    "Pr. areal [kg CO\u2082e/m\u00B2/år]",
                ]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            oTable.setColumnHidden(0, True)  # Id skjules
            oTable.setWordWrap(True)

            idxName = self.layerBuildings.fields().lookupField("id")
            idxAnvendelse = self.layerBuildings.fields().lookupField("Building_usage")
            idxEmissionsTotal = self.layerBuildings.fields().lookupField("EmissionsTotal")
            idxEmissionsTotalM2Year = self.layerBuildings.fields().lookupField("EmissionsTotalM2Year")
            if idxName < 0 or idxAnvendelse < 0 or idxEmissionsTotal < 0 or idxEmissionsTotalM2Year < 0:
                utils.Error('Et af disse felter mangler i laget med bygninger: id, Building_usage, EmissionsTotal, EmissionsTotalM2Year')
                return

            i = 0
            featList = []
            colId = 0
            colName = 1
            colAnvendelse = 2
            colEmissionsTotal = 3
            colEmissionsTotalM2Year = 4

            for feat in self.layerBuildings.getFeatures():
                featList.append(
                    [
                        feat.id(),
                        feat[idxName],
                        feat[idxAnvendelse],
                        feat[idxEmissionsTotal],
                        feat[idxEmissionsTotalM2Year],
                    ]
                )
            if self.sortColBuildingsResult == colEmissionsTotal or self.sortColBuildingsResult == colEmissionsTotalM2Year:
                #Lidt langsomt - overvej at indsætte 0 i stedet for NULL i listen
                if self.sortOrderBuildingsResult == 0:
                    featList.sort(key=lambda x:float(str(x[self.sortColBuildingsResult]).replace("NULL", "0")), reverse=False)
                else:
                    featList.sort(key=lambda x:float(str(x[self.sortColBuildingsResult]).replace("NULL", "0")), reverse=True)
            else:
                if self.sortOrderBuildingsResult == 0:
                    featList.sort(key=lambda x: (str(x[self.sortColBuildingsResult]).replace("NULL", "").lower(), x[0]), reverse=False)
                else:
                    featList.sort(key=lambda x: (str(x[self.sortColBuildingsResult]).replace("NULL", "").lower(), x[0]), reverse=True)

            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Navn
                oName = QLineEdit()
                if feat[colName] == None:
                    oName.setText("")
                else:
                    oName.setText(feat[colName])
                oName.setAlignment(Qt.AlignCenter)
                oName.setEnabled(False)
                oTable.setCellWidget(i, 1, oName)
                # Anvendelse
                oAnvendelse = QLineEdit()
                if feat[colAnvendelse] == None:
                    oAnvendelse.setText("")
                else:
                    oAnvendelse.setText(feat[colAnvendelse])
                oAnvendelse.setAlignment(Qt.AlignCenter)
                oAnvendelse.setEnabled(False)
                oTable.setCellWidget(i, 2, oAnvendelse)
                # EmissionsTotal
                oEmissionsTotal = QLineEdit()
                if feat[colEmissionsTotal] == None:
                    oEmissionsTotal.setText("")
                else:
                    oEmissionsTotal.setText(str(feat[colEmissionsTotal]))
                oEmissionsTotal.setAlignment(Qt.AlignRight)
                oEmissionsTotal.setEnabled(False)
                oTable.setCellWidget(i, 3, oEmissionsTotal)
                # EmissionsTotalM2Year
                oEmissionsTotalM2Year = QLineEdit()
                if feat[colEmissionsTotalM2Year] == None:
                    oEmissionsTotalM2Year.setText("")
                else:
                    oEmissionsTotalM2Year.setText(str(feat[colEmissionsTotalM2Year]))
                oEmissionsTotalM2Year.setAlignment(Qt.AlignRight)
                oEmissionsTotalM2Year.setEnabled(False)
                oTable.setCellWidget(i, 4, oEmissionsTotalM2Year)

                i += 1

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count() - 1):
                lstRowHeader.append("")
            lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i FillTableBuildingsResult. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def sortingChangedBuildingsResult(self, col, sortorder):
        if col <= 0:
            self.dlgMain.tblBuildingsResult.horizontalHeader().setSortIndicatorShown(False)
        else:
            self.sortColBuildingsResult = col
            self.sortOrderBuildingsResult = sortorder
            self.FillTableBuildingsResult()

    # ---------------------------------------------------------------------------------------------------------
    # Når der klikkes i 'verticalHeader' vælges polygon
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedBuildingsResult(self, index):

        try:
            self.selectRowBuildingsResult(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedBuildingsResult(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            self.selectRowBuildingsResult(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Række og polygon highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowBuildingsResult(self, index):

        try:
            oTable = self.dlgMain.tblBuildingsResult
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)
                # Polygon vælges i kortet
                sId = oTable.item(index, 0).text()
                if sId != "":
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig bygning valgt.")
                        return
                    id = int(sId)
                    self.layerBuildings.selectByIds([id])
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # 
    # ---------------------------------------------------------------------------------------------------------
    def FillTableRoadsResult(self):
        try:
            if self.layerOpenSurfacesLines == None:
                return

            if self.sortColRoadsResult <= 0:
                return

            oTable = self.dlgMain.tblRoadsResult
            oTable.setRowCount(0)
            oTable.setColumnCount(5)
            oTable.setSortingEnabled(True)
            oTable.sortByColumn(self.sortColRoadsResult, self.sortOrderRoadsResult)

            oTable.setHorizontalHeaderLabels(
                [
                    "",
                    "ID",
                    "Anvendelse",
                    "Total [kg CO\u2082e]",
                    "Pr. areal [kg CO\u2082e/m\u00B2/år]",
                ]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            oTable.setColumnHidden(0, True)  # Id skjules
            oTable.setWordWrap(True)

            idxName = self.layerOpenSurfacesLines.fields().lookupField("id")
            idxType = self.layerOpenSurfacesLines.fields().lookupField("type")
            idxEmissionsTotal = self.layerOpenSurfacesLines.fields().lookupField("EmissionsTotal")
            idxEmissionsTotalM2Year = self.layerOpenSurfacesLines.fields().lookupField("EmissionsTotalM2Year")
            if idxName < 0 or idxType < 0 or idxEmissionsTotal < 0 or idxEmissionsTotalM2Year < 0:
                utils.Error('Et af disse felter mangler i laget med veje: id, type, EmissionsTotal, EmissionsTotalM2Year')
                return

            i = 0
            featList = []
            colId = 0
            colName = 1
            colType = 2
            colEmissionsTotal = 3
            colEmissionsTotalM2Year = 4

            for feat in self.layerOpenSurfacesLines.getFeatures():
                featList.append(
                    [
                        feat.id(),
                        feat[idxName],
                        feat[idxType],
                        feat[idxEmissionsTotal],
                        feat[idxEmissionsTotalM2Year],
                    ]
                )
            if self.sortColRoadsResult == colEmissionsTotal or self.sortColRoadsResult == colEmissionsTotalM2Year:
                #Lidt langsomt - overvej at indsætte 0 i stedet for NULL i listen
                if self.sortOrderRoadsResult == 0:
                    featList.sort(key=lambda x:float(str(x[self.sortColRoadsResult]).replace("NULL", "0")), reverse=False)
                else:
                    featList.sort(key=lambda x:float(str(x[self.sortColRoadsResult]).replace("NULL", "0")), reverse=True)
            else:
                if self.sortOrderRoadsResult == 0:
                    featList.sort(key=lambda x: (str(x[self.sortColRoadsResult]).replace("NULL", "").lower(), x[0]), reverse=False)
                else:
                    featList.sort(key=lambda x: (str(x[self.sortColRoadsResult]).replace("NULL", "").lower(), x[0]), reverse=True)

            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Navn
                oName = QLineEdit()
                if feat[colName] == None:
                    oName.setText("")
                else:
                    oName.setText(feat[colName])
                oName.setAlignment(Qt.AlignCenter)
                oName.setEnabled(False)
                oTable.setCellWidget(i, 1, oName)
                # Type
                oType = QLineEdit()
                if feat[colType] == None:
                    oType.setText("")
                else:
                    oType.setText(feat[colType])
                oType.setAlignment(Qt.AlignCenter)
                oType.setEnabled(False)
                oTable.setCellWidget(i, 2, oType)
                # EmissionsTotal
                oEmissionsTotal = QLineEdit()
                if feat[colEmissionsTotal] == None:
                    oEmissionsTotal.setText("")
                else:
                    oEmissionsTotal.setText(str(feat[colEmissionsTotal]))
                oEmissionsTotal.setAlignment(Qt.AlignRight)
                oEmissionsTotal.setEnabled(False)
                oTable.setCellWidget(i, 3, oEmissionsTotal)
                # EmissionsTotalM2Year
                oEmissionsTotalM2Year = QLineEdit()
                if feat[colEmissionsTotalM2Year] == None:
                    oEmissionsTotalM2Year.setText("")
                else:
                    oEmissionsTotalM2Year.setText(str(feat[colEmissionsTotalM2Year]))
                oEmissionsTotalM2Year.setAlignment(Qt.AlignRight)
                oEmissionsTotalM2Year.setEnabled(False)
                oTable.setCellWidget(i, 4, oEmissionsTotalM2Year)

                i += 1

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count() - 1):
                lstRowHeader.append("")
            lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i FillTableRoadsResult. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def sortingChangedRoadsResult(self, col, sortorder):

        try:
            if col <= 0:
                self.dlgMain.tblRoadsResult.horizontalHeader().setSortIndicatorShown(False)
            else:
                self.sortColRoadsResult = col
                self.sortOrderRoadsResult = sortorder
                self.FillTableRoadsResult()

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der klikkes i 'verticalHeader' vælges polygon
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedRoadsResult(self, index):

        try:
            self.selectRowRoadsResult(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedRoadsResult(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            self.selectRowRoadsResult(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Række og polygon highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowRoadsResult(self, index):

        try:
            oTable = self.dlgMain.tblRoadsResult
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)
                # Polygon vælges i kortet
                sId = oTable.item(index, 0).text()
                if sId != "":
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig bygning valgt.")
                        return
                    id = int(sId)
                    self.layerOpenSurfacesLines.selectByIds([id])
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # 
    # ---------------------------------------------------------------------------------------------------------
    def FillTableSurfacesResult(self):
        try:
            if self.layerOpenSurfacesAreas == None:
                return

            if self.sortColSurfacesResult <= 0:
                return

            oTable = self.dlgMain.tblSurfacesResult
            oTable.setRowCount(0)
            oTable.setColumnCount(5)
            oTable.setSortingEnabled(True)
            oTable.sortByColumn(self.sortColSurfacesResult, self.sortOrderSurfacesResult)

            oTable.setHorizontalHeaderLabels(
                [
                    "",
                    "ID",
                    "Anvendelse",
                    "Total [kg CO\u2082e]",
                    "Pr. areal [kg CO\u2082e/m\u00B2/år]",
                ]
            )
            oTable.horizontalHeader().setFixedHeight(50)  # Header height
            oTable.verticalHeader().setVisible(True)
            oTable.horizontalHeader().setDefaultSectionSize(80)
            oTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            oTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
            oTable.verticalHeader().setDefaultSectionSize(25)  # row height
            oTable.verticalHeader().setFixedWidth(20)
            oTable.setColumnHidden(0, True)  # Id skjules
            oTable.setWordWrap(True)

            idxName = self.layerOpenSurfacesAreas.fields().lookupField("id")
            idxType = self.layerOpenSurfacesAreas.fields().lookupField("type")
            idxEmissionsTotal = self.layerOpenSurfacesAreas.fields().lookupField("EmissionsTotal")
            idxEmissionsTotalM2Year = self.layerOpenSurfacesAreas.fields().lookupField("EmissionsTotalM2Year")
            if idxName < 0 or idxType < 0 or idxEmissionsTotal < 0 or idxEmissionsTotalM2Year < 0:
                utils.Error('Et af disse felter mangler i laget med veje: id, type, EmissionsTotal, EmissionsTotalM2Year')
                return

            i = 0
            featList = []
            colId = 0
            colName = 1
            colType = 2
            colEmissionsTotal = 3
            colEmissionsTotalM2Year = 4

            for feat in self.layerOpenSurfacesAreas.getFeatures():
                featList.append(
                    [
                        feat.id(),
                        feat[idxName],
                        feat[idxType],
                        feat[idxEmissionsTotal],
                        feat[idxEmissionsTotalM2Year],
                    ]
                )
            #JCN/2024-11-08: Resultater fra belægninger vises også
            #Check af felter i Subareas
            idxNameSubarea = self.layerSubareas.fields().indexOf("Name")
            if idxNameSubarea < 0:
                utils.Error("Et af disse felter mangler i laget med delområder: Name")
                return 
            #Findes nye felter i Subareas?
            idxEmTotSubarea = self.layerSubareas.fields().indexOf("EmissionsTotal")
            idxEmTotM2YearSubarea = self.layerSubareas.fields().indexOf("EmissionsTotalM2Year")
            if idxEmTotSubarea < 0 or idxEmTotM2YearSubarea < 0:
                utils.Error('Et af disse felter mangler i laget med delområder: EmissionsTotal, EmissionsTotalM2Year')
                return
            for feat in self.layerSubareas.getFeatures():
                featList.append(
                    [
                        feat.id(),
                        feat[idxNameSubarea],
                        'Belægninger',
                        feat[idxEmTotSubarea],
                        feat[idxEmTotM2YearSubarea],
                    ]
                )

            if self.sortColSurfacesResult == colEmissionsTotal or self.sortColSurfacesResult == colEmissionsTotalM2Year:
                #Lidt langsomt - overvej at indsætte 0 i stedet for NULL i listen
                if self.sortOrderSurfacesResult == 0:
                    featList.sort(key=lambda x:float(str(x[self.sortColSurfacesResult]).replace("NULL", "0")), reverse=False)
                else:
                    featList.sort(key=lambda x:float(str(x[self.sortColSurfacesResult]).replace("NULL", "0")), reverse=True)
            else:
                if self.sortOrderSurfacesResult == 0:
                    featList.sort(key=lambda x: (str(x[self.sortColSurfacesResult]).replace("NULL", "").lower(), x[0]), reverse=False)
                else:
                    featList.sort(key=lambda x: (str(x[self.sortColSurfacesResult]).replace("NULL", "").lower(), x[0]), reverse=True)

            for feat in featList:
                oTable.insertRow(i)
                # Id
                oId = QTableWidgetItem(str(feat[colId]))
                oId.setFlags(Qt.NoItemFlags)  # not enabled
                oTable.setItem(i, 0, oId)
                # Navn
                oName = QLineEdit()
                if feat[colName] == None:
                    oName.setText("")
                else:
                    oName.setText(feat[colName])
                oName.setAlignment(Qt.AlignCenter)
                oName.setEnabled(False)
                oTable.setCellWidget(i, 1, oName)
                # Type
                oType = QLineEdit()
                if feat[colType] == None:
                    oType.setText("")
                else:
                    oType.setText(feat[colType])
                oType.setAlignment(Qt.AlignCenter)
                oType.setEnabled(False)
                oTable.setCellWidget(i, 2, oType)
                # EmissionsTotal
                oEmissionsTotal = QLineEdit()
                if feat[colEmissionsTotal] == None:
                    oEmissionsTotal.setText("")
                else:
                    oEmissionsTotal.setText(str(feat[colEmissionsTotal]))
                oEmissionsTotal.setAlignment(Qt.AlignRight)
                oEmissionsTotal.setEnabled(False)
                oTable.setCellWidget(i, 3, oEmissionsTotal)
                # EmissionsTotalM2Year
                oEmissionsTotalM2Year = QLineEdit()
                if feat[colEmissionsTotalM2Year] == None:
                    oEmissionsTotalM2Year.setText("")
                else:
                    oEmissionsTotalM2Year.setText(str(feat[colEmissionsTotalM2Year]))
                oEmissionsTotalM2Year.setAlignment(Qt.AlignRight)
                oEmissionsTotalM2Year.setEnabled(False)
                oTable.setCellWidget(i, 4, oEmissionsTotalM2Year)

                i += 1

            lstRowHeader = []
            for i in range(0, oTable.verticalHeader().count()):
                lstRowHeader.append("")
            #lstRowHeader.append("*")
            oTable.setVerticalHeaderLabels(lstRowHeader)

        except Exception as e:
            utils.Error("Fejl i FillTableSurfacesResult. Fejlbesked: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def sortingChangedSurfacesResult(self, col, sortorder):
        if col <= 0:
            self.dlgMain.tblSurfacesResult.horizontalHeader().setSortIndicatorShown(False)
        else:
            self.sortColSurfacesResult = col
            self.sortOrderSurfacesResult = sortorder
            self.FillTableSurfacesResult()

    # ---------------------------------------------------------------------------------------------------------
    # Når der klikkes i 'verticalHeader' vælges polygon
    # ---------------------------------------------------------------------------------------------------------
    def rowClickedSurfacesResult(self, index):

        try:
            self.selectRowSurfacesResult(index)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Når der skiftes række gemmes den gamle række og den nye række highlightes
    # ---------------------------------------------------------------------------------------------------------
    def rowChangedSurfacesResult(self, indexNew, indexOld):

        try:
            self.dlgMain.lblMessage.setText('')
            self.dlgMain.lblMessage.setStyleSheet('color: black')

            newrow = indexNew.row()
            if newrow < 0:
                return
            self.selectRowSurfacesResult(newrow)

        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    # Række og polygon highlightes
    # ---------------------------------------------------------------------------------------------------------
    def selectRowSurfacesResult(self, index):

        try:
            oTable = self.dlgMain.tblSurfacesResult
            if index < oTable.verticalHeader().count():
                # Highlight selected row
                for i in range(0, oTable.verticalHeader().count() - 1):
                    oTable.setVerticalHeaderItem(i, QTableWidgetItem())
                item = QTableWidgetItem("")
                item.setBackground(QColor(200, 200, 200))
                oTable.setVerticalHeaderItem(index, item)
                # Polygon vælges i kortet
                sId = oTable.item(index, 0).text()
                if sId != "":
                    if not utils.isInt(sId):
                        utils.Error("Ugyldig bygning valgt.")
                        return
                    id = int(sId)
                    self.layerOpenSurfacesAreas.selectByIds([id])
            
        except Exception as e:
            utils.Error(e)

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def Calculate(self):
        
        try:
            if self.dlgMain.cmbStudyPeriod.currentText() == "":
                utils.Error("Der er ikke valgt beregningsperiode under indstillinger")
                return
            if self.dlgMain.txtConstructionYear.text() == "":
                utils.Error("Der er ikke indtastet årstal for opførelse under indstillinger")
                return
            #Buildings
            sumTotalEmissionsBuildings, sumBuildingArea = calculate_utils.CalculateAllBuildings(self.dlgMain, self.layerBuildings, self.dict_translate)
            if sumTotalEmissionsBuildings == -1:
                return
            #Roads and paths
            sumRoads, sumTreesLines = calculate_utils.CalculateAllOpenSurfacesLines(self.dlgMain, self.layerOpenSurfacesLines, self.dict_translate, self.dict_surface_width)
            if sumRoads == -1 and sumTreesLines == -1:
                return
            #Areas
            sumHardSurfaces, sumGreenSurfaces, sumTreesAreas = calculate_utils.CalculateAllOpenSurfacesAreas(self.dlgMain, self.layerOpenSurfacesAreas, self.dict_translate)
            if sumHardSurfaces == -1 and sumGreenSurfaces == -1 and sumTreesAreas == -1:
                return
            #Subareas - belægninger
            sumHardSurfacesExtra, sumGreenSurfacesExtra = calculate_utils.CalculateAllSurfaces(self.dlgMain, self.layerSubareas, self.layerSurfaces, self.layerBuildings,self.layerOpenSurfacesLines, self.layerOpenSurfacesAreas, self.dict_translate, self.dict_surface_width)
            if sumHardSurfacesExtra == -999 and sumGreenSurfacesExtra == -999:
                return
            sumHardSurfaces += sumHardSurfacesExtra
            sumGreenSurfaces += sumGreenSurfacesExtra

            studyPeriod = int(self.dlgMain.cmbStudyPeriod.currentText())
            sumTotalEmissions = sumTotalEmissionsBuildings + sumRoads + sumTreesLines + sumHardSurfaces + sumGreenSurfaces + sumTreesAreas
            if sumBuildingArea == 0:
                sumTotalEmissionsM2Year = 0
            else:
                sumTotalEmissionsM2Year = sumTotalEmissions / (sumBuildingArea * studyPeriod)
            #Hvis der er angivet total number of residents/workplaces under indstillinger, skal disse bruges til beregning pr. person
            iTotalNumberofPersons = 0
            if self.dlgMain.txtTotalNumWorkplaces.text() != "":
                iTotalNumberofPersons = int(self.dlgMain.txtTotalNumWorkplaces.text())
            if self.dlgMain.txtTotalNumResidents.text() != "":
                iTotalNumberofPersons += int(self.dlgMain.txtTotalNumResidents.text())
            sumTotalEmissionsPersonYear = 0
            if iTotalNumberofPersons > 0:
                sumTotalEmissionsPersonYear = sumTotalEmissions /(iTotalNumberofPersons * studyPeriod)
            iSumTotalEmissionsTONS = round(sumTotalEmissions/1000)

            self.dlgMain.result_total_co2.setText(str(iSumTotalEmissionsTONS))
            self.dlgMain.result_sqm_co2.setText(format(sumTotalEmissionsM2Year, ",.2f"))
            self.dlgMain.result_person_co2.setText(format(sumTotalEmissionsPersonYear, ",.2f"))
            
            utils.SaveSetting("SumTotalEmissions", round(sumTotalEmissions))
            utils.SaveSetting("SumTotalEmissionsM2Year", sumTotalEmissionsM2Year)
            utils.SaveSetting("SumTotalEmissionsPersonYear", sumTotalEmissionsPersonYear)

            self.dlgMain.lblMessage.setText("Totale emissioner fra lokalplanen [tons CO\u2082e]: " + str(iSumTotalEmissionsTONS))

            #Tabeller med resultater opdateres
            self.FillTableBuildingsResult()
            self.FillTableRoadsResult()
            self.FillTableSurfacesResult()

            #Grafer opdateres
            #self.dlgMain.lblMessage.setText("Opdaterer grafer...")
            QApplication.processEvents() 
            plotly_utils.updateWidget(self.dlgMain.tabLayout_1, sumTotalEmissionsBuildings, sumRoads + sumHardSurfaces, sumGreenSurfaces + sumTreesLines + sumTreesAreas)
            
        except Exception as e:
            utils.Error("Fejl i Calculate: " + str(e))
        
    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def GogglesM2(self):
        try:
            utils.Message('Ikke implementeret')

        except Exception as e:
            utils.Error("Fejl i GogglesM2: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def GogglesTot(self):
        try:
            utils.Message('Ikke implementeret')

        except Exception as e:
            utils.Error("Fejl i GogglesTot: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def ExportReport(self):

        try:
            #Kontrollerer at bygningslag mv. findes
            if self.layerBuildings == None:
                utils.Message("Kan ikke fastlægge laget med bygninger")
                return

            from ..tools import report_utils
            report_utils.ExportReport(self, self.iface, self.dlgMain, self.dict_translate, self.layerBuildings, self.layerOpenSurfacesLines, self.layerOpenSurfacesAreas, self.layerSubareas, self.dict_surface_width)

        except Exception as e:
            utils.Error("Fejl i ExportReport: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def __onDialogHidden(self):

        try:
            self.iface.mapCanvas().unsetMapTool(self)

        except Exception as e:
            utils.Error("Fejl i __onDialogHidden: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def __cleanup(self):

        try:
            if self.rubberband:
                self.iface.mapCanvas().scene().removeItem(self.rubberband)
            self.rubberband = None
            if self.rubberbandLine:
                self.iface.mapCanvas().scene().removeItem(self.rubberbandLine)
            self.rubberbandLine = None

            if (self.layerSubareas != None and self.layerSubareas.selectedFeatureCount() > 0):
                self.layerSubareas.removeSelection()
            if (self.layerBuildings != None and self.layerBuildings.selectedFeatureCount() > 0):
                self.layerBuildings.removeSelection()
            if (self.layerOpenSurfacesLines != None and self.layerOpenSurfacesLines.selectedFeatureCount() > 0):
                self.layerOpenSurfacesLines.removeSelection()
            if (self.layerOpenSurfacesAreas != None and self.layerOpenSurfacesAreas.selectedFeatureCount() > 0):
                self.layerOpenSurfacesAreas.removeSelection()

        except Exception as e:
            #utils.Error("Fejl i __cleanup: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def canvasPressEvent(self, e):

        try:
            # if e.button() != Qt.LeftButton:
            #     return
            ActivePoint = self.toMapCoordinates(e.pos())
            if self.ActiveTool.find("Subarea") > -1:
                if self.layerSubareas == None:
                    self.layerSubareas = self.dlgMain.wInputLayerSubareas.currentLayer()
                if self.layerSubareas == None:
                    utils.Error(
                        "Vælg et lag som indeholder delområder (under indstillinger)"
                    )
                    return
                if not self.layerSubareas.isEditable():
                    self.layerSubareas.startEditing()
                if self.ActiveTool == "CreateSubarea":
                    if e.button() == Qt.RightButton:
                        # Højreklik: så gemmes polygonen
                        oTable = self.dlgMain.tblSubareas
                        feat = QgsFeature(self.layerSubareas.fields())
                        feat.setGeometry(QgsGeometry.fromPolygonXY([self.polyline]))
                        self.layerSubareas.addFeatures(
                            [feat], QgsFeatureSink.FastInsert
                        )
                        # Den nyoprettede feature genfindes
                        newfId = 0
                        for general_feat in self.layerSubareas.getFeatures():
                            if str(feat.geometry()) == str(general_feat.geometry()):
                                newfId = general_feat.id()
                                break
                        oId = QTableWidgetItem(str(newfId))
                        oId.setFlags(Qt.NoItemFlags)  # not enabled
                        oTable.setItem(oTable.currentRow(), 0, oId)
                        self.saveRowSubareas(oTable.currentRow())
                        self.InsertBlankRowSubareas(oTable)
                        #self.FillTableSurfaces(newfId)
                        oTable.selectRow(oTable.verticalHeader().count() - 2)
                        self.selectRowSubareas(oTable.verticalHeader().count() - 2)
                        self.ActiveTool = ""
                    else:
                        self.polyline.append(ActivePoint)
                elif self.ActiveTool == "EditSubarea":
                    if self.layerSubareas.selectedFeatureCount() < 1:
                        utils.Error("Ingen delområder valgt")
                        return
                    feat = self.layerSubareas.selectedFeatures()[0]
                    if e.button() == Qt.RightButton:
                        self.isEditing = False
                        self.polyline = []
                        self.vertex = -1
                    else:
                        if self.isEditing:
                            # Så gemmes rettelsen
                            if self.vertex == 0 or self.vertex == len(self.polyline)-1:
                                self.polyline[0] = ActivePoint
                                self.polyline[len(self.polyline)-1] = ActivePoint
                            else:
                                self.polyline[self.vertex] = ActivePoint
                            self.layerSubareas.changeGeometry(
                                feat.id(), QgsGeometry.fromPolygonXY([self.polyline])
                            )
                            self.isEditing = False
                            self.polyline = []
                            self.vertex = -1
                        else:
                            minDist = self.iface.mapCanvas().mapUnitsPerPixel() * 5
                            self.polyline, self.vertex = utils.selectVertex(ActivePoint, feat, minDist)
                            if self.vertex >= 0:
                                self.isEditing = True
            elif self.ActiveTool.find("Building") > -1:
                if self.layerBuildings == None:
                    self.layerBuildings = (
                        self.dlgMain.wInputLayerBuildings.currentLayer()
                    )
                if self.layerBuildings == None:
                    utils.Error(
                        "Vælg et lag som indeholder bygninger (under indstillinger)"
                    )
                    return
                if not self.layerBuildings.isEditable():
                    self.layerBuildings.startEditing()
                if self.ActiveTool == "CreateBuilding":
                    if e.button() == Qt.RightButton:
                        # Højreklik: så gemmes polygonen
                        oTable = self.dlgMain.tblBuildings
                        feat = QgsFeature(self.layerBuildings.fields())
                        feat.setGeometry(QgsGeometry.fromPolygonXY([self.polyline]))
                        self.layerBuildings.addFeatures(
                            [feat], QgsFeatureSink.FastInsert
                        )
                        # Den nyoprettede feature genfindes
                        newfId = 0
                        for general_feat in self.layerBuildings.getFeatures():
                            if str(feat.geometry()) == str(general_feat.geometry()):
                                newfId = general_feat.id()
                                break
                        if newfId == 0:
                            utils.Message('Fejl ved oprettelse af geometri')
                        else:
                            oId = QTableWidgetItem(str(newfId))
                            oId.setFlags(Qt.NoItemFlags)  # not enabled
                            oTable.setItem(oTable.currentRow(), 0, oId)
                            self.saveRowBuildings(oTable.currentRow())
                            self.EnableDisableBuildingParameters(True)
                            self.InsertBlankRowBuildings(oTable)
                            oTable.selectRow(oTable.verticalHeader().count() - 2)
                        self.ActiveTool = ""
                    else:
                        self.polyline.append(ActivePoint)
                elif self.ActiveTool == "EditBuilding":
                    if self.layerBuildings.selectedFeatureCount() < 1:
                        utils.Error("Ingen bygning valgt")
                        return
                    feat = self.layerBuildings.selectedFeatures()[0]
                    if e.button() == Qt.RightButton:
                        self.isEditing = False
                        self.polyline = []
                        self.vertex = -1
                    else:
                        if self.isEditing:
                            # Så gemmes rettelsen
                            if self.vertex == 0 or self.vertex == len(self.polyline)-1:
                                self.polyline[0] = ActivePoint
                                self.polyline[len(self.polyline)-1] = ActivePoint
                            else:
                                self.polyline[self.vertex] = ActivePoint
                            self.layerBuildings.changeGeometry(
                                feat.id(), QgsGeometry.fromPolygonXY([self.polyline])
                            )
                            self.isEditing = False
                            self.polyline = []
                            self.vertex = -1
                        else:
                            minDist = self.iface.mapCanvas().mapUnitsPerPixel() * 5
                            self.polyline, self.vertex = utils.selectVertex(ActivePoint, feat, minDist)
                            if self.vertex >= 0:
                                self.isEditing = True
            elif self.ActiveTool.find("SurfaceLine") > -1:
                if self.layerOpenSurfacesLines == None:
                    self.layerOpenSurfacesLines = (
                        self.dlgMain.wInputLayerOpenSurfacesLines.currentLayer()
                    )
                if self.layerOpenSurfacesLines == None:
                    utils.Error(
                        "Vælg et lag som indeholder åbne overflader med linier (under indstillinger)"
                    )
                    return
                if not self.layerOpenSurfacesLines.isEditable():
                    self.layerOpenSurfacesLines.startEditing()
                if self.ActiveTool == "CreateSurfaceLine":
                    if e.button() == Qt.RightButton:
                        # Højreklik: så gemmes linien
                        oTable = self.dlgMain.tblOpenSurfacesLines
                        feat = QgsFeature(self.layerOpenSurfacesLines.fields())
                        feat.setGeometry(QgsGeometry.fromPolylineXY(self.polyline))
                        self.layerOpenSurfacesLines.addFeatures(
                            [feat], QgsFeatureSink.FastInsert
                        )
                        # Den nyoprettede feature genfindes
                        newfId = 0
                        for general_feat in self.layerOpenSurfacesLines.getFeatures():
                            if str(feat.geometry()) == str(general_feat.geometry()):
                                newfId = general_feat.id()
                                break
                        oId = QTableWidgetItem(str(newfId))
                        oId.setFlags(Qt.NoItemFlags)  # not enabled
                        oTable.setItem(oTable.currentRow(), 0, oId)
                        self.saveRowOpenSurfacesLines(oTable.currentRow())
                        self.InsertBlankRowOpenSurfacesLines(oTable)
                        oTable.selectRow(oTable.verticalHeader().count() - 2)
                        self.ActiveTool = ""
                    else:
                        self.polyline.append(ActivePoint)
                elif self.ActiveTool == "EditSurfaceLine":
                    if self.layerOpenSurfacesLines.selectedFeatureCount() < 1:
                        utils.Error("Ingen linie/vej valgt")
                        return
                    feat = self.layerOpenSurfacesLines.selectedFeatures()[0]
                    if e.button() == Qt.RightButton:
                        self.isEditing = False
                        self.polyline = []
                        self.vertex = -1
                    else:
                        if self.isEditing:
                            # Så gemmes rettelsen
                            self.polyline[self.vertex] = ActivePoint
                            self.layerOpenSurfacesLines.changeGeometry(
                                feat.id(), QgsGeometry.fromPolylineXY(self.polyline)
                            )
                            self.isEditing = False
                            self.polyline = []
                            self.vertex = -1
                        else:
                            minDist = self.iface.mapCanvas().mapUnitsPerPixel() * 5
                            self.polyline, self.vertex = utils.selectVertex(ActivePoint, feat, minDist)
                            if self.vertex >= 0:
                                self.isEditing = True
            elif self.ActiveTool.find("SurfaceArea") > -1:
                if self.layerOpenSurfacesAreas == None:
                    self.layerOpenSurfacesAreas = (
                        self.dlgMain.wInputLayerOpenSurfacesAreas.currentLayer()
                    )
                if self.layerOpenSurfacesAreas == None:
                    utils.Error(
                        "Vælg et lag som indeholder åbne overflader med polygoner (under indstillinger)"
                    )
                    return
                if not self.layerOpenSurfacesAreas.isEditable():
                    self.layerOpenSurfacesAreas.startEditing()
                if self.ActiveTool == "CreateSurfaceArea":
                    if e.button() == Qt.RightButton:
                        # Højreklik: så gemmes polygonen
                        oTable = self.dlgMain.tblOpenSurfacesAreas
                        feat = QgsFeature(self.layerOpenSurfacesAreas.fields())
                        feat.setGeometry(QgsGeometry.fromPolygonXY([self.polyline]))
                        self.layerOpenSurfacesAreas.addFeatures(
                            [feat], QgsFeatureSink.FastInsert
                        )
                        # Den nyoprettede feature genfindes
                        newfId = 0
                        for general_feat in self.layerOpenSurfacesAreas.getFeatures():
                            if str(feat.geometry()) == str(general_feat.geometry()):
                                newfId = general_feat.id()
                                break
                        oId = QTableWidgetItem(str(newfId))
                        oId.setFlags(Qt.NoItemFlags)  # not enabled
                        oTable.setItem(oTable.currentRow(), 0, oId)
                        self.saveRowOpenSurfacesAreas(oTable.currentRow())
                        self.InsertBlankRowOpenSurfacesAreas(oTable)
                        oTable.selectRow(oTable.verticalHeader().count() - 2)
                        self.ActiveTool = ""
                    else:
                        self.polyline.append(ActivePoint)
                elif self.ActiveTool == "EditSurfaceArea":
                    if self.layerOpenSurfacesAreas.selectedFeatureCount() < 1:
                        utils.Error("Ingen polygon valgt")
                        return
                    feat = self.layerOpenSurfacesAreas.selectedFeatures()[0]
                    if e.button() == Qt.RightButton:
                        self.isEditing = False
                        self.polyline = []
                        self.vertex = -1
                    else:
                        if self.isEditing:
                            # Så gemmes rettelsen
                            if self.vertex == 0 or self.vertex == len(self.polyline)-1:
                                self.polyline[0] = ActivePoint
                                self.polyline[len(self.polyline)-1] = ActivePoint
                            else:
                                self.polyline[self.vertex] = ActivePoint
                            self.layerOpenSurfacesAreas.changeGeometry(
                                feat.id(), QgsGeometry.fromPolygonXY([self.polyline])
                            )
                            self.isEditing = False
                            self.polyline = []
                            self.vertex = -1
                        else:
                            minDist = self.iface.mapCanvas().mapUnitsPerPixel() * 5
                            self.polyline, self.vertex = utils.selectVertex(ActivePoint, feat, minDist)
                            if self.vertex >= 0:
                                self.isEditing = True

        except Exception as e:
            utils.Error("Fejl i canvasPressEvent: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def canvasMoveEvent(self, e):

        try:
            if self.rubberband == None and self.rubberbandLine == None:
                return
            self.rubberband.reset(QgsWkbTypes.PolygonGeometry)
            self.rubberbandLine.reset(QgsWkbTypes.LineGeometry)
            ActivePoint = self.toMapCoordinates(e.pos())
            if self.ActiveTool.find("Subarea") > -1 or self.ActiveTool.find("Building") > -1 or self.ActiveTool.find("SurfaceArea") > -1:
                if self.ActiveTool[0:6] == "Create":
                    if self.polyline != []:
                        for pnt in self.polyline:
                            self.rubberband.addPoint(pnt, False)
                        self.rubberband.addPoint(ActivePoint, True)
                elif self.ActiveTool[0:4] == "Edit":
                    if self.polyline != [] and self.vertex >= 0:
                        if self.vertex == 0:
                            self.rubberband.addPoint(self.polyline[-2], False)
                            self.rubberband.addPoint(self.toMapCoordinates(e.pos()), False)
                            self.rubberband.addPoint(self.polyline[1], True)
                        elif self.vertex == len(self.polyline) -1:
                            self.rubberband.addPoint(self.polyline[-1], False)
                            self.rubberband.addPoint(self.toMapCoordinates(e.pos()), False)
                            self.rubberband.addPoint(self.polyline[2], True)
                        else:
                            self.rubberband.addPoint(self.polyline[self.vertex - 1], False)
                            self.rubberband.addPoint(self.toMapCoordinates(e.pos()), False)
                            self.rubberband.addPoint(self.polyline[self.vertex + 1], True)
            elif self.ActiveTool.find("SurfaceLine") > -1:
                if self.ActiveTool == "CreateSurfaceLine":
                    if self.polyline != []:
                        for pnt in self.polyline:
                            self.rubberbandLine.addPoint(pnt, False)
                        self.rubberbandLine.addPoint(ActivePoint, True)
                elif self.ActiveTool == "EditSurfaceLine":
                    if self.polyline != []:
                        if self.vertex == 0:
                            self.rubberbandLine.addPoint(self.polyline[1], False)
                            self.rubberbandLine.addPoint(self.toMapCoordinates(e.pos()), True)
                        elif self.vertex == len(self.polyline) - 1:
                            self.rubberbandLine.addPoint(self.polyline[len(self.polyline) - 2], False)
                            self.rubberbandLine.addPoint(self.toMapCoordinates(e.pos()), True)
                        elif self.vertex > 0:
                            self.rubberbandLine.addPoint(self.polyline[self.vertex - 1], False)
                            self.rubberbandLine.addPoint(self.toMapCoordinates(e.pos()), False)
                            self.rubberbandLine.addPoint(self.polyline[self.vertex + 1], True)

        except Exception as e:
            utils.Error("Fejl i canvasMoveEvent: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def canvasReleaseEvent(self, e):

        try:
            if e.button() == Qt.RightButton:
                # savetool = ""
                # if self.isEditing:
                #     savetool = self.ActiveTool
                savetool = self.ActiveTool
                self.resetGeometryEdit()
                if savetool != "":
                    self.ActiveTool = savetool

        except Exception as e:
            utils.Error("Fejl i canvasReleaseEvent: " + str(e))
            return

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def resetGeometryEdit(self):

        try:
            # Reset map tools
            self.ActiveTool = ""
            self.polyline = []
            self.vertex = -1
            self.isEditing = False
            if self.rubberband == None:
                #Andre QGIS værktøjer kan have været brugt.
                self.rubberband = QgsRubberBand(self.iface.mapCanvas(), QgsWkbTypes.PolygonGeometry)
                self.rubberband.setColor(QColor(127, 127, 255, 127))
                self.rubberbandLine = QgsRubberBand(self.iface.mapCanvas(), QgsWkbTypes.LineGeometry)
                self.rubberbandLine.setColor(QColor(0, 0, 255, 255))
                self.iface.mapCanvas().setMapTool(self)
            else:            
                self.rubberband.reset()
                self.rubberbandLine.reset()

        except Exception as e:
            utils.Error("Fejl i resetGeometryEdit: " + str(e))
            return


# ---------------------------------------------------------------------------------------------------------
#
# ---------------------------------------------------------------------------------------------------------
