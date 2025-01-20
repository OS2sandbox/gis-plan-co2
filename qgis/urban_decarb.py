# -*- coding: utf-8 -*-
"""
/***************************************************************************
 PlanCO2
                                 A QGIS plugin
 PlanCO2
                              -------------------
        begin                : 2024-02-29
        git sha              : $Format:%H$
        copyright            : (C) 2024 by Rambøll
        email                : jcn@ramboll.dk
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""

from qgis.PyQt import uic
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtCore import QCoreApplication, QSettings, Qt
from qgis.PyQt.QtWidgets import (
    QMenu,
    QAction,
    QMenuBar,
    QToolBar,
    QDockWidget,
    QToolButton,
    QFileDialog,
)
from qgis.PyQt.QtXml import *

from qgis.core import QgsProject, QgsApplication, QgsFeatureRequest

# Initialize Qt resources from file resources.py
from .tools import utils

from .ui.frmAbout import AboutDialog

import os.path, sys


# ---------------------------------------------------------------------------------------------------------
#
# ---------------------------------------------------------------------------------------------------------
class PlanCO2:
    """QGIS Plugin Implementation."""

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface
        self.canvas = self.iface.mapCanvas()
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        try:
            # initialize locale
            if QSettings().value("locale/userLocale") != None:
                locale = QSettings().value("locale/userLocale")[0:2]
                locale_path = os.path.join(
                    self.plugin_dir, "i18n", "PlanCO2{}.qm".format(locale)
                )
                if os.path.exists(locale_path):
                    self.translator = QTranslator()
                    self.translator.load(locale_path)
                    if qVersion() > "4.3.3":
                        QCoreApplication.installTranslator(self.translator)
        except Exception as e:
            pass

        # Declare instance attributes
        self.actions = []

        # Add toolbar
        self.toolbar = self.iface.addToolBar("Urban Decarb Toolbar")
        self.toolbar.setObjectName("mPlanCO2Toolbar")

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None,
    ):
        # Create the dialog (after translation) and keep reference

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.menu.addAction(action)

        self.actions.append(action)

        return action

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def initGui(self):

        try:
            showMenu = False
            self.menu = None
            
            if showMenu:
                # Create Urban Decarb menu
                self.menu = self.iface.mainWindow().findChild(
                    QMenu, utils.tr("&Urban Decarb")
                )

                if not self.menu:
                    self.menu = QMenu(
                        utils.tr("&Urban Decarb"), self.iface.mainWindow().menuBar()
                    )
                    actions = self.iface.mainWindow().menuBar().actions()
                    lastAction = actions[-1]
                    self.iface.mainWindow().menuBar().insertMenu(lastAction, self.menu)

                # ---------------------------------------------------------------------------------------------------------------
                # Menubar og toolbar opbygges
                # ---------------------------------------------------------------------------------------------------------------
                # Om Urban Decarb
                self.menu.addSeparator()
                self.add_action(
                    icon_path=os.path.join(os.path.dirname(__file__), "icons/about.png"),
                    text=utils.tr("Om PlanCO2"),
                    callback=self.aboutPlanCO2,
                    add_to_menu=True,
                    add_to_toolbar=False,
                    parent=self.iface.mainWindow(),
                )

            # ---------------------------------------------------------------------------------------------------------
            #
            # ---------------------------------------------------------------------------------------------------------
            self.add_action(
                icon_path=os.path.join(os.path.dirname(__file__), "icons/PlanCO2.png"),
                text=utils.tr("Åben PlanCO2"),
                callback=self.ShowMainForm,
                add_to_menu=showMenu,
                add_to_toolbar=True,
                parent=self.iface.mainWindow(),
            )

        except Exception as e:
            utils.Error("Fejl i initGui. Fejlbesked: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def unload(self):

        try:
            """Removes the plugin menu item and icon from QGIS GUI."""
            for action in self.actions:
                self.iface.removeToolBarIcon(action)
            # remove the toolbar
            if self.toolbar:
                del self.toolbar
            if self.menu:
                self.menu.deleteLater()

        except Exception as e:
            utils.Error("Fejl i unload. Fejlbesked: " + str(e))

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------

    def aboutPlanCO2(self):
        self.dlgAbout = AboutDialog(self.iface.mainWindow())
        self.dlgAbout.initGui()
        self.dlgAbout.show()

    # ---------------------------------------------------------------------------------------------------------
    #
    # ---------------------------------------------------------------------------------------------------------
    def ShowMainForm(self):
        from .ui.frmMain import MainDialogTool

        MainDialogTool(self.iface)


# ---------------------------------------------------------------------------------------------------------
#
# ---------------------------------------------------------------------------------------------------------
