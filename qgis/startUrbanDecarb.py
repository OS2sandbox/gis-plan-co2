# -*- coding: utf-8 -*-
"""
/***************************************************************************
 startup script - eksekveres først efter urban_decarb.py
 ***************************************************************************/
"""

import os
os.putenv('UrbanDecarbStarted', '22')

from qgis.PyQt.QtCore import QSettings
s = QSettings()
s.setValue("qgis/projOpenAtLaunch", 3)
from qgis.PyQt.QtWidgets import QMessageBox
QMessageBox.information(None, u'besked',  'startUrbanDecarb.py')

from qgis.utils import iface
iface.newProject()
