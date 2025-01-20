# -*- coding: utf-8 -*-
"""
/***************************************************************************
 About
 ***************************************************************************/
"""

import os

from qgis.PyQt import uic
from qgis.PyQt.QtWidgets import QDialog 
from qgis.PyQt.QtGui import QPixmap

from qgis.core import QgsProject

import configparser
from ..tools import utils

FORM_CLASS, _ = uic.loadUiType(os.path.join(os.path.dirname(__file__), 'frmAbout.ui'))

#---------------------------------------------------------------------------------------------------------   
# 
#---------------------------------------------------------------------------------------------------------   
class AboutDialog(QDialog, FORM_CLASS):
    #---------------------------------------------------------------------------------------------------------   
    # 
    #---------------------------------------------------------------------------------------------------------   
    def __init__(self, parent=None):
        super(AboutDialog, self).__init__(parent)
        self.setupUi(self)

    #---------------------------------------------------------------------------------------------------------   
    # 
    #---------------------------------------------------------------------------------------------------------   
    def initGui(self):

        try:
                
            self.setWindowTitle('Om Urban Decarb')
            config = configparser.ConfigParser()
            config.read(os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')), 'metadata.txt'))
            version = config.get('general', 'version')
            versiondate = config.get('general', 'versiondate')
            qgisMinVersion = config.get('general', 'qgisMinimumVersion')
            qgisMaxVersion = config.get('general', 'qgisMaximumVersion')

            sStyle0 = '<style type="text/css">a:link {color:#00B0F0; text-decoration: none; }a:visited { color:#00B0F0; text-decoration: none; }a:hover { color:#00B0F0; text-decoration: none; }a:active { color:#00B0F0; text-decoration: underline; }</style>'
            sStyle1 = 'style="font-family:Arial Rounded MT Bold;font-size:14pt;color:#00B0F0;"'
            sStyle2 = 'style="font-family:Arial Rounded MT Bold;font-size:9pt;color:#00B0F0;"'

            self.lblVersionDate.setText('<html><head/><body><p><span ' + sStyle1 + '>Dato: ' + versiondate + '</span></p></body></html>')

            prjpath = QgsProject.instance().fileName()
            self.lblActiveProject.setText('<html><head/><body><p><span ' + sStyle2 + '>' + prjpath+ '</span></p></body></html>')

            pixmap = QPixmap(os.path.join(os.path.dirname(__file__), 'templates/RambollLogo2.gif')).scaledToWidth(150)
            self.lblLogo.setPixmap(pixmap)

        except Exception as e:
            utils.Error("Fejl i initGui: " + str(e))
            return 

#---------------------------------------------------------------------------------------------------------   
# 
#---------------------------------------------------------------------------------------------------------   
