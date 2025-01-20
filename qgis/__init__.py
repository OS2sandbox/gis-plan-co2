# -*- coding: utf-8 -*-
"""
/***************************************************************************
 PlanCO2
                                 A QGIS plugin
 PlanCO2
                             -------------------
        begin                : 2024-02-28
        copyright            : (C) 2024 by Rambøll
        email                : jcn@ramboll.dk
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load UrbanDecarb class from file UrbanDecarb.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .urban_decarb import PlanCO2
    return PlanCO2(iface)
