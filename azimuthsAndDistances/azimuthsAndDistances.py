# -*- coding: utf-8 -*-
"""
/***************************************************************************
 AzimuthDistanceCalculator
                                 A QGIS plugin
 Calculates azimuths and distances
                              -------------------
        begin                : 2014-09-24
        copyright            : (C) 2014 by Luiz Andrade
        email                : luiz.claudio@dsg.eb.mil.br
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

from builtins import str
from builtins import range

import os

from qgis.PyQt import uic
from qgis.PyQt.QtWidgets import QDialog, QTableWidgetItem, QMessageBox
from qgis.core import QgsWkbTypes, QgsGeometry, QgsFeature

import processing
import math
from decimal import Decimal

FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'ui_azimuthsAndDistances.ui'))

from .memorialGenerator import MemorialGenerator
from ..kappaAndConvergence.calculateKappaAndConvergence import CalculateKappaAndConvergenceDialog


class AzimuthsAndDistancesDialog(QDialog, FORM_CLASS):
    """Class that calculates azimuths and distances among vertexes in a linestring.
    """
    def __init__(self, iface, geometry):
        """Constructor
        :param iface:
        :param geometry:
        """
        QDialog.__init__( self )
        self.setupUi( self )

        self.geom = geometry
        self.iface = iface
        self.points = None
        self.distancesAndAzimuths = None
        self.area = self.geom.area()

        # Connecting SIGNAL/SLOTS for the Output button
        self.calculateButton.clicked.connect(self.fillTable)
        self.clearButton.clicked.connect(self.clearTable)
        self.saveFilesButton.clicked.connect(self.saveFiles)
        self.convergenceButton.clicked.connect(self.calculateConvergence)
        self.spinBox.valueChanged.connect(self.fillTable)

    def calculateConvergence(self):
        """Calculates convergence
        :param:
        :return:
        """
        convergenceCalculator = CalculateKappaAndConvergenceDialog(self.iface)
        (a, b) = convergenceCalculator.getSemiMajorAndSemiMinorAxis()

        currentLayer = self.iface.mapCanvas().currentLayer()
        if currentLayer:
            selectedFeatures = len(currentLayer.selectedFeatures())
            if selectedFeatures == 1:
                selectedFeature = currentLayer.selectedFeatures()[0]
                centroid = selectedFeature.geometry().centroid()
                geoPoint = convergenceCalculator.getGeographicCoordinates(centroid.asPoint().x(), centroid.asPoint().y())
                self.centralMeridian = convergenceCalculator.getCentralMeridian(geoPoint.x())
                convergence = convergenceCalculator.calculateConvergence2(geoPoint.x(), geoPoint.y(), a, b)
                self.lineEdit.setText(str(convergence))

    def setClockWiseRotation(self, points):
        """set clockwise Rotation
        :param points:
        :return: points
        """
        sum = 0
        for i in range(len(points) - 1):
            sum += (points[i+1].x() - points[i].x())*(points[i+1].y() + points[i].y())

        if sum > 0:
            return points
        else:
            return points[::-1]

    def setFirstPointToNorth(self, coords, yMax):
        """set first point to north
        :param coords:
        :param yMax:
        :return: vector sum
        """
        if coords[0].y() == yMax:
            return coords

        coords.pop()
        firstPart = []
        for i in range(len(coords)):
            firstPart.append(coords[i])
            if coords[i].y() == yMax:
                break

        return coords[i:] + firstPart

    def saveFiles(self):
        """save files
        :param:
        :return:
        """
        if (not self.distancesAndAzimuths) or (not self.points):
            QMessageBox.information(self.iface.mainWindow(), self.tr("Warning!"), self.tr("Click on calculate button first to generate the needed data."))
        else:
            confrontingList = list()
            for i in range(self.tableWidget.rowCount()):
                item = self.tableWidget.item(i, 7)
                confrontingList.append(item.text())

            d = MemorialGenerator(self.iface.mapCanvas().currentLayer().crs().description(), self.centralMeridian, self.lineEdit.text(), self.tableWidget, self.area, self.perimeter)
            d.getConfigurationMemorial()
            d.exec_()
            #d.getConfigurationMemorial()

    def isValidType(self):
        """Verifies the geometry type.
        :param:
        :return:
        """
        if self.geom.isMultipart():
            #QMessageBox.information(self.iface.mainWindow(), self.tr("Warning!"), self.tr("The limit of a patrimonial area must be a single part geometry."))
            #return False
            self.points =  self.azimuthPoints()
            print("self.points", self.points)
            polygon = QgsGeometry.fromPolygonXY([self.points])
            self.geom = polygon
            print("Oia", polygon)

        #if self.geom.type() == QGis.Line:
        if self.geom.type() == QgsWkbTypes.LineGeometry:
            self.points = self.geom.asPolyline()
            if self.points[0].y() < self.points[-1].y():
                self.points = self.points[::-1]
            return True
        elif self.geom.type() == QgsWkbTypes.PolygonGeometry:
            points = self.setClockWiseRotation(self.geom.asPolygon()[0])
            yMax = self.geom.boundingBox().yMaximum()
            self.points = self.setFirstPointToNorth(points, yMax)
            return True
        else:
            QMessageBox.information(self.iface.mainWindow(), self.tr("Warning!"), self.tr("The selected geometry should be a Line or a Polygon."))
            return False


    def calculate(self):
        """Calculates and Constructs a list with distances and azimuths.
        :param:
        :return: list with distances and azimuths
        """
        points = self.points
        self.perimeter = 0
        self.distancesAndAzimuths = list()

        for i in range(0, len(points)-1):
            before = points[i]
            after = points[i+1]
            distance = math.sqrt(before.sqrDist(after))
            azimuth = before.azimuth(after)
            if azimuth < 0:
                azimuth += 360
            self.distancesAndAzimuths.append((distance, azimuth))
            self.perimeter += distance

        return self.distancesAndAzimuths


    def fillTable(self):
        """fills and makes the CSV.
        :param:
        :return:
        """
        decimalPlaces = self.spinBox.value()
        q = Decimal(10)**-decimalPlaces

        distancesAndAzimuths = list()
        isValid = self.isValidType()
        if isValid:
            distancesAndAzimuths = self.calculate()
        try:
            convergence = float(self.lineEdit.text())
        except ValueError:
            QMessageBox.information(self.iface.mainWindow(), self.tr("Warning!"), self.tr("Please, insert the meridian convergence."))
            return

        isClosed = False
        if self.points[0] == self.points[len(self.points) - 1]:
            isClosed = True


        if self.tableWidget.rowCount() == 0:
            self.tableWidget.setRowCount(len(distancesAndAzimuths))

            for i in range(0,len(distancesAndAzimuths)):
                azimuth = self.dd2dms(distancesAndAzimuths[i][1])
                realAzimuth = self.dd2dms(distancesAndAzimuths[i][1] + convergence)

                itemVertex = QTableWidgetItem("P"+str(i))
                self.tableWidget.setItem(i, 0, itemVertex)
                e = Decimal(self.points[i].x()).quantize(q)
                itemE = QTableWidgetItem(str(e))
                self.tableWidget.setItem(i, 1, itemE)
                n = Decimal(self.points[i].y()).quantize(q)
                itemN = QTableWidgetItem(str(n))
                self.tableWidget.setItem(i, 2, itemN)

                if (i == len(distancesAndAzimuths) - 1) and isClosed:
                    itemSide = QTableWidgetItem("P"+str(i)+"-P0")
                    self.tableWidget.setItem(i, 3, itemSide)
                else:
                    itemSide = QTableWidgetItem("P"+str(i)+"-P"+str(i+1))
                    self.tableWidget.setItem(i, 3, itemSide)

                itemAz = QTableWidgetItem(azimuth)
                self.tableWidget.setItem(i, 4, itemAz)
                itemRealAz = QTableWidgetItem(realAzimuth)
                self.tableWidget.setItem(i, 5, itemRealAz)
                dist = "%0.2f"%(distancesAndAzimuths[i][0])
                itemDistance = QTableWidgetItem(dist)
                self.tableWidget.setItem(i, 6, itemDistance)
                itemConfronting = QTableWidgetItem("")
                self.tableWidget.setItem(i, 7, itemConfronting)
        else:

            for i in range(0,len(distancesAndAzimuths)):
                azimuth = self.dd2dms(distancesAndAzimuths[i][1])
                realAzimuth = self.dd2dms(distancesAndAzimuths[i][1] + convergence)

                itemVertex = QTableWidgetItem("P"+str(i))
                self.tableWidget.setItem(i, 0, itemVertex)
                e = Decimal(self.points[i].x()).quantize(q)
                itemE = QTableWidgetItem(str(e))
                self.tableWidget.setItem(i, 1, itemE)
                n = Decimal(self.points[i].y()).quantize(q)
                itemN = QTableWidgetItem(str(n))
                self.tableWidget.setItem(i, 2, itemN)

                if (i == len(distancesAndAzimuths) - 1) and isClosed:
                    itemSide = QTableWidgetItem("P"+str(i)+"-P0")
                    self.tableWidget.setItem(i, 3, itemSide)
                else:
                    itemSide = QTableWidgetItem("P"+str(i)+"-P"+str(i+1))
                    self.tableWidget.setItem(i, 3, itemSide)

                itemAz = QTableWidgetItem(azimuth)
                self.tableWidget.setItem(i, 4, itemAz)
                itemRealAz = QTableWidgetItem(realAzimuth)
                self.tableWidget.setItem(i, 5, itemRealAz)
                dist = "%0.2f"%(distancesAndAzimuths[i][0])
                itemDistance = QTableWidgetItem(dist)
                self.tableWidget.setItem(i, 6, itemDistance)
                itemConfronting = QTableWidgetItem("")

                # if str(self.tableWidget.item(i, 7).text()) != "":
                #     print (str(self.tableWidget.item(i, 7).text()))
                #     self.tableWidget.setItem(i, 7, itemConfronting)


    def azimuthPoints(self):
        """Colecta os pontos de uma geoemtria na ordem dos seus elemetos
        :param:
        :return: ponits list
        """
        listPoint = []
        geom = self.geom
        listMultiPolygon = geom.asMultiPolygon()
        for multiPolygon in listMultiPolygon:
            for polygon in multiPolygon:
                for ponto in polygon:
                    listPoint.append(ponto)

        return listPoint


    def clearTable(self):
        """clear table
        :param
        :return:
        """
        self.tableWidget.setRowCount(0)

    def dd2dms(self, dd):
        """
        :param dd:
        :return: cnoncated string of degree and time
        """
        is_positive = dd >= 0
        dd = abs(dd)
        minutes, seconds = divmod(dd*3600,60)
        degrees, minutes = divmod(minutes,60)

        degrees = str(int(degrees)) if is_positive else '-' + str(int(degrees))
        minutes = int(minutes)

        return degrees + u"\u00b0" + str(minutes).zfill(2) + "'" + "%0.2f"%(seconds) + "''"
