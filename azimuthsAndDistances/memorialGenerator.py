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
import shutil
import os
import time
import sys
import locale

from builtins import str
from builtins import range

from qgis.PyQt import uic
from qgis.PyQt.QtCore import QFile, QIODevice
from qgis.PyQt.QtWidgets import QFileDialog, QMessageBox, QDialog
from qgis.PyQt.QtXml import QDomDocument

from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table as TablePDF, TableStyle
from datetime import date
from reportlab.lib.units import mm

from importlib import reload


from odf.opendocument import OpenDocumentText
from odf.draw import Frame,Image as image_odf,Page
from odf.style import Style, TextProperties, ParagraphProperties, PageLayoutProperties, PageLayout, MasterPage, GraphicProperties, TableColumnProperties
from odf.text import H, P, Span
from odf.table import Table, TableColumn, TableRow, TableCell

FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'ui_memorialGenerator.ui'))

reload(sys)
#sys.setdefaultencoding('utf-8')

class MemorialGenerator(QDialog, FORM_CLASS):

    def __init__(self, crsDescription, centralMeridian, convergence, tableWidget, geomArea, geomPerimeter):
        """Constructor.
        """
        QDialog.__init__(self)
        self.setupUi( self )

        # Connecting SIGNAL/SLOTS for the Output button
        self.folderButton.clicked.connect(self.setDirectory)

        # Connecting SIGNAL/SLOTS for the Output button
        self.createButton.clicked.connect(self.createFiles)
        self.closeButton.clicked.connect(self.closeWindows)

        self.convergenciaEdit.setText(convergence)

        self.tableWidget = tableWidget
        self.geomArea = geomArea
        self.geomPerimeter = geomPerimeter

        #print crsDescription
        self.meridianoEdit.setText(str(centralMeridian))
        self.projectionEdit.setText(crsDescription.split('/')[-1])
        self.datumEdit.setText(crsDescription.split('/')[0])

        self.FULL_MONTHS = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro', 'outubro','novembro','dezembro']
        self.pathlogo = os.path.dirname(__file__)
        self.logo = os.path.join(self.pathlogo, 'templates/template_memorial_pdf/rep_of_brazil2.jpg')
        self.formattedTime = date.today().timetuple()
        self.subTitle = 'Memorial Descritivo'

    def setDirectory(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.folderEdit.setText(folder)

    def closeWindows(self):
        self.close()

    def copyAndRenameFiles(self):
        currentPath = os.path.dirname(__file__)
        templatePath = os.path.join(currentPath, "templates")
        simpleMemorialTemplate = os.path.join(templatePath, "template_sintetico.html")
        areaTemplate = os.path.join(templatePath, "template_area.csv")
        nameImovel= self.imovelEdit.text()
        prevNameFile = nameImovel.replace(" ", "_")
        folder=''
        folder = self.folderEdit.text()

        if folder =='':
            QMessageBox.information(self, self.tr('Attention!'), self.tr('A directory should be selected!'))
        else:

            if self.memorialSinteticHtml.isChecked():
                self.simpleMemorial = os.path.join(folder, prevNameFile + "_sintetico.html")
                shutil.copy2(simpleMemorialTemplate, self.simpleMemorial)

            if self.tableAreaCsv.isChecked():
                self.area = os.path.join(folder, prevNameFile + "_pontos.csv")
                shutil.copy2(areaTemplate, self.area)

            if self.memorialDescritivoPdf.isChecked():
                #print folder
                self.fullMemorialPdf = os.path.join(folder, prevNameFile + "_memorial.pdf")

            if self.memorialDescritivoOdt.isChecked():
                self.fullMemorialOdt = os.path.join(folder, prevNameFile + "_memorial.odt")

    def createFiles(self):
        #dados do orgão expeditor
        self.title = self.OrgaoExpeditorEdit.text()
        self.title2= self.secretariaEdit.text()
        self.superinte = self.superintenciaEdit.text()
        self.division = self.divisaoEdit.text()
        self.adresstitle = self.enderecoOrgaoEdit.text()
        self.numberControl = self.numMemorialEdit.text()
        self.numberSei = self.numeroSeiEdit.text()

        #dados do imovel
        self.denominationAreaImovel = self.imovelEdit.text()
        self.proprietarioImovel = self.proprietarioEdit.text()
        self.adressImovel = self.enderecoEdit.text()
        self.cityImovel = self.municipioEdit.text()
        self.ufImovel = self.ufEdit.text()
        self.comarca = self.comarcaEdit.text()
        self.matricula = self.matriculaEdit.text()
        self.ripImovel = self.ripEdit.text()
        self.nbpImovel = self.nbpEdit.text()
        self.codeIncra = self.codIncraEdit.text()
        self.plaintCor = self.plantaCorrespondenteEdit.text()
        self.kappa = float(self.kappaEdit.text())
        self.geomPerimeter = self.geomPerimeter/self.kappa
        self.geomArea = self.geomArea/(self.kappa*self.kappa)
        self.projection = self.projectionEdit.text()
        self.meridianCenter=self.meridianoEdit.text()
        self.datum = self.datumEdit.text()
        self.perimeter = "%0.2f"%(self.geomPerimeter)
        self.areaMetroQuad = "%0.2f"%(self.geomArea)

        #dados do Responsavel tecnico
        self.responsibletecName = self.autorEdit.text()
        self.officeResponsible = self.officeResponsibleEdit.text()
        self.addressBrCityDoc= self.mucipioResponsavelEdit.text()
        self.tipeIdResponsible = self.creaCau.currentText()
        self.identification = self.creaEdit.text()

        self.copyAndRenameFiles()
        try:
            if self.memorialSinteticHtml.isChecked():
                self.createSimpleMemorial()

            if self.tableAreaCsv.isChecked():
                self.createArea()

            if self.memorialDescritivoPdf.isChecked():
                self.createFullMemorialPdf()

            if self.memorialDescritivoOdt.isChecked():
                self.createFullMemorialOdt()

            if self.memorialSinteticHtml.isChecked() == self.tableAreaCsv.isChecked() == self.memorialDescritivoPdf.isChecked() == self.memorialDescritivoOdt.isChecked()==0:
                QMessageBox.information(self, self.tr('Attention!'), self.tr('Select at least one file type!'))
            else:
                QMessageBox.information(self, self.tr('Information!'), self.tr('Files created successfully!'))

        except IOError as e:
            QMessageBox.information(self, self.tr('ERROR!'), self.tr("You must be trying to modify or replace an existing file that is being used by another program."))

    def createCellElement(self, tempDoc, text, colspan, rowspan):
        td = tempDoc.createElement("td")
        p = tempDoc.createElement("p")
        span = tempDoc.createElement("span")

        if colspan > 0:
            td.setAttribute("colspan", colspan)
        if rowspan > 0:
            td.setAttribute("rowspan", rowspan)
        td.setAttribute("style", "border-color : #000000 #000000 #000000 #000000; border-style: solid;")
        p.setAttribute("style", " text-align: center; text-indent: 0px; padding: 0px 0px 0px 0px; margin: 0px 0px 0px 0px;")
        span.setAttribute("style", " font-size: 10pt; font-family: 'Times new roman', 'Helvetica', sans-serif; font-style: normal; font-weight: normal; color: #000000; background-color: transparent; text-decoration: none;")

        textElement = tempDoc.createTextNode(text)

        span.appendChild(textElement)
        p.appendChild(span)
        td.appendChild(p)

        return td

    def createSimpleMemorial(self):
        tempDoc = QDomDocument()
        simple = QFile(self.simpleMemorial)
        simple.open(QIODevice.ReadOnly)
        loaded = tempDoc.setContent(simple)
        simple.close()

        element = tempDoc.documentElement()

        nodes = element.elementsByTagName("table")

        table = nodes.item(0).toElement()

        tr = tempDoc.createElement("tr")
        tr.appendChild(self.createCellElement(tempDoc, u"MEMORIAL DESCRITIVO SINTÉTICO", 7, 0))
        table.appendChild(tr)

        tr = tempDoc.createElement("tr")

        tr.appendChild(self.createCellElement(tempDoc, u"VÉRTICE", 0, 2))
        tr.appendChild(self.createCellElement(tempDoc, "COORDENADAS", 2, 0))
        tr.appendChild(self.createCellElement(tempDoc, "LADO", 0, 2))
        tr.appendChild(self.createCellElement(tempDoc, "AZIMUTES", 2, 0))
        tr.appendChild(self.createCellElement(tempDoc, u"DISTÂNCIA", 0, 0))

        table.appendChild(tr)

        tr = tempDoc.createElement("tr")
        tr.appendChild(self.createCellElement(tempDoc, "E", 0, 0))
        tr.appendChild(self.createCellElement(tempDoc, "N", 0, 0))
        tr.appendChild(self.createCellElement(tempDoc, "PLANO", 0, 0))
        tr.appendChild(self.createCellElement(tempDoc, "REAL", 0, 0))
        tr.appendChild(self.createCellElement(tempDoc, "(m)", 0, 0))
        table.appendChild(tr)

        convergence = float(self.convergenciaEdit.text())

        rowCount = self.tableWidget.rowCount()

        for i in range(0,rowCount):
            lineElement = tempDoc.createElement("tr")

            lineElement.appendChild(self.createCellElement(tempDoc, self.tableWidget.item(i,0).text(), 0, 0))

            lineElement.appendChild(self.createCellElement(tempDoc, self.tableWidget.item(i,1).text(), 0, 0))
            lineElement.appendChild(self.createCellElement(tempDoc, self.tableWidget.item(i,2).text(), 0, 0))

            lineElement.appendChild(self.createCellElement(tempDoc, self.tableWidget.item(i,3).text(), 0, 0))

            lineElement.appendChild(self.createCellElement(tempDoc, self.tableWidget.item(i,4).text(), 0, 0))
            lineElement.appendChild(self.createCellElement(tempDoc, self.tableWidget.item(i,5).text(), 0, 0))
            lineElement.appendChild(self.createCellElement(tempDoc, self.tableWidget.item(i,6).text(), 0, 0))

            table.appendChild(lineElement)

        simple = open(self.simpleMemorial, "w")
        simple.write(tempDoc.toString())
        simple.close()

    def createArea(self):
        newData = ""
        # "Estação    Vante    Coordenada E    Coordenada N    Az Plano    Az Real    Distância\n"

        rowCount = self.tableWidget.rowCount()

        for i in range(0,rowCount):
            line  = str()
            side = self.tableWidget.item(i,3).text()
            sideSplit = side.split("-")
            line += '"'+ sideSplit[0]+'";"'+sideSplit[1]+'";'
            line += self.tableWidget.item(i,1).text()+";"
            line += self.tableWidget.item(i,2).text()+";"
            line += self.tableWidget.item(i,4).text()+";"
            line += self.tableWidget.item(i,5).text()+";"
            line += self.tableWidget.item(i,6).text()+"\n"

            newData += line

        fileCsv = open(self.area, "r")
        content = fileCsv.readlines()

        content.append(newData)
        fileCsv = open(self.area, "w")
        fileCsv.writelines(content)
        fileCsv.close()

    def addPageNumber(canvas, doc):
        page_num = canvas.getPageNumber()
        text = "%s" % page_num
        canvas.drawRightString(184*mm, 8*mm, text)

    #guilherme: funcção para criar um memorial descritivo completo
    def createFullMemorialPdf(self):
        #print 1
        doc = SimpleDocTemplate(self.fullMemorialPdf,pagesize=letter,rightMargin=85,leftMargin=85,topMargin=51,bottomMargin=35)

        Story=[]

        im = Image(self.logo, 1*inch, 1*inch)

        Story.append(im)

        styles=getSampleStyleSheet()

        #incerinto titulo
        Story.append(Spacer(1, 11))
        styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER, fontName="Times-Bold"))
        ptext = '<font size=10.5>%s</font>' %self.title.upper()
        Story.append(Paragraph(ptext, styles["Center"]))

        ptext = '<font size=10.5>%s</font>' %self.title2.upper()
        Story.append(Paragraph(ptext, styles["Center"]))

        #incerindo subtitulo
        ptext = '<font size=10>%s</font>' %self.superinte
        Story.append(Paragraph(ptext, styles["Center"]))

        ptext = '<font size=10>%s</font>' %self.division
        Story.append(Paragraph(ptext, styles["Center"]))

        styles.add(ParagraphStyle(name='adress', alignment=TA_CENTER, fontName="Times-Roman"))
        ptext = '<font size=9>%s</font>' %self.adresstitle
        Story.append(Paragraph(ptext, styles["adress"]))
        Story.append(Spacer(1, 11))

        #incerindo Metadados
        styles.add(ParagraphStyle(name='titlemem', alignment=TA_CENTER, fontName="Times-Bold"))

        if self.numMemorialEdit.text():
            subTitle = self.subTitle +' '+ self.numberControl
        else:
            subTitle = self.subTitle

        ptext = '<font size=12><u>%s</u></font>' % subTitle
        Story.append(Paragraph(ptext, styles["titlemem"]))
        Story.append(Spacer(1, 11))


        data=[["Imovel: " + self.denominationAreaImovel], ["Proprietario: " + self.proprietarioImovel], ["Endereço: " + self.adressImovel]]

        t = TablePDF(data, colWidths=(156*mm), rowHeights=(5*mm, 5*mm, 5*mm))

        t.setStyle(TableStyle([('ALIGN',(0,0),(0,2),'LEFT'),
                               #('BOX',(0,0),(0,2),0,colors.black),
                               #('GRID',(0,0),(0,2),0.5,colors.black),
                               ('FONT',(0,0),(0,2),'Times-Roman',10.5)]))
        Story.append(t)

        data=[["Município/UF: " + self.cityImovel + "/"+ self.ufImovel, "Matrícula: " + self.matricula],
                ["Perímetro(m): " + self.perimeter.replace('.', ','),"NBP: " + self.nbpImovel],
                ["Área(m²): " + self.areaMetroQuad.replace('.', ','), "RIP: " + self.ripImovel ],
                ["Comarca: " + self.comarca, "Código INCRA: " + self.codeIncra]]

        t = TablePDF(data, colWidths=(78* mm, 78*mm), rowHeights=(5*mm, 5*mm, 5*mm, 5*mm ) )
        t.setStyle(TableStyle([('ALIGN',(0,0),(1,3),'LEFT'),
                               #('BOX',(0,0),(1,3),0,colors.black),
                               #('GRID',(0,0),(1,3),0.5,colors.black),
                               ('FONT',(0,0),(1,3),'Times-Roman',10.5)]))

        Story.append(t)

        Story.append(Spacer(1, 11))
        ptext = '<font size=12>%s</font>' %"DESCRIÇÂO"
        Story.append(Paragraph(ptext, styles["Center"]))

        Story.append(Spacer(1, 11))
        Story.append(Spacer(1, 12))

        styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY, fontName="Times-Roman"))

        #add texto do memorial.
        ptext = self.insertDescriptionPDF()

        Story.append(Paragraph(ptext, styles["Justify"]))
        Story.append(Spacer(1, 12))

        #Add data e local
        textdataLocal = self.addressBrCityDoc + ", " + str(self.formattedTime[2]) + " de " + self.FULL_MONTHS[self.formattedTime[1]-1] + " de " + str(self.formattedTime[0])

        styles.add(ParagraphStyle(name='dateLocal', alignment=TA_RIGHT, fontName="Times-Roman"))
        ptext = '<font size=12> %s</font>' % textdataLocal
        Story.append(Paragraph(ptext, styles["dateLocal"]))
        Story.append(Spacer(1, 12))

        #addlocal assinatura
        Story.append(Spacer(1, 11))
        Story.append(Spacer(1, 11))

        styles.add(ParagraphStyle(name='style01', alignment=TA_CENTER, fontName="Times-Roman"))
        ptext = '<font size=11>%s</font>' %self.responsibletecName
        Story.append(Paragraph(ptext, styles["Center"]))

        ptext = '<font size=11>%s</font>' %self.officeResponsible
        Story.append(Paragraph(ptext, styles["style01"]))
        ptext = '<font size=11>%s</font>' % self.tipeIdResponsible + ': ' + self.identification
        Story.append(Paragraph(ptext, styles["style01"]))

        # try:
            #doc.build(Story, onFirstPage=addPageNumber, onLaterPages=addPageNumber)
        doc.build(Story)
        # except IOError as e:
            # QMessageBox.information(self, self.tr('Attention!'), self.tr(str(e)))


    def createFullMemorialOdt(self):

        self.textdoc = OpenDocumentText()
        s = self.textdoc.styles

        pagelayout = PageLayout(name="Mpm1")
        pagelayout.addElement(PageLayoutProperties(marginbottom="-1.25cm", marginright="3cm", marginleft="3cm"))
        self.textdoc.automaticstyles.addElement(pagelayout)

        masterpage = MasterPage(stylename="da", name="Default", pagelayoutname=pagelayout)
        self.textdoc.masterstyles.addElement(masterpage)

        h1style = Style(name="Heading 1", family="paragraph")
        h1style.addElement(TextProperties(attributes={'fontsize':"10.5pt",'fontweight':"bold", 'fontfamily':"Times New Roman"}))
        h1style.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(h1style)

        h1style2 = Style(name="Heading2 1", family="paragraph")
        h1style2.addElement(TextProperties(attributes={'fontsize':"10.5pt",'fontweight':"bold", 'fontfamily':"Times New Roman"}))
        h1style2.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(h1style2)

        h1style2a = Style(name="Heading2a 1", family="paragraph")
        h1style2a.addElement(TextProperties(attributes={'fontsize':"10pt",'fontweight':"bold", 'fontfamily':"Times New Roman"}))
        h1style2a.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(h1style2a)

        addressTitle = Style(name="addressTitle", family="paragraph")
        addressTitle.addElement(TextProperties(attributes={'fontsize':"9pt", 'fontfamily':"Times New Roman"}))
        addressTitle.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(addressTitle)

        h1style3 = Style(name="Heading3 1", family="paragraph")
        h1style3.addElement(TextProperties(attributes={'fontsize':"12pt",'fontweight':"bold", 'fontfamily':"Times New Roman", 'textunderlinewidth':"auto", 'textunderlinestyle':"solid"}))
        h1style3.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(h1style3)

        h1style4 = Style(name="Heading4 1", family="paragraph")
        h1style4.addElement(TextProperties(attributes={'fontsize':"12pt",'fontweight':"bold", 'fontfamily':"Times New Roman"}))
        h1style4.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(h1style4)

        texttable = Style(name="texttable", family="paragraph")
        texttable.addElement(TextProperties(attributes={'fontsize':"10.5pt", 'fontfamily':"Times New Roman"}))
        texttable.addElement(ParagraphProperties(attributes={"textalign":"left"}))
        s.addElement(texttable)

        # An automatic style
        self.bodystyle = Style(name="Body", family="paragraph")
        self.bodystyle.addElement(TextProperties(attributes={'fontsize':"11pt", 'fontfamily':"Times New Roman"}))
        self.bodystyle.addElement(ParagraphProperties(attributes={"textalign":"justify"}))
        s.addElement(self.bodystyle)
        self.textdoc.automaticstyles.addElement(self.bodystyle)

        bodystyle2 = Style(name="Body2", family="paragraph")
        bodystyle2.addElement(TextProperties(attributes={'fontsize':"11pt", 'fontfamily':"Times New Roman"}))
        bodystyle2.addElement(ParagraphProperties(attributes={"textalign":"right"}))
        s.addElement(bodystyle2)
        self.textdoc.automaticstyles.addElement(bodystyle2)

        bodystyle3 = Style(name="Body3", family="paragraph")
        bodystyle3.addElement(TextProperties(attributes={'fontsize':"11pt", 'fontweight':"bold", 'fontfamily':"Times New Roman"}))
        bodystyle3.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(bodystyle3)
        self.textdoc.automaticstyles.addElement(bodystyle3)

        bodystyle4 = Style(name="Body4", family="paragraph")
        bodystyle4.addElement(TextProperties(attributes={'fontsize':"11pt", 'fontfamily':"Times New Roman"}))
        bodystyle4.addElement(ParagraphProperties(attributes={"textalign":"center"}))
        s.addElement(bodystyle4)
        self.textdoc.automaticstyles.addElement(bodystyle4)

        #imgStile
        imgstyle = Style(name="Mfr1", family="graphic")
        imgprop = GraphicProperties(horizontalrel="paragraph", horizontalpos="center", verticalrel="paragraph-content", verticalpos="top")
        imgstyle.addElement(imgprop)
        self.textdoc.automaticstyles.addElement(imgstyle)

        p=P()
        self.textdoc.text.addElement(p)
        href = self.textdoc.addPicture(self.logo)
        imgframe = Frame(name="fig1", anchortype="paragraph", width="2.24cm", height="2.22cm", zindex="0", stylename=imgstyle)
        p.addElement(imgframe)
        img =image_odf(href=href, type="simple", show="embed", actuate="onLoad")

        imgframe.addElement(img)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=h1style, text=self.title.upper())
        self.textdoc.text.addElement(h)
        h=H(outlinelevel=1, stylename=h1style, text=self.title2.upper())
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=h1style2, text=self.superinte)
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=h1style2a, text=self.division)
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=addressTitle, text=self.adresstitle)
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        if self.numMemorialEdit.text():
            subTitle = self.subTitle +' '+ self.numberControl
        else:
            subTitle = self.subTitle

        h=H(outlinelevel=1, stylename=h1style3, text=subTitle)
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        # Create automatic styles for the column widths.
        widewidth = Style(name="co1", family="table-column")
        widewidth.addElement(TableColumnProperties(columnwidth="8cm"))
        self.textdoc.automaticstyles.addElement(widewidth)

        # Start the table, and describe the columns
        table = Table(name="Currency colours")

        table.addElement(TableColumn(stylename=widewidth, defaultcellstylename="co1"))
        tr = TableRow()
        table.addElement(tr)

        # Create a cell with a negative value. It should show as red.
        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text=u"Imóvel: " + str(self.denominationAreaImovel), stylename=texttable))
        tr.addElement(cell)

        tr = TableRow()
        table.addElement(tr)

        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text=u"Proprietário: " + self.proprietarioImovel, stylename=texttable))
        tr.addElement(cell)

        tr = TableRow()
        table.addElement(tr)

        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text=u"Endereço: " + self.adressImovel, stylename=texttable))
        tr.addElement(cell)

        # Create a column (same as <col> in HTML) Make all cells in column default to currency
        table.addElement(TableColumn(numbercolumnsrepeated=0, stylename=widewidth, defaultcellstylename="co1"))
        table.addElement(TableColumn(numbercolumnsrepeated=1, stylename=widewidth, defaultcellstylename="co1"))
        # Create a row (same as <tr> in HTML)
        tr = TableRow()
        table.addElement(tr)
        # Create a cell with a negative value. It should show as red.
        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text=u"Município/UF: " + self.cityImovel + '/' + self.ufImovel, stylename=texttable))
        tr.addElement(cell)

        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text=u"Matrícula: " + self.matricula, stylename=texttable))
        tr.addElement(cell)

        tr = TableRow()
        table.addElement(tr)

        cell = TableCell(valuetype="text", currency="AUD", value="123")
        cell.addElement(P(text=u"Perímetro (m): " + str(self.perimeter).replace('.', ','), stylename=texttable))
        tr.addElement(cell)

        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text="NBP: " + self.nbpImovel, stylename=texttable))
        tr.addElement(cell)

        tr = TableRow()
        table.addElement(tr)

        cell = TableCell(valuetype="text", currency="AUD", value="123")
        cell.addElement(P(text=u"Área (m²): " + str(self.areaMetroQuad).replace('.', ','), stylename=texttable))
        tr.addElement(cell)

        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text=u"RIP: " + self.ripEdit.text(), stylename=texttable))
        tr.addElement(cell)

        tr = TableRow()
        table.addElement(tr)

        cell = TableCell(valuetype="text", currency="AUD")
        cell.addElement(P(text=u"Comarca: " + self.comarca, stylename=texttable))
        tr.addElement(cell)

        cell = TableCell(valuetype="text", currency="AUD", value="123")
        cell.addElement(P(text=u"Código INCRA: " + self.codIncraEdit.text(), stylename=texttable))
        tr.addElement(cell)

        self.textdoc.text.addElement(table)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=h1style4, text='DESCRIÇÃO')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        self.insertDescriptionodt()

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        p = P(text=self.addressBrCityDoc + ", " + str(self.formattedTime[2]) + " de " + self.FULL_MONTHS[self.formattedTime[1]-1] + " de " + str(self.formattedTime[0]), stylename=bodystyle2)
        self.textdoc.text.addElement(p)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        h=H(outlinelevel=1, stylename=self.bodystyle, text='\n')
        self.textdoc.text.addElement(h)

        p = P(text=self.responsibletecName, stylename=bodystyle3)
        self.textdoc.text.addElement(p)

        p = P(text=self.officeResponsible, stylename=bodystyle4)
        self.textdoc.text.addElement(p)

        p = P(text=self.creaCau.currentText() + ": " + self.identification, stylename=bodystyle4)
        self.textdoc.text.addElement(p)

        self.textdoc.save(self.fullMemorialOdt)

    def getDescription(self):
        description = str()
        description += "O imóvel descrito abaixo corresponde a terreno um de" + self.areaMetroQuad + "m², localizado à" + self.adressImovel + ", no município de" + self.cityImovel +"/" + self.ufImovel + "representado na planta" + self.plaintCor + ", processo SEI: " + self.numberSei+ "."
        description += "\n"
        description += "Inicia-se a descrição deste perímetro no vértice "+self.tableWidget.item(0,0).text()+", de coordenadas "
        description += "N "+self.tableWidget.item(0,2).text()+" m e "
        description += "E "+self.tableWidget.item(0,1).text()+" m, "
        description += "Datum " + self.datumEdit.text()+ " com Meridiano Central " +self.meridianoEdit.text()+ ", localizado à "+self.enderecoEdit.text()

        if self.codIncraEdit.text():
            description += ", Código INCRA " + self.codIncraEdit.text()+ "; "
        else:
            description += (";")

        rowCount = self.tableWidget.rowCount()

        for i in range(0,rowCount):
            side = self.tableWidget.item(i,3).text()
            sideSplit = side.split("-")
            if self.tableWidget.item(i,7).text():
                description += " deste, segue confrontando com "+ self.tableWidget.item(i,7).text()+", "
                description += "com os seguintes azimute plano e distância:"
                description += self.tableWidget.item(i,4).text()+" e "
                description += self.tableWidget.item(i,6).text()+"; até o vértice "
            else:
                description += " deste, segue sem confrontante até o vértice "


            if (i == rowCount - 1):
                description += sideSplit[1]+", de coordenadas "
                description += "N "+self.tableWidget.item(0,2).text()+" m e "
                description += "E "+self.tableWidget.item(0,1).text()+" m, encerrando esta descrição."
                description += " Todas as coordenadas aqui descritas estão georrefereciadas ao Sistema Geodésico Brasileiro "

                if self.rbmcOrigemEdit.text():
                    description += ", a partir da estação RBMC de "+self.rbmcOrigemEdit.text()+" de coordenadas "
                    description += "E "+self.rbmcEsteEdit.text()+" m e N "+self.rbmcNorteEdit.text()+" m, "
                    description += "localizada em "+self.localRbmcEdit.text()+", "

                description += "e encontram-se representadas no sistema UTM, referenciadas ao Meridiano Central "+self.meridianoEdit.text()
                description += ", tendo como DATUM "+self.datumEdit.text()+"."
                description += "Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM."
            else:
                description += sideSplit[1]+", de coordenadas "
                description += "N "+self.tableWidget.item(i+1,2).text()+" m e "
                description += "E "+self.tableWidget.item(i+1,1).text()+" m;"

        return description

    def insertDescriptionodt(self):
            #locale.setlocale(locale.LC_ALL, ("pt_BR",""))
            boldstyle = Style(name="Bold", family="text")
            boldprop = TextProperties(fontweight="bold")
            boldstyle.addElement(boldprop)
            self.textdoc.automaticstyles.addElement(boldstyle)


            description = P(stylename=self.bodystyle)
            description.addText("O imóvel descrito abaixo corresponde a um terreno de " + self.areaMetroQuad.replace('.', ',') + " m², localizado à " + self.adressImovel + ", no município de " + self.cityImovel +"/" + self.ufImovel + ", representado na planta " + self.plaintCor + ", processo SEI: " + self.numberSei)

            if self.codIncraEdit.text():
                description.addText(", código INCRA "+ self.codIncraEdit.text()+ ".")
                # self.textdoc.text.addElement(description)
            else:
                description.addText(".")

            self.textdoc.text.addElement(description)

            description = P(stylename=self.bodystyle)
            description.addText("\n\n")
            self.textdoc.text.addElement(description)

            description = P(stylename=self.bodystyle)
            description.addText("\n\n")
            self.textdoc.text.addElement(description)

            description.addText("Inicia-se a descrição deste perímetro no vértice ")
            description.addElement(Span(stylename=boldstyle, text=self.tableWidget.item(0,0).text()))

            description.addText(", de coordenadas ")
            description.addElement(Span(stylename=boldstyle, text="N "+ self.tableWidget.item(0,2).text().replace('.', ',')))

            description.addText(" m e ")
            description.addElement(Span(stylename=boldstyle, text="E " + self.tableWidget.item(0,1).text().replace('.', ',')))

            self.textdoc.text.addElement(description)

            rowCount = self.tableWidget.rowCount()
            itemPrev =''

            for i in range(0,rowCount):
                side = self.tableWidget.item(i,3).text()
                sideSplit = side.split("-")

                if self.tableWidget.item(i,7).text():
                    if self.tableWidget.item(i,7).text() != itemPrev:

                        description.addText("; deste, segue confrontando com ")
                        description.addElement(Span(stylename=boldstyle, text=self.tableWidget.item(i,7).text().upper()))
                        description.addText(", ")
                    else:
                        description.addText("; deste, segue ")
                else:

                    description.addText("; deste, segue ")

                description.addText("com os seguintes azimute plano e distância: ")
                description.addElement(Span(stylename=boldstyle, text=self.tableWidget.item(i,4).text().replace('.', ',')))
                description.addText(" e ")

                description.addElement(Span(stylename=boldstyle, text=self.tableWidget.item(i,6).text().replace('.', ',') + " m"))
                description.addText("; até o vértice ")
                itemPrev = self.tableWidget.item(i,7).text()
                if (i == rowCount - 1):
                    description.addElement(Span(stylename=boldstyle, text=sideSplit[1]))
                    description.addText(", de coordenadas ")

                    description.addElement(Span(stylename=boldstyle, text="N "+ self.tableWidget.item(0,2).text().replace('.', ',')+" m"))
                    description.addText(" e ")

                    description.addElement(Span(stylename=boldstyle, text="E "+self.tableWidget.item(0,1).text().replace('.', ',')+" m"))
                    description.addText(", encerrando esta descrição.")

                    description = P(stylename=self.bodystyle)
                    description.addText("\n\n")
                    self.textdoc.text.addElement(description)

                    description = P(stylename=self.bodystyle)
                    description.addText("\n\n")
                    self.textdoc.text.addElement(description)

                    description.addText(" Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro")

                    if self.rbmcOrigemEdit.text():

                        description.addText(" , a partir da estação RBMC de " + self.rbmcOrigemEdit.text() + " de coordenadas ")
                        description.addElement(Span(stylename=boldstyle, text="E " + self.rbmcEsteEdit.text().replace('.', ',') + " m"))
                        description.addText(" e ")
                        description.addElement(Span(stylename=boldstyle, text="N " + self.rbmcNorteEdit.text().replace('.', ',')+ " m"))
                        description.addText(" , ")
                        description.addText("localizada em " + self.localRbmcEdit.text()+", ")

                    sp = self.projectionEdit.text().split(" ")[3]
                    #print "tai: " + sp
                    description.addText(" e encontram-se representadas no sistema UTM, referenciadas ao Meridiano Central ")
                    description.addElement(Span(stylename=boldstyle, text=self.meridianoEdit.text()))
                    description.addText(", Fuso ")
                    description.addElement(Span(stylename=boldstyle, text=str(sp)))



                    description.addText(", tendo como DATUM ")
                    description.addElement(Span(stylename=boldstyle, text=self.datumEdit.text()))
                    description.addText(". Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM.")
                else:
                    description.addElement(Span(stylename=boldstyle, text=sideSplit[1]))
                    description.addText(", de coordenadas ")

                    description.addElement(Span(stylename=boldstyle, text="N "+ self.tableWidget.item(i+1,2).text().replace('.', ',') + " m"))
                    description.addText(" e ")
                    description.addElement(Span(stylename=boldstyle, text="E "+self.tableWidget.item(i+1,1).text().replace('.', ',') + " m"))
                    #print "ta aqui"

    def insertDescriptionPDF(self):
        #locale.setlocale(locale.LC_ALL, ("pt_BR",""))
        ptex = "O imóvel descrito abaixo corresponde a um terreno de " + self.areaMetroQuad.replace('.', ',') + " m², localizado à " + self.adressImovel + ", no município de " + self.cityImovel +"/" + self.ufImovel + ", representado na planta " + self.plaintCor + ", processo SEI: " + self.numberSei

        if self.codIncraEdit.text():
            ptex += ", código INCRA "+ self.codIncraEdit.text()+ "."
            # self.textdoc.text.addElement(description)
        else:
            ptex += "."

        ptex += "<br></br><br></br>"
        ptex += "Inicia-se a descrição deste perímetro no vértice "
        ptex +='<font size=11 name="Times-Bold"> %s </font>' %self.tableWidget.item(0,0).text()
        ptex +=", de coordenadas "
        ptex +='<font size=11 name="Times-Bold">N %s m</font>' %self.tableWidget.item(0,2).text().replace('.', ',') + " e " +'<font size=11 name="Times-Bold">E %s m</font>' %self.tableWidget.item(0,1).text().replace('.', ',')
        rowCount = self.tableWidget.rowCount()
        itemPrev =''

        for i in range(0,rowCount):
            side = self.tableWidget.item(i,3).text()
            sideSplit = side.split("-")

            if self.tableWidget.item(i,7).text():
                if self.tableWidget.item(i,7).text() != itemPrev:

                    ptex += "; deste, segue confrontando com "
                    ptex += '<font size=11 name="Times-Bold"> %s </font>' %self.tableWidget.item(i,7).text().upper()
                    ptex += ", "
                else:
                    ptex += "; deste, segue "
            else:

                ptex += "; deste, segue "

            ptex +="com os seguintes azimute plano e distância: "
            ptex += '<font size=11 name="Times-Bold"> %s </font>'%self.tableWidget.item(i,4).text().replace('.', ',')
            ptex += " e "
            ptex += '<font size=11 name="Times-Bold"> %s m</font>' %self.tableWidget.item(i,6).text().replace('.', ',')
            ptex += "; até o vértice "
            itemPrev = self.tableWidget.item(i,7).text()

            if (i == rowCount - 1):
                ptex += '<font size=11 name="Times-Bold"> %s </font>' % sideSplit[1]
                ptex += ", de coordenadas "
                ptex += '<font size=11 name="Times-Bold">N %s m</font>' %self.tableWidget.item(0,2).text().replace('.', ',')
                ptex +=" e "
                ptex += '<font size=11 name="Times-Bold"> E %s m</font>' %self.tableWidget.item(0,1).text().replace('.', ',')
                ptex +=", encerrando esta descrição."
                ptex += "<br></br><br></br>"
                ptex += " Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro"

                if self.rbmcOrigemEdit.text():
                    ptex +=" , a partir da estação RBMC de " + self.rbmcOrigemEdit.text() + " de coordenadas "
                    ptex += '<font size=11 name="Times-Bold">E %s m</font>' %self.rbmcEsteEdit.text().replace('.', ',')
                    ptex += " e "
                    ptex +='<font size=11 name="Times-Bold">N %s m</font>' % self.rbmcNorteEdit.text().replace('.', ',')
                    ptex += " , "
                    ptex +="localizada em " + self.localRbmcEdit.text()+", "

                sp = self.projectionEdit.text().split(" ")[3]
                ptex += " e encontram-se representadas no sistema UTM, referenciadas ao Meridiano Central "
                ptex += '<font size=11 name="Times-Bold"> %s </font>' %self.meridianoEdit.text().replace('.', ',')
                ptex += ", Fuso " + '<font size=11 name="Times-Bold"> %s </font>' %str(sp)
                ptex += ", tendo como DATUM "
                ptex += '<font size=11 name="Times-Bold"> %s </font>' %self.datumEdit.text()
                ptex += ". Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM."

            else:
                ptex += '<font size=11 name="Times-Bold"> %s </font>' %sideSplit[1]
                ptex +=", de coordenadas "
                ptex += '<font size=11 name="Times-Bold">N %s m</font>' %self.tableWidget.item(i+1,2).text().replace('.', ',')
                ptex += " e "
                ptex += '<font size=11 name="Times-Bold">E %s m</font>' %self.tableWidget.item(i+1,1).text().replace('.', ',')

        return ptex
