# ============================================================================================================
# ============================================================================================================
# ==                                                                                                        ==
# ==                                           alias Manager                                                ==
# ==                                                                                                        ==
# ============================================================================================================
# ============================================================================================================
# ABOUT
# ============================================================================================================
# version v1.0
# Macro developed for FreeCAD (http://www.freecadweb.org/).
# This macro helps managing aliases inside FreeCAD Spreadsheet workbench. It is able to:
#        - create aliases based on first column (A column)
#        - delete aliases placed on a column
#        - move aliases from one column to other one
#        - create a "part family", that's it, create different files for each column in a range. It will add
#          to the original name a suffix based on first row (1)
# More information might be found on FreeCAD forums: http://forum.freecadweb.org/
#
#
# LICENSE
# ============================================================================================================
# Original work done by tarihatari (https://github.com/tarihatari/FreeCAD_Macros)
# Improved by Pablo Gil Fernandez
#
# Copyright (c) 2016 tarihatari & Pablo Gil Fernandez
#
# This work is licensed under GNU Lesser General Public License (LGPL).
# To view a copy of this license, visit https://www.gnu.org/licenses/lgpl-3.0.html.
#
# ============================================================================================================
__title__   = "alias manager"
__author__  = "Pablo Gil Fernandez"
__version__ = "01.00"
__date__    = "20/11/2016"
 
__Comment__ = "This macro helps managing aliases inside FreeCAD Spreadsheet workbench. It is able to create, delete, move aliases and create a 'part family' group of files"
 
__Wiki__ = "https://github.com/pgilfernandez/FreeCAD_AliasManager"
__Help__ = "https://github.com/pgilfernandez/FreeCAD_AliasManager"
__Status__ = "stable"
__Requires__ = "FreeCAD 0.16"

from PySide import QtGui, QtCore
from FreeCAD import Gui
import os
import string
import traceback

App = FreeCAD
Gui = FreeCADGui

# ==
# person = Object(name = "Bernhard", gender = "male", age = 42)
# https://stackoverflow.com/questions/652276/is-it-possible-to-create-anonymous-objects-in-python
Object = lambda **kwargs: type("Object", (), kwargs)

# ========================================================
# ===== Info popup window ================================
# ========================================================
class infoPopup(QtGui.QDialog):
    def __init__(self, parent=None):
        self.dialog = None
        self.dialog = QtGui.QDialog()
        self.dialog.resize(360,400)
        self.dialog.setWindowTitle("About...")

        info = QtGui.QTextEdit("<h2>INFORMATION</h2><hr><br>This macro helps managing aliases inside FreeCAD Spreadsheet workbench. It is able to create, delete, move aliases and create a 'part family' group of files.<br><h2>USAGE</h2><hr><ul><li><b>set aliases</b>: based on first column (A column), it will create aliases for the range given. If an alias is already set for any other cell the command won't work, it will be needed to clear them before setting them again.</li><li><b>Clear aliases</b>: it will clear all aliases inside the given range of cells (only one column).</li><li><b>Move aliases</b>: it will clear and set aliases <b>from</b> a given column <b>to</b> a new one.</li><li><b>Generate part family</b>: it will create different files for each column in a range. It will add to the original name a suffix based on first row. If an alias is already set for any other cell the command won't work, it will be needed to clear them before running the command.</li></ul><h2>LICENCE</h2><hr>Original work done by <b>tarihatari</b> (<a href='https://github.com/tarihatari/FreeCAD_Macros'>https://github.com/tarihatari/FreeCAD_Macros</a>)<br>Improved by <b>Pablo Gil Fernandez</b><br><br>Copyright (c) 2016 tarihatari & Pablo Gil Fernandez<br><br>This work is licensed under GNU Lesser General Public License (LGPL).<br>To view a copy of this license, visit <a href='https://www.gnu.org/licenses/lgpl-3.0.html'>https://www.gnu.org/licenses/lgpl-3.0.html</a>.<br>")
        info.setReadOnly(True)
        info.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)

        okbox = QtGui.QDialogButtonBox(self.dialog)
        okbox.setOrientation(QtCore.Qt.Horizontal)
        okbox.setStandardButtons(QtGui.QDialogButtonBox.Close)
        okbox.setFocus()

        grid2 = QtGui.QGridLayout()
        grid2.setSpacing(10)
        grid2.addWidget(info, 0, 0)
        grid2.addWidget(okbox, 1, 0)

        self.dialog.setLayout(grid2)

        QtCore.QObject.connect(okbox, QtCore.SIGNAL("rejected()"), self.close)
        QtCore.QMetaObject.connectSlotsByName(self.dialog)
        self.dialog.show()
        self.dialog.exec_()

    def close(self):
        self.dialog.close()


# ========================================================
# ===== Main code ========================================
# ========================================================


# ===== Global variables ==============================================
alphabet_list = list(string.ascii_uppercase)

property_definitions = [
    Object(name = "AliasColumn", type = "App::PropertyString", tooltip = "The colom where the aliasses reside"),
    Object(name = "ColumnFrom", type = "App::PropertyString", tooltip = "Last used column"),
    Object(name = "RowFrom", type = "App::PropertyInteger", tooltip = "Last used row from"),
    Object(name = "RowTo", type = "App::PropertyInteger", tooltip = "Last used row to")
]

class p():
    def aliasManager(self):
        try:

# ===== Variables ==============================================
            mode = self.d1.currentText()
            alias_column = self.d0.text()
            column_from = self.d2.text()
            column_to = self.d3.text()
            row_from = self.d4.value()
            row_to = self.d5.value()

# ===== Save Variables to ease further run =====================
            for prop in property_definitions:
                if not App.ActiveDocument.Spreadsheet.PropertiesList.__contains__(prop.name):
                    App.ActiveDocument.Spreadsheet.addProperty(prop.type, prop.name, "Alias Manager", prop.tooltip)

            App.ActiveDocument.Spreadsheet.AliasColumn = alias_column
            App.ActiveDocument.Spreadsheet.ColumnFrom = column_from
            App.ActiveDocument.Spreadsheet.RowFrom = row_from
            App.ActiveDocument.Spreadsheet.RowTo = row_to

# ===== Mode - Set ==============================================
            if mode == "Set aliases":
                for i in range(row_from,row_to+1):
                    cell_from = str(alias_column) + str(i)
                    cell_to = str(column_from) + str(i)
                    App.ActiveDocument.Spreadsheet.setAlias(cell_to, '')
                    App.ActiveDocument.Spreadsheet.setAlias(cell_to, App.ActiveDocument.Spreadsheet.getContents(cell_from))
                    App.ActiveDocument.recompute()

                FreeCAD.Console.PrintMessage("\nAliases set\n")


# ===== Mode - Clear ==============================================
            elif mode == "Clear aliases":
                for i in range(row_from,row_to+1):
                    cell_to = str(column_from) + str(i)
                    App.ActiveDocument.Spreadsheet.setAlias(cell_to, '')
                    App.ActiveDocument.recompute()

                FreeCAD.Console.PrintMessage("\nAliases cleared\n")


# ===== Mode - Move ==============================================
            elif mode == "Move aliases":
                # Next time we probably want to use the moved column as the from column
                App.ActiveDocument.Spreadsheet.ColumnFrom = column_to

                self.d3.setDisabled(False)
                for i in range(row_from,row_to+1):
                    cell_reference = str(alias_column) + str(i)
                    cell_from = column_from + str(i)
                    cell_to = column_to + str(i)
                    App.ActiveDocument.Spreadsheet.setAlias(cell_from, '')
                    App.ActiveDocument.recompute()
                    App.ActiveDocument.Spreadsheet.setAlias(cell_to, App.ActiveDocument.Spreadsheet.getContents(cell_reference))
                    App.ActiveDocument.recompute()
                FreeCAD.Console.PrintMessage("\nAliases moved\n")


# ===== Mode - Generate part family ==============================================
            elif mode == "Generate part family":
                doc = FreeCAD.ActiveDocument    
                if not doc.FileName:
                    FreeCAD.Console.PrintError('\nMust save project first\n')
                    
                docDir, docFilename = os.path.split(doc.FileName)
                filePrefix = os.path.splitext(docFilename)[0]

                def char_range(c1, c2):
                    """Generates the characters from `c1` to `c2`, inclusive."""
                    for c in xrange(ord(c1), ord(c2)+1):
                        yield str.capitalize(chr(c))

                fam_range = []
                for c in char_range(str(column_from), str(column_to)):
                    fam_range.append(c)

                for index in range(len(fam_range)):
                    # set aliases
                    for i in range(row_from,row_to+1):
                        cell_reference = str(alias_column) + str(i)
                        cell_from = str(fam_range[index-1]) + str(i)
                        cell_to = str(fam_range[index]) + str(i)
                        App.ActiveDocument.Spreadsheet.setAlias(cell_from, '')
                        App.ActiveDocument.recompute()
                        App.ActiveDocument.Spreadsheet.setAlias(cell_to, App.ActiveDocument.Spreadsheet.getContents(cell_reference))
                        App.ActiveDocument.recompute()
                        sfx = str(fam_range[index]) + '1'

                    # save file
                    suffix = App.ActiveDocument.Spreadsheet.getContents(sfx)
                    filename = filePrefix + '_' + suffix + '.fcstd'
                    filePath = os.path.join(docDir, filename)
                
                    FreeCAD.Console.PrintMessage("\nSaving view to %s\n" % filePath)
                    App.getDocument(filePrefix).saveAs(filePath)

                # Clear last aliases created:
                for j in range(row_from,row_to+1):
                    cell_to = str(column_to) + str(j)
                    App.ActiveDocument.Spreadsheet.setAlias(cell_to, '')
                App.ActiveDocument.recompute()

                # Turn working file to original naming:
                filename = filePrefix + '.fcstd'
                filePath = os.path.join(docDir, filename)
                FreeCAD.Console.PrintMessage("\nSaving original view to %s\n" % filePath)
                App.getDocument(filePrefix).saveAs(filePath)
                FreeCAD.Console.PrintMessage("\nPart family files generated\n")



# ===== If errors ==============================================
            else:
                FreeCAD.Console.PrintError("\nError or 'TODO'\n")


        except:
            FreeCAD.Console.PrintError("\nUnable to complete task\n")
            traceback.print_exc()

            self.close()
 

    def close(self):
        self.dialog.hide()


# ========================================================
# ===== GUI menu ========================================
# ========================================================
    def __init__(self):
        infoIcon = ['16 16 3 1',
                '   c None',
                '+  c #444444',
                '.  c #e6e6e6',
                '     ......    ',
                '   ..........  ',
                '  ......++.... ',
                ' .......++.....',
                ' ..............',
                '.....+++++......',
                '....+++++.......',
                '.......++.......',
                '.......++.......',
                '.......+........',
                '......++........',
                ' .....++.+.....',
                ' .....++++.....',
                '  .....++..... ',
                '   ..........  ',
                '     ......    ']

        # Hide/show "to column" label and spinbox based on mode type
        def disableWidget(currentIndex):
            # Hide by default
            iN0.hide()
            self.d0.setEnabled(False)
            self.d0.hide()

            iN3.hide()
            self.d3.setEnabled(False)
            self.d3.hide()

            if self.d1.currentText() != "Clear aliases":
                iN0.show()
                self.d0.setEnabled(True)
                self.d0.show()
            if not (self.d1.currentText() == "Set aliases" or self.d1.currentText() == "Clear aliases"):
                iN3.show()
                self.d3.setEnabled(True)
                self.d3.show()

        self.dialog = None
        
        self.dialog = QtGui.QDialog()
        self.dialog.resize(400,140)
        
        self.dialog.setWindowTitle("Alias manager")

        iN0 = QtGui.QLabel("alias column:")
        iN0.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.d0 = QtGui.QLineEdit()
        d0DefaultValue = "A"
        if App.ActiveDocument.Spreadsheet.PropertiesList.__contains__("AliasColumn"):
            d0DefaultValue = App.ActiveDocument.Spreadsheet.AliasColumn
        self.d0.setText(d0DefaultValue) # set default item

        iN1 = QtGui.QLabel("mode:")
        iN1.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.d1 = QtGui.QComboBox()
        self.d1.addItem("Set aliases")
        self.d1.addItem("Clear aliases")
        self.d1.addItem("Move aliases")
        self.d1.addItem("Generate part family")
        self.d1.setCurrentIndex(0) # set default item
        self.d1.currentIndexChanged['QString'].connect(disableWidget)
        
        iN2a = QtGui.QLabel("column:")
        iN2a.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        iN2b = QtGui.QLabel("from")
        iN2b.setAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)
        self.d2 = QtGui.QLineEdit()
        d2DefaultValue = "B"
        if App.ActiveDocument.Spreadsheet.PropertiesList.__contains__("ColumnFrom"):
            d2DefaultValue = App.ActiveDocument.Spreadsheet.ColumnFrom
        self.d2.setText(d2DefaultValue) # set default item

        iN3 = QtGui.QLabel("to")
        iN3.setAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)
        self.d3 = QtGui.QLineEdit()
        d3DefaultValue = "C"
        self.d3.setText(d3DefaultValue) # set default item

        iN4 = QtGui.QLabel("from row:")
        iN4.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.d4 = QtGui.QSpinBox()
        d4DefaultValue = 2
        if App.ActiveDocument.Spreadsheet.PropertiesList.__contains__("RowFrom"):
            d4DefaultValue = App.ActiveDocument.Spreadsheet.RowFrom
        self.d4.setValue(d4DefaultValue) # set default item
        self.d4.setSingleStep(1.0)
        self.d4.setMinimum(1.0)

        iN5 = QtGui.QLabel("to row:")
        iN5.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.d5 = QtGui.QSpinBox()
        d5DefaultValue = 4
        if App.ActiveDocument.Spreadsheet.PropertiesList.__contains__("RowTo"):
            d5DefaultValue = App.ActiveDocument.Spreadsheet.RowTo
        self.d5.setValue(d5DefaultValue) # set default item
        self.d5.setSingleStep(1.0)
        self.d5.setMinimum(1.0)

        # Info button
        self.d6 = QtGui.QPushButton("")
        self.d6.setFixedWidth(40)
        self.d6.setIcon(QtGui.QIcon(QtGui.QPixmap(infoIcon)))
        self.d6.clicked.connect(self.popup)

        okbox = QtGui.QDialogButtonBox(self.dialog)
        okbox.setOrientation(QtCore.Qt.Horizontal)
        okbox.setStandardButtons(QtGui.QDialogButtonBox.Ok|QtGui.QDialogButtonBox.Close)

        grid = QtGui.QGridLayout()
        grid.setSpacing(10)

        # Mode
        grid.addWidget(self.d1, 0, 0, 1, 3)

        # Alias Column
        grid.addWidget(iN0, 1, 0, 1, 1)
        grid.addWidget(self.d0, 1, 1, 1, 2)

        # column, column from
        grid.addWidget(iN2a,    3, 0, 1, 1)
        grid.addWidget(iN2b,    2, 1, 1, 1)
        grid.addWidget(self.d2, 3, 1, 1, 1)
        # column to
        grid.addWidget(iN3,     2, 2, 1, 1)
        grid.addWidget(self.d3, 3, 2, 1, 1)
        # from row
        grid.addWidget(iN4,     4, 0, 1, 1)
        grid.addWidget(self.d4, 4, 1, 1, 1)
        # to row
        grid.addWidget(iN5,     5, 0, 1, 1)
        grid.addWidget(self.d5, 5, 1)
        # + info
        grid.addWidget(self.d6, 7, 0, 1, 1)
        # cancel, OK
        grid.addWidget(okbox,   7, 1, 1, 2)

        self.dialog.setLayout(grid)

        # update widget to hide/show what we need.
        disableWidget(self.d1.currentIndex)

        # # Set Tab order (not needed anymore because of enabling/disabling spinboxes)
        # self.dialog.setTabOrder(self.d3, self.d1)
        # self.dialog.setTabOrder(self.d1, self.d2)
        # self.dialog.setTabOrder(self.d2, self.d4)
        # self.dialog.setTabOrder(self.d4, self.d5)
        # self.dialog.setTabOrder(self.d5, self.d3)
 
        QtCore.QObject.connect(okbox, QtCore.SIGNAL("rejected()"), self.close)
        QtCore.QObject.connect(okbox, QtCore.SIGNAL("accepted()"), self.aliasManager)
        QtCore.QMetaObject.connectSlotsByName(self.dialog)
        self.dialog.show()
        self.dialog.exec_()


    def popup(self):
        self.dialog2 = infoPopup()
        self.dialog2.exec_()
p()