# -*- coding: utf-8 -*-

import sys

from PyQt5 import QtCore, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        # Form.resize(1100, 700)
        Form.setMinimumSize(QtCore.QSize(900, 600))
        Form.setMaximumSize(QtCore.QSize(900, 600))
        Form.setStyleSheet(self._get_stylesheet())

        # Main Layout Root
        self.mainLayout = QtWidgets.QHBoxLayout(Form)
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.mainLayout.setSpacing(0)
        self.mainLayout.setObjectName("mainLayout")

        # ----------------------------------------------------
        # ASIDE / SIDEBAR
        # ----------------------------------------------------
        self.wAside = QtWidgets.QWidget(Form)
        self.wAside.setObjectName("wAside")
        self.wAside.setFixedWidth(200)

        self.layoutAside = QtWidgets.QVBoxLayout(self.wAside)
        self.layoutAside.setContentsMargins(16, 24, 16, 24)
        self.layoutAside.setSpacing(12)

        # Brand / App Title
        self.brandLabel = QtWidgets.QLabel("TaxProcess")
        self.brandLabel.setObjectName("brandLabel")
        self.layoutAside.addWidget(self.brandLabel)

        # Nav Buttons Group
        self.iva = QtWidgets.QPushButton("  IVA Acreditable")
        self.iva.setObjectName("navBtn")
        self.iva.setCheckable(True)
        self.iva.setChecked(True)
        self.layoutAside.addWidget(self.iva)

        self.seals = QtWidgets.QPushButton("  Ventas")
        self.seals.setObjectName("navBtn")
        self.seals.setCheckable(True)
        self.layoutAside.addWidget(self.seals)

        self.process = QtWidgets.QPushButton("  Procesos")
        self.process.setObjectName("navBtn")
        self.process.setCheckable(True)
        self.layoutAside.addWidget(self.process)

        # Navigation Button Group logic
        self.navGroup = QtWidgets.QButtonGroup(Form)
        self.navGroup.setExclusive(True)
        self.navGroup.addButton(self.iva, 0)
        self.navGroup.addButton(self.seals, 1)
        self.navGroup.addButton(self.process, 2)

        self.layoutAside.addStretch()

        # Exit Button
        self.exit = QtWidgets.QPushButton("  Cerrar")
        self.exit.setObjectName("exitBtn")
        self.layoutAside.addWidget(self.exit)

        self.mainLayout.addWidget(self.wAside)

        # ----------------------------------------------------
        # MAIN CONTENT VIEW (QStackedWidget)
        # ----------------------------------------------------
        self.wMain = QtWidgets.QStackedWidget(Form)
        self.wMain.setObjectName("wMain")

        # PAGE 1: IVA WIDGET
        self.ivaWidget = QtWidgets.QWidget()
        self.layoutIvaPage = QtWidgets.QVBoxLayout(self.ivaWidget)
        self.layoutIvaPage.setContentsMargins(32, 32, 32, 32)
        self.layoutIvaPage.setSpacing(24)

        # Header IVA
        self.wHeaderIva = QtWidgets.QWidget()
        self.layoutHeaderIva = QtWidgets.QVBoxLayout(self.wHeaderIva)
        self.layoutHeaderIva.setContentsMargins(0, 0, 0, 0)

        self.ivaTitle = QtWidgets.QLabel("IVA Acreditable")
        self.ivaTitle.setObjectName("pageTitle")
        self.ivaDescription = QtWidgets.QLabel(
            "Conciliación para el auxiliar de IVA acreditable en el ejercicio en curso."
        )
        self.ivaDescription.setObjectName("pageSubtitle")

        self.layoutHeaderIva.addWidget(self.ivaTitle)
        self.layoutHeaderIva.addWidget(self.ivaDescription)
        self.layoutIvaPage.addWidget(self.wHeaderIva)

        # Content Grid Layout (Two Column Dashboard Style)
        self.gridIvaContent = QtWidgets.QHBoxLayout()
        self.gridIvaContent.setSpacing(20)

        # Card Left: File Selection
        self.cardLoad = QtWidgets.QFrame()
        self.cardLoad.setObjectName("card")
        self.layoutCardLoad = QtWidgets.QVBoxLayout(self.cardLoad)
        self.layoutCardLoad.setContentsMargins(24, 24, 24, 24)

        self.wTitleLoad = QtWidgets.QLabel("Procesamiento de Información")
        self.wTitleLoad.setObjectName("cardTitle")
        self.wDescriptionLoad = QtWidgets.QLabel(
            "Selecciona la carpeta de trabajo para cargar archivos (.xlsx)"
        )
        self.wDescriptionLoad.setObjectName("cardSubtitle")

        self.layoutCardLoad.addWidget(self.wTitleLoad)
        self.layoutCardLoad.addWidget(self.wDescriptionLoad)
        self.layoutCardLoad.addSpacing(16)

        # Path Selector Row
        self.layoutPathRow = QtWidgets.QHBoxLayout()
        self.labelPath = QtWidgets.QLineEdit()
        self.labelPath.setPlaceholderText("Selecciona una ruta de trabajo...")
        self.labelPath.setReadOnly(True)

        self.pathButton = QtWidgets.QPushButton("Buscar ruta")
        self.pathButton.setObjectName("secondaryBtn")

        self.layoutPathRow.addWidget(self.labelPath)
        self.layoutPathRow.addWidget(self.pathButton)
        self.layoutCardLoad.addLayout(self.layoutPathRow)

        # Directory Tree View
        self.wWorkDirectory = QtWidgets.QTreeView()
        self.wWorkDirectory.setHeaderHidden(True)
        self.layoutCardLoad.addWidget(self.wWorkDirectory)

        self.gridIvaContent.addWidget(self.cardLoad, stretch=1)

        # Card Right: Progress & Execution
        self.cardProcess = QtWidgets.QFrame()
        self.cardProcess.setObjectName("card")
        self.layoutCardProcess = QtWidgets.QVBoxLayout(self.cardProcess)
        self.layoutCardProcess.setContentsMargins(24, 24, 24, 24)

        self.processCardTitle = QtWidgets.QLabel("Estado de Ejecución")
        self.processCardTitle.setObjectName("cardTitle")
        self.layoutCardProcess.addWidget(self.processCardTitle)
        self.layoutCardProcess.addStretch()

        self.progressDescription = QtWidgets.QLabel("Esperando inicio de proceso...")
        self.progressDescription.setAlignment(QtCore.Qt.AlignCenter)

        self.progressBar = QtWidgets.QProgressBar()
        self.progressBar.setValue(0)

        self.layoutCardProcess.addWidget(self.progressDescription)
        self.layoutCardProcess.addWidget(self.progressBar)
        self.layoutCardProcess.addStretch()

        # Action Buttons
        self.layoutButtons = QtWidgets.QHBoxLayout()
        self.formatButton = QtWidgets.QPushButton("Formato")
        self.formatButton.setObjectName("secondaryBtn")

        self.processButton = QtWidgets.QPushButton("Procesar")
        self.processButton.setObjectName("primaryBtn")

        self.layoutButtons.addWidget(self.formatButton)
        self.layoutButtons.addWidget(self.processButton)
        self.layoutCardProcess.addLayout(self.layoutButtons)

        self.gridIvaContent.addWidget(self.cardProcess, stretch=1)
        self.layoutIvaPage.addLayout(self.gridIvaContent)

        self.wMain.addWidget(self.ivaWidget)

        # PAGE 2: SALES WIDGET
        self.salesWidget = QtWidgets.QWidget()
        self.layoutSales = QtWidgets.QVBoxLayout(self.salesWidget)
        self.layoutSales.setContentsMargins(32, 32, 32, 32)
        self.salesTitle = QtWidgets.QLabel("Ventas")
        self.salesTitle.setObjectName("pageTitle")
        self.layoutSales.addWidget(self.salesTitle)
        self.layoutSales.addStretch()
        self.wMain.addWidget(self.salesWidget)

        # PAGE 3: PROCESS WIDGET
        self.processWidget = QtWidgets.QWidget()
        self.layoutProcessPage = QtWidgets.QVBoxLayout(self.processWidget)
        self.layoutProcessPage.setContentsMargins(32, 32, 32, 32)
        self.layoutProcessPage.setSpacing(20)

        self.processTitle = QtWidgets.QLabel("Procesos")
        self.processTitle.setObjectName("pageTitle")
        self.processDescription = QtWidgets.QLabel(
            "Sección dedicada a diversos procesos genéricos."
        )
        self.processDescription.setObjectName("pageSubtitle")

        self.comboProcess = QtWidgets.QComboBox()
        self.comboProcess.addItems(
            [
                "-- Selecciona un proceso --",
                "Unir archivos (.xlsx)",
                "Descarga masiva",
                "Crear comprimido (.7z)",
                "Flujograma",
            ]
        )

        self.layoutProcessPage.addWidget(self.processTitle)
        self.layoutProcessPage.addWidget(self.processDescription)
        self.layoutProcessPage.addWidget(self.comboProcess)
        self.layoutProcessPage.addStretch()

        self.wMain.addWidget(self.processWidget)

        # Add Stacked Widget to Root Layout
        self.mainLayout.addWidget(self.wMain)

        # Connect Signals
        self.retranslateUi(Form)
        self.navGroup.buttonClicked[int].connect(self.wMain.setCurrentIndex)
        self.exit.clicked.connect(Form.close)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "TaxProcess - Gestión Fiscal"))

    def _get_stylesheet(self):
        return """
            /* General Theme Dark Minimalist */
            QWidget {
                background-color: #0f172a;
                color: #f8fafc;
                font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
                font-size: 13px;
            }

            /* Aside / Sidebar */
            #wAside {
                background-color: #1e293b;
                border-right: 1px solid #334155;
            }

            #brandLabel {
                font-size: 18px;
                font-weight: bold;
                color: #38bdf8;
                padding-bottom: 12px;
            }

            /* Buttons Navigation */
            #navBtn {
                background-color: transparent;
                color: #94a3b8;
                border: none;
                border-radius: 8px;
                text-align: left;
                padding: 10px 14px;
                font-weight: 500;
            }
            #navBtn:hover {
                background-color: #334155;
                color: #f8fafc;
            }
            #navBtn:checked {
                background-color: #0284c7;
                color: #ffffff;
                font-weight: 600;
            }

            #exitBtn {
                background-color: transparent;
                color: #f43f5e;
                border: 1px solid #f43f5e;
                border-radius: 8px;
                padding: 8px 14px;
                font-weight: 500;
            }
            #exitBtn:hover {
                background-color: #f43f5e;
                color: #ffffff;
            }

            /* Headers */
            #pageTitle {
                font-size: 24px;
                font-weight: 700;
                color: #ffffff;
            }
            #pageSubtitle {
                font-size: 13px;
                color: #94a3b8;
            }

            /* Cards */
            #card {
                background-color: #1e293b;
                border: 1px solid #334155;
                border-radius: 12px;
            }
            #cardTitle {
                font-size: 15px;
                font-weight: 600;
                color: #f1f5f9;
            }
            #cardSubtitle {
                font-size: 12px;
                color: #64748b;
            }

            /* Forms & Inputs */
            QLineEdit, QComboBox {
                background-color: #0f172a;
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 8px 12px;
                color: #f8fafc;
            }
            QLineEdit:focus, QComboBox:focus {
                border-color: #38bdf8;
            }

            /* Buttons */
            #primaryBtn {
                background-color: #0284c7;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: 600;
            }
            #primaryBtn:hover {
                background-color: #0369a1;
            }

            #secondaryBtn {
                background-color: #334155;
                color: #f8fafc;
                border: none;
                border-radius: 6px;
                padding: 10px 16px;
                font-weight: 500;
            }
            #secondaryBtn:hover {
                background-color: #475569;
            }

            /* TreeView */
            QTreeView {
                background-color: #0f172a;
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 6px;
            }

            /* Progress Bar */
            QProgressBar {
                border: none;
                background-color: #0f172a;
                border-radius: 6px;
                text-align: center;
                color: #ffffff;
                font-weight: bold;
                height: 12px;
            }
            QProgressBar::chunk {
                background-color: #38bdf8;
                border-radius: 6px;
            }
        """


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
