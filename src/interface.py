# -*- coding: utf-8 -*-

import os
import sys
import time
from datetime import datetime
from html import escape

import resources_rc


# ⭐ NUEVA FUNCIÓN - Detectar si estamos en .exe o en desarrollo
def get_base_path():
    """
    Devuelve la ruta base correcta según el entorno:
    - En desarrollo: carpeta del proyecto
    - En .exe compilado: carpeta temporal de Nuitka (_MEIPASS)
    """
    if getattr(sys, "frozen", False):
        # Estamos en un .exe compilado (PyInstaller o Nuitka)
        return sys._MEIPASS
    else:
        # Estamos en desarrollo
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def resource_path(relative_path):
    """
    Obtiene la ruta completa de un archivo de recurso.

    Uso:
        icon_path = resource_path("icons/favicon.ico")
        logo_path = resource_path("icons/logo.png")
    """
    base_path = get_base_path()
    return os.path.join(base_path, relative_path)


from flujograma_process import main_flujograma_process
from header_process import main_header_process
from iva_format_process import main_iva_format_process
from iva_process import main_iva_process
from merge_process import main_merge_process
from PyQt5 import QtCore, QtGui, QtWidgets
from underline_aux_process import main_underline_aux_process
from underline_bank_process import main_underline_bank_process
from v0_process import main_v0_process
from v16_process import main_v16_process

MAP_SUFIX_IDX = {
    "1": "iva",
    "2": "v0",
    "3": "v16",
}


class ModernIconProvider(QtWidgets.QFileIconProvider):
    def __init__(self):
        super().__init__()
        self.custom_icons = {
            ".xlsx": ":/icons/xlsx.png",
            ".xls": ":/icons/xlsx.png",
            ".pdf": ":/icons/pdf.png",
            ".txt": ":/icons/txt.png",
            ".csv": ":/icons/txt.png",
            # "folder": ":/icons/folder_open.svg",
            "file": ":/icons/archivo.png",
        }

    def icon(self, info):
        # Intentar cargar icono personalizado
        if info.isDir():
            icon = QtGui.QIcon(self.custom_icons.get("folder", ""))
        else:
            ext = info.suffix().lower()
            icon_path = self.custom_icons.get(
                f".{ext}", self.custom_icons.get("file", "")
            )
            icon = QtGui.QIcon(icon_path)

        # Si el icono está vacío (porque la ruta :/icons falló), usar el del sistema
        if icon.isNull():
            return super().icon(info)
        return icon


class DirectoryManager:
    def __init__(self, tree_view):
        self.tree = tree_view
        self.model = QtWidgets.QFileSystemModel()
        self.model.setIconProvider(ModernIconProvider())
        self.tree.setModel(self.model)

        # --- AJUSTES VISUALES MODERNOS ---
        self.tree.setIconSize(QtCore.QSize(24, 24))  # Iconos más grandes y legibles
        self.tree.setIndentation(15)  # Espacio para que el icono no pegue al borde
        self.tree.setAnimated(True)
        self.tree.setSortingEnabled(True)

        # Eliminar el encabezado (opcional, se ve más limpio)
        self.tree.header().hide()

        # Bloqueo de expansión
        self.tree.setRootIsDecorated(False)
        self.tree.setItemsExpandable(False)
        self.tree.setExpandsOnDoubleClick(False)

        # Ocultar columnas sobrantes
        for i in range(1, 4):
            self.tree.hideColumn(i)

    def load_directory(self, path, label_path):
        if not os.path.exists(path):
            label_path.setText("❌ Ruta inválida")
            return

        self.model.setRootPath(path)
        self.tree.setRootIndex(self.model.index(path))
        label_path.setText(f"Directorio actual: {os.path.basename(path)}")


class Ui_fedApp(object):
    def setupUi(self, fedApp):
        self.main_window = fedApp
        fedApp.setObjectName("fedApp")
        fedApp.resize(1200, 700)

        # 🔒 Tamaño fijo basado en el contenido
        fedApp.setFixedSize(fedApp.size())  # o self.setFixedSize(800, 600)
        fedApp.setWindowFlags(
            fedApp.windowFlags() | QtCore.Qt.MSWindowsFixedSizeDialogHint
        )

        fedApp.setWindowTitle("FED - Sistema de Devoluciones")

        # Variable para guardar la posición del mouse
        # Variables para el arrastre
        self.dragging = False
        self.offset = QtCore.QPoint()
        self.old_pos = None

        self.wWrapper = QtWidgets.QWidget(fedApp)
        self.wWrapper.setObjectName("wWrapper")
        wrapperLayout = QtWidgets.QHBoxLayout(self.wWrapper)
        wrapperLayout.setContentsMargins(10, 10, 10, 10)
        wrapperLayout.setSpacing(10)
        self.first_directory = None
        self.second_directory = None

        # ============ SIDEBAR (ASIDE) ============
        self.wAsider = QtWidgets.QWidget(self.wWrapper)
        self.wAsider.setObjectName("wAsider")
        self.wAsider.setMaximumWidth(180)
        self.wAsider.setMinimumWidth(180)

        asideLayout = QtWidgets.QVBoxLayout(self.wAsider)
        asideLayout.setContentsMargins(10, 10, 10, 10)
        asideLayout.setSpacing(8)

        # Botón Directorio de Trabajo
        self.rootButton = QtWidgets.QPushButton(self.wAsider)
        self.rootButton.setObjectName("rootButton")
        self.rootButton.setMinimumHeight(45)
        rootIcon = QtGui.QIcon()
        rootIcon.addPixmap(QtGui.QPixmap(":/icons/folder_open.svg"))
        self.rootButton.setIcon(rootIcon)
        self.rootButton.setIconSize(QtCore.QSize(24, 24))
        self.rootButton.setText("Directorio de\nTrabajo")
        self.rootButton.setToolTip("Configurar directorio de trabajo")
        asideLayout.addWidget(self.rootButton)

        # Botón FED
        self.fedButton = QtWidgets.QPushButton(self.wAsider)
        self.fedButton.setObjectName("fedButton")
        self.fedButton.setMinimumHeight(45)
        fedIcon = QtGui.QIcon()
        fedIcon.addPixmap(QtGui.QPixmap(":/icons/file_document.svg"))
        self.fedButton.setIcon(fedIcon)
        self.fedButton.setIconSize(QtCore.QSize(24, 24))
        self.fedButton.setText("FED")
        self.fedButton.setToolTip("Formato Electrónico de Devoluciones")
        asideLayout.addWidget(self.fedButton)

        # Botón Procesos
        self.processButton = QtWidgets.QPushButton(self.wAsider)
        self.processButton.setObjectName("processButton")
        self.processButton.setMinimumHeight(45)
        processIcon = QtGui.QIcon()
        processIcon.addPixmap(QtGui.QPixmap(":/icons/settings.svg"))
        self.processButton.setIcon(processIcon)
        self.processButton.setIconSize(QtCore.QSize(24, 24))
        self.processButton.setText("Procesos")
        self.processButton.setToolTip("Gestionar procesos de trabajo")
        asideLayout.addWidget(self.processButton)

        # Espaciador
        spacerItem = QtWidgets.QSpacerItem(
            20, 400, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding
        )
        asideLayout.addItem(spacerItem)

        # Botón Salir
        self.exitButton = QtWidgets.QPushButton(self.wAsider)
        self.exitButton.setObjectName("exitButton")
        self.exitButton.setMinimumHeight(45)
        exitIcon = QtGui.QIcon()
        exitIcon.addPixmap(QtGui.QPixmap(":/icons/exit.svg"))
        self.exitButton.setIcon(exitIcon)
        self.exitButton.setIconSize(QtCore.QSize(24, 24))
        self.exitButton.setText("Salir")
        self.exitButton.setToolTip("Cerrar aplicación")
        asideLayout.addWidget(self.exitButton)

        wrapperLayout.addWidget(self.wAsider)

        # ============ CONTENEDOR PRINCIPAL ============
        self.wMain = QtWidgets.QWidget(self.wWrapper)
        self.wMain.setObjectName("wMain")
        mainLayout = QtWidgets.QVBoxLayout(self.wMain)
        mainLayout.setContentsMargins(0, 0, 0, 0)
        mainLayout.setSpacing(0)

        self.stackedWidget = QtWidgets.QStackedWidget(self.wMain)
        self.stackedWidget.setObjectName("stackedWidget")
        mainLayout.addWidget(self.stackedWidget)

        # ============ PÁGINA 1: DIRECTORIO DE TRABAJO ============
        self.setRootDirectory = QtWidgets.QWidget()
        self.setRootDirectory.setObjectName("setRootDirectory")
        rootLayout = QtWidgets.QVBoxLayout(self.setRootDirectory)
        rootLayout.setContentsMargins(0, 0, 0, 0)
        rootLayout.setSpacing(0)

        # Header
        self.rootHeader = QtWidgets.QWidget(self.setRootDirectory)
        self.rootHeader.setObjectName("rootHeader")
        self.rootHeader.setMaximumHeight(120)
        headerLayout = QtWidgets.QVBoxLayout(self.rootHeader)
        headerLayout.setContentsMargins(20, 15, 20, 15)

        self.wrapperRootTitle = QtWidgets.QWidget()
        self.wrapperRootTitle.setObjectName("wrapperRootTitle")
        titleLayout = QtWidgets.QVBoxLayout(self.wrapperRootTitle)
        titleLayout.setContentsMargins(0, 0, 0, 0)
        titleLayout.setSpacing(5)

        self.rootTitle = QtWidgets.QLabel("📁 Directorio de Trabajo")
        self.rootTitle.setObjectName("rootTitle")
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        self.rootTitle.setFont(font)
        titleLayout.addWidget(self.rootTitle)

        self.rootDescription = QtWidgets.QLabel(
            "Selecciona el directorio raíz donde se procesarán todos tus archivos"
        )
        self.rootDescription.setObjectName("rootDescription")
        font = QtGui.QFont()
        font.setPointSize(10)
        self.rootDescription.setFont(font)
        titleLayout.addWidget(self.rootDescription)

        headerLayout.addWidget(self.wrapperRootTitle)
        rootLayout.addWidget(self.rootHeader)

        # Contenedor de búsqueda y árbol
        self.wWrapperDirectory = QtWidgets.QWidget()
        self.wWrapperDirectory.setObjectName("wWrapperDirectory")
        dirLayout = QtWidgets.QVBoxLayout(self.wWrapperDirectory)
        dirLayout.setContentsMargins(20, 10, 20, 20)
        dirLayout.setSpacing(12)

        # Barra de búsqueda
        self.wrapperSearchBotton = QtWidgets.QWidget()
        self.wrapperSearchBotton.setObjectName("wrapperSearchBotton")
        self.wrapperSearchBotton.setMaximumHeight(60)
        searchLayout = QtWidgets.QHBoxLayout(self.wrapperSearchBotton)
        searchLayout.setContentsMargins(0, 0, 0, 0)
        searchLayout.setSpacing(10)

        self.iconPath = QtWidgets.QLabel()
        self.iconPath.setObjectName("iconPath")
        self.iconPath.setMaximumSize(QtCore.QSize(40, 40))
        self.iconPath.setPixmap(QtGui.QPixmap(":/icons/folder_open.svg"))
        self.iconPath.setScaledContents(True)
        searchLayout.addWidget(self.iconPath)

        self.path = QtWidgets.QLabel("Selecciona tu directorio de trabajo")
        self.path.setObjectName("path")
        self.path.setAlignment(QtCore.Qt.AlignVCenter)
        searchLayout.addWidget(self.path, 1)

        self.buttonPath = QtWidgets.QPushButton("Establecer")
        self.buttonPath.setObjectName("buttonPath")
        self.buttonPath.setMaximumWidth(120)
        self.buttonPath.setMinimumHeight(40)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/icons/search.svg"))
        self.buttonPath.setIcon(icon)
        self.buttonPath.setIconSize(QtCore.QSize(20, 20))
        searchLayout.addWidget(self.buttonPath)

        dirLayout.addWidget(self.wrapperSearchBotton)

        # TreeWidget
        self.wRootDiretory = QtWidgets.QTreeView()
        self.wRootDiretory.setObjectName("wRootDiretory")
        # self.wRootDiretory.setHeaderLabel("Archivos del Directorio")
        self.wRootDiretory.setHeaderHidden(True)
        # self.wRootDiretory.setColumnCount(1)
        # self.wRootDiretory.setMinimumHeight(300)
        dirLayout.addWidget(self.wRootDiretory)

        rootLayout.addWidget(self.wWrapperDirectory, 1)
        self.stackedWidget.addWidget(self.setRootDirectory)

        # Inicializar el Manager
        self.dir_manager = DirectoryManager(self.wRootDiretory)

        # ============ PÁGINA 2: FED ============
        self.fedProcess = QtWidgets.QWidget()
        self.fedProcess.setObjectName("fedProcess")
        fedLayout = QtWidgets.QVBoxLayout(self.fedProcess)
        fedLayout.setContentsMargins(0, 0, 0, 0)
        fedLayout.setSpacing(0)

        # Header FED
        self.fedHeader = QtWidgets.QWidget()
        self.fedHeader.setObjectName("fedHeader")
        self.fedHeader.setMaximumHeight(140)
        fedHeaderLayout = QtWidgets.QVBoxLayout(self.fedHeader)
        fedHeaderLayout.setContentsMargins(20, 15, 20, 15)

        self.wrapperFedTitle = QtWidgets.QWidget()
        self.wrapperFedTitle.setObjectName("wrapperFedTitle")
        fedTitleLayout = QtWidgets.QVBoxLayout(self.wrapperFedTitle)
        fedTitleLayout.setContentsMargins(0, 0, 0, 0)
        fedTitleLayout.setSpacing(5)

        self.fedTitle = QtWidgets.QLabel("📄 Formato Electrónico de Devoluciones")
        self.fedTitle.setObjectName("fedTitle")
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        self.fedTitle.setFont(font)
        fedTitleLayout.addWidget(self.fedTitle)

        self.fedDescription = QtWidgets.QLabel(
            "Selecciona el tipo de proceso (IVA, V0, V16) y la acción a realizar"
        )
        self.fedDescription.setObjectName("fedDescription")
        font = QtGui.QFont()
        font.setPointSize(10)
        self.fedDescription.setFont(font)
        fedTitleLayout.addWidget(self.fedDescription)

        fedHeaderLayout.addWidget(self.wrapperFedTitle)
        fedLayout.addWidget(self.fedHeader)

        # Contenedor de acciones y specs
        mainContentWidget = QtWidgets.QWidget()
        mainContentLayout = QtWidgets.QHBoxLayout(mainContentWidget)
        mainContentLayout.setContentsMargins(15, 10, 15, 15)
        mainContentLayout.setSpacing(15)

        # Panel izquierdo: Acciones
        self.fedWrapperAction = QtWidgets.QWidget()
        self.fedWrapperAction.setObjectName("fedWrapperAction")
        actionLayout = QtWidgets.QVBoxLayout(self.fedWrapperAction)
        actionLayout.setContentsMargins(15, 15, 15, 15)
        actionLayout.setSpacing(15)

        # Selección de proceso y acción
        selectLayout = QtWidgets.QHBoxLayout()
        selectLayout.setSpacing(10)

        # Combo Proceso
        processWidget = QtWidgets.QWidget()
        processVLayout = QtWidgets.QVBoxLayout(processWidget)
        processVLayout.setContentsMargins(0, 0, 0, 0)
        processVLayout.setSpacing(8)

        self.labelProcess = QtWidgets.QLabel("Proceso")
        self.labelProcess.setObjectName("labelProcess")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        self.labelProcess.setFont(font)
        processVLayout.addWidget(self.labelProcess)

        self.boxProcess = QtWidgets.QComboBox()
        self.boxProcess.setObjectName("boxProcess")
        self.boxProcess.addItems(
            [
                "-- Selecciona un proceso --",
                "IVA Acreditable 100%",
                "Ventas al 0%",
                "Ventas al 16%",
            ]
        )
        self.boxProcess.setMinimumHeight(38)
        processVLayout.addWidget(self.boxProcess)
        selectLayout.addWidget(processWidget, 1)

        self.boxProcess.currentIndexChanged.connect(self.toggle_action_items)

        # Combo Acción
        actionComboWidget = QtWidgets.QWidget()
        actionVLayout = QtWidgets.QVBoxLayout(actionComboWidget)
        actionVLayout.setContentsMargins(0, 0, 0, 0)
        actionVLayout.setSpacing(8)

        self.labelAction = QtWidgets.QLabel("Acción")
        self.labelAction.setObjectName("labelAction")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        self.labelAction.setFont(font)
        actionVLayout.addWidget(self.labelAction)

        self.boxAction = QtWidgets.QComboBox()
        self.boxAction.setObjectName("boxAction")
        self.boxAction.addItems(["-- Selecciona una acción --"])
        self.boxAction.setMinimumHeight(38)
        actionVLayout.addWidget(self.boxAction)
        selectLayout.addWidget(actionComboWidget, 1)

        actionLayout.addLayout(selectLayout)

        # Botón Procesar
        self.buttonProcess = QtWidgets.QPushButton("▶ Procesar")
        self.buttonProcess.setObjectName("buttonProcess")
        self.buttonProcess.setMinimumHeight(40)
        self.buttonProcess.setMaximumWidth(150)
        processIcon = QtGui.QIcon()
        processIcon.addPixmap(QtGui.QPixmap(":/icons/play.svg"))
        self.buttonProcess.setIcon(processIcon)
        self.buttonProcess.setIconSize(QtCore.QSize(18, 18))
        self.buttonProcess.clicked.connect(self.ejecutar_proceso)
        actionLayout.addWidget(self.buttonProcess)

        # Consola
        consoleLabel = QtWidgets.QLabel("📋 Consola")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        consoleLabel.setFont(font)
        actionLayout.addWidget(consoleLabel)

        self.textConsole = QtWidgets.QTextEdit()
        self.textConsole.setObjectName("textConsole")
        self.textConsole.setReadOnly(True)
        self.textConsole.setMinimumHeight(250)
        self.textConsole.setMaximumHeight(280)
        actionLayout.addWidget(self.textConsole)

        mainContentLayout.addWidget(self.fedWrapperAction, 1)

        # Panel derecho: Especificaciones
        self.wWrapperSpecs = QtWidgets.QWidget()
        self.wWrapperSpecs.setObjectName("wWrapperSpecs")
        specsVLayout = QtWidgets.QVBoxLayout(self.wWrapperSpecs)
        specsVLayout.setContentsMargins(15, 15, 15, 15)
        specsVLayout.setSpacing(10)

        specsLabel = QtWidgets.QLabel("📊 Especificaciones")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        specsLabel.setFont(font)
        specsVLayout.addWidget(specsLabel)

        self.textSpecs = QtWidgets.QTextEdit()
        self.textSpecs.setObjectName("textSpecs")
        self.textSpecs.setReadOnly(True)
        self.textSpecs.setMinimumWidth(320)

        specsVLayout.addWidget(self.textSpecs)

        mainContentLayout.addWidget(self.wWrapperSpecs, 1)

        fedLayout.addWidget(mainContentWidget, 1)
        self.stackedWidget.addWidget(self.fedProcess)

        # ============ PÁGINA 3: PROCESOS ============
        self.processWork = QtWidgets.QWidget()
        self.processWork.setObjectName("processWork")
        processWorkLayout = QtWidgets.QVBoxLayout(self.processWork)
        processWorkLayout.setContentsMargins(0, 0, 0, 0)
        processWorkLayout.setSpacing(0)

        # Header Procesos
        self.processWorkHeader = QtWidgets.QWidget()
        self.processWorkHeader.setObjectName("processWorkHeader")
        self.processWorkHeader.setMaximumHeight(140)
        processHeaderLayout = QtWidgets.QVBoxLayout(self.processWorkHeader)
        processHeaderLayout.setContentsMargins(20, 15, 20, 15)

        self.wapperProcessWorkTitle = QtWidgets.QWidget()
        self.wapperProcessWorkTitle.setObjectName("wapperProcessWorkTitle")
        processWorkTitleLayout = QtWidgets.QVBoxLayout(self.wapperProcessWorkTitle)
        processWorkTitleLayout.setContentsMargins(0, 0, 0, 0)
        processWorkTitleLayout.setSpacing(5)

        self.processWorkTitle = QtWidgets.QLabel("⚙️ Procesos de Trabajo")
        self.processWorkTitle.setObjectName("processWorkTitle")
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        self.processWorkTitle.setFont(font)
        processWorkTitleLayout.addWidget(self.processWorkTitle)

        self.processWorkDescription = QtWidgets.QLabel(
            "Ejecuta procesos automáticos: unir archivos, generar flujogramas, comprimir, descargas masivas"
        )
        self.processWorkDescription.setObjectName("processWorkDescription")
        font = QtGui.QFont()
        font.setPointSize(10)
        self.processWorkDescription.setFont(font)
        processWorkTitleLayout.addWidget(self.processWorkDescription)

        processHeaderLayout.addWidget(self.wapperProcessWorkTitle)
        processWorkLayout.addWidget(self.processWorkHeader)

        # TabWidget
        self.tabWidgetProcess = QtWidgets.QTabWidget()
        self.tabWidgetProcess.setObjectName("tabWidgetProcess")
        self.tabWidgetProcess.setDocumentMode(False)
        tabLayout = QtWidgets.QVBoxLayout()
        tabLayout.setContentsMargins(15, 10, 15, 15)

        # ===== TAB 1: UNIR ARCHIVOS =====
        self.mergeFiles = QtWidgets.QWidget()
        self.mergeFiles.setObjectName("mergeFiles")
        mergeLayout = QtWidgets.QVBoxLayout(self.mergeFiles)
        mergeLayout.setContentsMargins(15, 15, 15, 15)
        mergeLayout.setSpacing(15)

        # Primer archivo
        firstFileLabel = QtWidgets.QLabel("📄 Selecciona el primer archivo")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        firstFileLabel.setFont(font)
        mergeLayout.addWidget(firstFileLabel)

        firstFileWidget = QtWidgets.QWidget()
        firstFileHLayout = QtWidgets.QHBoxLayout(firstFileWidget)
        firstFileHLayout.setContentsMargins(0, 0, 0, 0)
        firstFileHLayout.setSpacing(10)

        self.labelSearchFF = QtWidgets.QLabel("C:\\")
        self.labelSearchFF.setObjectName("labelSearchFF")
        self.labelSearchFF.setAlignment(QtCore.Qt.AlignVCenter)
        firstFileHLayout.addWidget(self.labelSearchFF, 1)

        self.buttonSearchFF = QtWidgets.QPushButton("📁 Seleccionar")
        self.buttonSearchFF.setObjectName("buttonSearchFF")
        self.buttonSearchFF.setMaximumWidth(140)
        self.buttonSearchFF.setMinimumHeight(38)
        firstFileHLayout.addWidget(self.buttonSearchFF)
        mergeLayout.addWidget(firstFileWidget)

        # Segundo archivo
        secondFileLabel = QtWidgets.QLabel("📄 Selecciona el segundo archivo")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        secondFileLabel.setFont(font)
        mergeLayout.addWidget(secondFileLabel)

        secondFileWidget = QtWidgets.QWidget()
        secondFileHLayout = QtWidgets.QHBoxLayout(secondFileWidget)
        secondFileHLayout.setContentsMargins(0, 0, 0, 0)
        secondFileHLayout.setSpacing(10)

        self.labelSearchSF = QtWidgets.QLabel("C:\\")
        self.labelSearchSF.setObjectName("labelSearchSF")
        self.labelSearchSF.setAlignment(QtCore.Qt.AlignVCenter)
        secondFileHLayout.addWidget(self.labelSearchSF, 1)

        self.buttonSearchSF = QtWidgets.QPushButton("📁 Seleccionar")
        self.buttonSearchSF.setObjectName("buttonSearchSF")
        self.buttonSearchSF.setMaximumWidth(140)
        self.buttonSearchSF.setMinimumHeight(38)
        secondFileHLayout.addWidget(self.buttonSearchSF)
        mergeLayout.addWidget(secondFileWidget)

        # Botón procesar
        self.bProcessMergeFile = QtWidgets.QPushButton("▶ Procesar Unión")
        self.bProcessMergeFile.setObjectName("bProcessMergeFile")
        self.bProcessMergeFile.setMinimumHeight(40)
        self.bProcessMergeFile.setMaximumWidth(160)
        self.bProcessMergeFile.clicked.connect(self.execute_work_process)
        mergeLayout.addWidget(self.bProcessMergeFile)
        mergeLayout.addStretch()

        self.tabWidgetProcess.addTab(self.mergeFiles, "📑 Unir Archivos")

        # ===== TAB 2: FLUJOGRAMA =====
        self.flowchart = QtWidgets.QWidget()
        self.flowchart.setObjectName("flowchart")
        flowLayout = QtWidgets.QVBoxLayout(self.flowchart)
        flowLayout.setContentsMargins(15, 15, 15, 15)
        flowLayout.setSpacing(15)

        flowcharLabel = QtWidgets.QLabel(
            "📊 Establece directorio para generar flujograma"
        )
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        flowcharLabel.setFont(font)
        flowLayout.addWidget(flowcharLabel)

        flowFileWidget = QtWidgets.QWidget()
        flowFileHLayout = QtWidgets.QHBoxLayout(flowFileWidget)
        flowFileHLayout.setContentsMargins(0, 0, 0, 0)
        flowFileHLayout.setSpacing(10)

        self.labelSearchF = QtWidgets.QLabel("C:\\")
        self.labelSearchF.setObjectName("labelSearchF")
        self.labelSearchF.setAlignment(QtCore.Qt.AlignVCenter)
        flowFileHLayout.addWidget(self.labelSearchF, 1)

        self.buttonSearchF = QtWidgets.QPushButton("📁 Seleccionar")
        self.buttonSearchF.setObjectName("buttonSearchF")
        self.buttonSearchF.setMaximumWidth(140)
        self.buttonSearchF.setMinimumHeight(38)
        flowFileHLayout.addWidget(self.buttonSearchF)
        flowLayout.addWidget(flowFileWidget)

        # Botón procesar
        self.bProcessFlowchart = QtWidgets.QPushButton("▶ Generar Flujograma")
        self.bProcessFlowchart.setObjectName("bProcessFlowchart")
        self.bProcessFlowchart.setMinimumHeight(40)
        self.bProcessFlowchart.setMaximumWidth(160)
        self.bProcessFlowchart.clicked.connect(self.execute_work_process)
        flowLayout.addWidget(self.bProcessFlowchart)
        flowLayout.addStretch()

        self.tabWidgetProcess.addTab(self.flowchart, "🔄 Flujograma")

        # ===== TAB 3: COMPRIMIR =====
        self.compressFile = QtWidgets.QWidget()
        self.compressFile.setObjectName("compressFile")
        compressLayout = QtWidgets.QVBoxLayout(self.compressFile)
        compressLayout.setContentsMargins(15, 15, 15, 15)
        compressLayout.setSpacing(15)

        compressLabel = QtWidgets.QLabel("📦 Establece directorio a comprimir")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        compressLabel.setFont(font)
        compressLayout.addWidget(compressLabel)

        compressFileWidget = QtWidgets.QWidget()
        compressFileHLayout = QtWidgets.QHBoxLayout(compressFileWidget)
        compressFileHLayout.setContentsMargins(0, 0, 0, 0)
        compressFileHLayout.setSpacing(10)

        self.labelSearchCompressPath = QtWidgets.QLabel("C:\\")
        self.labelSearchCompressPath.setObjectName("labelSearchCompressPath")
        self.labelSearchCompressPath.setAlignment(QtCore.Qt.AlignVCenter)
        compressFileHLayout.addWidget(self.labelSearchCompressPath, 1)

        self.buttonSearchCompressPath = QtWidgets.QPushButton("📁 Seleccionar")
        self.buttonSearchCompressPath.setObjectName("buttonSearchCompressPath")
        self.buttonSearchCompressPath.setMaximumWidth(140)
        self.buttonSearchCompressPath.setMinimumHeight(38)
        compressFileHLayout.addWidget(self.buttonSearchCompressPath)
        compressLayout.addWidget(compressFileWidget)

        # Botón procesar
        self.buttonProcessCP = QtWidgets.QPushButton("▶ Crear Comprimido (.7z)")
        self.buttonProcessCP.setObjectName("buttonProcessCP")
        self.buttonProcessCP.setMinimumHeight(40)
        self.buttonProcessCP.setMaximumWidth(200)
        compressLayout.addWidget(self.buttonProcessCP)
        compressLayout.addStretch()

        # self.tabWidgetProcess.addTab(self.compressFile, "📦 Comprimir")

        # ===== TAB 4: DESCARGA MASIVA =====
        self.bulkDownload = QtWidgets.QWidget()
        self.bulkDownload.setObjectName("bulkDownload")
        bulkLayout = QtWidgets.QVBoxLayout(self.bulkDownload)
        bulkLayout.setContentsMargins(15, 15, 15, 15)
        bulkLayout.setSpacing(15)

        # PDF
        pdfLabel = QtWidgets.QLabel("📑 Selecciona el PDF (pólizas)")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        pdfLabel.setFont(font)
        bulkLayout.addWidget(pdfLabel)

        pdfWidget = QtWidgets.QWidget()
        pdfHLayout = QtWidgets.QHBoxLayout(pdfWidget)
        pdfHLayout.setContentsMargins(0, 0, 0, 0)
        pdfHLayout.setSpacing(10)

        self.labelPdfBD = QtWidgets.QLabel("C:\\")
        self.labelPdfBD.setObjectName("labelPdfBD")
        self.labelPdfBD.setAlignment(QtCore.Qt.AlignVCenter)
        pdfHLayout.addWidget(self.labelPdfBD, 1)

        self.buttonPdfBD = QtWidgets.QPushButton("📁 Seleccionar PDF")
        self.buttonPdfBD.setObjectName("buttonPdfBD")
        self.buttonPdfBD.setMaximumWidth(140)
        self.buttonPdfBD.setMinimumHeight(38)
        pdfHLayout.addWidget(self.buttonPdfBD)
        bulkLayout.addWidget(pdfWidget)

        # XLSX
        xlsxLabel = QtWidgets.QLabel("📊 Selecciona XLSX (referencias)")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        xlsxLabel.setFont(font)
        bulkLayout.addWidget(xlsxLabel)

        xlsxWidget = QtWidgets.QWidget()
        xlsxHLayout = QtWidgets.QHBoxLayout(xlsxWidget)
        xlsxHLayout.setContentsMargins(0, 0, 0, 0)
        xlsxHLayout.setSpacing(10)

        self.labelXlsxBD = QtWidgets.QLabel("C:\\")
        self.labelXlsxBD.setObjectName("labelXlsxBD")
        self.labelXlsxBD.setAlignment(QtCore.Qt.AlignVCenter)
        xlsxHLayout.addWidget(self.labelXlsxBD, 1)

        self.buttonXlsxBD = QtWidgets.QPushButton("📁 Seleccionar XLSX")
        self.buttonXlsxBD.setObjectName("buttonXlsxBD")
        self.buttonXlsxBD.setMaximumWidth(140)
        self.buttonXlsxBD.setMinimumHeight(38)
        xlsxHLayout.addWidget(self.buttonXlsxBD)
        bulkLayout.addWidget(xlsxWidget)

        # Botón procesar
        self.bProcessBulkDownload = QtWidgets.QPushButton("▶ Descargar en Masa")
        self.bProcessBulkDownload.setObjectName("bProcessBulkDownload")
        self.bProcessBulkDownload.setMinimumHeight(40)
        self.bProcessBulkDownload.setMaximumWidth(180)
        bulkLayout.addWidget(self.bProcessBulkDownload)
        bulkLayout.addStretch()

        # self.tabWidgetProcess.addTab(self.bulkDownload, "⬇️ Descarga Masiva")

        # ===== TAB 5: ENCABEZADO =====
        self.setHeading = QtWidgets.QWidget()
        self.setHeading.setObjectName("setHeading")
        headingLayout = QtWidgets.QVBoxLayout(self.setHeading)
        headingLayout.setContentsMargins(15, 15, 15, 15)
        headingLayout.setSpacing(15)

        headingLabel = QtWidgets.QLabel("📝 Selecciona archivo para añadir encabezado")
        font = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(11)
        headingLabel.setFont(font)
        headingLayout.addWidget(headingLabel)

        headingFileWidget = QtWidgets.QWidget()
        headingFileHLayout = QtWidgets.QHBoxLayout(headingFileWidget)
        headingFileHLayout.setContentsMargins(0, 0, 0, 0)
        headingFileHLayout.setSpacing(10)

        self.labelSearchSh = QtWidgets.QLabel("C:\\")
        self.labelSearchSh.setObjectName("labelSearchSh")
        self.labelSearchSh.setAlignment(QtCore.Qt.AlignVCenter)
        headingFileHLayout.addWidget(self.labelSearchSh, 1)

        self.buttonSearchSH = QtWidgets.QPushButton("📁 Seleccionar")
        self.buttonSearchSH.setObjectName("buttonSearchSH")
        self.buttonSearchSH.setMaximumWidth(140)
        self.buttonSearchSH.setMinimumHeight(38)
        headingFileHLayout.addWidget(self.buttonSearchSH)
        headingLayout.addWidget(headingFileWidget)

        # Botón procesar
        self.bProcessSetHeading = QtWidgets.QPushButton("▶ Añadir Encabezado")
        self.bProcessSetHeading.setObjectName("bProcessSetHeading")
        self.bProcessSetHeading.setMinimumHeight(40)
        self.bProcessSetHeading.setMaximumWidth(160)
        headingLayout.addWidget(self.bProcessSetHeading)
        headingLayout.addStretch()

        # self.tabWidgetProcess.addTab(self.setHeading, "📝 Encabezado")

        processWorkLayout.addWidget(self.tabWidgetProcess, 1)
        self.stackedWidget.addWidget(self.processWork)

        wrapperLayout.addWidget(self.wMain, 1)
        fedApp.setCentralWidget(self.wWrapper)

        # Conexiones de StackWidget
        self.rootButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.fedButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(1))
        self.processButton.clicked.connect(
            lambda: self.stackedWidget.setCurrentIndex(2)
        )
        self.exitButton.clicked.connect(lambda: fedApp.close())

        # ✅ CONEXIÓN DEL BOTÓN "ESTABLECER" PARA CARGAR DIRECTORIO
        self.buttonPath.clicked.connect(lambda: self.open_directory_dialog(fedApp))

        self.buttonSearchFF.clicked.connect(
            lambda: self.cargar_archivo_xlsx("merge_ff")
        )
        self.buttonSearchSF.clicked.connect(
            lambda: self.cargar_archivo_xlsx("merge_sf")
        )
        self.buttonSearchF.clicked.connect(
            lambda: self.open_directory_dialog(fedApp, False, "flowchart")
        )
        self.buttonSearchCompressPath.clicked.connect(
            lambda: self.open_directory_dialog(fedApp, False, "compress")
        )

        self.retranslateUi(fedApp)
        self.stackedWidget.setCurrentIndex(
            0
        )  # Mostrar página de directorio por defecto
        self.tabWidgetProcess.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(fedApp)

    def toggle_action_items(self, idx):
        mapa_acciones = {
            "1": ["Procesar", "Formato", "Referencia Bancos", "Encabezado"],
            "2": ["Procesar", "Subrayado Auxiliar", "Referencia Bancos", "Encabezado"],
            "3": ["Procesar", "Subrayado Auxiliar", "Referencia Bancos", "Encabezado"],
        }

        """Carga acciones en el combo box con un ítem placeholder."""
        self.boxAction.clear()
        self.boxAction.addItem("-- Selecciona una acción --")

        idx_seleccion = self.boxProcess.currentIndex()
        acciones = mapa_acciones.get(str(idx_seleccion), [])

        if acciones:
            self.dynami_feature_content(str(idx_seleccion))
            self.boxAction.addItems(acciones)

        self.boxAction.setCurrentIndex(0)

    def text_console_log(self, message, level="INFO"):
        # 1. Obtener la hora actual para un look profesional
        now = datetime.now().strftime("%H:%M:%S")

        # 2. Definir colores según el nivel (Moderno)
        color = "#ABB2BF"  # Gris claro por defecto
        if level == "SUCCESS":
            color = "#98C379"  # Verde
        elif level == "ERROR":
            color = "#E06C75"  # Rojo
        elif level == "PROCESS":
            color = "#61AFEF"  # Azul

        # 3. Crear el formato HTML para el mensaje
        log_entry = f'<span style="color: #5C6370;">[{now}]</span> <span style="color: {color};">{message}</span>'

        # 4. Agregar a la consola sin borrar lo anterior
        self.textConsole.append(log_entry)

        # Fuerza a la aplicación a actualizar la interfaz antes de seguir
        QtWidgets.QApplication.processEvents()

        # 5. Auto-scroll al final
        self.textConsole.verticalScrollBar().setValue(
            self.textConsole.verticalScrollBar().maximum()
        )

    def bloquear_pantalla(self):
        # 1. Aplicar efecto de Blur al contenedor principal
        self.blur_effect = QtWidgets.QGraphicsBlurEffect()
        self.blur_effect.setBlurRadius(10)  # Intensidad del blur
        self.wWrapper.setGraphicsEffect(self.blur_effect)

        # 2. Crear la capa de bloqueo (Overlay)
        self.overlay = QtWidgets.QWidget(self.main_window)
        self.overlay.setGeometry(self.main_window.rect())
        # Color negro semitransparente
        self.overlay.setStyleSheet("background-color: rgba(0, 0, 0, 80);")

        # 3. Añadir un mensaje o spinner en el centro
        layout = QtWidgets.QVBoxLayout(self.overlay)
        self.loading_label = QtWidgets.QLabel(
            "⌛ Procesando...\nPor favor espere", self.overlay
        )
        self.loading_label.setAlignment(QtCore.Qt.AlignCenter)
        self.loading_label.setStyleSheet(
            "color: white; font-size: 18px; font-weight: bold; background: none;"
        )
        layout.addWidget(self.loading_label)

        self.overlay.show()
        # Forzar a la interfaz a procesar el dibujo antes de seguir
        QtWidgets.QApplication.processEvents()

    def desbloquear_pantalla(self):
        # 1. Quitar el efecto de blur
        if hasattr(self, "blur_effect"):
            self.wWrapper.setGraphicsEffect(None)

        # 2. Eliminar la capa de bloqueo
        if hasattr(self, "overlay"):
            self.overlay.deleteLater()

    def obtener_proceso_fed(self):
        indice_seleccionado = self.boxProcess.currentIndex()
        texto_seleccionado = MAP_SUFIX_IDX.get(str(indice_seleccionado), "")

        # Validar que no se haya seleccionado la opción por defecto (índice 0)
        if indice_seleccionado == 0:
            print("Por favor, selecciona un proceso válido.")
            return 0, ""

        return indice_seleccionado, texto_seleccionado

    def obtener_accion_fed(self):
        texto_seleccionado = self.boxAction.currentText()
        indice_seleccionado = self.boxAction.currentIndex()

        # Validar que no se haya seleccionado la opción por defecto (índice 0)
        if indice_seleccionado == 0:
            print("Por favor, selecciona una accion válida.")
            return 0, ""

        return indice_seleccionado, texto_seleccionado

    def open_xlsx_dialog(self, parent):
        """
        Abre un diálogo para seleccionar un archivo XLSX

        Returns:
            str: Ruta del archivo, o None si se cancela
        """
        file_filter = "Excel Files (*.xlsx);;All Files (*.*)"

        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            parent, "Seleccionar Archivo Excel", "", file_filter
        )

        if not file_path:
            return None

        # Validar extensión
        if not file_path.lower().endswith(".xlsx"):
            QtWidgets.QMessageBox.warning(
                parent, "Archivo Inválido", "Solo se aceptan archivos .xlsx"
            )
            return None

        return file_path

    def open_directory_dialog(self, parent, show_tree=True, current_tab=""):
        """
        Abre un diálogo para seleccionar un directorio y lo carga en el TreeWidget
        """
        directory = QtWidgets.QFileDialog.getExistingDirectory(
            parent,
            "Seleccionar Directorio de Trabajo",
            "",
            QtWidgets.QFileDialog.ShowDirsOnly,
        )

        if directory:
            # Cargar el directorio en el TreeWidget
            # self.buttonProcess.clicked.connect(lambda: self.procesar_boton(directory))
            self.first_directory = directory
            if show_tree and not current_tab:
                self.dir_manager.load_directory(directory, self.path)

            if not show_tree and current_tab:
                self.work_process(current_tab)

    def cargar_archivo_xlsx(self, button_type):
        """Abre el diálogo y carga el archivo en la variable"""
        archivo_xlsx = self.open_xlsx_dialog(fedApp)

        if archivo_xlsx:
            print(f"Archivo cargado: {archivo_xlsx}")
            # Ahora puedes usar self.archivo_xlsx en cualquier otro método
            if button_type == "merge_ff":
                self.first_directory = archivo_xlsx
            if button_type == "merge_sf":
                self.second_directory = archivo_xlsx
            self.work_process(button_type)

    def work_process(self, current_tab):
        current_index_tab = self.tabWidgetProcess.currentIndex()

        if current_index_tab == 0:
            if current_tab == "merge_ff":
                self.labelSearchFF.setText(self.first_directory)
            if current_tab == "merge_sf":
                self.labelSearchSF.setText(self.second_directory)

        if current_index_tab == 1:
            self.labelSearchF.setText(self.first_directory)

        if current_index_tab == 2:
            self.labelSearchCompressPath.setText(self.first_directory)

    def retranslateUi(self, fedApp):
        _translate = QtCore.QCoreApplication.translate
        fedApp.setWindowTitle(_translate("fedApp", "FED - Sistema de Devoluciones"))

    def message_box(self, title, mssg, err_type=""):
        # Mapeo de strings a Iconos reales de Qt
        icons = {
            "err": QtWidgets.QMessageBox.Critical,
            "inf": QtWidgets.QMessageBox.Information,
            "wrg": QtWidgets.QMessageBox.Warning,
            "ok": QtWidgets.QMessageBox.Information,  # 'ok' no es un icono, suele ser 'inf'
        }

        msg = QtWidgets.QMessageBox(self.main_window)
        msg.setWindowTitle(title)
        msg.setText(mssg)

        # Obtener el icono correcto o NoIcon si no existe el tipo
        msg.setIcon(icons.get(err_type, QtWidgets.QMessageBox.NoIcon))

        # Importante: se usa .exec_() para mostrarlo.
        msg.exec_()

    def ejecutar_proceso(self):
        # Si la variable está vacía o es None, detonamos el mensaje
        if not self.first_directory:
            self.message_box(
                "Algo ha pasado!",
                "No se ha seleccionado un directorio de trabajo",
                "wrg",
            )
            return  # Salimos de la función para no procesar nada

        # Si pasa la validación, llamamos a la lógica real
        self.procesar_boton(self.first_directory)

    def procesar_boton(self, path):
        idx_process, vle_process = self.obtener_proceso_fed()
        idx_action, vle_action = self.obtener_accion_fed()

        if idx_process == 0 or idx_action == 0:
            return

        self.bloquear_pantalla()

        try:
            #############################################################################
            # Procesos genericos tanto para IVA
            #############################################################################
            if vle_process.casefold() == "iva".casefold():
                if vle_action.casefold() == "Procesar".casefold():
                    self.message_box("IVA", "Se inicio proceso de IVA...")
                    self.text_console_log("Se inicio proceso de IVA...", "INFO")
                    main_iva_process(self, path)

                if vle_action.casefold() == "Formato".casefold():
                    self.text_console_log("Formato final al archivo...", "PROCESS")
                    file_name = f"{path}/cedula_iva_acreditable_100.xlsx"
                    main_iva_format_process(self, path, file_name)

            #############################################################################
            # Procesos genericos tanto para V0
            #############################################################################
            if vle_process.casefold() == "v0".casefold():
                if vle_action.casefold() == "Procesar".casefold():
                    self.text_console_log("Iniciando proceso V0...", "PROCESS")
                    main_v0_process(self, path)

            #############################################################################
            # Procesos genericos tanto para V16
            #############################################################################
            if vle_process.casefold() == "v16".casefold():
                if vle_action.casefold() == "Procesar".casefold():
                    self.text_console_log("Iniciando proceso V16...", "PROCESS")
                    main_v16_process(path)

            #############################################################################
            # Procesos genericos tanto para IVA como V0 y V16
            #############################################################################
            if vle_action.casefold() == "Subrayado Auxiliar".casefold():
                self.text_console_log(
                    "Iniciando proceso Subrayado Auxiliar...", "PROCESS"
                )
                main_underline_aux_process(self, path, vle_process.casefold())

            if vle_action.casefold() == "Referencia Bancos".casefold():
                self.text_console_log(
                    "Procesando referencia en estados de cuenta...", "PROCESS"
                )
                file_name = (
                    "cedula_iva_acreditable_100"
                    if vle_process.casefold() == "iva".casefold()
                    else vle_process.casefold()
                )
                main_underline_bank_process(
                    self, vle_process.casefold(), path, file_name
                )
                print("Proceso de referencia")

            if vle_action.casefold() == "Encabezado".casefold():
                self.text_console_log(
                    "Procesando referencia en estados de cuenta...", "PROCESS"
                )
                file_path = (
                    f"{path}/cedula_iva_acreditable_100.xlsx"
                    if vle_process.casefold() == "iva".casefold()
                    else f"{path}/{vle_process}.xlsx"
                )
                logo_path = resource_path(":/icons/logo.png")
                main_header_process(file_path, vle_process, ":/icons/logo.png")
                print("Proceso de referencia")

        except Exception as e:
            self.message_box(
                "Error Crítico", f"Ocurrió un error inesperado: {e}", "err"
            )

        finally:
            # --- DESBLOQUEAR ---
            # Se pone en 'finally' para que si el proceso falla (da error),
            # la pantalla no se quede bloqueada para siempre.
            self.desbloquear_pantalla()

    def execute_work_process(self):
        idx_tab_process = self.tabWidgetProcess.currentIndex()

        logic_process = (
            not self.first_directory
            if idx_tab_process == 1
            else not self.first_directory or not self.second_directory
        )

        # Si la variable está vacía o es None, detonamos el mensaje
        if logic_process:
            archivo_faltante = (
                "Primer Archivo" if not self.first_directory else "Segundo Archivo"
            )
            self.message_box(
                "Algo ha pasado!",
                f"No se ha seleccionado un directorio para el proceso. Falta {archivo_faltante}",
                "wrg",
            )
            return  # Salimos de la función para no procesar nada

        # Si pasa la validación, llamamos a la lógica real
        self.logic_work_process(self.first_directory, idx_tab_process)

    def logic_work_process(self, path, idx_tab_process):
        self.bloquear_pantalla()

        try:
            if idx_tab_process == 0:
                self.message_box(
                    "Procesos de Trabajo", "Se ha iniciado el proceso Unir Archivos..."
                )
                self.text_console_log(
                    "Se ha iniciado el proceso Unir Archivos...", "INFO"
                )
                main_merge_process(self, self.first_directory, self.second_directory)

            if idx_tab_process == 1:
                self.message_box(
                    "Procesos de Trabajo", "Se ha iniciado el proceso Flujograma..."
                )
                self.text_console_log("Se ha iniciado el proceso Flujograma...", "INFO")
                main_flujograma_process(self, path)

            if idx_tab_process == 2:
                print("ejecutando")

            if idx_tab_process == 3:
                print("ejecutando")

            if idx_tab_process == 4:
                print("ejecutando")

        except Exception as e:
            self.message_box(
                "Error Crítico", f"Ocurrió un error inesperado: {e}", "err"
            )
        finally:
            self.desbloquear_pantalla()

    def dynami_feature_content(self, suff):
        features = {
            "1": [
                ("diot.xlsx", "sheets.png"),
                ("egresos.xlsx", "sheets.png"),
                ("cambio_obligaciones.csv", "sheets.png"),
                ("HSBC8881.pdf", "pdf.png"),
                ("BASE2018.pdg", "pdf.png"),
            ],
            "2": [
                ("mxn.xlsx", "sheets.png"),
                ("usd.xlsx", "sheets.png"),
                ("cambio_obligaciones.csv", "sheets.png"),
                ("aux.pdf", "pdf.png"),
                ("HSBC5430.pdf", "pdf.png"),
            ],
            "3": [
                ("rete.xlsx", "sheets.png"),
                ("tras.xlsx", "sheets.png"),
                ("cambio_obligaciones.csv", "sheets.png"),
                ("ventas.xlsx", "sheets.png"),
                ("aux.pdf", "pdf.png"),
                ("HSBC5430.pdf", "pdf.png"),
                ("BASE2018.pdf", "pdf.png"),
                ("HSBC8881.pdf", "pdf.png"),
            ],
        }

        result = {
            "1": [
                ("cedula_iva_acreditable_100.xlsx", "sheets.png"),
                ("HSBC8881_underlined.pdf", "pdf.png"),
                ("BASE2018_underlined.pdg", "pdf.png"),
            ],
            "2": [
                ("v0.xlsx", "sheets.png"),
                ("aux_underlined.pdf", "pdf.png"),
                ("HSBC5430_underlined.pdf", "pdf.png"),
            ],
            "3": [
                ("v16.xlsx", "sheets.png"),
                ("aux_underlined.pdf", "pdf.png"),
                ("HSBC5430_underlined.pdf", "pdf.png"),
                ("BASE2018_underlined.pdf", "pdf.png"),
                ("HSBC8881_underlined.pdf", "pdf.png"),
            ],
        }

        item_template = (
            '<p style="margin-top:0px; margin-bottom:0px; margin-left:0px; '
            "margin-right:0px; -qt-block-indent:0; text-indent:0px; "
            'background-color:#f8f9fa;">'
            "<span style=\"font-family:'Segoe UI','Arial','sans-serif'; "
            'font-size:10pt; color:#212529; background-color:#f8f9fa;">    </span>'
            '<img src=":/icons/{icon}" alt="Archivo" width="20" height="20" '
            'style="vertical-align: middle;" />'
            "<span style=\"font-family:'Segoe UI','Arial','sans-serif'; "
            'font-size:10pt; color:#212529;">    </span>'
            "<span style=\"font-family:'Segoe UI','Arial','sans-serif'; "
            'font-size:10pt; font-weight:500; color:#495057;">{file}</span>'
            "</p>"
        )

        items = features.get(suff.casefold(), [])
        items_result = result.get(suff.casefold(), [])

        content = "".join(
            item_template.format(file=escape(file_name), icon=escape(icon_name))
            for file_name, icon_name in items
        )

        content_result = "".join(
            item_template.format(file=escape(file_name), icon=escape(icon_name))
            for file_name, icon_name in items_result
        )

        sufix = MAP_SUFIX_IDX.get(str(suff), "").upper()

        html = f"""<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">
    <html>
    <head>
        <meta name="qrichtext" content="1" />
        <meta charset="utf-8" />
        <style type="text/css">
            p, li {{ white-space: pre-wrap; }}
            hr {{ height: 1px; border-width: 0; }}
            li.unchecked::marker {{ content: "\\2610"; }}
            li.checked::marker {{ content: "\\2612"; }}
        </style>
    </head>
    <body style="font-family:'Sans Serif'; font-size:9pt; font-weight:400; font-style:normal;">
        <p style="margin-top:0px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">
            <span style="font-size:12pt;"></span>
            <span style="font-size:12pt; font-weight:500;">Directorio para el proceso</span>
            <span style="font-size:12pt;"></span> <span style="font-size:12pt; font-weight:500;">{escape(sufix)}</span>
        </p>
        {content}
        <p style="margin-top:0px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">
            <span style="font-size:12pt;"></span>
            <span style="font-size:12pt; font-weight:500;">Este proceso te devolverá</span>
        </p>
        {content_result}
    </body>
    </html>"""

        _translate = QtCore.QCoreApplication.translate
        self.textSpecs.setHtml(_translate("fedApp", html))

    def mousePressEvent(self, event):
        # Detectar si se presiona Meta (o Ctrl en macOS por ciertos comportamientos)
        modifiers = event.modifiers()

        # En macOS, Meta es la tecla "Command", pero a veces hay que ajustar
        is_meta_pressed = (
            modifiers & QtCore.Qt.MetaModifier  # Meta/Command
            or modifiers
            & QtCore.Qt.ControlModifier  # A veces en macOS se usa Ctrl como fallback
        )

        if event.button() == QtCore.Qt.RightButton and is_meta_pressed:
            self.dragging = True
            self.offset = event.globalPos() - self.pos()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.dragging:
            # Actualizamos la posición de la ventana
            new_pos = event.globalPos() - self.offset
            self.move(new_pos)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.RightButton:
            self.dragging = False
        else:
            super().mouseReleaseEvent(event)


def applyModernStyle(app):
    """
    Aplica un estilo moderno y profesional sin modificar el diseño original
    """
    app.setStyleSheet("""
        /* Fuente base */
        QWidget {
            font-family: 'Segoe UI', 'Roboto', 'Helvetica', sans-serif;
            font-size: 10pt;
            background-color: #f5f7fa;
            color: #1f2937;
        }

        /* Botones generales */
        QPushButton {
            background-color: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 8px 12px;
            color: #1f2937;
            font-weight: 500;
            font-size: 10pt;
        }

        QPushButton:hover {
            background-color: #f8fafc;
            border-color: #94a3b8;
        }

        QPushButton:pressed {
            background-color: #e2e8f0;
        }

        /* Botón Salir */
        #exitButton {
            background-color: #fee2e2;
            color: #b91c1c;
            border-color: #fca5a5;
        }

        #exitButton:hover {
            background-color: #fecaca;
        }

        /* Sidebar */
        #wAsider {
            background-color: #ffffff;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            padding: 10px;
        }

        /* Títulos */
        QLabel#rootTitle,
        QLabel#fedTitle,
        QLabel#processWorkTitle {
            color: #1e40af;
            font-weight: 600;
        }

        QLabel#rootDescription,
        QLabel#fedDescription,
        QLabel#processWorkDescription {
            color: #64748b;
            font-weight: 400;
        }

        /* Contenedores principales */
        #wWrapperDirectory,
        #wWrapperSpecs,
        #fedWrapperAction {
            background-color: #ffffff;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
        }

        /* TreeWidget */
        QTreeWidget {
            background-color: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 8px;
            alternate-background-color: #f8fafc;
        }

        QTreeWidget::item {
            padding: 6px;
            margin: 2px 0;
        }

        QTreeWidget::item:selected {
            background-color: #dbeafe;
            color: #1e40af;
            border-radius: 6px;
        }

        QTreeWidget::item:hover {
            background-color: #f0f9ff;
        }

        /* ComboBox */
        QComboBox {
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 8px 12px;
            background-color: #ffffff;
            selection-background-color: #3b82f6;
        }

        QComboBox:hover {
            border-color: #94a3b8;
        }

        QComboBox::drop-down {
            width: 28px;
            border-left: 1px solid #cbd5e1;
            background-color: #f8fafc;
        }

        QComboBox::down-arrow {
            width: 14px;
            height: 14px;
        }

        /* TabWidget */
        QTabWidget::pane {
            border: none;
        }

        QTabBar::tab {
            background-color: #f8fafc;
            border: 1px solid #cbd5e1;
            border-bottom: none;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            padding: 10px 16px;
            margin-right: 2px;
            color: #64748b;
            font-weight: 500;
        }

        QTabBar::tab:selected {
            background-color: #ffffff;
            border-top: 2px solid #3b82f6;
            color: #1f2937;
        }

        QTabBar::tab:hover:!selected {
            background-color: #ffffff;
        }

        /* Botones de acción (azules) */
        QPushButton#buttonProcess,
        QPushButton#bProcessMergeFile,
        QPushButton#bProcessFlowchart,
        QPushButton#buttonProcessCP,
        QPushButton#bProcessBulkDownload,
        QPushButton#bProcessSetHeading,
        QPushButton#buttonPath,
        QPushButton#buttonSearchFF,
        QPushButton#buttonSearchSF,
        QPushButton#buttonSearchF,
        QPushButton#buttonSearchCompressPath,
        QPushButton#buttonPdfBD,
        QPushButton#buttonXlsxBD,
        QPushButton#buttonSearchSH {
            background-color: #3b82f6;
            color: white;
            border: none;
            font-weight: 600;
        }

        QPushButton#buttonProcess:hover,
        QPushButton#bProcessMergeFile:hover,
        QPushButton#bProcessFlowchart:hover,
        QPushButton#buttonProcessCP:hover,
        QPushButton#bProcessBulkDownload:hover,
        QPushButton#bProcessSetHeading:hover,
        QPushButton#buttonPath:hover,
        QPushButton[objectName^="buttonSearch"]:hover {
            background-color: #2563eb;
        }

        QPushButton#buttonProcess:pressed,
        QPushButton#bProcessMergeFile:pressed,
        QPushButton#bProcessFlowchart:pressed,
        QPushButton#buttonProcessCP:pressed,
        QPushButton#bProcessBulkDownload:pressed,
        QPushButton#bProcessSetHeading:pressed,
        QPushButton#buttonPath:pressed,
        QPushButton[objectName^="buttonSearch"]:pressed {
            background-color: #1d4ed8;
        }

        /* Labels de ruta */
        QLabel#path,
        QLabel#labelSearchFF,
        QLabel#labelSearchSF,
        QLabel#labelSearchF,
        QLabel#labelSearchCompressPath,
        QLabel#labelPdfBD,
        QLabel#labelXlsxBD,
        QLabel#labelSearchSh {
            background-color: #f9fafb;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 10px;
            color: #64748b;
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 10pt;
        }

        /* TextEdit */
        QTextEdit {
            background-color: #f9fafb;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 10px;
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 10pt;
        }

        /* Scroll bar */
        QScrollBar:vertical {
            background-color: #f8fafc;
            width: 12px;
            border-radius: 6px;
        }

        QScrollBar::handle:vertical {
            background-color: #cbd5e1;
            border-radius: 6px;
        }

        QScrollBar::handle:vertical:hover {
            background-color: #94a3b8;
        }

        /* Header backgrounds */
        #rootHeader,
        #fedHeader,
        #processWorkHeader {
            background-color: #ffffff;
            border-bottom: 1px solid #e2e8f0;
        }

        /* QTreeView styles */
        QTreeView {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            outline: none;
            font-size: 11pt;
        }

        QTreeView::item {
            padding: 8px; /* Más espacio vertical entre archivos */
            border-radius: 6px;
            margin: 2px 5px;
            color: #334155;
        }

        QTreeView::item:hover {
            background-color: #f1f5f9;
            color: #1e40af;
        }

        QTreeView::item:selected {
            background-color: #e0e7ff;
            color: #4338ca;
            font-weight: bold;
            border-left: 4px solid #4338ca; /* Indicador lateral de selección */
        }

        /* Ajuste para el icono */
        QTreeView::icon {
            margin-right: 10px;
        }

        #textConsole {
            background-color: #1e1e1e;
            color: #dcdcdc;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 12px;
            border: 1px solid #3e3e42;
            border-radius: 8px;
            padding: 10px;
        }

        /* Esto aplica solo a la barra de desplazamiento dentro de textConsole */
        #textConsole QScrollBar:vertical {
            border: none;
            background: #2d2d30;
            width: 10px;
            margin: 0px;
        }

        #textConsole QScrollBar::handle:vertical {
            background: #3e3e42;
            min-height: 20px;
            border-radius: 5px;
        }

        #textConsole QScrollBar::add-line:vertical,
        #textConsole QScrollBar::sub-line:vertical {
            background: none;
            height: 0px;
        }
    """)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    fedApp = QtWidgets.QMainWindow()
    ui = Ui_fedApp()
    ui.setupUi(fedApp)

    # --- AÑADE ESTA LÍNEA AQUÍ ---
    # fedApp.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.WindowTitleHint | QtCore.Qt.CustomizeWindowHint)
    # fedApp.setWindowFlags(QtCore.Qt.FramelessWindowHint)
    # -----------------------------

    # ASIGNAR ICONO A LA VENTANA (Usando tu recurso compilado)
    icon_path = ":/icons/favicon.png"  # Asegúrate que el nombre coincida con tu .qrc
    fedApp.setWindowIcon(QtGui.QIcon(icon_path))

    # Aplicar estilo moderno
    applyModernStyle(fedApp)

    fedApp.show()
    sys.exit(app.exec_())
