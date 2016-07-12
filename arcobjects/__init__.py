#Most of this code adopted and modified from:
#
#    https://bitbucket.org/maphew/canvec/src/eaf2678de06f/Canvec/Scripts/parco.py
#
# For Bolton & Menk, Inc. use
#
# Snippets.py
# ************************************************
# Updated for ArcGIS 10.2
# ************************************************
# Requires installation of the comtypes package
# Available at: http://sourceforge.net/projects/comtypes/
# Once comtypes is installed, the following modifications
# need to be made for compatibility with ArcGIS 10.2:
# 1) Delete automation.pyc, automation.pyo, safearray.pyc, safearray.pyo
# 2) Edit automation.py
# 3) Add the following entry to the _ctype_to_vartype dictionary (line 794):
#    POINTER(BSTR): VT_BYREF|VT_BSTR,
# ************************************************
# Most of this code adopted and modified from:
#   https://bitbucket.org/maphew/canvec/src/eaf2678de06f/Canvec/Scripts/parco.py

# make sure it is being ran in 32 bit
import struct
import comtypes
from comtypes.client import GetModule, CreateObject
if struct.calcsize('P') * 8 != 32:
    raise ImportError('Must use ArcObjects in 32 bit!')
import os
import glob
import sys
import datetime

ACCESS_MODE = {
                0: 'unknown',
                2: 'write',
                4: 'read only',
                6: 'read/write'
              }

def load_mod(filterer=None):
    '''loads a list of all modules (*.olb files)

    Optional:
    filterer -- wildcard to search by.  If want to find all modules that start
        with "G", simply use "g" for the filterer
    '''
    olb = '*.olb'
    if filterer:
        olb = 'esri{0}*.olb'.format(filterer)
    mods = glob.glob(os.path.join(GetLibPath(), olb))
    if mods:
        all_mods = dict((i,os.path.basename(m)) for i,m in enumerate(mods))
        for k,v in sorted(all_mods.iteritems()):
            print '{0}: {1}'.format(k,v)
        return getModule(all_mods[input('\nChoose number for module\n')])
    print 'Did not find any modules!'
    return None

def load_all():
    '''loads all object libraries'''
    mods = glob.glob(os.path.join(GetLibPath(), '*.olb'))
    for mod in mods:
        GetModule(mod)
    return

def InstallInfo():
    """Gets ArcGIS Install Info"""
    # Get ArcObjects version
    g = comtypes.GUID("{6FCCEDE0-179D-4D12-B586-58C88D26CA78}")
    GetModule((g, 1, 0))
    import comtypes.gen.ArcGISVersionLib as esriVersion
    pVM = NewObj(esriVersion.VersionManager, esriVersion.IArcGISVersion)
    return pVM.GetVersions().Next()

def GetLibPath():
    '''Reference to com directory which houses ArcObjects
    Ojbect Libraries (*.OLB)'''
    return os.path.join(InstallInfo()[2], 'com')

def GetVersion():
    """returns ArcGIS Version"""
    return InstallInfo()[1]

def getModule(sModuleName):
    ''' loads the object library by name'''
    olb = os.path.abspath(os.path.join(GetLibPath(), sModuleName))
    return GetModule(olb)

def GetStandaloneModules():
    """Import commonly used ArcGIS libraries for standalone scripts"""
    getModule("esriSystem.olb")
    getModule("esriGeometry.olb")
    getModule("esriCarto.olb")
    getModule("esriDisplay.olb")
    getModule("esriGeoDatabase.olb")
    getModule("esriDataSourcesGDB.olb")
    getModule("esriDataSourcesFile.olb")
    getModule("esriOutput.olb")

def GetDesktopModules():
    """Import basic ArcGIS Desktop libraries"""
    getModule("esriFramework.olb")
    getModule("esriArcMapUI.olb")
    getModule("esriArcCatalogUI.olb")

#**** Helper Functions ****

def GetCurrentApp():
    """Returns the Application if the script is being run from
    within the application boundary of an ArcGIS application
    (i.e. from the Python window in ArcMap)
    """
    import comtypes.gen.esriFramework as esriFramework
    return NewObj(esriFramework.AppRef, esriFramework.IApplication)

def GetApp(app="ArcMap"):
    InitStandalone()
    """app must be 'ArcMap' (default) or 'ArcCatalog'\n\
    Execute GetDesktopModules() first"""
    if not (app == "ArcMap" or app == "ArcCatalog"):
        print "app must be 'ArcMap' or 'ArcCatalog'"
        return None
    import comtypes.gen.esriFramework as esriFramework
    import comtypes.gen.esriArcMapUI as esriArcMapUI
    import comtypes.gen.esriCatalogUI as esriCatalogUI
    pAppROT = NewObj(esriFramework.AppROT, esriFramework.IAppROT)
    iCount = pAppROT.Count
    if iCount == 0:
        return None
    for i in range(iCount):
        pApp = pAppROT.Item(i)
        print pApp
        if app == "ArcCatalog":
            if CType(pApp, esriCatalogUI.IGxApplication):
                return pApp
            continue
        if CType(pApp, esriArcMapUI.IMxApplication):
            return pApp
    return None

def NewObj(COMClass, COMInterface):
    """Creates a new comtypes POINTER object where\n\
    MyClass is the class to be instantiated,\n\
    MyInterface is the interface to be assigned"""
    try:
        ptr = CreateObject(COMClass, interface=COMInterface)
        return ptr
    except:
        return None

def CType(obj, interface):
    """Casts obj to interface and returns comtypes POINTER or None"""
    try:
        newobj = obj.QueryInterface(interface)
        return newobj
    except:
        return None

def CLSID(MyClass):
    """Return CLSID of MyClass as string

    CLSID is the GUID for the COM Class (CoClass) corresponding
    to instances of the object class
    """
    return str(MyClass._reg_clsid_)

def GetMxDoc(pMxDoc):
    """Get an IMapDocument pointer from variety of inputs

    pMxDoc -- can be one of the following types:
        str -- full path to unopened mxd on disk or "current" to reference current
            open mxd
        IApplication -- esriFramework.IApplication pointer to current open application


    """
    import comtypes.gen.esriDisplay as esriDisplay
    import comtypes.gen.esriArcMapUI as esriArcMapUI
    import comtypes.gen.esriCarto as esriCarto
    import comtypes.gen.esriFramework as esriFramework

    # validate/get proper map document instance
    if isinstance(pMxDoc, esriFramework.IApplication):
        pDoc = pApp.Document
        pMxDoc = CType(pDoc, esriArcMapUI.IMxDocument)

     # pMap validation
    if isinstance(pMxDoc, esriCarto.IMapDocument):
        pass

    # it is either the current map open or mxd on disc
    elif isinstance(pMxDoc, basestring):
        if pMxDoc.lower() == 'current':
            pApp = GetCurrentApp()
            pDoc = pApp.Document
            pMxDoc = CType(pDoc, esriArcMapUI.IMxDocument)

        elif os.path.exists(pMxDoc):
            mapDoc = pMxDoc
            InitStandalone()
            pMxDoc = NewObj(esriCarto.MapDocument, esriCarto.IMapDocument)
            pMxDoc.Open(mapDoc)

    return pMxDoc

#********* Stand alone ****

def InitStandalone():
    """Init standalone ArcGIS license"""
    # Set ArcObjects version
    g = comtypes.GUID("{6FCCEDE0-179D-4D12-B586-58C88D26CA78}")
    GetModule((g, 1, 0))
    import comtypes.gen.ArcGISVersionLib as esriVersion
    import comtypes.gen.esriSystem as esriSystem
    pVM = NewObj(esriVersion.VersionManager, esriVersion.IArcGISVersion)
    # make sure version matches
    if not pVM.LoadVersion(esriVersion.esriArcGISDesktop, GetVersion()):
        return False
    # Get license
    pInit = NewObj(esriSystem.AoInitialize, esriSystem.IAoInitialize)
    ProductList = [esriSystem.esriLicenseProductCodeAdvanced, \
                   esriSystem.esriLicenseProductCodeStandard, \
                   esriSystem.esriLicenseProductCodeBasic]
    for eProduct in ProductList:
        licenseStatus = pInit.IsProductCodeAvailable(eProduct)
        if licenseStatus != esriSystem.esriLicenseAvailable:
            continue
        licenseStatus = pInit.Initialize(eProduct)
        return (licenseStatus == esriSystem.esriLicenseCheckedOut)
    return False


def check_extension(ext_code):
    """Determines whether or not an extension is checked out, returns True or False

    Required:
        ext_code -- extension code for license (int),
            for example, Spatial Analyst is code 10.

    Some common extension codes:
        6: Geostatistical Analyst
        8: Network Analyst
        9: 3D Analyst
       10: Spatial Analyst

    http://resources.arcgis.com/en/help/arcobjects-net/componenthelp/index.html#//004200000021000000
    """
    from comtypes.gen import esriSystem
    # now call AoInitialize
    pInit = NewObj(esriSystem.AoInitialize,
                    esriSystem.IAoInitialize)

    # check extension
    return pInit.IsExtensionCheckedOut(int(ext_code))

def mxd_version(mxd):
    """return mxd version information as tuple"""
    InitStandalone()
    getModule('esriCarto')
    import comtypes.gen.esriCarto as esriCarto
    pMapDoc = NewObj(esriCarto.MapDocument, esriCarto.IMapDocument)
    pMapDoc.Open(mxd)
    ver_info = pMapDoc.GetVersionInfo()
    pMapDoc.Close()
    if not ver_info[0]:
        return '.'.join(map(str, ver_info[1:3]))
    print 'No Version info avaliable or the mxd was saved as a newer document'
    return None

def Msg(message="Hello world", title="ARDemo"):
    """Opens a dialog box with OK button

    Required:
    message -- text for message box
    title -- title of message box
    """
    from ctypes import c_int, WINFUNCTYPE, windll
    from ctypes.wintypes import HWND, LPCSTR, UINT
    prototype = WINFUNCTYPE(c_int, HWND, LPCSTR, LPCSTR, UINT)
    fn = prototype(("MessageBoxA", windll.user32))
    return fn(0, message, title, 0)

def create_mxd(mapDoc):
    """creates a new map document (.mxd)

    mapDoc -- path to new mxd
    """
    InitStandalone()
    getModule('esriCarto.olb')
    import comtypes.gen.esriCarto as esriCarto
    mxd = NewObj(esriCarto.MapDocument, esriCarto.IMapDocument)
    mxd.New(mapDoc)
    mxd.Close()
    return mapDoc

def Standalone_OpenSDE(server, instance, database=None, mode='OSA', version=None):
    """open SDE database

    instance example: "sde:oracle10g:/;LOCAL=PRODUCTION_TUCSON"
    """
    GetStandaloneModules()
    InitStandalone()
    import comtypes.gen.esriSystem as esriSystem
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    import comtypes.gen.esriDataSourcesGDB as esriDataSourcesGDB

    pPropSet = NewObj(esriSystem.PropertySet, esriSystem.IPropertySet)
    props = {'SERVER': server,
             'INSTANCE': instance,
             'AUTHENTICATION_MODE': mode,
             'DATABASE': database,
             'VERSION': version}

    for prop, val in props.iteritems():
        if val:
            pPropSet.SetProperty(prop, val)
    pWSF = NewObj(esriDataSourcesGDB.SdeWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory)
    pWS = pWSF.Open(pPropSet, 0)
    pDS = CType(pWS, esriGeoDatabase.IDataset)
    print "Workspace name: " + pDS.BrowseName
    print "Workspace category: " + pDS.Category
    return pWS


def Standalone_OpenFileGDB(sPath):
    """open file GDB"""
    GetStandaloneModules()
    if not InitStandalone():
        print "We've got lumps of it 'round the back..."
        return
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    import comtypes.gen.esriDataSourcesGDB as esriDataSourcesGDB

    pWSF = NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory)
    pWS = pWSF.OpenFromFile(sPath, 0)
    pDS = CType(pWS, esriGeoDatabase.IDataset)
    print "Workspace name: " + pDS.BrowseName
    print "Workspace category: " + pDS.Category
    return pWS

def CheckGDBRelease(sPath):
    """open file GDB"""
    if not os.path.exists(sPath):
        raise IOError('"{}" does not exist!'.format(sPath))
    InitStandalone()
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    import comtypes.gen.esriDataSourcesGDB as esriDataSourcesGDB

    if sPath.strip().endswith('.gdb'):
        pWSF = NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory2)
    elif sPath.strip().endswith('.mdb'):
        pWSF = NewObj(esriDataSourcesGDB.AccessWorkspaceFactory, \
                      esriGeoDatabase.IWorkspaceFactory2)

    pWS = pWSF.OpenFromFile(sPath, 0)
    pGDBRelease = CType(pWS, esriGeoDatabase.IGeodatabaseRelease2)
    return pGDBRelease.CurrentRelease, pGDBRelease.MajorVersion, pGDBRelease.MinorVersion

def OpenFeatureClass(sFileGDB, sFCName):
    """Opens a feature class from a File Geodatabase
    sFileGDB -- path to File Geodatabase
    sFCName -- name of feature class
    """
    InitStandalone()
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    import comtypes.gen.esriDataSourcesGDB as esriDataSourcesGDB
    pWSF = NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory2)
    pWS = pWSF.OpenFromFile(sFileGDB, 0)
    pFWS = CType(pWS, esriGeoDatabase.IFeatureWorkspace)

    # determine if FC exists before attempting to open
    # http://edndoc.esri.com/arcobjects/9.2/ComponentHelp/esriGeoDatabase/IWorkspace2_NameExists.htm
    #   5 = feature class datatype
    pWS2 = CType(pWS, esriGeoDatabase.IWorkspace2)
    if pWS2.NameExists(5, sFCName):
        pFC = pFWS.OpenFeatureClass(sFCName)
    else:
        pFC = None
        print '** %s not found' % sFCName

    return pFC

def iterFeatures(fc, recycle=True):
    """generator that iterates over features and returns an IFeature interface

    fc -- IFeatureClass pointer
    recycle -- option to recycle the row object.  If you want to build a list of
        IFeature row objects, set this to False.

    # example usage:
    pars = r'C:\TEMP\frontage_test.gdb\parcels'
    fc = arcobjects.OpenFeatureClass(*os.path.split(pars))

    for ft in enum_features(fc):
        print ft.OID
    """
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    qf = CreateObject(progid=esriGeoDatabase.QueryFilter, interface=esriGeoDatabase.IQueryFilter)
    count = fc.FeatureCount(qf)
    cur = fc.Search(qf, recycle)
    for i in xrange(count):
        yield cur.NextFeature()

def SearchCursor(fc, fields):
    """arcpy.da. style search cursor (does not support with statement)

    fc -- IFeatureCleass pointer
    fields -- list of field names
    """
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    count = fc.FeatureCount(None)
    tableFields = fc.Fields
    indices = [tableFields.FindField(f) for f in fields]
    cur = fc.Search(None, True)
    for i in xrange(count):
        row = cur.NextFeature()
        yield tuple(row.Value(fi) for fi in indices)

def unregisterReplica(ws, replicaName=None, replicaID=None, replicaGUID=None):
    """Unregisters a replica from a database

    Required:
        ws -- Pointer to IWorkspace Interface
    Optional (need to use one of these):
        replicaName -- name of replica to unregister
        replicaID -- id of replica
        replicaGUID -- Guid of replica
    """
    wsReplicasAdmin = CType(ws, esriGeoDatabase.IWorkspaceReplicasAdmin2)
    wsReps = CType(ws, esriGeoDatabase.IWorkspaceReplicas)
    replica = None
    if replicaName:
        replica = wsReps.ReplicaByName(replicaName)
    elif replicaID:
        replica = wsReps.ReplicaByID(replicaID)
    elif replicaGUID:
        replica = wsReps.ReplicaByGuid(replicaGUID)
    else:
        print 'Not a valid name, ID, or Guid for replica!'
        return False

    if isinstance(replica, esriGeoDatabase.IReplica):
        rep_name = replica.Name
        wsReplicasAdmin.UnregisterReplica(replica, True)
        print 'Successfully unregistered: "{}"'.format(rep_name)
        return True
    else:
        print 'Could not find replica!'
        return False

def getUnitSize(my_size, rounding=1):
    """will return a size in bytes to human readable format"""
    if not isinstance(my_size, (float, int, long)):
        my_size = sys.getsizeof(my_size)
    theSize = '0 KB'
    if my_size == 0:
        theSize = '0 KB'
    if my_size <= 1024:
        theSize = '1 KB'
    elif my_size > 1024 and my_size <= 1048576:
        theSize = '%s KB' %round(my_size/1024.0, rounding)
    elif my_size > 1048576 and my_size <= 1073741824:
        theSize = '%s MB' %round(my_size/1048576.0, rounding)
    elif my_size > 1073741824 and my_size <= 1099511627776:
        theSize = '%s GB' %round(my_size/1073741824.0, rounding)
    elif my_size >= 1099511627776:
        theSize = '%s TB' %round(my_size/1099511627776.0, rounding)
    else:
        # default return size in MB
        return '%s MB' %round(my_size/1048576.0, rounding)
    return theSize

def GetModifiedDate(fc, statType=2):
    """Gets information on feature class and returns a tuple of (mode, size, time)

    Required:
        fc -- feature class to check

    Optional:
        statType -- type of time info as shown below.  Default is 2
            0 	Return the time last accessed.
            1 	Return the time last created.
            2 	Return the time last modified.
    modified from:
        https://geonet.esri.com/thread/74409
    """
    InitStandalone()
    import comtypes.gen.esriSystem as esriSystem
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    import comtypes.gen.esriDataSourcesGDB as esriDataSourcesGDB

    # Open the FGDB
    gdb, tableName = os.path.split(fc)
    pWS = Standalone_OpenFileGDB(gdb)

    # Create empty Properties Set
    pPropSet = NewObj(esriSystem.PropertySet, esriSystem.IPropertySet)
    pPropSet.SetProperty("database", gdb)

    # Cast the FGDB as IFeatureWorkspace
    pFW = CType(pWS, esriGeoDatabase.IFeatureWorkspace)

    # Open the table
    pTab = pFW.OpenTable(tableName)

    # Cast the table as a IDatasetFileStat
    pDFS = CType(pTab, esriGeoDatabase.IDatasetFileStat)

    # Get the date modified
    mod = datetime.datetime.fromtimestamp(pDFS.StatTime(statType)).strftime('%Y-%m-%d %H:%M:%S')
    return (ACCESS_MODE[pDFS.StatMode], getUnitSize(pDFS.StatSize), mod)

#**** ArcMap ****

def MapRefresh(current=True):
    """refreshes ArcMap's TOC and active view"""
    from comtypes.gen import esriArcMapUI

    # reresh active view and TOC
    if current:
        pApp = GetCurrentApp()
    else:
        pApp = GetApp()
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, esriArcMapUI.IMxDocument)
    pMxDoc.UpdateContents()
    pMxDoc.ActiveView.Refresh()
    del pApp, pMxDoc
    return

def ArcMap_GetSelectedGeometry(bStandalone=False):
    """ gets the selected geometry in an ArcMap Session"""

    GetDesktopModules()
    if bStandalone:
        InitStandalone()
        pApp = GetApp()
    else:
        pApp = GetCurrentApp()
    if not pApp:
        print "We found this spoon, sir."
        return
    import comtypes.gen.esriFramework as esriFramework
    import comtypes.gen.esriArcMapUI as esriArcMapUI
    import comtypes.gen.esriSystem as esriSystem
    import comtypes.gen.esriCarto as esriCarto
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    import comtypes.gen.esriGeometry as esriGeometry

    # Get selected feature geometry

    pDoc = pApp.Document
    pMxDoc = CType(pDoc, esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pFeatSel = pMap.FeatureSelection
    pEnumFeat = CType(pFeatSel, esriGeoDatabase.IEnumFeature)
    pEnumFeat.Reset()
    pFeat = pEnumFeat.Next()
    if not pFeat:
        print "No selection found."
        return
    pShape = pFeat.ShapeCopy
    eType = pShape.GeometryType
    if eType == esriGeometry.esriGeometryPoint:
        print "Geometry type = Point"
    elif eType == esriGeometry.esriGeometryPolyline:
        print "Geometry type = Line"
    elif eType == esriGeometry.esriGeometryPolygon:
        print "Geometry type = Poly"
    else:
        print "Geometry type = Other"
    return pShape

def ArcMap_GetEditWorkspace(bStandalone=False):

    GetDesktopModules()
    if bStandalone:
        InitStandalone()
        pApp = GetApp()
    else:
        pApp = GetCurrentApp()
    GetModule("esriEditor.olb")
    import comtypes.gen.esriSystem as esriSystem
    import comtypes.gen.esriEditor as esriEditor
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    pID = NewObj(esriSystem.UID, esriSystem.IUID)
    pID.Value = CLSID(esriEditor.Editor)
    pExt = pApp.FindExtensionByCLSID(pID)
    pEditor = CType(pExt, esriEditor.IEditor)
    if pEditor.EditState == esriEditor.esriStateEditing:
        pWS = pEditor.EditWorkspace
        pDS = CType(pWS, esriGeoDatabase.IDataset)
        print "Workspace name: " + pDS.BrowseName
        print "Workspace category: " + pDS.Category
    return

def ArcMap_GetSelectedTable(bStandalone=False):

    GetDesktopModules()
    if bStandalone:
        InitStandalone()
        pApp = GetApp()
    else:
        pApp = GetCurrentApp()
    import comtypes.gen.esriFramework as esriFramework
    import comtypes.gen.esriArcMapUI as esriArcMapUI
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, esriArcMapUI.IMxDocument)
    pUnk = pMxDoc.SelectedItem
    if not pUnk:
        print "Nothing selected."
        return
    pTable = CType(pUnk, esriGeoDatabase.ITable)
    if not pTable:
        print "No table selected."
        return
    pDS = CType(pTable, esriGeoDatabase.IDataset)
    print "Selected table: " + pDS.Name

#**** ArcCatalog ****

def ArcCatalog_GetSelectedTable(bStandalone=False):

    GetDesktopModules()
    if bStandalone:
        InitStandalone()
        pApp = GetApp("ArcCatalog")
    else:
        pApp = GetCurrentApp()
    import comtypes.gen.esriFramework as esriFramework
    import comtypes.gen.esriCatalogUI as esriCatalogUI
    import comtypes.gen.esriCatalog as esriCatalog
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase
    pGxApp = CType(pApp, esriCatalogUI.IGxApplication)
    pGxObj = pGxApp.SelectedObject
    if not pGxObj:
        print "Nothing selected."
        return
    pGxDS = CType(pGxObj, esriCatalog.IGxDataset)
    if not pGxDS:
        print "No dataset selected."
        return
    eType = pGxDS.Type
    if not (eType == esriGeoDatabase.esriDTFeatureClass or eType == esriGeoDatabase.esriDTTable):
        print "No table selected."
        return
    pDS = pGxDS.Dataset
    pTable = CType(pDS, esriGeoDatabase.ITable)
    print "Selected table: " + pDS.Name

# **** custom helper functions ****
def getLayer(pApp, layer_name):
    """ get a layer by name"""
    pApp = GetMxDoc(pApp)
    for lyr in iterLayers(pApp):
        if lyr.Name == layer_name:
            return lyr


def iterLayers(pApp=None, filterer=[], iterGroups=True):
    """creates a generator for layers"""
    from comtypes.gen import esriArcMapUI, esriCarto

    # nested func
    def getSubLayers(obj, filterer, iterGroups):
        """returns a layer at an index from a group layer or map"""
        if isinstance(obj, esriCarto.IMap):
            attr = 'LayerCount'
        elif isinstance(obj, esriCarto.ICompositeLayer):
            attr = 'Count'

        for i in range(getattr(obj, attr)):
            layer = obj.Layer(i)
            gl = CType(layer, esriCarto.IGroupLayer)
            if iterGroups and isinstance(gl, esriCarto.IGroupLayer):
                cl = CType(gl, esriCarto.ICompositeLayer)

                # iter sub layers
                for lyr in getSubLayers(cl, filterer, True):
                    if len(filterer) and lyr.Name in filterer:
                        yield lyr

                    elif not filterer:
                        yield lyr

            else:
                # its just a layer
                if len(filterer) and layer.Name in filterer:
                    yield layer

                elif not filterer:
                    yield layer

    pApp = GetMxDoc(pApp)
    if isinstance(pApp, esriArcMapUI.IMxDocument):
        pMxDoc = pApp

    else:
        if not pApp:
            pApp = GetApp()
        pDoc = pApp.Document
        pMxDoc = CType(pDoc, esriArcMapUI.IMxDocument)

    if isinstance(pMxDoc, esriArcMapUI.IMxDocument):
        pMap = pMxDoc.FocusMap

        # reference layer
        return getSubLayers(pMap, filterer, iterGroups)

def clearReferenceScale(layer):
    """clears a refernce scale for a layer"""
    from comtypes.gen import esriArcMapUI, esriCarto
    if isinstance(layer, esriCarto.ILayer):
        layer = CType(layer, esriCarto.IGeoFeatureLayer)

    if isinstance(layer, esriCarto.IGeoFeatureLayer):
        layer.ScaleSymbols = False
    else:
        raise ValueError('{} is not an esriCarto.ILayer or esriCarto.IGeoFeatureLayer!'.format(layer.__class__.__name__))

def importSymbologyFromLayer(pMxDoc, target_layer, target_field, symbol_layer, normalization_field=None, minVal=None, maxVal=None):
    """imports the symbology from a layer in TOC

    Required:
        pApp -- reference to current app
        target_layer -- layer to symbolize
        target_field -- name of target field to symbolize
        symbol_layer -- symbology layer

    Optional:
        normalization_field -- field to normalize by
        minVal -- minimum value for classification, only used for Proportional Symbols
        maxVal -- maximum value for classification, only used for Proportional Symbols
    """
    import comtypes.gen.esriDisplay as esriDisplay
    import comtypes.gen.esriArcMapUI as esriArcMapUI
    import comtypes.gen.esriCarto as esriCarto
    import comtypes.gen.esriFramework as esriFramework
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase

    # validate/get proper map document instance
    pMxDoc = GetMxDoc(pMxDoc)

    # get IMapPointer
    pMap = pMxDoc.FocusMap

    # get layer references
    targ_lyr = [pMap.Layer(i) for i in range(pMap.LayerCount) if pMap.Layer(i).Name == target_layer][0]
    symb_lyr = [pMap.Layer(i) for i in range(pMap.LayerCount) if pMap.Layer(i).Name == symbol_layer][0]

    # cast to IGeoFeatureLayer interface and get renderers
    targGeo = CType(targ_lyr, esriCarto.IGeoFeatureLayer)
    symbGeo = CType(symb_lyr, esriCarto.IGeoFeatureLayer)
    symbRenderer = symbGeo.Renderer

    #**********************************************************************************
    #
    # ClassBreaksRenderer
    symbClassBreaks = CType(symbRenderer, esriCarto.IClassBreaksRenderer)
    symbUnique = CType(symbRenderer, esriCarto.IUniqueValueRenderer)
    symbProp = CType(symbRenderer, esriCarto.IProportionalSymbolRenderer)

    if isinstance(symbClassBreaks, esriCarto.IClassBreaksRenderer):

        # cast to IClassBreaksRenderer
        targSymb = NewObj(esriCarto.ClassBreaksRenderer, esriCarto.IClassBreaksRenderer)

        # set target symbology field
        targSymb.Field = target_field
        targSymb.NormField = normalization_field
        targSymb.BreakCount = symbClassBreaks.BreakCount

        # iterate through class beraks from symbol layer and apply to target
        for i in range(symbClassBreaks.BreakCount):
            targSymb.Break[i] = symbClassBreaks.Break[i]
            targSymb.Description[i] = symbClassBreaks.Description[i]
            targSymb.Symbol[i] = symbClassBreaks.Symbol[i]
            targSymb.Label[i] = symbClassBreaks.Label[i]

    # ***********************************************************************************
    #
    # UniqueValuesRenderer
    elif isinstance(symbUnique, esriCarto.IUniqueValueRenderer):
        targSymb = NewObj(esriCarto.UniqueValueRenderer, esriCarto.IUniqueValueRenderer)

        # add all values from symbol layer
        targGeo.Field = target_field

        for i in range(symbUnique.ValueCount):
            targSymb.AddValue(symbUnique.Value[i], symbUnique.Heading(symbUnique.Value[i]), symbUnique.Symbol(symbUnique.Value[i]))

    # ***********************************************************************************
    #
    # ProportionalSymbolRenderer
    elif isinstance(symbProp, esriCarto.IProportionalSymbolRenderer):
        targSymb = NewObj(esriCarto.ProportionalSymbolRenderer, esriCarto.IProportionalSymbolRenderer)
        targSymb.Field = target_field
        targSymb.NormField = normalization_field

        # set min and max symbols
        for attr in ('FlanneryCompensation', 'MinDataValue', 'MaxDataValue', 'BackGroundSymbol',
                     'MinSymbol', 'ValueUnit', 'ValueRepresentation', 'LegendSymbolCount'):
            if hasattr(symbProp, attr):
                setattr(targSymb, attr, getattr(symbProp, attr))

        targSymb.MinDataValue = minVal
        targSymb.MaxDataValue = maxVal

        # create legend symbolss
        targSymb.CreateLegendSymbols()

    targGeo.Renderer = targSymb
    pMxDoc.ActiveView.Refresh()
    pMxDoc.UpdateContents()
    return targSymb

def setSymbolSize(pMapDocument, layer_names=[], pointSize=12, lineWidth=1, autoSave=True, do_all=False, clearRefScale=True):
    """sets the size/width of simple esri Marker or Line symbols

    Required:
        pMapDocument -- IMapDocument pointer (esriCarto) or string map document path

    Optional:
        layer_names -- list of layer names to set size of as they appear in TOC.
        pointSize -- size for all point layers specified.  Default is 12
        lineWidth -- width for all line layers specified.  Default is 1
        autoSave -- option to save document automatically after making changes.
            Default is True.
        do_all -- option to do all point/line layers in map.  Default is False.
    """
    import comtypes.gen.esriDisplay as esriDisplay
    import comtypes.gen.esriArcMapUI as esriArcMapUI
    import comtypes.gen.esriCarto as esriCarto

    # pMap validation
    if isinstance(pMapDocument, esriCarto.IMapDocument):
        pass
    elif isinstance(pMapDocument, basestring):
        InitStandalone()
        mapDoc = pMapDocument
        pMapDocument = NewObj(esriCarto.MapDocument, esriCarto.IMapDocument)
        pMapDocument.Open(mapDoc)

    # get IMap pointer
    pMap = pMapDocument.Map(0)

    if isinstance(layer_names, (basestring)):
        layer_names = [layer_names]

    # defaults for point (shapeType 1) is 12, default for line (shapeType 3) is 2
    defaultSymbol = {1: esriDisplay.IMarkerSymbol,
                     3: esriDisplay.ILineSymbol}

    # iterate through layer list
    for i in range(pMap.LayerCount):

        lyr = pMap.Layer(i)
        if lyr.Name in layer_names or do_all in (True, 1):

            # cast to esriCarto.IGeoFeatureLayer to get renderer
            flyr = CType(lyr, esriCarto.IFeatureLayer2)
            iGeo = CType(lyr, esriCarto.IGeoFeatureLayer)
            if clearRefScale:
                iGeo.ScaleSymbols = False
            renderer = iGeo.Renderer

            # cast to simple renderer and then to simple symbol
            simpleRenderer = CType(renderer, esriCarto.ISimpleRenderer)
            if isinstance(simpleRenderer, esriCarto.ISimpleRenderer):
                simple = simpleRenderer.Symbol

                # cast to marker symbol interface and set size or width
                if flyr.ShapeType in (1,3):
                    marker = CType(simple, defaultSymbol[flyr.ShapeType])
                    if flyr.ShapeType == 1 and pointSize is not None:
                        marker.Size = pointSize
                    elif flyr.ShapeType == 3 and lineWidth is not None:
                        marker.Width = lineWidth

    # refresh Map
    pMapDocument.ActiveView.Refresh()

    # save and close map
    if autoSave in (True, 1):
        pMapDocument.Save()
    pMapDocument.Close()
    del pMap, pMapDocument
    return

def alter_alias(fc, f_dict):
    """Change field names at the database level

    Required:
    fc -- feature class (must be in gdb)
    f_dict -- fields dictionary {field_alias : new_alias, ...}
    """
    getModule('esriGeoDatabase.olb')
    getModule('esriDataSourcesGDB.olb')
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase

    # get at fc properties
    pFC = OpenFeatureClass(*os.path.split(fc))

    # get at schema editing library
    pSE = CType(pFC, esriGeoDatabase.IClassSchemaEdit)

    # iterate through dict and update aliases
    for field, alias in f_dict.iteritems():
        try:
            pSE.AlterFieldAliasName(field, alias)
            print 'Changed field "{0}"\'s alias to: "{1}"'.format(field, alias)
        except:
            print 'Error changing field "{0}"\'s alias to: "{1}"'.format(field, alias)
    return

def alter_fieldName(fc, f_dict):
    """Change field names at the database level

    Required:
    fc -- feature class (must be in gdb)
    f_dict -- fields dictionary {field_name : new_name, ...}
    """
    import comtypes.gen.esriGeoDatabase as esriGeoDatabase

    # get at fc properties
    pFC = OpenFeatureClass(*os.path.split(fc))

    # get at schema editing library
    pSE = CType(pFC, esriGeoDatabase.IClassSchemaEdit4)

    # iterate through dict and update aliases
    for field, new_name in f_dict.iteritems():
        try:
            pSE.AlterFieldName(field, new_name)
            print 'Changed field "{0}" name to: "{1}"'.format(field, new_name)
        except:
            print 'Error changing field "{0}" name to: "{1}"'.format(field, new_name)
    return


if __name__ == '__main__':
    instance = 'sde:sqlserver:ArcSQL;LOCAL=BMI_PhotoApp'
##    sde = Standalone_OpenSDE('ArcSQL', instance, version='dbo.DEFAULT')
##    Standalone_OpenFileGDB(r'C:\TEMP\TemplateData.gdb')
    gdb = r'E:\HamiltonCo\Soil_Library\AgLand_Adjustment\CSR_AgLand.gdb'
    gdb = r'C:\TEMP\temp.gdb'
    tup = CheckGDBRelease(gdb)
    print tup
    print CheckGDBRelease(r'C:\TEMP\mdb_test.mdb')
##    fc = OpenFeatureClass(r'C:\TEMP\TemplateData.gdb', 'cities')
##    fc = OpenFeatureClass(*os.path.split(r'C:\TEMP\test.gdb\StmCB'))
##    print fc
