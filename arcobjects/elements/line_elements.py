import arcobjects
import comtypes.gen.esriFramework as esriFramework
import comtypes.gen.esriArcMapUI as esriArcMapUI
import comtypes.gen.esriSystem as esriSystem
import comtypes.gen.esriGeometry as esriGeometry
import comtypes.gen.esriCarto as esriCarto
import comtypes.gen.esriDisplay as esriDisplay
import comtypes.gen.stdole as stdole

def add_line(pApp=None, name='Line', x=None, y=None, end_x=None, end_y=None,
             x_len=0, y_len=0, anchor=0, rgb=(0,0,0), view='layout'):
    '''adds a line to an ArcMap Document

    Required:
    pApp -- reference to either open ArcMap document or path on disk
    name -- name of line element

    Optional:
    x -- start x coordinate, if none, middle of the extent will be used (data view)
    y -- start y coordinate, if none, middle of the extent will be used (data view)
    end_x -- end x coordinate, if making straight lines use x_len
    end_y -- end y coordinate, if making straight lines use y_len
    x_len -- length of line in east/west direction
    y_len -- length of line in north/south direction
    anchor -- anchor point for line element
    rgb -- tuple for red, green, blue color values.  Default is (0,0,0) for black.
    view -- choose view for text element (layout|data)

    Anchor Points:
        esriTopLeftCorner 	0 	Anchor to the top left corner.
        esriTopMidPoint 	1 	Anchor to the top mid point.
        esriTopRightCorner 	2 	Anchor to the top right corner.
        esriLeftMidPoint 	3 	Anchor to the left mid point.
        esriCenterPoint 	4 	Anchor to the center point.
        esriRightMidPoint 	5 	Anchor to the right mid point.
        esriBottomLeftCorner 	6 	Anchor to the bottom left corner.
        esriBottomMidPoint 	7 	Anchor to the bottom mid point.
        esriBottomRightCorner 	8 	Anchor to the botton right corner.
    '''

    # set mxd
    if not pApp:
        pApp = arcobjects.GetApp()
    if str(pApp).lower() == 'current':
        pApp = arcobjects.GetCurrentApp()
    pDoc = pApp.Document
    pMxDoc = arcobjects.CType(pDoc, esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pMapL = pMap
    if view.lower() == 'layout':
        pMapL = pMxDoc.PageLayout
    pAV = arcobjects.CType(pMapL, esriCarto.IActiveView)
    pSD = pAV.ScreenDisplay

    # set coords for elment
    pFact = arcobjects.CType(pApp, esriFramework.IObjectFactory)
    if view.lower() == 'data':
        pEnv = pAV.Extent
        if x == None:
            x = (pEnv.XMin + pEnv.XMax) / 2
        if y == None:
            y = (pEnv.YMin + pEnv.YMax) / 2
    else:
        # default layout position, move off page
        if x == None: x = -4
        if y == None: y = 4

    # from point
    pUnk_pt = pFact.Create(arcobjects.CLSID(esriGeometry.Point))
    pPt = arcobjects.CType(pUnk_pt, esriGeometry.IPoint)
    pPt.PutCoords(x, y)

    # to point
    pUnk_pt2 = pFact.Create(arcobjects.CLSID(esriGeometry.Point))
    pPt2 = arcobjects.CType(pUnk_pt2, esriGeometry.IPoint)
    if x_len or y_len:
        pPt2.PutCoords(x + x_len, y + y_len)
    elif end_x or end_y:
        pPt2.PutCoords(end_x, end_y)

    # line (from point - to point)
    pUnk_line = pFact.Create(arcobjects.CLSID(esriGeometry.Polyline))
    pLg = arcobjects.CType(pUnk_line, esriGeometry.IPolyline)
    pLg.FromPoint = pPt
    pLg.ToPoint = pPt2

    # preset color according to RGB values
    pUnk_color = pFact.Create(arcobjects.CLSID(esriDisplay.RgbColor))
    pColor = arcobjects.CType(pUnk_color, esriDisplay.IRgbColor)
    pColor.Red, pColor.Green, pColor.Blue = rgb

    # set line properties
    pUnk_line = pFact.Create(arcobjects.CLSID(esriDisplay.SimpleLineSymbol))
    pLineSymbol = arcobjects.CType(pUnk_line, esriDisplay.ISimpleLineSymbol)
    pLineSymbol.Color = pColor

    # create the actual element
    pUnk_elm = pFact.Create(arcobjects.CLSID(esriCarto.LineElement))
    pLineElement = arcobjects.CType(pUnk_elm, esriCarto.ILineElement)
    pLineElement.Symbol = pLineSymbol
    pElement = arcobjects.CType(pLineElement, esriCarto.IElement)

    # elm properties
    pElmProp = arcobjects.CType(pElement, esriCarto.IElementProperties3)
    pElmProp.Name = name
    pElmProp.AnchorPoint = esriCarto.esriAnchorPointEnum(anchor)
    pElement.Geometry = pLg

    # add to map
    pGC = arcobjects.CType(pMapL, esriCarto.IGraphicsContainer)
    pGC.AddElement(pElement, 0)
    pGCSel = arcobjects.CType(pMapL, esriCarto.IGraphicsContainerSelect)
    pGCSel.SelectElement(pElement)
    iOpt = esriCarto.esriViewGraphics + \
           esriCarto.esriViewGraphicSelection
    pAV.PartialRefresh(iOpt, None, None)
    return pElement

if __name__ == '__main__':

    # testing (make a triangle)
    add_line(name='hypot', end_x=-2, end_y=2, anchor=3)
    add_line(name='vertLine', y_len=-2, anchor=1)
    add_line(name='bottom', y=2, end_x=-2, end_y=2)

