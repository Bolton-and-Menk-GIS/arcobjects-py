import arcobjects
import textwrap
import comtypes.gen.esriFramework as esriFramework
import comtypes.gen.esriArcMapUI as esriArcMapUI
import comtypes.gen.esriSystem as esriSystem
import comtypes.gen.esriGeometry as esriGeometry
import comtypes.gen.esriCarto as esriCarto
import comtypes.gen.esriDisplay as esriDisplay
import comtypes.gen.stdole as stdole

def add_text(pApp=None, text='Hello, World!', name='textElm',
             size=10, font='Arial', bold=False, rgb=(0,0,0),
             x=None, y=None, wrapping=None, angle=0, anchor=0,
             mask=False, mask_size=1, view='layout'):
    '''Creates and adds a text element to an mxd

    Required:
    pApp -- mxd application reference, if none specified
            will use current open map document
    text -- text for text element

    Optional:
    name -- name of element
    size -- font size
    font -- name of font
    bold -- set bold (bool)
    rgb -- tuple of red, green, blue color values. Default is (0,0,0) for black.
    x -- x-axis position for text (in map units)
    y -- y-axis position for text (in map units)
    wrapping -- option to wrap text in text element (int)
    angle -- angle for text (0-360)
    anchor -- anchor point for line element
    mask -- if True, will create a halo around text element
    mask_size -- only applies if mask is True, size of halo
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

    # get object factory
    if not pApp:
        pApp = arcobjects.GetApp()
    if str(pApp).lower() == 'current':
        pApp = arcobjects.GetCurrentApp()
    pFact = arcobjects.CType(pApp, esriFramework.IObjectFactory)

    # set mxd
    pDoc = pApp.Document
    pMxDoc = arcobjects.CType(pDoc, esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pMapL = pMap
    if view.lower() == 'layout':
        pMapL = pMxDoc.PageLayout
    pAV = arcobjects.CType(pMapL, esriCarto.IActiveView)
    pSD = pAV.ScreenDisplay

    # set coords for text elm
    pUnk = pFact.Create(arcobjects.CLSID(esriGeometry.Point))
    pPt = arcobjects.CType(pUnk, esriGeometry.IPoint)
    if view.lower() == 'data':
        pEnv = pAV.Extent
        if not x:
            x = (pEnv.XMin + pEnv.XMax) / 2
        if not y:
            y = (pEnv.YMin + pEnv.YMax) / 2
    else:
        # default layout, move off page
        if x == None: x = -4
        if y == None: y = 4
    pPt.PutCoords(x, y)

    # preset color according to RGB values
    pUnk = pFact.Create(arcobjects.CLSID(esriDisplay.RgbColor))
    pColor = arcobjects.CType(pUnk, esriDisplay.IRgbColor)
    pColor.Red, pColor.Green, pColor.Blue = rgb

    # set text properties
    pUnk = pFact.Create(arcobjects.CLSID(stdole.StdFont))
    pFontDisp = arcobjects.CType(pUnk, stdole.IFontDisp)
    pFontDisp.Name = font
    pFontDisp.Bold = bold
    pUnk = pFact.Create(arcobjects.CLSID(esriDisplay.TextSymbol))
    pTextSymbol = arcobjects.CType(pUnk, esriDisplay.ITextSymbol)
    pTextSymbol.Font = pFontDisp
    pTextSymbol.Color = pColor
    pTextSymbol.Size = size
    pTextSymbol.Angle = angle

    # create mask
    if mask:
        pMask = arcobjects.CType(pTextSymbol, esriDisplay.IMask)
        pMask.MaskStyle = 1
        pMask.MaskSize = mask_size

    # create the actual element
    pUnk = pFact.Create(arcobjects.CLSID(esriCarto.TextElement))
    pTextElement = arcobjects.CType(pUnk, esriCarto.ITextElement)
    pTextElement.Symbol = pTextSymbol
    pTextElement.ScaleText = False
    if wrapping:
        pTextElement.Text = textwrap.fill(text, wrapping)
    else:
        pTextElement.Text = text
    pElement = arcobjects.CType(pTextElement, esriCarto.IElement)
    pElement.Geometry = pPt

    # elm properties
    pElmProp = arcobjects.CType(pElement, esriCarto.IElementProperties3)
    pElmProp.Name = name
    pElmProp.AnchorPoint = esriCarto.esriAnchorPointEnum(anchor)

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

    add_text(text='Header', name='headerTxt', size=30, font='Arial', bold=True, x=2,y=8.5,
             anchor=0, mask=True, mask_size=4, view='layout')
##    add_text(text='Cell', name='cellTxt', size=8, font='Arial', anchor=4)

