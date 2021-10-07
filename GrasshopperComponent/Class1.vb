Imports System.IO
Imports System.Collections.Generic
Imports Grasshopper.Kernel
Imports Rhino.Geometry
Imports Grasshopper.Kernel.Parameters
Imports Grasshopper.Kernel.Types

'newly added
Imports EGIS.ShapeFileLib
Imports DotSpatial.Projections
Imports System.Drawing
Imports System.Text
Imports System.Math


Public Class Meerkat1

    Inherits GH_Component
    ''' <summary>
    ''' Initializes a new instance of the $safeitemrootname$ class.
    ''' </summary>
    ''' 
    Public Sub New()
        MyBase.New("Import Shapefile", "Shapefile", "Opens dialog for loading and cropping shapefiles.", "Extra", "Meerkat")
    End Sub

    ''' <summary>
    ''' Registers all the input parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterInputParams(pManager As GH_Component.GH_InputParamManager)
        'pManager.AddIntegerParameter("Number A", "A", "First number for operation", GH_ParamAccess.item, 4)
        'pManager.AddIntegerParameter("Number B", "B", "Second number for operation", GH_ParamAccess.item, 5)
        'pManager.AddTextParameter("Number C", "C", "Third number for operation", GH_ParamAccess.item, "none")
        pManager.AddBooleanParameter("Power", "On-Off", "To Turn on-off", GH_ParamAccess.item, False)
    End Sub
    ''' <summary>
    ''' Registers all the output parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterOutputParams(pManager As GH_Component.GH_OutputParamManager)
        'pManager.AddIntegerParameter("Product", "P", "Product of all inputs", GH_ParamAccess.item)
    End Sub

    ''' <summary>
    ''' This is the method that actually does the work.
    ''' </summary>
    ''' <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
    Protected Overrides Sub SolveInstance(DA As IGH_DataAccess)

        Dim D As Boolean = False

        If Not DA.GetData(0, D) Then Return

        If D = False Then
            Exit Sub
        End If
        If formopen = False Then
            If D = True Then
                Dim myForm As New Form2()
                AddHandler myForm.Button1.Click, AddressOf Bob
                AddHandler myForm.Closed, AddressOf Sue
                myForm.Show()
                formopen = True
            End If
        End If

        'If formopen = True Then
        '    DA.SetData(0, z)

        'End If

    End Sub

    Sub Bob()

        ExpireSolution(True)

    End Sub
    Sub Sue()

        formopen = False

    End Sub
    ''' <summary>
    ''' Provides an Icon for every component that will be visible in the User Interface.
    ''' Icons need to be 24x24 pixels.
    ''' </summary>
    Protected Overrides ReadOnly Property Icon() As System.Drawing.Bitmap
        Get
            'You can add image files to your project resources and access them like this:
            Return My.Resources.shapefile
            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets the unique ID for this component. Do not change this ID after release.
    ''' </summary>
    Public Overrides ReadOnly Property ComponentGuid() As Guid
        Get
            Return New Guid("{B209ACD1-B880-B11F-B516-DFE90EE4588C}")
        End Get
    End Property


End Class
Public Class Meerkat2
    Inherits GH_Component
    ''' <summary>
    ''' Initializes a new instance of the $safeitemrootname$ class.
    ''' </summary>
    ''' 

    Public Sub New()
        MyBase.New("Parse Meerkat File", "Parse .mkgis", "Provides access to mkgis geometry and attributes from shapefile.", "Extra", "Meerkat")
    End Sub

    ''' <summary>
    ''' Registers all the input parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterInputParams(pManager As GH_Component.GH_InputParamManager)

        pManager.AddTextParameter("Meerkat File", "F", "Connect .mkgis file path.", GH_ParamAccess.item)

    End Sub
    ''' <summary>
    ''' Registers all the output parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterOutputParams(pManager As GH_Component.GH_OutputParamManager)

        '0
        pManager.AddTextParameter("Date Cropped", "Date Cropped", "", GH_ParamAccess.item)
        '1
        pManager.AddTextParameter("User and Machine Named", "User and Machine Named", "", GH_ParamAccess.item)
        '2
        pManager.AddTextParameter("Path for files", "Path for files", "", GH_ParamAccess.item)
        '3
        pManager.AddTextParameter("Shape Type", "shape Type", "", GH_ParamAccess.item)
        '4
        pManager.AddIntegerParameter("Shape Count in Source", "Shape Count in Source", "", GH_ParamAccess.item)
        '5
        pManager.AddCurveParameter("Bounds in Point Space", "Bounds in Point Space", "", GH_ParamAccess.item)
        '6
        pManager.AddTextParameter("Bounds in Lat-Long", "Bounds in Lat-Long", "", GH_ParamAccess.item)
        '7
        pManager.AddTextParameter("Field Names", "Field Names", "", GH_ParamAccess.list)
        '8
        pManager.AddIntegerParameter("Shape Count in Crop", "Shape Count in Crop", "", GH_ParamAccess.item)
        '9
        pManager.AddTextParameter("Shape ID in source", "Shape ID in source", "", GH_ParamAccess.tree)
        '10
        pManager.AddTextParameter("Field Values per Shape", "Field Values per Shape", "", GH_ParamAccess.tree)
        '11
        pManager.AddPointParameter("Geometry per Shape", "Geometry per Shape", "", GH_ParamAccess.tree)


    End Sub


    Dim branchcontrol As Integer = 0
    ''' <summary>
    ''' This is the method that actually does the work.
    ''' </summary>
    ''' <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
    Protected Overrides Sub SolveInstance(DA As IGH_DataAccess)

        branchcontrol = DA.Iteration
        Dim D As String = Nothing

        If Not DA.GetData(0, D) Then Return
        If D = "" Then Return

        Dim sr As New StreamReader(D)

        Dim OP0 As String = Replace(sr.ReadLine, "Date Cropped: ", "")
        Dim OP1 As String = sr.ReadLine
        Dim OP2 As String = Replace(sr.ReadLine, "File Path: ", "")
        Dim OP3 As String = Replace(sr.ReadLine, "Shape Type: ", "")
        Dim OP4 As Integer = CInt(Replace(sr.ReadLine, "Record Count: ", ""))
        Dim OP5 As String = Replace(sr.ReadLine, "Bounds: ", "")

        Dim bounds() As String

        bounds = Split(OP5, "|")

        Dim SWpt As New Rhino.Geometry.Point3d(CDbl(bounds(1)), CDbl(bounds(0)), 0)
        Dim NWpt As New Rhino.Geometry.Point3d(CDbl(bounds(1)), CDbl(bounds(2)), 0)
        Dim NEpt As New Rhino.Geometry.Point3d(CDbl(bounds(3)), CDbl(bounds(2)), 0)
        Dim SEpt As New Rhino.Geometry.Point3d(CDbl(bounds(3)), CDbl(bounds(0)), 0)

        Dim boundpts As New Rhino.Collections.Point3dList()

        boundpts.Add(SWpt)
        boundpts.Add(NWpt)
        boundpts.Add(NEpt)
        boundpts.Add(SEpt)
        boundpts.Add(SWpt)

        Dim boundsGH As New Rhino.Geometry.PolylineCurve(boundpts)

        Dim OP6 As String = Replace(sr.ReadLine, "Bounds (SW corner and NE corner Lat/Lng ): ", "")
        Dim OP7 As String = Replace(sr.ReadLine, "Field Names: ", "")

        Dim fieldnames() As String
        fieldnames = Split(OP7, "|")

        Dim OP8 As Integer = CInt(Replace(sr.ReadLine, "Record Count in Crop: ", ""))

        'read blank line
        sr.ReadLine()

        Dim firstDT As New Grasshopper.Kernel.Data.GH_Structure(Of Grasshopper.Kernel.Types.GH_Integer)
        Dim secondDT As New Grasshopper.Kernel.Data.GH_Structure(Of Grasshopper.Kernel.Types.GH_String)
        Dim thirdDT As New Grasshopper.Kernel.Data.GH_Structure(Of Grasshopper.Kernel.Types.GH_Point)

        Dim Counter As Integer = 0

        Do Until sr.EndOfStream



            Dim GHint As Grasshopper.Kernel.Types.GH_Integer = Nothing
            Dim GHpath As New Grasshopper.Kernel.Data.GH_Path()

            GHpath.FromString(branchcontrol.ToString & ":" & Counter.ToString)


            Grasshopper.Kernel.GH_Convert.ToGHInteger(CInt(sr.ReadLine()), GH_Conversion.Both, GHint)
            firstDT.Insert(GHint, GHpath, 0)


            Dim second() As String
            second = Split(sr.ReadLine, "|")

            Dim scounter As Integer = 0

            For Each s As String In second

                Dim GHstr1 As Grasshopper.Kernel.Types.GH_String = Nothing
                Grasshopper.Kernel.GH_Convert.ToGHString(s, GH_Conversion.Both, GHstr1)
                secondDT.Insert(GHstr1, GHpath, scounter)
                GHstr1 = Nothing
                scounter = scounter + 1
            Next

            Dim third As String = sr.ReadLine

            Dim scounter1 As Integer = 0


            third = Replace(third, "{X=", "")
            third = Replace(third, " Y=", "")
            third = Replace(third, "}", "")

            Dim thirds() As String
            thirds = Split(third, " ")

            For Each s As String In thirds

                If s = "" Then GoTo theEnd

                Dim xy() As String
                xy = Split(s, ",")

                Dim x As Double
                Dim y As Double

                x = xy(0)
                y = xy(1)

                Dim rhinopt As New Rhino.Geometry.Point3d(x, y, 0)
                Dim GHpt As New Grasshopper.Kernel.Types.GH_Point(rhinopt)
                thirdDT.Insert(GHpt, GHpath, scounter1)
                scounter1 = scounter1 + 1
theEnd:
            Next
            Counter = Counter + 1

        Loop

        DA.SetData(0, OP0)
        DA.SetData(1, OP1)
        DA.SetData(2, OP2)
        DA.SetData(3, OP3)
        DA.SetData(4, OP4)
        DA.SetData(5, boundsGH)
        DA.SetData(6, OP6)
        DA.SetDataList(7, fieldnames)
        DA.SetData(8, OP8)
        DA.SetDataTree(9, firstDT)
        DA.SetDataTree(10, secondDT)
        DA.SetDataTree(11, thirdDT)


        sr.Close()
        branchcontrol = branchcontrol + 1
    End Sub

    ''' <summary>
    ''' Provides an Icon for every component that will be visible in the User Interface.
    ''' Icons need to be 24x24 pixels.
    ''' </summary>
    Protected Overrides ReadOnly Property Icon() As System.Drawing.Bitmap
        Get

            Return My.Resources.parseshape
            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets the unique ID for this component. Do not change this ID after release.
    ''' </summary>
    Public Overrides ReadOnly Property ComponentGuid() As Guid
        Get
            Return New Guid("{B209ACD1-B880-B11F-B516-DFE90EE4588D}")
        End Get
    End Property


End Class
Public Class Meerkat3

    Inherits GH_Component
    ''' <summary>
    ''' Initializes a new instance of the $safeitemrootname$ class.
    ''' </summary>
    ''' 
    Public Sub New()
        MyBase.New("Lat-Long to Point", "Lat-Long to Point", "Projects latitude and longitude to point.", "Extra", "Meerkat")
    End Sub

    ''' <summary>
    ''' Registers all the input parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterInputParams(pManager As GH_Component.GH_InputParamManager)
        pManager.AddNumberParameter("Latitude", "Lat", "Latitude to project", GH_ParamAccess.item)
        pManager.AddNumberParameter("Longitude", "Long", "Longitude to project", GH_ParamAccess.item)
        pManager.AddTextParameter("Projection File", "Prj", "Projection file", GH_ParamAccess.item)
    End Sub
    ''' <summary>
    ''' Registers all the output parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterOutputParams(pManager As GH_Component.GH_OutputParamManager)
        'pManager.AddIntegerParameter("Product", "P", "Product of all inputs", GH_ParamAccess.item)
        pManager.AddPointParameter("Projected Point", "p", "Point projected from lat and long.", GH_ParamAccess.item)
    End Sub

    ''' <summary>
    ''' This is the method that actually does the work.
    ''' </summary>
    ''' <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
    Protected Overrides Sub SolveInstance(DA As IGH_DataAccess)

        Dim Lat As Double = 0
        Dim Lng As Double = 0
        Dim Prj As String = Nothing

        If Not DA.GetData(0, Lat) Then Return
        If Not DA.GetData(1, Lng) Then Return
        If Not DA.GetData(2, Prj) Then Return

        Dim X As Double = CoordstoPts(Lat, Lng, Prj).Var1
        Dim Y As Double = CoordstoPts(Lat, Lng, Prj).Var2

        Dim newPoint As New Point3d(X, Y, 0)

        DA.SetData(0, newPoint)

    End Sub

    ''' <summary>
    ''' Provides an Icon for every component that will be visible in the User Interface.
    ''' Icons need to be 24x24 pixels.
    ''' </summary>
    Protected Overrides ReadOnly Property Icon() As System.Drawing.Bitmap
        Get
            'You can add image files to your project resources and access them like this:
            Return My.Resources.PP
            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets the unique ID for this component. Do not change this ID after release.
    ''' </summary>
    Public Overrides ReadOnly Property ComponentGuid() As Guid
        Get
            Return New Guid("{B209ACD1-B880-B11F-B516-DFE90EE4588E}")
        End Get
    End Property

    Function CoordstoPts(ByVal Lat1 As Double, ByVal Lng1 As Double, ByVal prjfile As String) As TwoVars

        Dim Latitude As Double
        Dim Longitude As Double

        Latitude = Lat1
        Longitude = Lng1

        'Sets up a array to contain the x and y coordinates
        Dim xy(1) As Double
        xy(0) = Longitude
        xy(1) = Latitude
        'An array for the z coordinate
        Dim z(0) As Double
        z(0) = 0
        'Defines the starting coordiante system
        Dim pend As ProjectionInfo = KnownCoordinateSystems.Geographic.World.WGS1984
        'Defines the ending coordiante system
        Dim pstart As New ProjectionInfo()

        'initiates a StreamReader to read the ESRI .prj file
        Dim re As StreamReader = File.OpenText(prjfile)
        'sets the ending PCS to the ESRI .prj file
        pstart.ParseEsriString(re.ReadLine())
        Reproject.ReprojectPoints(xy, z, pend, pstart, 0, 1)
        'MessageBox.Show(xy(0).ToString & ", " & xy(1).ToString & ", " & z(0))
        Dim newnums As TwoVars

        newnums.Var1 = xy(0)
        newnums.Var2 = xy(1)

        Return newnums

    End Function

    Function Projectpoints(ByVal xpt As Long, ByVal ypt As Long, ByVal prjfile As String) As TwoVars

        'declares a new ProjectionInfo for the startind and ending coordinate systems
        'sets the start GCS to WGS_1984
        Dim pStart As ProjectionInfo = KnownCoordinateSystems.Geographic.World.WGS1984
        Dim pESRIEnd As New ProjectionInfo()
        'declares the point(s) that will be reprojected
        Dim xy As Double() = New Double(1) {}
        Dim z As Double() = New Double(0) {}

        xy(0) = xpt
        xy(1) = ypt
        z(0) = 0

        'initiates a StreamReader to read the ESRI .prj file
        Dim re As StreamReader = File.OpenText(prjfile)
        'sets the ending PCS to the ESRI .prj file
        pESRIEnd.ParseEsriString(re.ReadLine())
        'calls the reprojection function
        Reproject.ReprojectPoints(xy, z, pESRIEnd, pStart, 0, 1)


        Dim newnums As TwoVars

        newnums.Var1 = xy(1)
        newnums.Var2 = xy(0)
        Return newnums

    End Function

    Public Structure TwoVars

        Public Var1 As Double
        Public Var2 As Double

    End Structure

End Class

Public Class Meerkat4

    Inherits GH_Component
    ''' <summary>
    ''' Initializes a new instance of the $safeitemrootname$ class.
    ''' </summary>
    ''' 
    Public Sub New()
        MyBase.New("Point to Lat-Long", "Point to Lat-Long", "Projects point to Latitude and Longitude.", "Extra", "Meerkat")
    End Sub

    ''' <summary>
    ''' Registers all the input parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterInputParams(pManager As GH_Component.GH_InputParamManager)

        pManager.AddPointParameter("Projected Point", "p", "Point projected from lat and long.", GH_ParamAccess.item)
        pManager.AddTextParameter("Projection File", "prj", "Projection file", GH_ParamAccess.item)
    End Sub
    ''' <summary>
    ''' Registers all the output parameters for this component.
    ''' </summary>
    Protected Overrides Sub RegisterOutputParams(pManager As GH_Component.GH_OutputParamManager)
        'pManager.AddIntegerParameter("Product", "P", "Product of all inputs", GH_ParamAccess.item)
        pManager.AddNumberParameter("Latitude", "Lat", "Latitude to project", GH_ParamAccess.item)
        pManager.AddNumberParameter("Longitude", "Long", "Longitude to project", GH_ParamAccess.item)
    End Sub

    ''' <summary>
    ''' This is the method that actually does the work.
    ''' </summary>
    ''' <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
    Protected Overrides Sub SolveInstance(DA As IGH_DataAccess)

        Dim SPoint As New Point3d()
        Dim Prj As String = Nothing

        If Not DA.GetData(0, SPoint) Then Return
        If Not DA.GetData(1, Prj) Then Return

        Dim Lat As Double
        Dim Lng As Double

        Lat = Projectpoints(SPoint.X, SPoint.Y, Prj).Var1
        Lng = Projectpoints(SPoint.X, SPoint.Y, Prj).Var2

        DA.SetData(0, Lat)
        DA.SetData(1, Lng)

    End Sub

    ''' <summary>
    ''' Provides an Icon for every component that will be visible in the User Interface.
    ''' Icons need to be 24x24 pixels.
    ''' </summary>
    Protected Overrides ReadOnly Property Icon() As System.Drawing.Bitmap
        Get
            'You can add image files to your project resources and access them like this:
            Return My.Resources.LL
            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets the unique ID for this component. Do not change this ID after release.
    ''' </summary>
    Public Overrides ReadOnly Property ComponentGuid() As Guid
        Get
            Return New Guid("{B209ACD1-B880-B11F-B516-DFE90EE4588F}")
        End Get
    End Property


    Function Projectpoints(ByVal xpt As Long, ByVal ypt As Long, ByVal prjfile As String) As TwoVars

        'declares a new ProjectionInfo for the startind and ending coordinate systems
        'sets the start GCS to WGS_1984
        Dim pStart As ProjectionInfo = KnownCoordinateSystems.Geographic.World.WGS1984
        Dim pESRIEnd As New ProjectionInfo()
        'declares the point(s) that will be reprojected
        Dim xy As Double() = New Double(1) {}
        Dim z As Double() = New Double(0) {}

        xy(0) = xpt
        xy(1) = ypt
        z(0) = 0

        'initiates a StreamReader to read the ESRI .prj file
        Dim re As StreamReader = File.OpenText(prjfile)
        'sets the ending PCS to the ESRI .prj file
        pESRIEnd.ParseEsriString(re.ReadLine())
        'calls the reprojection function
        Reproject.ReprojectPoints(xy, z, pESRIEnd, pStart, 0, 1)


        Dim newnums As TwoVars

        newnums.Var1 = xy(1)
        newnums.Var2 = xy(0)
        Return newnums

    End Function

    Public Structure TwoVars

        Public Var1 As Double
        Public Var2 As Double

    End Structure

End Class