Imports System.IO
Imports System
Imports System.Collections
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports EGIS.ShapeFileLib
Imports DotSpatial.Projections
Imports System.Drawing
Imports System.Text
Imports System.Math
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Grasshopper

<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
Public Class Form2

    Private treeViewWithToolTips As TreeView

    Private Sub Form2_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Call removehtm()
    End Sub


    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        TreeView1.ShowNodeToolTips = True
        WebBrowser1.ObjectForScripting = Me
        WebBrowser1.WebBrowserShortcutsEnabled = False
        Call loadhtm()
        WebBrowser1.DocumentText = My.Computer.FileSystem.ReadAllText(Grasshopper.Folders.DefaultAssemblyFolder & "GoogleMap.htm")

        Button2.Hide()

        If Not My.Settings.StoreSW = Nothing Then
            TextBox1.Text = My.Settings.StoreSW
            TextBox2.Text = My.Settings.StoreNE
        End If

        'If Now > CDate("1-1-14 12:00:00") Then
        '    MessageBox.Show("Trial period expired. See Food4Rhino for update.", "Identical File Path", MessageBoxButtons.OK)
        '    System.Windows.Forms.Application.Exit()
        'End If
        'MessageBox.Show("File expires on 1/1/2014")

        DeleteMToolStripMenuItem.Enabled = False
        ToolStripTextBox1.ForeColor = Color.LightGray
        TreeView1.Nodes.Add("| Shape Files |")
        HideCheckBox(TreeView1, TreeView1.Nodes(0))

        TreeView1.Nodes.Add("| Views |")
        HideCheckBox(TreeView1, TreeView1.Nodes(1))

    End Sub
    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Dim InitialZoom As Integer
        Dim InitialLatitude As Double
        Dim InitialLongitude As Double
        Dim InitialMapType As GoogleMapType

        InitialZoom = 14
        InitialLatitude = 47.6037
        InitialLongitude = -122.3347
        InitialMapType = GoogleMapType.Satellite

        WebBrowser1.Document.InvokeScript("Initialize", New Object() {InitialZoom, InitialLatitude, InitialLongitude, CInt(InitialMapType)})
    End Sub

    Public Sub Map_MouseMove(ByVal lat As Double, ByVal lng As Double)

        'Called from the GoogleMap.htm script when ever the mouse is moved.
        ToolStripStatusLabel1.Text = "lat/lng: " & CStr(Math.Round(lat, 4)) & " , " & CStr(Math.Round(lng, 4))

    End Sub

    Public Enum GoogleMapType
        None
        RoadMap
        Terrain
        Hybrid
        Satellite
    End Enum


    Public Sub Map_Click(ByVal lat As Double, ByVal lng As Double)
        'Add a marker to the map.

        'Dim MarkerName As String = InputBox("Enter a Marker Name", "New Marker")
        'If Not String.IsNullOrEmpty(MarkerName) Then
        '    'The script function AddMarker takes three parameters
        '    'name,lat,lng.  Call the script passing the parameters in.
        '    WebBrowser1.Document.InvokeScript("AddMarker", New Object() {MarkerName, lat, lng})
        'End If
    End Sub

    Public Sub Map_Idle()
        'Would be a good place to load your own custom markers
        'from data source
    End Sub

    Public Sub ShowShpBounds(ByVal S As Double, ByVal W As Double, ByVal N As Double, ByVal E As Double, ByVal colors As Color, ByVal zoomBool As String)
        'ByVal latSW As Double, ByVal lngSW As Double, ByVal latNE As Double, ByVal lngNE As Double
        'WebBrowser1.Document.InvokeScript("AddRectangle", New Object() {latSW, lngSW, latNE, lngNE})

        WebBrowser1.Document.InvokeScript("AddRectangle", New Object() {S, W, N, E, ColorTranslator.ToHtml(colors).ToString, zoomBool})

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click


        Dim FD1 As New FolderBrowserDialog
        FD1.Description = "Choose location for output files.  New files will overwrite files with the same name in the same folder."
        If Not My.Settings.WorkFolder = Nothing Then
            FD1.SelectedPath = My.Settings.WorkFolder
        End If

        If FD1.ShowDialog() = DialogResult.OK Then
            My.Settings.WorkFolder = FD1.SelectedPath
            My.Settings.Save()
        Else
            Exit Sub
        End If

        For Each n As TreeNode In GetCheck(TreeView1.Nodes)

            For Each s As ShapeFile In shapefiles
                If n.Name = s.Names Then

                    Dim sw As New StreamWriter(My.Settings.WorkFolder & "\" & Path.GetFileNameWithoutExtension(n.Name) & ".mkgis")
                    sw.Write(HTMLcoordsToCrop(s))
                    sw.Close()

                End If
            Next

        Next
    End Sub
    Public Sub Rect_Call(ByVal SWhtml As String, ByVal NEhtml As String)

        'Called from the GoogleMap.htm script when selection rectangle is drawn.

        SWpass = SWhtml
        NEpass = NEhtml

        My.Settings.StoreSW = SWpass
        My.Settings.StoreNE = NEpass
        My.Settings.Save()

        TextBox1.Text = SWhtml
        TextBox2.Text = NEhtml
        Button2.Show()

    End Sub

    Private Function GetCheck(ByVal node As TreeNodeCollection) As ArrayList

        Dim lN As New ArrayList
        For Each n As TreeNode In node
            If n.Checked Then lN.Add(n)
            lN.AddRange(GetCheck(n.Nodes))
        Next
        Return lN

    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Add shapefiles to treeview

        Dim OF1 As New OpenFileDialog


        With OF1

            If My.Settings.LastFolder = Nothing Then

                .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                My.Settings.LastFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            Else
                .InitialDirectory = My.Settings.LastFolder
            End If
            .Multiselect = True
            .Title = "Please select a shape file to add to the map."
            .Filter = "Shape files (*.shp)|*.shp"

            If .ShowDialog = DialogResult.OK Then

                My.Settings.LastFolder = Path.GetDirectoryName(.FileName)
                My.Settings.Save()

                For Each s As String In OF1.FileNames

                    For Each sf As ShapeFile In shapefiles
                        If s = sf.Names Then
                            MessageBox.Show("Only one instance of a shapefile allowed!", "Identical File Path", MessageBoxButtons.OK)
                            GoTo Skip_Add
                        End If
                    Next

                    If Not File.Exists(Path.GetDirectoryName(s) & "\" & Path.GetFileNameWithoutExtension(s) & ".shx") Then
                        MessageBox.Show(Path.GetFileName(s) & " is missing its .shx file.", "Shape File associated files", MessageBoxButtons.OK)
                        GoTo Skip_Add
                    End If

                    If Not File.Exists(Path.GetDirectoryName(s) & "\" & Path.GetFileNameWithoutExtension(s) & ".dbf") Then
                        MessageBox.Show(Path.GetFileName(s) & " is missing its .dbf file.", "Shape File associated files", MessageBoxButtons.OK)
                        GoTo Skip_Add
                    End If

                    Dim prjPass As String

                    If File.Exists(Path.GetDirectoryName(s) & "\" & Path.GetFileNameWithoutExtension(s) & ".prj") Then
                        prjPass = Path.GetDirectoryName(s) & "\" & Path.GetFileNameWithoutExtension(s) & ".prj"
                    Else
                        Dim OF2 As New OpenFileDialog

                        With OF2


                            If My.Settings.LastFolder2 = Nothing Then

                                .InitialDirectory = My.Settings.LastFolder
                                My.Settings.LastFolder2 = My.Settings.LastFolder
                            Else
                                .InitialDirectory = My.Settings.LastFolder2
                            End If
                            .Multiselect = False
                            .Title = "Please select a projection file (.prj) for " & Path.GetFileName(s)
                            .Filter = "Shape files (*.prj)|*.prj"

                            If .ShowDialog = DialogResult.OK Then

                                My.Settings.LastFolder2 = Path.GetDirectoryName(.FileName)
                                My.Settings.Save()

                                Dim s2 As String
                                s2 = OF2.FileName
                                prjPass = s2
                            Else

                                MessageBox.Show("Could not locate .prj for " & Path.GetFileName(s), "Shape File projection file", MessageBoxButtons.OK)
                                GoTo Skip_Add

                            End If
                        End With
                    End If

                    Dim sChecktype As New EGIS.ShapeFileLib.ShapeFile

                    Try

                        sChecktype.LoadFromFile(s)

                    Catch ex As Exception

                        GoTo Failed_Add

                    End Try

                    GoTo Successful
Failed_Add:
                    MessageBox.Show(Path.GetFileName(s) & " is not a supported shape file format.", "Shape File type", MessageBoxButtons.OK)
                    GoTo Skip_Add

Successful:
                    'Dim ok_Types As New List(Of String)

                    ''Easy GIS currently accepts Point, PointZ, Polygon, PolygonZ, PolyLine, PolyLineM, MultiPoint And MultiPointZ
                    'ok_Types.Add("Point")
                    'ok_Types.Add("PointZ")
                    'ok_Types.Add("Polygon")
                    'ok_Types.Add("PolygonZ")
                    'ok_Types.Add("PolyLine")
                    'ok_Types.Add("PolyLineM")
                    'ok_Types.Add("MultiPoint")
                    'ok_Types.Add("MultiPointZ")

                    Dim shape = New ShapeFile(s, prjPass)
                    Dim NewNode As TreeNode = Me.TreeView1.Nodes(0).Nodes.Add(Path.GetFileName(shape.Names))

                    NewNode.ToolTipText = "Right Click to Delete"
                    NewNode.ForeColor = shape.colors
                    NewNode.Name = shape.Names
                    NewNode.Text = Path.GetFileName(shape.Names)
                    TreeView1.Nodes(0).Expand()
                    shapefiles.Add(shape)
Skip_Add:

                Next
            End If
        End With

    End Sub

    Private Sub TreeView1_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterCheck


    End Sub




    Private Sub TreeView1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TreeView1.KeyPress

        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Delete) Then

            If TreeView1.SelectedNode.Text = "| Shape Files |" Or TreeView1.SelectedNode.Text = "| Views |" Then

                Exit Sub

            End If

            If Not TreeView1.SelectedNode Is Nothing Then
                TreeView1.SelectedNode.Remove()
            End If
        End If
    End Sub


    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick


        If e.Node.Text = "| Shape Files |" Or e.Node.Text = "| Views |" Then
            Exit Sub
        End If

        If e.Node.Name.StartsWith("View:: ") And e.Button() = Windows.Forms.MouseButtons.Right Then
            Dim result2 = MessageBox.Show("Remove reference view from Meerkat?", "Remove", MessageBoxButtons.YesNo)
            If result2 = DialogResult.No Then
                Exit Sub
            ElseIf result2 = DialogResult.Yes Then
                e.Node.Remove()
                Exit Sub
            End If
        ElseIf e.Node.Name.StartsWith("View:: ") And e.Button() = Windows.Forms.MouseButtons.Left Then
            WebBrowser1.Document.InvokeScript("codeAddress", New Object() {e.Node.Text.Remove(0, 6)})
            Exit Sub
        End If


        If e.Button() = Windows.Forms.MouseButtons.Right Then
            Dim result = MessageBox.Show("Remove shape file from Meerkat?", "Remove", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                Exit Sub
            ElseIf result = DialogResult.Yes Then

                'REMOVE NODE IN TREE VIEW
                e.Node.Remove()

                'REMOVE CLASS INSTANCE OF SHAPEFILE
                For Each sf As ShapeFile In shapefiles
                    If e.Node.Name() = sf.Names Then

                        sf.Dispose()

                    End If
                Next

                'REMOVE ITEM IN LIST OF SHAPEFILES
                shapefiles.RemoveAll(Function(item) item.Names = e.Node.Name)

            End If
        End If

        If e.Button() = Windows.Forms.MouseButtons.Left Then
            Dim zoomBool As Boolean

            WebBrowser1.Document.InvokeScript("DeleteRects")
            Dim n As TreeNode


            zoomBool = True


            For Each n In TreeView1.Nodes(0).Nodes

                If n.Checked = True Then

                    If n.Name.StartsWith("View:: ") Then
                        n.Checked = False
                        GoTo Nextnode
                    End If

                    For Each sf As ShapeFile In shapefiles


                        If n.Name() = sf.Names Then
                            Dim sfx As New EGIS.ShapeFileLib.ShapeFile
                            sfx.LoadFromFile(sf.Names)

                            Dim shpbds As New RectangleF
                            shpbds = sfx.GetActualExtent()
                            Dim S1 As Double
                            Dim W1 As Double
                            Dim N1 As Double
                            Dim E1 As Double

                            W1 = Projectpoints(shpbds.Left, shpbds.Bottom, sf.prj).Var2
                            S1 = Projectpoints(shpbds.Left, shpbds.Bottom, sf.prj).Var1
                            N1 = Projectpoints(shpbds.Right, shpbds.Top, sf.prj).Var1
                            E1 = Projectpoints(shpbds.Right, shpbds.Top, sf.prj).Var2

                            ShowShpBounds(S1, W1, N1, E1, sf.colors, zoomBool)
                        End If
                    Next
                End If
Nextnode:
            Next
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim SW() As String
        Dim NE() As String

        Dim S As Double
        Dim W As Double
        Dim N As Double
        Dim E1 As Double

        Try
            SW = Split(TextBox1.Text, ",")
            NE = Split(TextBox2.Text, ",")

            S = CDbl(SW(0))
            W = CDbl(SW(1))
            N = CDbl(NE(0))
            E1 = CDbl(NE(1))

        Catch

            MessageBox.Show("Latitude and Longitude are not formatted correctly... Returning to last successful bounds.", "Latitude and Longitude format", MessageBoxButtons.OK)
            TextBox1.Text = My.Settings.StoreSW
            TextBox2.Text = My.Settings.StoreNE
            Exit Sub
        End Try

        WebBrowser1.Document.InvokeScript("adjustRect", New Object() {S, W, N, E1})
    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click


        If ToolStripTextBox1.Text = "<enter location here>" Then
            ToolStripTextBox1.Text = ""
            ToolStripTextBox1.ForeColor = Color.Black

        End If


    End Sub


    Private Sub ToolStripTextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ToolStripTextBox1.KeyPress

        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then
            SendAddress()


            Dim stext As String = ToolStripTextBox1.Text
            Dim NewNode As TreeNode = Me.TreeView1.Nodes(1).Nodes.Add(stext)

            NewNode.ToolTipText = "Left Click to Zoom to Location, Right Click to Delete"
            NewNode.ForeColor = Color.Black
            NewNode.Name = "View:: " & stext
            NewNode.Checked = False
            NewNode.Text = "view: " & stext
            HideCheckBox(TreeView1, NewNode)
            TreeView1.Nodes(1).Expand()
        End If
    End Sub

    Public Sub SendAddress()
        'ByVal latSW As Double, ByVal lngSW As Double, ByVal latNE As Double, ByVal lngNE As Double
        'WebBrowser1.Document.InvokeScript("AddRectangle", New Object() {latSW, lngSW, latNE, lngNE})

        WebBrowser1.Document.InvokeScript("codeAddress", New Object() {ToolStripTextBox1.Text})
        DeleteMToolStripMenuItem.Enabled = True
    End Sub


    Private Sub DeleteMToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteMToolStripMenuItem.Click
        WebBrowser1.Document.InvokeScript("DeleteMarkers")
        DeleteMToolStripMenuItem.Enabled = False
    End Sub
    Private Const TVIF_STATE As Integer = &H8
    Private Const TVIS_STATEIMAGEMASK As Integer = &HF000
    Private Const TV_FIRST As Integer = &H1100
    'Friend WithEvents Button1 As System.Windows.Forms.Button
    Private Const TVM_SETITEM As Integer = TV_FIRST + 63

    <StructLayout(LayoutKind.Sequential, Pack:=8, CharSet:=CharSet.Auto)>
    Private Structure TVITEM
        Public mask As Integer
        Public hItem As IntPtr
        Public state As Integer
        Public stateMask As Integer
        <MarshalAs(UnmanagedType.LPTStr)>
        Public lpszText As String
        Public cchTextMax As Integer
        Public iImage As Integer
        Public iSelectedImage As Integer
        Public cChildren As Integer
        Public lParam As IntPtr
    End Structure

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByRef lParam As TVITEM) As IntPtr
    End Function

    ''' <summary>
    ''' Hides the checkbox for the specified node on a TreeView control.
    ''' </summary>
    Private Sub HideCheckBox(ByVal tvw As TreeView, ByVal node As TreeNode)
        Dim tvi As New TVITEM()
        tvi.hItem = node.Handle
        tvi.mask = TVIF_STATE
        tvi.stateMask = TVIS_STATEIMAGEMASK
        tvi.state = 0
        SendMessage(tvw.Handle, TVM_SETITEM, IntPtr.Zero, tvi)
    End Sub

    Public shapefiles As New List(Of ShapeFile)

    Public Structure sfInfo

        Public shapefileBounds As RectangleF

    End Structure

    Public Structure TwoVars

        Public Var1 As Double
        Public Var2 As Double

    End Structure

    Public Function GetShapefileInfo(ByVal source As String) As sfInfo

        Dim activeShape As sfInfo
        Dim sf As New EGIS.ShapeFileLib.ShapeFile

        sf.LoadFromFile(source)
        Dim ShapeRect As RectangleF
        ShapeRect = sf.GetActualExtent()
        activeShape.shapefileBounds = ShapeRect

        Return activeShape

    End Function

    Function CoordstoPts(ByVal LatitudeLongitude As String, ByVal prjfile As String) As TwoVars

        Dim LatLong As Array
        Dim Latitude As Double
        Dim Longitude As Double

        LatLong = Split(LatitudeLongitude, ",")

        Latitude = CDbl(LatLong(0))
        Longitude = CDbl(LatLong(1))


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
    Public SWpass As String
    Public NEpass As String

    Function HTMLcoordsToCrop(ByRef cnode As ShapeFile)

        Dim SW1 As String
        Dim NE1 As String

        SW1 = SWpass
        NE1 = NEpass

        Dim W As Double = CoordstoPts(SW1, cnode.prj).Var1
        Dim S As Double = CoordstoPts(SW1, cnode.prj).Var2
        Dim E As Double = CoordstoPts(NE1, cnode.prj).Var1
        Dim N As Double = CoordstoPts(NE1, cnode.prj).Var2

        Dim sf As New EGIS.ShapeFileLib.ShapeFile
        Dim sb As New StringBuilder
        Dim sbfinal As New StringBuilder

        Dim db As New DbfReader(Path.GetDirectoryName(cnode.Names) & "\" & Path.GetFileNameWithoutExtension(cnode.Names) & ".dbf")

        sf.LoadFromFile(cnode.Names)

        Dim Rboundary As New RectangleF(W, S, Abs(E - W), Abs(N - S))
        Dim sfEnum As ShapeFileEnumerator = sf.GetShapeFileEnumerator(Rboundary, ShapeFileEnumerator.IntersectionType.Intersects)
        Dim countIncluded As Integer = 0

        'for cleaned below
        Dim fieldnames As String = String.Join("|", db.GetFieldNames())
        Dim cleaned2 As String = Regex.Replace(fieldnames, "\s{2,}", " ")

        sbfinal.AppendLine("Date Cropped: " & Now)
        sbfinal.AppendLine("By User: " & Environment.UserName & " On Machine Name: " & Environment.MachineName)
        sbfinal.AppendLine("File Path: " & sf.FilePath.ToString)
        sbfinal.AppendLine("Shape Type: " & sf.ShapeType.ToString)
        sbfinal.AppendLine("Record Count: " & sf.RecordCount.ToString)
        sbfinal.AppendLine("Bounds: " & S & "|" & W & "|" & N & "|" & E)
        sbfinal.AppendLine("Bounds (SW corner and NE corner Lat/Lng ): " & My.Settings.StoreSW & " " & My.Settings.StoreNE)
        sbfinal.AppendLine("Field Names: " & cleaned2)

        While sfEnum.MoveNext()
            Dim cItem As Integer = sfEnum.CurrentShapeIndex
            countIncluded = countIncluded + 1

            Dim points As PointD() = sfEnum.Current(0)
            Dim ptlist As New StringBuilder

            For j = 0 To points.Length - 1
                ptlist.Append(points(j).ToString & " ")
            Next

            Dim result As String = String.Join("|", db.GetFields(cItem))
            Dim cleaned As String = Regex.Replace(result, "\s{2,}", " ")

            sb.AppendLine(sfEnum.CurrentShapeIndex)
            sb.AppendLine(cleaned)
            sb.AppendLine(ptlist.ToString)

            ptlist.Clear()

        End While

        sbfinal.AppendLine("Record Count in Crop: " & countIncluded.ToString)
        sbfinal.AppendLine("")
        sbfinal.Append(sb)

        Return sbfinal

    End Function

    Public Sub loadhtm()
        Dim GoogleHTML As String = <![CDATA[
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="initial-scale=1.0, user-scalable=no" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <style type="text/css">
        html
        {
            height: 100%;
        }
        body
        {
            height: 100%;
            margin: 0;
            padding: 0;
        }
        #map_canvas
        {
            height: 100%;
        }
    </style>

  <script type="text/javascript"
       src="https://maps.googleapis.com/maps/api/js?key=_replace_me_&sensor=false&libraries=drawing">
  </script>
    <script type="text/javascript">

      var rect;
      var rectangle2;

      var dog = "nothing yet";
      var hog = "nothing yet";


      var geocoder;
      var map;
      var markers = [];
      var Rectangles = [];

      function Initialize(zoomLevel,lat,lng,type){

      geocoder = new google.maps.Geocoder();

      var MapType;
      switch (type)
      {
      case 1:
      MapType = google.maps.MapTypeId.ROADMAP;
      break;
      case 2:
      MapType = google.maps.MapTypeId.TERRAIN;
      break;
      case 3:
      MapType = google.maps.MapTypeId.HYBRID;
      break;
      case 4:
      MapType = google.maps.MapTypeId.SATELLITE;
      break;
      default:
      MapType = google.maps.MapTypeId.ROADMAP;
      };


      var myLatlng = new google.maps.LatLng(lat,lng);
      var myOptions = {zoom: zoomLevel,center: myLatlng,mapTypeId: MapType};
      var MarkerSize = new google.maps.Size(48,48);

      map = new google.maps.Map(document.getElementById("map_canvas"), myOptions);

      map.setTilt(0);

      var drawingManager = new google.maps.drawing.DrawingManager({

      drawingControl: true,
      drawingControlOptions: {
      position: google.maps.ControlPosition.TOP_CENTER,
      drawingModes: [

      google.maps.drawing.OverlayType.RECTANGLE
      ]
      },

      rectangleOptions: {
      fillColor: '#ffff00',
      fillOpacity: .2,
      draggable: true,
      strokeWeight: .5,
      clickable: true,
      editable: true,
      zIndex: 1
      }
      });

      drawingManager.setMap(map);


      google.maps.event.addListener(map, 'click', Map_Click);
      google.maps.event.addListener(map, 'mousemove', Map_MouseMove);
      google.maps.event.addListener(map, 'idle',Map_Idle);



      google.maps.event.addListener(drawingManager, 'rectanglecomplete', function(subject){

      if (rect) {
      google.maps.event.clearInstanceListeners(rect);
      rect.setMap(null);
      }

      if (rectangle2) {
      google.maps.event.clearInstanceListeners(rectangle2);
      rectangle2.setMap(null);
      }

      google.maps.event.addListener(subject, "bounds_changed", scheduleDelayedCallback);


      google.maps.event.addDomListener(window, 'load', Initialize);

      rect = subject;

      hog = rect.getBounds().getSouthWest().toUrlValue().toString();
      dog = rect.getBounds().getNorthEast().toUrlValue().toString();

      window.external.Rect_Call(hog,dog);

      });

      }

      function fireIfLastEvent()
      {
      if (lastEvent.getTime() + 200 <= new Date().getTime())
        {

        hog = rect.getBounds().getSouthWest().toUrlValue().toString();
        dog = rect.getBounds().getNorthEast().toUrlValue().toString();

        window.external.Rect_Call(hog,dog);
        }
        }

        function scheduleDelayedCallback()
        {
        lastEvent = new Date();
        setTimeout(fireIfLastEvent, 200);
        }


        function adjustRect(S,W,N,E)
        {

        if (rect) {
        google.maps.event.clearInstanceListeners(rect);
        rect.setMap(null);
        }

        if (rectangle2) {
        google.maps.event.clearInstanceListeners(rectangle2);
        rectangle2.setMap(null);
        }

        var RecLatLngSW1 = new google.maps.LatLng(S,W);
        var RecLatLngNE1 = new google.maps.LatLng(N,E);

        var RectangleOption2 = ({
        fillColor: '#ffff00',
        fillOpacity: .2,
        draggable: false,
        strokeWeight: .5,
        clickable: false,
        editable: false,
        map: map,
        bounds: new google.maps.LatLngBounds(RecLatLngSW1,RecLatLngNE1)
        });


        rectangle2 = new google.maps.Rectangle(RectangleOption2);


        hog = rectangle2.getBounds().getSouthWest().toUrlValue().toString();
        dog = rectangle2.getBounds().getNorthEast().toUrlValue().toString();


        window.external.Rect_Call(hog,dog);
        var bounds2 = new google.maps.LatLngBounds(RecLatLngSW1,RecLatLngNE1);
        map.fitBounds(bounds2);

        }


        function Map_Click(e){
        window.external.Map_Click(e.latLng.lat(),e.latLng.lng());
        }

        function Map_MouseMove(e){
        window.external.Map_MouseMove(e.latLng.lat(),e.latLng.lng());
        }

        function Map_Idle(){
        window.external.Map_Idle();
        }

        function DeleteMarkers()
        {
        if (markers){
        for (i in markers){
        markers[i].setMap(null);
        google.maps.event.clearInstanceListeners(markers[i]);
        markers[i] = null;
        }
        markers.length = 0;
        }
        }

        function DeleteRects()
        {
        if (Rectangles){
        for (i in Rectangles){
        Rectangles[i].setMap(null);
        google.maps.event.clearInstanceListeners(Rectangles[i]);
        Rectangles[i] = null;
        }
        Rectangles.length = 0;
        }
        }

        function AddRectangle(latSW,lngSW,latNE,lngNE,color,zoomBool)
        {
        var RecLatLngSW = new google.maps.LatLng(latSW, lngSW);
        var RecLatLngNE = new google.maps.LatLng(latNE, lngNE);

        if (zoomBool == "True")
        {
        var bounds1 = new google.maps.LatLngBounds(RecLatLngSW,RecLatLngNE);
        map.fitBounds(bounds1);
        }

        var RectangleOption = ({
        strokeColor: color,
        strokeOpacity: 0.8,
        strokeWeight: 2,
        fillColor: 'rgb(0,0,0)',
        fillOpacity: 0.0,
        map: map,
        bounds: new google.maps.LatLngBounds(
        RecLatLngSW,
        RecLatLngNE)
        });


        var rectangle = new google.maps.Rectangle(RectangleOption);
        Rectangles.push(rectangle);
        RecLatLngSW = null;
        RecLatLngNE = null;
        RectangleOption = null;
        }


        function codeAddress(formAddress) {
        //var address = document.getElementById('address').value;
        var address = formAddress;
        geocoder.geocode( { 'address': address}, function(results, status) {
        if (status == google.maps.GeocoderStatus.OK) {
        map.setZoom(16);
        map.setCenter(results[0].geometry.location);

        var marker = new google.maps.Marker({
        map: map,
        position: results[0].geometry.location,
        title: address
        });
        markers.push(marker);


        } else {
        alert('Geocode was not successful for the following reason: ' + status);
        }
        });
        }

      </script>
    
</head>
<body>
    
  <div id="map_canvas" style="width: 100%; height: 100%">
</body>
</html>


]]>.Value()

        Dim apiKey = File.ReadAllText(Grasshopper.Folders.DefaultAssemblyFolder & "google_api_key_mkgis.txt").Trim()
        Dim newGoogleHTML = Replace(GoogleHTML, "_replace_me_", apiKey)

        Dim fs As New StreamWriter(Grasshopper.Folders.DefaultAssemblyFolder & "GoogleMap.htm")
        fs.WriteLine(newGoogleHTML)
        fs.Close()


    End Sub
    Sub removehtm()
        If File.Exists(Grasshopper.Folders.DefaultAssemblyFolder & "GoogleMap.htm") Then
            File.Delete(Grasshopper.Folders.DefaultAssemblyFolder & "GoogleMap.htm")
        End If
    End Sub

    Private Sub ToolStripSplitButton2_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton2.ButtonClick

        MsgBox("Meerkat GIS makes extensive use of EasyGIS. Please visit http://www.easygisdotnet.com/About.aspx for more information about this great library.")
        MsgBox("Meerkat GIS also makes extensive use of DotSpatial. Please visit http://dotspatial.codeplex.com/ for more information about this great library.")

    End Sub
End Class


Public Class ShapeFile

    Implements IDisposable


    Private Name As String
    Private bounds As String
    Private mycolor As Color
    Private myPrj As String

    Dim m_Rnd As New Random(CInt(Date.Now.Ticks Mod Integer.MaxValue))
    Public Function RandomRGBColor() As Color
        Return Color.FromArgb(255, _
            m_Rnd.Next(0, 255), _
            m_Rnd.Next(0, 255), _
            m_Rnd.Next(0, 255))
    End Function

    Public Sub New(ByVal strfilepath As String, ByVal projPath As String)

        Me.Name = strfilepath
        Me.myPrj = projPath

        Dim chooseColors As New List(Of Color)


        Me.mycolor = RandomRGBColor()

    End Sub 'New 
    Public ReadOnly Property Names() As String
        Get
            Return Name
        End Get
    End Property


    Public ReadOnly Property colors() As Color
        Get
            Return mycolor
        End Get
    End Property
    Public ReadOnly Property prj() As String
        Get
            Return myPrj
        End Get
    End Property




#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class 'Shapefile
