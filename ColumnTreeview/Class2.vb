Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.ComponentModel
Imports Controls

Public Enum NodeRegion
    Expander
    Favorite
    CheckBox
    Image
    Node
End Enum
Public Class HitRegion
    Implements IEquatable(Of HitRegion)
    Public Property Region As NodeRegion
    Public Property Node As Node
    Public Overrides Function GetHashCode() As Integer
        Return Region.GetHashCode Xor Node.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As HitRegion) As Boolean Implements IEquatable(Of HitRegion).Equals
        If other Is Nothing Then
            Return Me Is Nothing
        Else
            Return Region = other.Region AndAlso Node Is other.Node
        End If
    End Function
    Public Shared Operator =(ByVal value1 As HitRegion, ByVal value2 As HitRegion) As Boolean
        If value1 Is Nothing Then
            Return value2 Is Nothing
        Else
            Return value1.Equals(value2)
        End If
    End Operator
    Public Shared Operator <>(ByVal value1 As HitRegion, ByVal value2 As HitRegion) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is HitRegion Then
            Return CType(obj, HitRegion) = Me
        Else
            Return False
        End If
    End Function
End Class
Public Class NodeEventArgs
    Inherits EventArgs
    Public ReadOnly Property Node As Node
    Public ReadOnly Property Nodes As List(Of Node)
    Public ReadOnly Property ProposedText As String = String.Empty
    Public Sub New(ByVal Value As Node)
        Node = Value
    End Sub
    Public Sub New(ByVal Values As List(Of Node))
        Nodes = Values
    End Sub
    Public Sub New(ByVal Value As Node, NewText As String)
        Node = Value
        ProposedText = NewText
    End Sub
End Class
Public Class TreeViewer
    Inherits Control
    Private WithEvents Karen As New Hooker
    Public WithEvents VScroll As New VScrollBar
    Public WithEvents HScroll As New HScrollBar
    Private WithEvents NodeTimer As New Timer With {.Interval = 200}
    Private WithEvents CursorTimer As New Timer With {.Interval = 300}
    Private WithEvents ScrollTimer As New Timer With {.Interval = 50}
#Region " TREEVIEW GLOBAL FUNCTIONS (CMS) "
    Private WithEvents TSDD_Options As New ToolStripDropDown With {.AutoClose = False, .Padding = New Padding(0), .DropShadowEnabled = True, .BackColor = Color.Transparent}
    Private WithEvents B_Arrow As New Button With {.Margin = New Padding(0), .FlatStyle = FlatStyle.Flat}
    Private WithEvents B_Book As New Button With {.Margin = New Padding(0), .FlatStyle = FlatStyle.Flat}
    Private WithEvents B_Default As New Button With {.Margin = New Padding(0), .FlatStyle = FlatStyle.Flat}
    Private WithEvents B_Light As New Button With {.Margin = New Padding(0), .FlatStyle = FlatStyle.Flat}
    Private WithEvents TSMI_ExpandCollapseStyles As New ToolStripMenuItem With {.Text = "Node ± Styles"}
    Private WithEvents TLP_NodeStyleButtons As New TableLayoutPanel With {.Size = New Size(36, 36), .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset}
    Private WithEvents TSMI_ExpandCollapseAll As New ToolStripMenuItem With {.Text = "Expand All"}

    Private WithEvents TSMI_Checkboxes As New ToolStripMenuItem With {.Text = "Checkboxes", .Checked = False, .CheckOnClick = True, .ImageScaling = ToolStripItemImageScaling.None}
    Private WithEvents TSMI_CheckUncheckAll As New ToolStripMenuItem With {.Text = "Check All", .Checked = False, .CheckOnClick = True, .ImageScaling = ToolStripItemImageScaling.None}

    Private WithEvents TSMI_MultiSelect As New ToolStripMenuItem With {.Text = "Multi-Select", .Checked = False, .CheckOnClick = True}
    Private WithEvents TSMI_SelectAll As New ToolStripMenuItem With {.Text = "Select All"}
    Private WithEvents TSMI_SelectNone As New ToolStripMenuItem With {.Text = "Select None"}

    Private WithEvents TSMI_Sort As New ToolStripMenuItem With {.Text = "Click to Sort Ascending"}

    Private WithEvents TSMI_NodeEditing As New ToolStripMenuItem With {.Text = "Node Editing Options"}
    Private ReadOnly TLP_NodePermissions As New TableLayoutPanel With {.CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
        .ColumnCount = 1,
        .RowCount = 3,
        .Size = New Size(200, 90),
        .Margin = New Padding(0)}
    Private WithEvents IC_NodeAdd As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "Add Child Node"}
    Private WithEvents IC_NodeRemove As New ImageCombo With {.Dock = DockStyle.Fill,
        .Text = "Remove Node",
        .Margin = New Padding(0),
        .Mode = ImageComboMode.ColorPicker}
    Private WithEvents IC_NodeEdit As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0)}
#End Region
    Private Const CheckHeight As Integer = 14
    Private Const VScrollWidth As Integer = 14
    Private Const HScrollHeight As Integer = 12
    Private DragData As New DragInfo
    Private _Cursor As Cursor
    Private VisibleIndex As Integer, RollingHeight As Integer, RollingWidth As Integer
    Private ExpandHeight As Integer = 10
    Private ExpandImage As Image, CollapseImage As Image
    Public Event Alert(sender As Object, e As AlertEventArgs)

#Region " STRUCTURES / ENUMS "
    Public Enum CheckState
        None
        All
        Mixed
    End Enum
    Private Structure DragInfo
        Friend MousePoints As List(Of Point)
        Friend IsDragging As Boolean
        Friend DragNode As Node
        Friend DropHighlightNode As Node
    End Structure
#End Region
    Public Sub New()

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)

        BackColor = Color.GhostWhite
#Region " GLOBAL OPTIONS SET-UP "
        With TSDD_Options
#Region " NODE EDITING "
            With TSMI_NodeEditing
                .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                .ImageScaling = ToolStripItemImageScaling.None
                .Image = Base64ToImage(Edit_String)
                With TLP_NodePermissions
                    .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 200})
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 28})
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 28})
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 28})
                    .Controls.Add(IC_NodeEdit, 0, 0)
                    .Controls.Add(IC_NodeAdd, 0, 1)
                    .Controls.Add(IC_NodeRemove, 0, 2)
                End With
                TLP.SetSize(TLP_NodePermissions)
                .DropDownItems.Add(New ToolStripControlHost(TLP_NodePermissions))
                IC_NodeEdit.Image = Base64ToImage(Edit_String)
                IC_NodeEdit.DropDown.SelectionColor = Color.Transparent
                IC_NodeAdd.Image = Base64ToImage(Add_String)
                IC_NodeAdd.DropDown.SelectionColor = Color.Transparent
                IC_NodeRemove.Image = Base64ToImage(Remove_String)
            End With
            .Items.Add(TSMI_NodeEditing)
#End Region
#Region " SORTING "
            .Items.Add(TSMI_Sort)
            TSMI_Sort.Image = Base64ToImage(SortString)
#End Region
#Region " EXPAND/COLLAPSE STYLES "
            With TSMI_ExpandCollapseStyles
                .Image = Base64ToImage(DefaultCollapsed)
                B_Arrow.Image = Base64ToImage(ArrowExpanded)
                B_Book.Image = Base64ToImage(BookOpen)
                B_Default.Image = Base64ToImage(DefaultCollapsed)
                B_Light.Image = Base64ToImage(LightOn)
                With TLP_NodeStyleButtons
                    .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 50})
                    .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 50})
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 50})
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 50})
                    .Controls.Add(B_Arrow, 0, 0)
                    .Controls.Add(B_Book, 1, 0)
                    .Controls.Add(B_Default, 0, 1)
                    .Controls.Add(B_Light, 1, 1)
                End With
                Dim TSCH As New ToolStripControlHost(TLP_NodeStyleButtons)
                .DropDownItems.Add(TSCH)
            End With
            .Items.Add(TSMI_ExpandCollapseStyles)
            TSMI_ExpandCollapseAll.Image = Base64ToImage(DefaultCollapsed)
            .Items.Add(TSMI_ExpandCollapseAll)
#End Region
#Region " MULTI-SELECT "
            TSMI_MultiSelect.DropDownItems.AddRange({TSMI_SelectAll, TSMI_SelectNone})
            .Items.Add(TSMI_MultiSelect)
#End Region
#Region " CHECKBOXES "
            With TSMI_Checkboxes
                .Image = Base64ToImage(CheckString)
                .DropDownItems.Add(TSMI_CheckUncheckAll)
            End With
            .Items.Add(TSMI_Checkboxes)
#End Region
        End With
#End Region
        Controls.AddRange({VScroll, HScroll})
        UpdateExpandImage()

    End Sub
    Protected Overrides Sub InitLayout()
        REM /// FIRES AFTER BEING ADDED TO ANOTHER CONTROL...ADD TREEVIEW AFTER LOADING NODES
        RequiresRepaint()
        MyBase.InitLayout()
    End Sub
    Private Sub WhenParentChanges() Handles Me.ParentChanged
        RequiresRepaint()
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If e IsNot Nothing Then
            With e.Graphics
                'Dim bColor As Color = If(UnRestrictedSize.Height > Height And Not VScroll.Visible, Color.Yellow, BackColor)
                .FillRectangle(New SolidBrush(BackColor), ClientRectangle)
                If BackgroundImage IsNot Nothing Then
                    Dim xOffset As Integer = {ClientRectangle.Width, Math.Abs(Convert.ToInt32((ClientRectangle.Width - BackgroundImage.Width) / 2))}.Min
                    Dim yOffset As Integer = {ClientRectangle.Height, Math.Abs(Convert.ToInt32((ClientRectangle.Height - BackgroundImage.Height) / 2))}.Min
                    Dim ImageBounds As New Rectangle(xOffset, yOffset, {ClientRectangle.Width, BackgroundImage.Width}.Min, {ClientRectangle.Height, BackgroundImage.Height}.Min)
                    .DrawImage(BackgroundImage, ImageBounds)
                End If
                If Ancestors.Any Then
                    Dim drawRootLines As Boolean = RootLines And Ancestors.Count > 1 'Doesn't make sense if only one root, when User wants RootLines
                    Dim firstNode As Node = Ancestors.First
                    Dim verticalLineLeft As Integer = CInt({
                                            firstNode.CheckBounds.Left,
                                            firstNode.FavoriteBounds.Left,
                                            firstNode.ImageBounds.Left,
                                            firstNode.ExpandCollapseBounds.Left
                                                                          }.Min / 2)
                    Using linePen As New Pen(LineColor) With {.DashStyle = LineStyle}
                        For Each Node As Node In Ancestors.Draw
                            With Node
                                If .ColumnIndex = 1 Then
                                    Dim firstHead As ColumnHead = ColumnHeaders(.HeaderLevel)(0)
                                    Dim delta As Integer = firstHead.Width - .Bounds.X - HScroll.Value
                                    ._Bounds.Offset(delta, 0)
                                    ._CheckBounds.Offset(delta, 0)
                                    ._FavoriteBounds.Offset(delta, 0)
                                    ._ExpandCollapseBounds.Offset(delta, 0)
                                    ._ImageBounds.Offset(delta, 0)
                                End If
                                Dim mouseInTip As Boolean = False
                                If .Separator = Node.SeparatorPosition.Above And Node IsNot Ancestors.First Then
                                    Using Pen As New Pen(Color.Blue, 1)
                                        Pen.DashStyle = DashStyle.DashDot
                                        e.Graphics.DrawLine(Pen, New Point(0, .Bounds.Top), New Point(ClientRectangle.Right, .Bounds.Top))
                                    End Using

                                ElseIf .Separator = Node.SeparatorPosition.Below And Node IsNot Ancestors.Last Then
                                    Using Pen As New Pen(Color.Blue, 1)
                                        Pen.DashStyle = DashStyle.DashDot
                                        e.Graphics.DrawLine(Pen, New Point(0, .Bounds.Bottom), New Point(ClientRectangle.Right, .Bounds.Bottom))
                                    End Using

                                End If
                                If .HasChildren Then
                                    If Node.Expanded Then
                                        e.Graphics.DrawImage(CollapseImage, Node.ExpandCollapseBounds)
                                    Else
                                        e.Graphics.DrawImage(ExpandImage, Node.ExpandCollapseBounds)
                                    End If
                                End If
                                If .CanFavorite Then e.Graphics.DrawImage(If(.Favorite, My.Resources.star, My.Resources.starEmpty), .FavoriteBounds)
                                If .CheckBox Then
                                    Using CheckFont As New Font("Marlett", 10)
                                        If .PartialChecked Then
                                            Using Brush As New SolidBrush(Color.FromArgb(192, Color.LightGray))
                                                e.Graphics.FillRectangle(Brush, .CheckBounds)
                                            End Using
                                            TextRenderer.DrawText(e.Graphics, "a".ToString(InvariantCulture), CheckFont, .CheckBounds, Color.DarkGray, TextFormatFlags.NoPadding Or TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
                                        Else
                                            Using Brush As New SolidBrush(Color.White)
                                                e.Graphics.FillRectangle(Brush, .CheckBounds)
                                            End Using
                                            TextRenderer.DrawText(e.Graphics, If(.Checked, "a".ToString(InvariantCulture), String.Empty), CheckFont, .CheckBounds, Color.Blue, TextFormatFlags.NoPadding Or TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
                                        End If
                                    End Using
                                    Using Pen As New Pen(Color.Blue, 1)
                                        e.Graphics.DrawRectangle(Pen, .CheckBounds)
                                    End Using
                                End If
                                If .TipText IsNot Nothing Then
                                    Dim triangleHeight As Single = 8
                                    Dim trianglePoints As New List(Of PointF) From {
                                            New PointF(.Bounds.Left, .Bounds.Top),
                                            New PointF(.Bounds.Left + triangleHeight, .Bounds.Top),
                                            New PointF(.Bounds.Left, .Bounds.Top + triangleHeight)
                                    }
                                    e.Graphics.FillPolygon(Brushes.DarkOrange, trianglePoints.ToArray)
                                    mouseInTip = InTriangle(MousePoint, trianglePoints.ToArray)
                                End If
                                Dim SelectionBounds As New Rectangle(.ImageBounds.Left, .Bounds.Top, .Bounds.Right - { .FavoriteBounds.Left, .ExpandCollapseBounds.Left}.Min, .Bounds.Height)
                                Using Brush As New SolidBrush(If(DragData.DropHighlightNode Is Node, DropHighlightColor, .BackColor))
                                    SelectionBounds.Inflate(-1, -1)
                                    e.Graphics.FillRectangle(Brush, SelectionBounds)
                                End Using
                                If .Image IsNot Nothing Then e.Graphics.DrawImage(.Image, .ImageBounds)
                                TextRenderer.DrawText(e.Graphics,
                                                      If(mouseInTip, .TipText, .Text),
                                                      .Font,
                                                      .Bounds,
                                                      .ForeColor,
                                                      .TextBackColor,
                                                      TextFormatFlags.LeftAndRightPadding Or TextFormatFlags.NoPadding Or TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                                If CurrentMouseNode Is Node And .TipText Is Nothing Then
                                    Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, MouseOverColor))
                                        e.Graphics.FillRectangle(SemiTransparentBrush, .Bounds)
                                    End Using
                                End If
                                If .Selected Then
                                    Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, SelectionColor))
                                        e.Graphics.FillRectangle(SemiTransparentBrush, SelectionBounds)
                                        SelectionBounds.Inflate(1, 1)
                                    End Using
                                    e.Graphics.DrawRectangle(Pens.Black, SelectionBounds)
                                End If

                                Dim ExpandCollapseCenter As Integer = CInt(.ExpandCollapseBounds.Height / 2)
                                Dim VerticalCenter As Integer = .ExpandCollapseBounds.Top + ExpandCollapseCenter
                                Dim NodeHorizontalLeftPoint As Point
                                Dim NodeHorizontalRightPoint As Point
                                Dim NodeVerticalTopPoint As Point
                                Dim NodeVerticalBottomPoint As Point
                                If .HasChildren And .Expanded Then
                                    REM /// VERTICAL LINE DOWN
                                    NodeVerticalTopPoint = New Point(.ExpandCollapseBounds.Left + ExpandCollapseCenter, .ExpandCollapseBounds.Bottom)
                                    NodeVerticalBottomPoint = New Point(.ExpandCollapseBounds.Left + ExpandCollapseCenter, { .Children.Last.ExpandCollapseBounds.Top + ExpandCollapseCenter, ClientRectangle.Height}.Min)
                                    e.Graphics.DrawLine(linePen, NodeVerticalTopPoint, NodeVerticalBottomPoint)
                                End If
                                Dim horizontalLineLeft As Integer = {
                                                .CheckBounds.Left,
                                                .FavoriteBounds.Left,
                                                .ExpandCollapseBounds.Left
                                                                              }.Min
                                If IsNothing(.Parent) Then
                                    If drawRootLines Then e.Graphics.DrawLine(linePen, New Point(verticalLineLeft, VerticalCenter), New Point(horizontalLineLeft, VerticalCenter))
                                Else
                                    REM /// HORIZONTAL LINES LEFT OF EXPAND/COLLAPSE
                                    NodeHorizontalLeftPoint = New Point(.Parent.ExpandCollapseBounds.Left + ExpandCollapseCenter + 1, VerticalCenter)
                                    NodeHorizontalRightPoint = New Point(horizontalLineLeft, VerticalCenter)
                                    e.Graphics.DrawLine(linePen, NodeHorizontalLeftPoint, NodeHorizontalRightPoint)
                                End If

                                If Not .HeaderLevel = 255 Then
                                    Dim leftWidth As Integer() = Node_LeftWidth(.HeaderLevel, .ColumnIndex)
                                    Dim nodeColumnBounds As New Rectangle(leftWidth.First - HScroll.Value, .Bounds.Top, leftWidth.Last, .Bounds.Height)
                                    Using dottedPen As New Pen(Brushes.Silver, 1) With {
                                            .DashStyle = DashStyle.Dot
                                        }
                                        e.Graphics.DrawRectangle(dottedPen, nodeColumnBounds)
                                        For Each fieldNode In .Fields
                                            leftWidth = Node_LeftWidth(fieldNode.HeaderLevel, fieldNode.ColumnIndex)
                                            nodeColumnBounds = New Rectangle(leftWidth.First - HScroll.Value, .Bounds.Top, leftWidth.Last, .Bounds.Height)
                                            fieldNode._Bounds = nodeColumnBounds
                                            If nodeColumnBounds.Left >= Width Then Exit For
                                            TextRenderer.DrawText(e.Graphics,
                                                        fieldNode.Text,
                                                        fieldNode.Font,
                                                        nodeColumnBounds,
                                                        fieldNode.ForeColor,
                                                        fieldNode.TextBackColor,
                                                        TextFormatFlags.LeftAndRightPadding Or TextFormatFlags.NoPadding Or TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                                            e.Graphics.DrawRectangle(dottedPen, nodeColumnBounds)
                                        Next
                                    End Using
                                End If
                            End With
                        Next

                        '/// VERTICAL LINE ... DON'T DRAW IF ONLY ONE ROOT NODE
                        If drawRootLines Then
                            Dim lastNode As Node = Ancestors.Last
                            Dim LineTop As Integer = 2 + firstNode.ExpandCollapseBounds.Top + CInt(firstNode.ExpandCollapseBounds.Height / 2)
                            Dim TopPoint As New Point(verticalLineLeft, {0, LineTop}.Max)
                            Dim LineBottom As Integer = lastNode.Bounds.Top + Convert.ToInt32(lastNode.Height / 2)
                            Dim BottomPoint As New Point(verticalLineLeft, {LineBottom, Height}.Min)
                            e.Graphics.DrawLine(linePen, TopPoint, BottomPoint)
                        End If

                        If ColumnHeaders.Any Then
                            Dim headers_LevelZero = ColumnHeaders(0)
                            Dim headerX As Integer = -HScroll.Value
                            headers_LevelZero.ForEach(Sub(header)
                                                          With header
                                                              Dim headerBounds As New Rectangle(headerX, 0, .Width, headers_LevelZero.Height)
                                                              If headerBounds.Right >= 0 Then

                                                              End If
                                                              If .Style.Theme = Theme.None Then
                                                                  If .Style.BackImage Is Nothing Then
                                                                      Using LinearBrush As New LinearGradientBrush(headerBounds, .Style.BackColor, .Style.ShadeColor, LinearGradientMode.Vertical)
                                                                          e.Graphics.FillRectangle(LinearBrush, headerBounds)
                                                                      End Using
                                                                  Else
                                                                      e.Graphics.DrawImage(.Style.BackImage, headerBounds)
                                                                  End If
                                                              Else
                                                                  e.Graphics.DrawImage(GlossyImages(.Style.Theme), headerBounds)
                                                              End If
                                                              Using headerTextBrush As New SolidBrush(.Style.ForeColor)
                                                                  e.Graphics.DrawString(
                                                .Text,
                                                .Style.Font,
                                                headerTextBrush,
                                                headerBounds,
                                                .Style.Alignment
                                            )
                                                              End Using
                                                              ControlPaint.DrawBorder3D(e.Graphics, headerBounds, Border3DStyle.Raised)
                                                              headerX += header.Width
                                                          End With
                                                      End Sub)
                        End If
                    End Using

                Else
                    HScroll.Hide()
                    VScroll.Hide()
                End If
            End With
            ControlPaint.DrawBorder3D(e.Graphics, ClientRectangle, Border3DStyle.Sunken)
        End If

    End Sub
    Private Function Node_LeftWidth(headLevel As Byte, columnIndex As Byte) As Integer()

        Dim columnHeads = ColumnHeaders(headLevel)
        Dim columnHead As ColumnHead = columnHeads(columnIndex)
        Dim headIndex As Integer = 0
        Dim ColumnLeft As Integer = 0
        Do While headIndex < columnIndex
            ColumnLeft += columnHeads(headIndex).Width
            headIndex += 1
        Loop
        Return {ColumnLeft, columnHead.Width}

    End Function

#Region " GLOBAL OPTIONS "
    Private Sub IC_TextChanged(sender As Object, e As EventArgs) Handles IC_NodeEdit.TextChanged, IC_NodeAdd.TextChanged

        With DirectCast(sender, ImageCombo)
            TLP_NodePermissions.ColumnStyles(0).Width = {200, .Image.Width + TextRenderer.MeasureText(.Text, .Font).Width + .Image.Width}.Max
            TLP.SetSize(TLP_NodePermissions)
        End With

    End Sub
    Private Sub NodeEditingMouseEnter() Handles TSMI_NodeEditing.MouseEnter
        If SelectedNodes.Any Then
            IC_NodeEdit.Text = SelectedNodes.First.Text
        Else
            IC_NodeEdit.Text = String.Empty
        End If
        TSMI_NodeEditing.ShowDropDown()
    End Sub
    Private Sub ToggleSelect()
        TSMI_SelectAll.Visible = MultiSelect
        If MultiSelect Then
            TSMI_MultiSelect.Text = "Multi-Select".ToString(InvariantCulture)
        Else
            TSMI_MultiSelect.Text = "Single-Select".ToString(InvariantCulture)
        End If
    End Sub
    Private Sub NodeEditingOptions_Opening() Handles TSMI_NodeEditing.DropDownOpening
        Karen.Subscribe()
    End Sub
    Private Sub Hook_Moused() Handles Karen.Moused

        Dim CoCOptions As String = If(CursorOverControl(TSDD_Options), "[Y]", "[N]") & " TSDD_Options"
        Dim CoCNodeEdit As String = If(CursorOverControl(IC_NodeEdit), "[Y]", "[N]") & " IC_NodeEdit"
        Dim CoCNodeAdd As String = If(CursorOverControl(IC_NodeAdd), "[Y]", "[N]") & " IC_NodeAdd"
        Dim CoCNodeRemove As String = If(CursorOverControl(IC_NodeRemove), "[Y]", "[N]") & " IC_NodeRemove"

        Dim OverStatus As New List(Of String) From {CoCOptions, CoCNodeEdit, CoCNodeAdd, CoCNodeRemove}
        Dim NotOvers = OverStatus.Where(Function(o) o.Contains("[N]")).Select(Function(n) Split(n, " ").Last)
        Dim Overs = OverStatus.Where(Function(o) o.Contains("[Y]")).Select(Function(n) Split(n, " ").Last)

        Dim MessageOver As String

        If Overs.Any Then
            MessageOver = "Over:" & Join(Overs.ToArray, ",") & ", Not over:" & Join(NotOvers.ToArray, ",")

        Else
            MessageOver = "Over none, Not over any:" & Join(NotOvers.ToArray, ",")
            HideOptions()
            Karen.Unsubscribe()
        End If
        'RaiseEvent Alert(Me, New AlertEventArgs(MessageOver & " *** " & Now.ToLongTimeString))

    End Sub
    Private Sub TreeviewGlobalOptions_Opening() Handles TSDD_Options.Opening

        _OptionsOpen = True

        REM /// TEST IF THE NODE CanS EDITING, ADDS, OR REMOVAL SO AS TO HIDE THE EDIT OPTION IF NONE EXIST
        REM /// TLP_NodePermissions.Controls={1.Edit, 2.Add, 3.Remove}
        Dim SelectedNode As Node = Nothing
        Dim EditVisible As Boolean = False
        Dim AddVisible As Boolean = False
        Dim RemoveVisible As Boolean = False

        ToggleSelect()

        If SelectedNodes.Any Then
            REM /// A NODE IS SELECTED- NOW CHECK IF THE PERMISSION PROPERTIES
            SelectedNode = SelectedNodes.First
            With SelectedNode
                EditVisible = .CanEdit
                AddVisible = .CanAdd
                RemoveVisible = .CanRemove
                TSMI_Sort.Visible = .HasChildren
            End With

        Else
            REM /// NOTHING SELECTED SO CAN ONLY POTENTIALLY ADD
            EditVisible = False
            AddVisible = CanAdd
            RemoveVisible = False

        End If
        REM /// NOW IT CAN BE DETERMINED IF TSMI_NodeEditing CAN BE VISIBLE
        If EditVisible = False And AddVisible = False And RemoveVisible = False Then
            TSMI_NodeEditing.Visible = False

        Else
            REM /// AT LEAST ONE ITEM IS VISIBLE
            TSMI_NodeEditing.Visible = True

            With TLP_NodePermissions
                .RowStyles(0).Height = If(EditVisible, 28, 0)
                If EditVisible Then
                    With IC_NodeEdit
                        .Text = SelectedNode.Text
                        .DataSource = SelectedNode.Options
                    End With
                End If

                .RowStyles(1).Height = If(AddVisible, 28, 0)
                If AddVisible Then
                    With IC_NodeAdd
                        If IsNothing(SelectedNode) Then
                            IC_NodeAdd.HintText = "Add Root Node"
                        Else
                            IC_NodeAdd.HintText = "Add Child Node"
                            .DataSource = SelectedNode.ChildOptions
                        End If
                    End With
                End If

                .RowStyles(2).Height = If(RemoveVisible, 28, 0)
                .Height = Convert.ToInt32({ .RowStyles(0).Height, .RowStyles(1).Height, .RowStyles(2).Height}.Sum)

            End With

        End If
        TSMI_CheckUncheckAll.Visible = TSMI_Checkboxes.Checked

    End Sub
    Private Sub OptionsClicked(sender As Object, e As EventArgs) Handles B_Arrow.Click, B_Book.Click, B_Default.Click, B_Light.Click, TSMI_ExpandCollapseAll.Click, TSMI_Checkboxes.Click, TSMI_CheckUncheckAll.Click, TSMI_MultiSelect.Click, TSMI_SelectAll.Click, TSMI_SelectNone.Click, TSMI_Sort.Click

        If sender Is B_Arrow Then ExpanderStyle = ExpandStyle.Arrow
        If sender Is B_Book Then ExpanderStyle = ExpandStyle.Book
        If sender Is B_Default Then ExpanderStyle = ExpandStyle.PlusMinus
        If sender Is B_Light Then ExpanderStyle = ExpandStyle.LightBulb
        If sender.GetType Is GetType(Button) Then TSMI_ExpandCollapseStyles.Image = DirectCast(sender, Button).Image

#Region " SORT FUNCTIONS "
        If sender Is TSMI_Sort Then

            Dim nodes As NodeCollection = If(SelectedNodes.Any, SelectedNodes.First.Children, Ancestors)
            If nodes.SortOrder = SortOrder.None Or nodes.SortOrder = SortOrder.Descending Then
                nodes.SortOrder = SortOrder.Ascending
                TSMI_Sort.Text = "Click to Sort Descending".ToString(InvariantCulture)

            ElseIf nodes.SortOrder = SortOrder.Ascending Then
                nodes.SortOrder = SortOrder.Descending
                TSMI_Sort.Text = "Click to Sort Ascending".ToString(InvariantCulture)

            End If

        End If
#End Region
#Region " EXPAND / COLLAPSE FUNCTIONS "
        If sender Is TSMI_ExpandCollapseAll Then
            If TSMI_ExpandCollapseAll.Text = "Expand All" Then
                TSMI_ExpandCollapseAll.Text = "Collapse All".ToString(InvariantCulture)
                TSMI_ExpandCollapseAll.Image = Base64ToImage(DefaultExpanded)
                ExpandNodes()
            Else
                TSMI_ExpandCollapseAll.Text = "Expand All".ToString(InvariantCulture)
                TSMI_ExpandCollapseAll.Image = Base64ToImage(DefaultCollapsed)
                CollapseNodes()
            End If
            VScroll.Value = 0
            HScroll.Value = 0
            RequiresRepaint()
        End If
#End Region
#Region " MULTI-SELECT FUNCTIONS "
        If sender Is TSMI_MultiSelect Then MultiSelect = TSMI_MultiSelect.Checked
        ToggleSelect()

        If sender Is TSMI_SelectAll Then
            If MultiSelect Then
                For Each Node In Ancestors.All
                    Node._Selected = True
                Next
                Invalidate()
            End If
        ElseIf sender Is TSMI_SelectNone Then
            For Each Node In Ancestors.All
                Node._Selected = False
            Next
            Invalidate()
        End If
#End Region
#Region " CHECKBOX FUNCTIONS "
        TSMI_CheckUncheckAll.Visible = TSMI_Checkboxes.Checked
        If Not TSMI_Checkboxes.Checked Then TSMI_Checkboxes.HideDropDown()
        If sender Is TSMI_Checkboxes Then
            If TSMI_Checkboxes.Checked Then
                CheckBoxes = CheckState.All
            Else
                CheckBoxes = CheckState.None
            End If
        End If
        If sender Is TSMI_CheckUncheckAll Then
            TSMI_CheckUncheckAll.Text = If(TSMI_CheckUncheckAll.Checked, "UnCheck All", "Check All").ToString(InvariantCulture)
            CheckAll = TSMI_CheckUncheckAll.Checked
        End If
#End Region

    End Sub
    Private Sub EditNodeText_ValueSubmitted() Handles IC_NodeEdit.ValueSubmitted

        If SelectedNodes.Count = 1 Then
            Dim editNode As Node = SelectedNodes.First
            If editNode.CanEdit And IC_NodeEdit.Text <> editNode.Text Then
                RaiseEvent NodeBeforeEdited(Me, New NodeEventArgs(editNode, IC_NodeEdit.Text))
                If Not editNode.CancelAction Then
                    editNode.Text = IC_NodeEdit.Text
                    RaiseEvent NodeAfterEdited(Me, New NodeEventArgs(editNode, IC_NodeAdd.Text))
                End If

            End If
            TSMI_NodeEditing.HideDropDown()
        End If
        Karen.Unsubscribe()

    End Sub
    Private Sub NodeAddRequested() Handles IC_NodeAdd.ValueSubmitted, IC_NodeAdd.ItemSelected

        Karen.Unsubscribe()

        Dim nodes As NodeCollection = Nothing
        If Not SelectedNodes.Any Then
            nodes = nodes

        ElseIf SelectedNodes.Count = 1 Then
            nodes = SelectedNodes.First.Children

        End If

        If Not IsNothing(nodes) Then
            Dim Items As New List(Of Node)({New Node With {.Text = IC_NodeAdd.Text, .BackColor = Color.Lavender}})
            Items.AddRange(From I In IC_NodeAdd.Items Where Not I.Text = IC_NodeAdd.Text And I.Checked Select New Node With {.Text = I.Text, .BackColor = Color.Lavender})
            If Items.Count = 1 Then
                Dim Item As Node = Items.First
                RaiseEvent NodeBeforeAdded(Me, New NodeEventArgs(Item, IC_NodeAdd.Text))
                If Item.CancelAction Then
                    Item.Dispose()
                Else
                    REM /// BEFORE ADDING CanS FOR SETTING TO NOTHING ∴ CANCELLING
                    nodes.Add(Item)
                    RaiseEvent NodeAfterAdded(Me, New NodeEventArgs(Item, IC_NodeAdd.Text))
                End If
            Else
                REM /// ADDING A RANGE DOES NOT Can FOR TESTING BEFORE ADD
                nodes.AddRange(Items)
                RaiseEvent NodeAfterAdded(Me, New NodeEventArgs(Items))
            End If
            HideOptions()
        End If

    End Sub
    Private Sub NodeRemoveRequested() Handles IC_NodeRemove.Click

        Dim nodes As NodeCollection '= TryCast(SelectedNodes, Children)
        If SelectedNodes.Any Then
            For Each Node As Node In SelectedNodes
                If Node.CanRemove Then
                    RaiseEvent NodeBeforeRemoved(Me, New NodeEventArgs(Node))
                    If Not Node.CancelAction Then
                        If IsNothing(Node.Parent) Then
                            nodes = Ancestors
                        Else
                            nodes = Node.Parent.Children
                        End If
                        nodes.Remove(Node)
                        RaiseEvent NodeAfterRemoved(Me, New NodeEventArgs(Node))
                    End If
                End If
            Next
            HideOptions()
            TSMI_NodeEditing.HideDropDown()
        End If

    End Sub
    Private Sub HideOptions()

        _OptionsOpen = False
        Karen.Unsubscribe()
        TSDD_Options.AutoClose = True
        TSDD_Options.Hide()
        TSMI_NodeEditing.HideDropDown()

    End Sub
#End Region

    Private ExpandBeforeText_ As Boolean = True
    Public Property ExpandBeforeText As Boolean
        Get
            Return ExpandBeforeText_
        End Get
        Set(value As Boolean)
            If ExpandBeforeText_ <> value Then
                ExpandBeforeText_ = value
                Invalidate()
            End If
        End Set
    End Property
    Public Enum ExpandStyle
        PlusMinus
        Arrow
        Book
        LightBulb
    End Enum
    Private Sub UpdateExpandImage()

        Select Case _ExpanderStyle
            Case ExpandStyle.Arrow
                ExpandImage = Base64ToImage(ArrowCollapsed)
                CollapseImage = Base64ToImage(ArrowExpanded)

            Case ExpandStyle.Book
                ExpandImage = Base64ToImage(BookClosed)
                CollapseImage = Base64ToImage(BookOpen)

            Case ExpandStyle.LightBulb
                ExpandImage = Base64ToImage(LightOff)
                CollapseImage = Base64ToImage(LightOn)

            Case ExpandStyle.PlusMinus
                ExpandImage = Base64ToImage(DefaultCollapsed)
                CollapseImage = Base64ToImage(DefaultExpanded)

        End Select
        ExpandHeight = ExpandImage.Height
        REM /// IF EXPAND/COLLAPSE HEIGHT CHANGES, THEN NODE.BOUNDS WILL BE AFFECTED 
        RequiresRepaint()

    End Sub

#Region " PROPERTIES "
    Public Property ColumnHeaders As New HeaderLevels(Me)
    Private ReadOnly Property LoadTime As TimeSpan
    Private Table_ As DataTable
    Public ReadOnly Property Table As DataTable
        Get
            Return Table_
        End Get
    End Property
    Private WithEvents BindingSource As New BindingSource
    Private _DataSource As Object
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <Category("Data")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property DataSource As Object
        Get
            Return _DataSource
        End Get
        Set(value As Object)
            If value IsNot _DataSource Then
                Dim startLoad As Date = Now
                Table_ = New DataTable
                _DataSource = value
                BindingSource.DataSource = value
#Region " FILL TABLE "
                If DataSource Is Nothing Then
                    Exit Property

                ElseIf TypeOf DataSource Is String Then
                    Exit Property

                ElseIf TypeOf DataSource Is IEnumerable Then
#Region " UNSTRUCTURED "
                    Select Case DataSource.GetType
                        Case GetType(List(Of String()))
#Region " LIST OF STRING() - LIST ITEMS=ROWS...STRING()=HeaderS "
                            Dim Rows As List(Of String()) = DirectCast(_DataSource, List(Of String()))
                            If Rows IsNot Nothing Then
                                Dim HeaderCount As Integer = (From C In Rows Select C.Count).Max
                                For Header As Integer = 1 To HeaderCount
                                    Table_.Columns.Add(New DataColumn With {.ColumnName = "Header" & Header, .DataType = GetType(String)})
                                Next
                                For Each Row As String() In Rows
                                    Table_.Rows.Add(Row)
                                Next
                            End If
#End Region
                        Case GetType(List(Of Object()))
#Region " LIST OF OBJECT() - LIST ITEMS=ROWS...STRING()=HeaderS "
                            Dim Rows As List(Of Object()) = DirectCast(_DataSource, List(Of Object()))
                            If Rows IsNot Nothing Then
                                Dim HeaderCount As Integer = (From C In Rows Select C.Count).Max
                                For Header As Integer = 1 To HeaderCount
                                    Table_.Columns.Add(New DataColumn With {.ColumnName = "Header" & Header, .DataType = GetType(String)})
                                Next
                                For Each Row As String() In Rows
                                    Table_.Rows.Add(Row)
                                Next
                            End If
#End Region
                        Case Else

                    End Select
#End Region
                ElseIf DataSource.GetType Is GetType(DataTable) Then
                    Table_ = DirectCast(DataSource, DataTable)
                    If ColumnHeaders.Any Then
                        RecursiveBuild(Ancestors, 0, Table)
                    End If
                    _LoadTime = Now.Subtract(startLoad)
                End If
#End Region
                Invalidate()
            End If
        End Set
    End Property
    Private Sub RecursiveBuild(nc As NodeCollection, lvl As Byte, tbl As DataTable)

        Dim headers As ColumnHeadCollection = ColumnHeaders(lvl)
        If lvl = 0 Then Offset = New Point(Offset.X + 0, Offset.Y + headers.Height)
        Dim headName As String = headers(0).Text 'Must be the first
        If tbl.Columns.Contains(headName) Then
            Dim tables = (From r In tbl.AsEnumerable Group r By groupX = r(headName).ToString Into headGroup = Group
                          Select New With {
                                                     .groupBy = groupX,
                                                     .table = headGroup.ToList
                                                     }).ToDictionary(Function(k) k.groupBy, Function(v) v.table)
            '/// Each table will be a subset of Rows
            For Each tableGroup In tables
                Dim headNode As Node = nc.Add(New Node(tableGroup.Key))
                With headNode
                    .HeaderLevel_ = lvl
                    .ColumnIndex_ = 0
                    .IsField_ = False
                    .Font = Font
                End With
                tableGroup.Value.ForEach(Sub(row)
                                             Dim subHeaders = headers.Skip(1).ToList
                                             If subHeaders.Any Then
                                                 Dim childNode As Node = headNode.Children.Add(row(subHeaders.First.Text).ToString)
                                                 With childNode
                                                     .HeaderLevel_ = lvl
                                                     .ColumnIndex_ = 1
                                                     .IsField_ = False
                                                     .Font = Font
                                                 End With
                                                 subHeaders.RemoveAt(0)
                                                 Dim columnIndex As Byte = 2
                                                 subHeaders.ForEach(Sub(cn)
                                                                        Dim fieldNode As Node = childNode.Fields.Add(row(cn.Text).ToString)
                                                                        With fieldNode
                                                                            .HeaderLevel_ = lvl
                                                                            .ColumnIndex_ = columnIndex
                                                                            .IsField_ = True
                                                                            .Font = Font
                                                                        End With
                                                                        columnIndex += CByte(1)
                                                                    End Sub)
                                             End If
                                             'If lvl + 1 < ColumnHeaders.Count Then RecursiveBuild(headNode.Children, lvl + 1, tableGroup.Value.CopyToDataTable)
                                         End Sub)
            Next
        Else
            Throw New Exception
        End If

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Property FavoriteImage As Image = Base64ToImage(StarString)
    Public ReadOnly Property OptionsOpen As Boolean
    Public Property FavoritesFirst As Boolean = True
    Private _ExpanderStyle As ExpandStyle = ExpandStyle.PlusMinus
    Public Property ExpanderStyle As ExpandStyle
        Get
            Return _ExpanderStyle
        End Get
        Set(ByVal value As ExpandStyle)
            _ExpanderStyle = value
            UpdateExpandImage()
        End Set
    End Property
    Public ReadOnly Property DropHighlightNode As Node
        Get
            Return DragData.DropHighlightNode
        End Get
    End Property
    Public Property MouseOverExpandsNode As Boolean = False
    Public Property CanAdd As Boolean = True
    Private _MultiSelect As Boolean
    Public Property MultiSelect As Boolean
        Get
            Return _MultiSelect
        End Get
        Set(value As Boolean)
            If _MultiSelect <> value Then
                If Not value Then
                    For Each Node In Ancestors.All
                        Node._Selected = False
                    Next
                End If
                _MultiSelect = value
                Invalidate()
            End If
        End Set
    End Property
    Public ReadOnly Property UnRestrictedSize As Size
        Get
            Return New Size(VScrollWidth + RollingWidth + Offset.X, RollingHeight + Offset.Y)
        End Get
    End Property
    Public ReadOnly Property NodeHeight As Integer
        Get
            Return CInt(Ancestors.All.Average(Function(a) a.Height))
        End Get
    End Property
    Public ReadOnly Property SelectedNodes As List(Of Node)
        Get
            Return Ancestors.All.Where(Function(n) n.Selected).ToList
        End Get
    End Property
    Public ReadOnly Property Ancestors As New NodeCollection(Me)
    Public Property MaxNodes As Integer
    Public Property LineStyle As DashStyle = DashStyle.Dot
    Public Property LineColor As Color = Color.Blue
    Public Property RootLines As Boolean = True
    Public Property DropHighlightColor As Color = Color.Gainsboro
    Public Property MouseOverColor As Color = Color.Gainsboro
    Public Property SelectionColor As Color = Color.Gainsboro
    Public Property Offset As New Point(5, 3)
    Public Overrides Property AutoSize As Boolean = True
    Public Property StopMe As Boolean
    Private _CheckBoxes As CheckState = CheckState.Mixed
    Public Property CheckBoxes As CheckState
        Get
            Return _CheckBoxes
        End Get
        Set(value As CheckState)
            If _CheckBoxes <> value Then
                If value = CheckState.All Then
                    For Each Node In Ancestors.All
                        Node.CheckBox = True
                    Next
                ElseIf value = CheckState.None Then
                    For Each Node In Ancestors.All
                        Node.CheckBox = False
                    Next
                End If
                _CheckBoxes = value
            End If
        End Set
    End Property
    Private _CheckAll As Boolean = False
    Public Property CheckAll As Boolean
        Get
            Return _CheckAll
        End Get
        Set(value As Boolean)
            If _CheckAll <> value Then
                Dim CheckNodes = Ancestors.All.Where(Function(c) c.CheckBox)
                If _CheckAll And Not CheckNodes.Any Then
                    CheckBoxes = CheckState.All
                End If
                For Each Node In Ancestors.All.Where(Function(c) c.CheckBox)
                    Node.Checked = value
                Next
                _CheckAll = value
            End If
        End Set
    End Property
#End Region
    Public Event NodesChanged(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeBeforeAdded(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeAfterAdded(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeBeforeRemoved(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeAfterRemoved(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeBeforeEdited(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeAfterEdited(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeDragStart(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeDragOver(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeDropped(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeChecked(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeExpanded(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeCollapsed(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeClicked(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeRightClicked(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeFavorited(sender As Object, e As NodeEventArgs)
    Public Event NodeDoubleClicked(ByVal sender As Object, ByVal e As NodeEventArgs)

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ M O U S E   E V E N T S
    Private LastMouseNode As Node = Nothing
    Private CurrentMouseNode As Node = Nothing
    Private MousePoint As Point
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

        If e IsNot Nothing Then
            Dim HitRegion As HitRegion = HitTest(e.Location)
            Dim HitNode As Node = HitRegion.Node
            DragData = New DragInfo With {.DragNode = HitNode,
                .IsDragging = False,
                .MousePoints = New List(Of Point)}

            If e.Button = MouseButtons.Right Then
                Dim ShowLocation As Point = Cursor.Position
                ShowLocation.Offset(10, 0)
                TSDD_Options.AutoClose = False
                TSDD_Options.Show(ShowLocation)
            Else
                HideOptions()
                If HitNode IsNot Nothing Then
                    With HitNode
                        Select Case HitRegion.Region
                            Case NodeRegion.Expander
                                ._Clicked = True
                                If .HasChildren Then
                                    ._Expanded = Not .Expanded
                                    If .Expanded Then
                                        .Expand()
                                        RaiseEvent NodeExpanded(Me, New NodeEventArgs(HitNode))
                                    Else
                                        .Collapse()
                                        RaiseEvent NodeCollapsed(Me, New NodeEventArgs(HitNode))
                                    End If
                                End If

                            Case NodeRegion.Favorite
                                .Favorite = Not .Favorite
                                RaiseEvent NodeFavorited(Me, New NodeEventArgs(HitNode))

                            Case NodeRegion.CheckBox
                                ._Clicked = True
                                .Checked = Not .Checked
                                RaiseEvent NodeChecked(Me, New NodeEventArgs(HitNode))

                            Case NodeRegion.Image, NodeRegion.Node
                                ._Clicked = True
                                If Not MultiSelect Then
                                    For Each Node In SelectedNodes.Except({HitNode})
                                        Node._Selected = False
                                    Next
                                End If
                                HitNode._Selected = Not HitNode.Selected
                                If e.Button = MouseButtons.Left Then
                                    RaiseEvent NodeClicked(Me, New NodeEventArgs(HitNode))

                                ElseIf e.Button = MouseButtons.Right Then
                                    RaiseEvent NodeRightClicked(Me, New NodeEventArgs(HitNode))

                                End If

                        End Select
                    End With
                    Invalidate()
                End If
            End If
        End If
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            MousePoint = e.Location
            If e.Button = MouseButtons.None Then
                CurrentMouseNode = HitTest(e.Location).Node
                If CurrentMouseNode IsNot LastMouseNode Then
                    If MouseOverExpandsNode Then
                        If CurrentMouseNode Is Nothing Then
                            LastMouseNode.Collapse()
                        Else
                            CurrentMouseNode.Expand()
                        End If
                    End If
                    LastMouseNode = CurrentMouseNode
                    Invalidate()
                End If

            ElseIf e.Button = MouseButtons.Left Then
                With DragData
                    If Not .MousePoints.Contains(e.Location) Then .MousePoints.Add(e.Location)
                    .IsDragging = .DragNode IsNot Nothing AndAlso (.MousePoints.Count >= 5 And Not .DragNode.Bounds.Contains(.MousePoints.Last))
                    If .IsDragging Then
                        OnDragStart()
                        Dim Data As New DataObject
                        Data.SetData(GetType(Node), .DragNode)
                        MyBase.OnDragOver(New DragEventArgs(Data, 0, e.X, e.Y, DragDropEffects.Copy Or DragDropEffects.Move, DragDropEffects.All))
                        DoDragDrop(Data, DragDropEffects.Copy Or DragDropEffects.Move)
                    End If
                End With

            End If
        End If
        MyBase.OnMouseMove(e)

    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)

        Cursor = Cursors.Default
        DragData.IsDragging = False
        DragData.MousePoints.Clear()
        ScrollTimer.Stop()
        MyBase.OnMouseUp(e)

    End Sub
    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)

        If e IsNot Nothing Then
            Dim HitRegion As HitRegion = HitTest(e.Location)
            Dim HitNode As Node = HitRegion.Node
            If Not IsNothing(HitNode) Then
                RaiseEvent NodeDoubleClicked(Me, New NodeEventArgs(HitNode))
            End If
        End If
        MyBase.OnMouseDoubleClick(e)

    End Sub

#Region " KEYPRESS EVENTS "
    Private Sub On_PreviewKeyDown(ByVal sender As Object, ByVal e As PreviewKeyDownEventArgs)

        If e IsNot Nothing Then
            Select Case e.KeyCode
                Case Keys.Up, Keys.Down, Keys.Left, Keys.Right, Keys.Tab, Keys.ControlKey
                    e.IsInputKey = True
            End Select
        End If

    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
        If e IsNot Nothing AndAlso e.Handled Then
        End If
        MyBase.OnKeyPress(e)
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

        If e IsNot Nothing AndAlso e.Modifiers = Keys.Control AndAlso e.KeyCode = Keys.C Then
            Dim Nodes2Clipboard As New List(Of String)(From sn In SelectedNodes Select sn.Text)
            Clipboard.SetText(Join(Nodes2Clipboard.ToArray, vbNewLine))
        End If
        MyBase.OnKeyDown(e)

    End Sub
#End Region
#Region " DRAG & DROP "
    Private Sub OnDragStart()

        With DragData
            If .DragNode.CanDragDrop Then
#Region " CUSTOM CURSOR WITH NODE.TEXT "
                Dim nodeFont As New Font(.DragNode.Font.FontFamily, .DragNode.Font.Size + 4, FontStyle.Bold)
                Dim textSize As Size = TextRenderer.MeasureText(.DragNode.Text, nodeFont)
                Dim cursorBounds As New Rectangle(New Point(0, 0), New Size(3 + textSize.Width + 3, 2 + textSize.Height + 2))
                Dim shadowDepth As Integer = 16
                cursorBounds.Inflate(shadowDepth, shadowDepth)
                Dim bmp As New Bitmap(cursorBounds.Width, cursorBounds.Height)
                Using Graphics As Graphics = Graphics.FromImage(bmp)
                    Graphics.SmoothingMode = SmoothingMode.AntiAlias
                    Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
                    Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
                    Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
                    Dim fadingRectangle As New Rectangle(0, 0, bmp.Width, bmp.Height)
                    Dim fadeFactor As Integer
                    For P = 0 To shadowDepth - 1
                        fadeFactor = 16 + (P * 5)
                        fadingRectangle.Inflate(-1, -1)
                        Using fadingBrush As New SolidBrush(Color.FromArgb(fadeFactor, .DragNode.CursorGlowColor)) '16, 21, 26, 31, 36
                            Using fadingPen As New Pen(fadingBrush, 1)
                                Graphics.DrawRectangle(fadingPen, fadingRectangle)
                            End Using
                        End Using
                    Next
                    Using fadedBrush As New SolidBrush(Color.FromArgb(fadeFactor, Color.Gainsboro))
                        Graphics.FillRectangle(fadedBrush, fadingRectangle)
                    End Using
                    Graphics.DrawRectangle(Pens.Silver, fadingRectangle)
                    Dim Format As New StringFormat With {
    .Alignment = StringAlignment.Center,
    .LineAlignment = StringAlignment.Center
}
                    Graphics.DrawString(.DragNode.Text, nodeFont, Brushes.Black, fadingRectangle, Format)
                End Using
                _Cursor = CursorHelper.CreateCursor(bmp, 0, Convert.ToInt32(bmp.Height / 2))
#End Region
                REM /// NOW SET DATA
                Dim Data As New DataObject
                Data.SetData(GetType(Node), .DragNode)
                AllowDrop = True
                RaiseEvent NodeDragStart(Me, New NodeEventArgs(.DragNode))
            Else
                CursorTimer.Start()
            End If

        End With

    End Sub
    Private Sub CursorTimerTick() Handles CursorTimer.Tick

        CursorTimer.Tag = If(CursorTimer.Tag, 0)
        Dim Index As Integer = DirectCast(If(CursorTimer.Tag, 0), Integer)
        Dim im As Integer = Index Mod 3
        'Using CursorImage As Bitmap = If(im = 0, My.Resources.FingerLeft, If(im = 1, My.Resources.Finger, My.Resources.FingerRight))
        '    Index += 1
        '    If Index >= 30 Then
        '        Index = 0
        '        CursorTimer.Stop()
        '        _Cursor = Cursors.Default
        '    Else
        '        _Cursor = CursorHelper.CreateCursor(CursorImage, 0, Convert.ToInt32(CursorImage.Height / 2))
        '    End If
        '    CursorTimer.Tag = Index
        'End Using

    End Sub
    Protected Overrides Sub OnDragLeave(ByVal e As EventArgs)

        DragData.DropHighlightNode = Nothing
        Invalidate()
        ScrollTimer.Tag = Nothing

        Dim MouseLocation As Point = PointToClient(New Point(0, 0))
        MouseLocation.Offset(Cursor.Position)
        If VScroll.Visible Then
            If MouseLocation.Y <= 0 Then
                REM /// EXITED TOP
                ScrollTimer.Tag = "Up"
                ScrollTimer.Start()

            ElseIf MouseLocation.Y >= Height Then
                REM /// EXITED BOTTOM
                ScrollTimer.Tag = "Down"
                ScrollTimer.Start()
            End If
        End If
        MyBase.OnDragLeave(e)

    End Sub
    Protected Overrides Sub OnDragEnter(ByVal e As DragEventArgs)

        ScrollTimer.Stop()
        Invalidate()
        MyBase.OnDragEnter(e)

    End Sub
    Protected Overrides Sub OnDragOver(ByVal e As DragEventArgs)

        If e IsNot Nothing Then
            e.Effect = DragDropEffects.All
            Dim Location As Point = PointToClient(New Point(e.X, e.Y))
            Dim HitRegion As HitRegion = HitTest(Location)
            Dim HitNode As Node = HitRegion?.Node
            If Not HitNode Is DragData.DropHighlightNode Then
                DragData.DropHighlightNode = HitNode
                Invalidate()
                If Not IsNothing(HitNode) Then
                    e.Data.SetData(GetType(Object), HitNode)
                    If HitRegion.Region = NodeRegion.Expander And HitNode.HasChildren Then
                        If HitNode.Expanded Then
                            HitNode.Collapse()
                        Else
                            HitNode.Expand()
                        End If
                    End If
                End If
            End If
        End If
        MyBase.OnDragOver(e)

    End Sub
    Protected Overrides Sub OnDragDrop(ByVal e As DragEventArgs)

        If e IsNot Nothing Then
            CursorTimer.Stop()
            Dim DragNode As Node = TryCast(e.Data.GetData(GetType(Node)), Node)
            Dim Location As Point = PointToClient(New Point(e.X, e.Y))
            Dim HitRegion As HitRegion = HitTest(Location)
            Dim HitNode As Node = HitRegion.Node
            If DragNode IsNot Nothing AndAlso DragNode.CanDragDrop And HitNode IsNot Nothing AndAlso HitNode.CanDragDrop Then
                e.Data.SetData(GetType(Node), HitNode)
                RaiseEvent NodeDropped(Me, New NodeEventArgs(HitNode))
            End If
            DragData.DropHighlightNode = Nothing
            RequiresRepaint()
            MyBase.OnDragDrop(e)
        End If

    End Sub
    Protected Overrides Sub OnGiveFeedback(ByVal e As GiveFeedbackEventArgs)

        If e IsNot Nothing Then
            e.UseDefaultCursors = False
            Cursor.Current = _Cursor
        End If
        MyBase.OnGiveFeedback(e)

    End Sub
    Private Sub DragScroll() Handles ScrollTimer.Tick

        If MouseButtons.HasFlag(MouseButtons.Left) And Not IsNothing(ScrollTimer.Tag) Then
            UpdateDrawNodes()
            Dim VScrollValue As Integer = VScroll.Value
            Dim Delta As Integer = 0
            If ScrollTimer.Tag.ToString = "Up" Then
                REM /// EXITED TOP
                VScrollValue -= VScroll.SmallChange
                VScrollValue = {VScroll.Minimum, VScrollValue}.Max
                Delta = VScroll.Value - VScrollValue
                VScroll.Value = VScrollValue
                VScrollUpDown(Delta)

            ElseIf ScrollTimer.Tag.ToString = "Down" Then
                REM /// EXITED TOP
                VScrollValue += VScroll.SmallChange
                VScrollValue = {VScroll.Maximum - Height, VScrollValue}.Min
                Delta = VScroll.Value - VScrollValue
                VScroll.Value = VScrollValue
                VScrollUpDown(Delta)

            End If
            If Delta = 0 Then ScrollTimer.Stop()
        Else
            ScrollTimer.Stop()

        End If

    End Sub
#End Region
#Region " INVALIDATION "
    Dim RR As Boolean = False
    Friend Sub RequiresRepaint()

        RR = True
        REM /// RESET INDEX / HEIGHT
        VisibleIndex = 0
        RollingWidth = Offset.X
        RollingHeight = Offset.Y

        REM /// ITERATE ALL NODES CHANGING BOUNDS
        RefreshNodesBounds_Lines(Ancestors)

        REM /// TOTAL SIZE + RESIZE THE CONTROL IF AUTOSIZE
#Region " DETERMINE THE MAXIMUM POSSIBLE SIZE OF THE CONTROL AND COMPARE TO THE UNRESTRICTED SIZE "
        Dim screenLocation As Point = PointToScreen(New Point(0, 0))
        Dim wa As Size = WorkingArea.Size
        Dim maxScreenSize As New Size(wa.Width - screenLocation.X, wa.Height - screenLocation.Y)
        Dim maxParentSize As New Size
        Dim maxUserSize As Size = MaximumSize
        Dim unboundedSize As Size = UnRestrictedSize 'This is strictly the space size required to fit all node text ( does NOT include ScrollBars )
#Region " DETERMINE IF A PARENT RESTRICTS THE SIZE OF THE TREEVIEWER - LOOK FOR <.AutoSize> IN PARENT CONTROL PROPERTIES "
        If Parent IsNot Nothing Then
            Dim controlType As Type = Parent.GetType
            Dim properties As Reflection.PropertyInfo() = controlType.GetProperties
            Dim growParent As Boolean = False
            For Each controlProperty In properties
                If controlProperty.Name = "AutoSize" Then
                    Dim propertyValue As Boolean = DirectCast(controlProperty.GetValue(Parent), Boolean)
                    If propertyValue Then growParent = True
                    Exit For
                End If
            Next
            If Not growParent Then maxParentSize = New Size(Parent.Width, Parent.Height)
            If Not Parent.MaximumSize.IsEmpty Then maxParentSize = Parent.MaximumSize
        End If
#End Region
        Dim maxSize As New Size
        Dim sizes As New List(Of Size) From {maxScreenSize, maxParentSize, maxUserSize}
        Dim nonZeroWidths As New List(Of Integer)(From s In sizes Where s.Width > 0 Select s.Width)
        Dim nonZeroHeights As New List(Of Integer)(From s In sizes Where s.Height > 0 Select s.Height)
        Dim maxWidth As Integer = nonZeroWidths.Min
        Dim maxHeight As Integer = nonZeroHeights.Min
        If AutoSize Then 'Can resize
            Dim proposedWidth = {UnRestrictedSize.Width, maxWidth}.Min
            Dim proposedHeight = {UnRestrictedSize.Height, maxHeight}.Min

            Dim hScrollVisible As Boolean = UnRestrictedSize.Width > maxWidth
            If hScrollVisible Then proposedHeight = {proposedHeight + HScrollHeight, maxHeight}.Min

            Dim vscrollVisible As Boolean = UnRestrictedSize.Height > maxHeight
            If vscrollVisible Then proposedWidth = {proposedWidth + VScrollWidth, maxWidth}.Min
            Width = proposedWidth
            Height = proposedHeight
            maxWidth = Width
            maxHeight = Height
        Else
            maxWidth = {Width, maxWidth}.Min
            maxHeight = {Height, maxHeight}.Min
        End If
        With HScroll
            .Minimum = 0
            .Maximum = {0, UnRestrictedSize.Width - 1}.Max
            .Visible = UnRestrictedSize.Width > maxWidth
            .Left = 0
            .Width = maxWidth
            .Top = Height - .Height
            .LargeChange = maxWidth
            If .Visible Then
                If .Value > maxWidth - Width Then .Value = {maxWidth - Width, .Minimum}.Max
                .Show()
            Else
                .Value = 0
                .Hide()
            End If
        End With
        With VScroll
            .Minimum = 0
            .Maximum = {0, UnRestrictedSize.Height - 1}.Max
            .Visible = UnRestrictedSize.Height > maxHeight
            .Left = maxWidth - VScrollWidth
            .Height = maxHeight
            .Top = 0
            .LargeChange = maxHeight
            If .Visible Then
                If .Value > maxHeight - Height Then .Value = {maxHeight - Height, .Minimum}.Max
                .Show()
            Else
                .Value = 0
                .Hide()
            End If
        End With
        If Ancestors.Any And StopMe Then Stop
#End Region
        UpdateDrawNodes()

        REM /// FINALLY- REPAINT
        Invalidate()
        RR = False

    End Sub
    Private Sub RefreshNodesBounds_Lines(Nodes As NodeCollection)

        Dim NodeIndex As Integer = 0
        Dim maxWidths As New Dictionary(Of Integer, Integer)
        For Each Node As Node In Nodes
            With Node
                ._Index = NodeIndex
                ._Visible = .Parent Is Nothing OrElse .Parent.Expanded
                If .Visible Then
                    RefreshNodeBounds_Lines(Node)
                    ._VisibleIndex = VisibleIndex
                    VisibleIndex += 1
                    If .Bounds.Right > RollingWidth Then RollingWidth = .Bounds.Right
                    RollingHeight += .Height
                    If .HasChildren Then RefreshNodesBounds_Lines(.Children)
                End If
                If Not .HeaderLevel = 255 Then
                    Dim columnHead As ColumnHead = ColumnHeaders(.HeaderLevel)(.ColumnIndex)
                    If .Bounds.Width > columnHead.Width Then columnHead.Width = .Bounds.Width
                    Dim headWidths As Integer = .Bounds.Right
                    For Each fieldNode In .Fields
                        columnHead = ColumnHeaders(.HeaderLevel)(fieldNode.ColumnIndex)
                        Dim textWidth As Integer = TextRenderer.MeasureText(fieldNode.Text, fieldNode.Font).Width
                        If textWidth > columnHead.Width Then columnHead.Width = textWidth
                        headWidths += columnHead.Width
                    Next
                    If headWidths > RollingWidth Then RollingWidth = headWidths
                End If
                NodeIndex += 1
            End With
            If FavoritesFirst And Node.CanFavorite Then Node.Children.SortAscending(False) 'Do not let the Sort require repaint as it cycles back here to an infinate loop
        Next

    End Sub
    Private Sub RefreshNodeBounds_Lines(Node As Node)

        Dim Y As Integer = RollingHeight - VScroll.Value
        Dim HorizontalSpacing As Integer = 3
        With Node
            If ExpandBeforeText Then
#Region "+- Icon precedes Text"
                REM FAVORITE
                ._FavoriteBounds.X = Offset.X + HorizontalSpacing + If(IsNothing(.Parent), If(RootLines, 6, 0), .Parent.ExpandCollapseBounds.Right + HorizontalSpacing)
                ._FavoriteBounds.Y = Y + CInt((.Height - FavoriteImage.Height) / 2)
                ._FavoriteBounds.Width = If(.CanFavorite, FavoriteImage.Width, 0)
                ._FavoriteBounds.Height = If(.CanFavorite, FavoriteImage.Height, .Height)

                REM CHECKBOX
                ._CheckBounds.X = ._FavoriteBounds.Right + If(._FavoriteBounds.Width = 0, 0, HorizontalSpacing)
                ._CheckBounds.Width = If(.CheckBox, CheckHeight, 0)
                ._CheckBounds.Height = CheckHeight
                ._CheckBounds.Y = Y + CInt((.Height - ._CheckBounds.Height) / 2)

                REM IMAGE
                ._ImageBounds.X = ._CheckBounds.Right + If(._CheckBounds.Width = 0, 0, HorizontalSpacing)
                ._ImageBounds.Height = If(IsNothing(.Image), 0, If(.ImageScaling, .Height, .Image.Height))
                'MAKE IMAGE SQUARE IF SCALING
                ._ImageBounds.Width = If(IsNothing(.Image), 0, If(.ImageScaling, ._ImageBounds.Height, .Image.Width))
                ._ImageBounds.Y = Y + CInt((.Height - ._ImageBounds.Height) / 2)

                REM EXPAND/COLLAPSE
                ._ExpandCollapseBounds.X = ._ImageBounds.Right + HorizontalSpacing '+ If(IsNothing(.Parent), If(RootLines, 6, 0), .Parent._ImageBounds.Right + HorizontalSpacing)
                ._ExpandCollapseBounds.Y = Y + CInt((.Height - ExpandHeight) / 2)
                ._ExpandCollapseBounds.Width = If(.HasChildren, ExpandHeight, 0)
                ._ExpandCollapseBounds.Height = ExpandHeight

                REM TEXT
                ._Bounds.X = ._ExpandCollapseBounds.Right + If(._ExpandCollapseBounds.Width = 0, 0, HorizontalSpacing)
                ._Bounds.Y = Y
                ._Bounds.Width = TextRenderer.MeasureText(.Text, .Font).Width
                ._Bounds.Height = .Height
#End Region
            Else
#Region "+- Icon follows Text"
                REM EXPAND/COLLAPSE
                ._ExpandCollapseBounds.X = Offset.X + HorizontalSpacing + If(IsNothing(.Parent), If(RootLines, 6, 0), .Parent.ExpandCollapseBounds.Right + HorizontalSpacing)
                ._ExpandCollapseBounds.Y = Y + CInt((.Height - ExpandHeight) / 2)
                ._ExpandCollapseBounds.Width = If(.HasChildren, ExpandHeight, 0)
                ._ExpandCollapseBounds.Height = ExpandHeight

                REM FAVORITE
                ._FavoriteBounds.X = ._ExpandCollapseBounds.Right + If(._ExpandCollapseBounds.Width = 0, 0, HorizontalSpacing)
                ._FavoriteBounds.Y = Y + CInt((.Height - FavoriteImage.Height) / 2)
                ._FavoriteBounds.Width = If(.CanFavorite, FavoriteImage.Width, 0)
                ._FavoriteBounds.Height = If(.CanFavorite, FavoriteImage.Height, .Height)

                REM CHECKBOX
                ._CheckBounds.X = ._FavoriteBounds.Right + If(._FavoriteBounds.Width = 0, 0, HorizontalSpacing)
                ._CheckBounds.Width = If(.CheckBox, CheckHeight, 0)
                ._CheckBounds.Height = CheckHeight
                ._CheckBounds.Y = Y + CInt((.Height - ._CheckBounds.Height) / 2)

                REM IMAGE
                ._ImageBounds.X = ._CheckBounds.Right + If(._CheckBounds.Width = 0, 0, HorizontalSpacing)
                ._ImageBounds.Height = If(IsNothing(.Image), 0, If(.ImageScaling, .Height, .Image.Height))
                'MAKE IMAGE SQUARE IF SCALING
                ._ImageBounds.Width = If(IsNothing(.Image), 0, If(.ImageScaling, ._ImageBounds.Height, .Image.Width))
                ._ImageBounds.Y = Y + CInt((.Height - ._ImageBounds.Height) / 2)

                REM TEXT
                ._Bounds.X = ._ImageBounds.Right + If(._ImageBounds.Width = 0, 0, HorizontalSpacing)
                ._Bounds.Y = Y
                ._Bounds.Width = TextRenderer.MeasureText(.Text, .Font).Width
                ._Bounds.Height = .Height
#End Region
            End If
            For Each fieldNode In .Fields
                fieldNode._Bounds.Y = .Bounds.Y
                fieldNode._Bounds.Height = .Bounds.Height
            Next
        End With

    End Sub

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ SCROLLING, SCROLLING, SCROLLING
    Private Sub OnScrolled(ByVal sender As Object, e As ScrollEventArgs) Handles HScroll.Scroll, VScroll.Scroll

        If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
            HScrollLeftRight(e.OldValue - e.NewValue)

        ElseIf e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            VScrollUpDown(e.OldValue - e.NewValue)

        End If
        UpdateDrawNodes()

    End Sub
    Private Sub HScrollLeftRight(X_Change As Integer)

        If X_Change = 0 Then Exit Sub
        Ancestors.Visible.ForEach(Sub(node)
                                      node._ExpandCollapseBounds.X += X_Change
                                      node._FavoriteBounds.X += X_Change
                                      node._CheckBounds.X += X_Change
                                      node._ImageBounds.X += X_Change
                                      node._Bounds.X += X_Change
                                  End Sub)
        Invalidate()

    End Sub
    Private Sub VScrollUpDown(Y_Change As Integer)

        If Y_Change = 0 Then Exit Sub
        Ancestors.Visible.ForEach(Sub(node)
                                      node._ExpandCollapseBounds.Y += Y_Change
                                      node._FavoriteBounds.Y += Y_Change
                                      node._CheckBounds.Y += Y_Change
                                      node._ImageBounds.Y += Y_Change
                                      node._Bounds.Y += Y_Change
                                  End Sub)
        Invalidate()

    End Sub

    Protected Overrides Sub OnFontChanged(e As EventArgs)

        Ancestors.All.ForEach(Sub(node)
                                  node.Font = Font
                                  node.Fields.ForEach(Sub(f)
                                                          f.Font = Font
                                                      End Sub)
                              End Sub)
        RequiresRepaint()
        MyBase.OnFontChanged(e)

    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)

        If Not RR Then RequiresRepaint()
        Invalidate()
        MyBase.OnSizeChanged(e)

    End Sub
    Friend Sub NodeTimer_Start(TickNode As Node)

        With NodeTimer
            .Tag = TickNode
            .Start()
        End With

    End Sub
    Private Sub NodeTimer_Tick() Handles NodeTimer.Tick

        With NodeTimer
            .Stop()
            Dim TickNode As Node = DirectCast(.Tag, Node)
            RaiseEvent NodesChanged(Me, New NodeEventArgs(TickNode))
            'If Name = "Scripts" Then Stop
            RequiresRepaint()
        End With

    End Sub
#End Region
#Region " METHODS / FUNCTIONS "
    Public Function HitTest(Location As Point) As HitRegion

        Dim Region As New HitRegion
        Dim expandBounds As New List(Of Node)(From N In Ancestors.Draw Where N.ExpandCollapseBounds.Contains(Location))
        Dim favoriteBounds As New List(Of Node)(From N In Ancestors.Draw Where N.FavoriteBounds.Contains(Location))
        Dim checkBounds As New List(Of Node)(From N In Ancestors.Draw Where N.CheckBounds.Contains(Location))
        Dim imageBounds As New List(Of Node)(From N In Ancestors.Draw Where N.ImageBounds.Contains(Location))
        Dim nodeBounds As New List(Of Node)(From N In Ancestors.Draw Where N.Bounds.Contains(Location))

        Dim HitBounds As New List(Of Node)(expandBounds.Union(favoriteBounds).Union(checkBounds).Union(nodeBounds))
        If HitBounds.Any Then
            With Region
                .Node = HitBounds.First
                If expandBounds.Any Then .Region = NodeRegion.Expander
                If favoriteBounds.Any Then .Region = NodeRegion.Favorite
                If checkBounds.Any Then .Region = NodeRegion.CheckBox
                If imageBounds.Any Then .Region = NodeRegion.Image
                If nodeBounds.Any Then .Region = NodeRegion.Node
            End With
        End If
        Return Region

    End Function
    Public Sub ExpandNodes()

        If Ancestors.Any Then
            ExpandCollapseNodes(Ancestors, True)
            Ancestors.First.Expand()
        End If

    End Sub
    Public Sub CollapseNodes()

        If Ancestors.Any Then
            ExpandCollapseNodes(Ancestors, False)
            Ancestors.First.Collapse()
        End If

    End Sub
    Private Sub ExpandCollapseNodes(ByVal Nodes As NodeCollection, State As Boolean)

        For Each Node As Node In Nodes
            If State Then
                Node.Expand()
            Else
                Node.Collapse()
            End If
            If Node.HasChildren Then ExpandCollapseNodes(Node.Children, State)
        Next

    End Sub
    Private Sub UpdateDrawNodes()

        Ancestors.Client.Clear()
        For Each Node In Ancestors.Visible
            If ClientRectangle.Contains(New Point(1, Node.Bounds.Top)) Or ClientRectangle.Contains(New Point(1, Node.Bounds.Bottom)) Then
                Ancestors.Client.Add(Node)
            End If
        Next

        Ancestors.Draw.Clear()
        Ancestors.Draw.AddRange(Ancestors.Client)
        Dim NodesDraw = (From N In Ancestors.Visible Group N By N.Parent Into Group Select Parent, Children =
                ((From C In Group Where C._Bounds.Bottom < 0 Order By C._Bounds.Y Descending).Take(1)).Union _
                ((From C In Group Where C._Bounds.Top > ClientRectangle.Height Order By C._Bounds.Y Ascending).Take(1)))
        For Each ParentGroup In NodesDraw
            Ancestors.Draw.AddRange(ParentGroup.Children)
        Next

    End Sub
    Public Sub AutoWidth()

        If Ancestors.Any Then
            Width = Ancestors.Max(Function(n) n.Bounds.Right)
            RequiresRepaint()
        End If

    End Sub
#End Region
End Class

REM //////////////// NODE COLLECTION \\\\\\\\\\\\\\\\
Public NotInheritable Class NodeCollection
    Inherits List(Of Node)
    Private _SortOrder As New SortOrder
    Public Property SortOrder As SortOrder
        Get
            Return _SortOrder
        End Get
        Set(value As SortOrder)
            If Not _SortOrder = value Then
                _SortOrder = value
                If value = SortOrder.Ascending Then SortAscending()
                If value = SortOrder.Descending Then SortDescending()
            End If
        End Set
    End Property
    Friend Sub New(TreeViewer As TreeViewer)
        _TreeViewer = TreeViewer
    End Sub
    Private _TreeViewer As TreeViewer
    Public ReadOnly Property TreeViewer As TreeViewer
        Get
            Return _TreeViewer
        End Get
    End Property
    Friend _Parent As Node
    Public ReadOnly Property Parent As Node
        Get
            Return _Parent
        End Get
    End Property
    Public ReadOnly Property All As List(Of Node)
        Get
            Dim Nodes As New List(Of Node)
            Dim Queue As New Queue(Of Node)
            Dim TopNode As Node, Node As Node
            For Each TopNode In Me
                Queue.Enqueue(TopNode)
            Next
            While Queue.Any
                TopNode = Queue.Dequeue
                Nodes.Add(TopNode)
                For Each Node In TopNode.Children
                    Queue.Enqueue(Node)
                Next
            End While
            Nodes.Sort(Function(x, y) x.VisibleIndex.CompareTo(y.VisibleIndex))
            Return Nodes
        End Get
    End Property
    Public ReadOnly Property Roots As List(Of Node)
        Get
            Return All.Where(Function(x) x.Level = 0).ToList
        End Get
    End Property
    Public ReadOnly Property Level(LevelIndex As Integer) As List(Of Node)
        Get
            Return All.Where(Function(x) x.Level = LevelIndex).ToList
        End Get
    End Property
    Public ReadOnly Property Visible As List(Of Node)
        Get
            Return All.Where(Function(m) m.Visible).ToList
        End Get
    End Property
    Public ReadOnly Property Parents As List(Of Node)
        Get
            Return All.Where(Function(m) m.Visible And m.HasChildren).ToList
        End Get
    End Property
    Public ReadOnly Property Client As New List(Of Node)
    Public ReadOnly Property Draw As New List(Of Node)
    Public ReadOnly Property CollectionDataType As Type
        Get
            Dim Types = From n In Me Select n.DataType
            Return GetDataType(Types)
        End Get
    End Property

    Public Overloads Function Contains(ByVal Name As String) As Boolean

        If Count = 0 Then
            Return False
        Else
            Dim Found As Boolean = False
            For Each Node In All
                If Node.Name = Name Then
                    Found = True
                    Exit For
                End If
            Next
            Return Found
        End If

    End Function
    Public Overloads Function Add(Name As String, Text As String, Checked As Boolean, Image As Image) As Node
        Return Add(New Node With {.Name = Name, .Text = Text, .Checked = Checked, .Image = Image})
    End Function
    Public Overloads Function Add(Name As String, Text As String, Checked As Boolean) As Node
        Return Add(New Node With {.Name = Name, .Text = Text, .Checked = Checked})
    End Function
    Public Overloads Function Add(Name As String, Text As String, Image As Image) As Node
        Return Add(New Node With {.Name = Name, .Text = Text, .Image = Image})
    End Function
    Public Overloads Function Add(Name As String, Text As String) As Node
        Return Add(New Node With {.Name = Name, .Text = Text})
    End Function
    Public Overloads Function Add(Text As String, Checked As Boolean, Image As Image) As Node
        Return Add(New Node With {.Text = Text, .Checked = Checked, .Image = Image})
    End Function
    Public Overloads Function Add(Text As String, Checked As Boolean) As Node
        Return Add(New Node With {.Text = Text, .Checked = Checked})
    End Function
    Public Overloads Function Add(Text As String, Image As Image) As Node
        Return Add(New Node With {.Text = Text, .Image = Image})
    End Function
    Public Overloads Function Add(Text As String) As Node
        Return Add(New Node With {.Text = Text})
    End Function
    Public Overloads Function Add(ByVal AddNode As Node) As Node

        If AddNode IsNot Nothing Then
            With AddNode
                ._Index = Count
                If Parent Is Nothing Then   ' *** ROOT NODE
                    'mTreeViewer was set when TreeViewer was created with New NodeCollection
                    ._TreeViewer = TreeViewer
                    ._Visible = True
                    ._Level = 0

                Else                        ' *** CHILD NODE
                    REM /// Get TreeViewer value from Parent. A Node Collection shares some Node properties
                    _TreeViewer = Parent.TreeViewer
                    ._TreeViewer = Parent.TreeViewer
                    ._Parent = Parent
                    ._Level = Parent.Level + 1
                    If Parent.Expanded Then
                        ._Visible = True
                        If Count = 0 Then
                            ._VisibleIndex = 0
                        Else
                            ._VisibleIndex = Last.VisibleIndex + 1
                        End If
                    Else
                        ._VisibleIndex = -1
                    End If

                End If
                If TreeViewer IsNot Nothing Then
                    .Font = TreeViewer.Font
                    TreeViewer.NodeTimer_Start(AddNode)
                End If
            End With
            MyBase.Add(AddNode)
        End If
        Return AddNode

    End Function
    Public Overloads Function AddRange(ByVal Nodes As List(Of Node)) As List(Of Node)

        If Nodes IsNot Nothing Then
            For Each Node As Node In Nodes
                Add(Node)
            Next
            If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        End If
        Return Nodes

    End Function
    Public Overloads Function AddRange(ByVal Nodes As Node()) As Node()

        If Nodes IsNot Nothing Then
            For Each Node As Node In Nodes
                Add(Node)
            Next
            If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        End If
        Return Nodes

    End Function
    Public Overloads Function AddRange(ByVal Nodes As String()) As Node()

        Dim NewNodes As New List(Of Node)
        If Nodes IsNot Nothing Then
            For Each NewNode As String In Nodes
                Add(NewNode)
                NewNodes.Add(New Node With {.Text = NewNode})
            Next
            If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        End If
        Return NewNodes.ToArray

    End Function
    Public Overloads Function Clear(ByVal Nodes As NodeCollection) As NodeCollection

        Clear()
        If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        Return Nodes

    End Function
    Public Overloads Function Insert(Index As Integer, InsertNode As Node) As Node

        If InsertNode IsNot Nothing Then
            If IsNothing(TreeViewer) Then
                MyBase.Insert(Index, InsertNode)
            Else
                If TreeViewer.Ancestors.All.Contains(InsertNode) Then
                    'Throw New ArgumentException("This node already exists in the Treeviewer. Try Removing the Node")

                Else
                    MyBase.Insert(Index, InsertNode)
                    InsertNode._TreeViewer = TreeViewer
                    InsertNode._Parent = Parent
                    TreeViewer.NodeTimer_Start(InsertNode)
                End If
            End If
        End If
        Return InsertNode

    End Function
    Public Overloads Function Remove(ByVal RemoveNode As Node) As Node

        If RemoveNode IsNot Nothing Then
            MyBase.Remove(RemoveNode)
            TreeViewer.NodeTimer_Start(RemoveNode)
        End If
        Return RemoveNode

    End Function
    Public Shadows Function Item(ByVal Name As String) As Node
        Dim Nodes As New List(Of Node)((From N In Me Where N.Name = Name).ToArray)
        Return If(Nodes.Any, Nodes.First, Nothing)
    End Function
    Public Shadows Function ItemByTag(ByVal TagObject As Object) As Node
        Dim Nodes As New List(Of Node)((From N In All Where N.Tag Is TagObject).ToArray)
        Return If(Nodes.Any, Nodes.First, Nothing)
    End Function
    Friend Sub SortAscending(Optional repaint As Boolean = True)

        If TreeViewer?.FavoritesFirst Then
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite) 'False=0, True=1 
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = String.Compare(x.Text, y.Text, StringComparison.InvariantCulture)
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Boolean)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Date)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)
            End Select
        Else
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(x, y) String.Compare(x.Text, y.Text, StringComparison.Ordinal))

                Case GetType(Boolean)
                    Sort(Function(x, y) Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture)))

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(x, y) Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture)))

                Case GetType(Date)
                    Sort(Function(x, y) Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture)))

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(x, y) Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture)))

            End Select
        End If
        If repaint Then TreeViewer?.RequiresRepaint()

    End Sub
    Friend Sub SortDescending(Optional repaint As Boolean = True)

        If TreeViewer?.FavoritesFirst Then
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = String.Compare(x.Text, y.Text, StringComparison.InvariantCulture)
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Boolean)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Date)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)
            End Select
        Else
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(y, x) String.Compare(x.Text, y.Text, StringComparison.Ordinal))

                Case GetType(Boolean)
                    Sort(Function(y, x) Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture)))

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(y, x) Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture)))

                Case GetType(Date)
                    Sort(Function(y, x) Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture)))

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(y, x) Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture)))

            End Select
        End If
        If repaint Then TreeViewer?.RequiresRepaint()

    End Sub
End Class

REM //////////////// NODE \\\\\\\\\\\\\\\\
Public Class Node
    Implements IDisposable
    Public Sub New()
        Children = New NodeCollection(_TreeViewer) With {._Parent = Me}
        Fields = New NodeCollection(_TreeViewer) With {._Parent = Me}
    End Sub
    Public Sub New(nodeText As String)
        Text = nodeText
        Children = New NodeCollection(_TreeViewer) With {._Parent = Me}
        Fields = New NodeCollection(_TreeViewer) With {._Parent = Me}
    End Sub
    Public Enum SeparatorPosition
        None
        Above
        Below
    End Enum

    Friend _TreeViewer As TreeViewer
    Public ReadOnly Property TreeViewer As TreeViewer
        Get
            Return _TreeViewer
        End Get
    End Property
    Friend _Parent As Node
    Public ReadOnly Property Parent As Node
        Get
            Return _Parent
        End Get
    End Property
    Public ReadOnly Property Children As NodeCollection
    Public ReadOnly Property Fields As NodeCollection
    Friend HeaderLevel_ As Byte = 255
    Friend ReadOnly Property HeaderLevel As Byte
        Get
            Return HeaderLevel_
        End Get
    End Property
    Friend ColumnIndex_ As Byte = 255
    Friend ReadOnly Property ColumnIndex As Byte
        Get
            Return ColumnIndex_
        End Get
    End Property
    Friend IsField_ As Boolean
    Friend ReadOnly Property IsField As Boolean
        Get
            Return IsField_
        End Get
    End Property
    Public Property CursorGlowColor As Color = Color.LimeGreen
    Public ReadOnly Property HasChildren As Boolean
        Get
            Return Children.Any
        End Get
    End Property
    Friend _Expanded As Boolean
    Public ReadOnly Property Expanded As Boolean
        Get
            Return _Expanded
        End Get
    End Property
    Public ReadOnly Property Collapsed As Boolean
        Get
            Return Not Expanded
        End Get
    End Property
    Private _CheckBox As Boolean = False
    Public Property CheckBox As Boolean
        Get
            If TreeViewer Is Nothing Then
                Return _CheckBox
            Else
                If TreeViewer.CheckBoxes = TreeViewer.CheckState.All Then
                    _CheckBox = True
                ElseIf TreeViewer.CheckBoxes = TreeViewer.CheckState.None Then
                    _CheckBox = False
                ElseIf TreeViewer.CheckBoxes = TreeViewer.CheckState.Mixed Then

                End If
                Return _CheckBox
            End If
        End Get
        Set(value As Boolean)
            If value <> _CheckBox Then
                _CheckBox = value
                If TreeViewer Is Nothing Then
                Else
                    If value Then
                        If TreeViewer.CheckBoxes = TreeViewer.CheckState.None Then
                            TreeViewer.CheckBoxes = TreeViewer.CheckState.Mixed
                        End If
                    End If
                End If
                RequiresRepaint()
            End If
        End Set
    End Property
    Private _Checked As Boolean = False
    Public Property Checked As Boolean
        Get
            If Descendants.Any Then
                _Checked = (From d In Descendants Where d.Checked).Count = Descendants.Count
            End If
            Return _Checked
        End Get
        Set(value As Boolean)
            If Not value = _Checked Then
                If value Then CheckBox = True
                _Checked = value
                For Each Child In Descendants
                    Child.Checked = value
                Next
                RequiresRepaint()
            End If
        End Set
    End Property
    Public ReadOnly Property PartialChecked As Boolean
        Get
            Dim _descendants As New List(Of Node)(Descendants)
            Dim descendantCheckCount As Integer = (From d In _descendants Where d.Checked).Count
            Return descendantCheckCount > 0 And descendantCheckCount < _descendants.Count
        End Get
    End Property
    Private _Image As Image
    Public Property Image As Image
        Get
            Return _Image
        End Get
        Set(value As Image)
            _Image = value
            RequiresRepaint()
        End Set
    End Property
    Public Property CanEdit As Boolean = True
    Public Property CanAdd As Boolean = True
    Public Property CanRemove As Boolean = True
    Public Property CancelAction As Boolean = False
    Private _ImageScaling As Boolean = False
    Public Property ImageScaling As Boolean
        Get
            Return _ImageScaling
        End Get
        Set(value As Boolean)
            _ImageScaling = value
            RequiresRepaint()
        End Set
    End Property
    Public ReadOnly Property Root As Node
        Get
            Dim RootNode As Node = Me
            Do While Not IsNothing(RootNode.Parent)
                RootNode = RootNode.Parent
            Loop
            Return RootNode
        End Get
    End Property
    Public ReadOnly Property IsRoot As Boolean
        Get
            Return Parent Is Nothing
        End Get
    End Property
    Public Property Name As String
    Public Property Tag As Object
    Private _Text As String
    Public Property Text As String
        Get
            If ShowNodeIndex Then
                Return _Text & " [" & Index.ToString(InvariantCulture) & "]"
            Else
                Return _Text
            End If
        End Get
        Set(value As String)
            If Not _Text = value Then
                _Text = Replace(value, "&", "&&")
                _Bounds.Width = TextRenderer.MeasureText(value, Font).Width
                RequiresRepaint()
            End If
        End Set
    End Property
    Public Property TipText As String
    Public ReadOnly Property Options As New List(Of Object)
    Public ReadOnly Property ChildOptions As New List(Of Object)
    Private _Font As New Font("Calibri", 9)
    Public Property Font As Font
        Get
            Return _Font
        End Get
        Set(value As Font)
            _Font = value
            RequiresRepaint()
        End Set
    End Property
    Private _ForeColor As Color = Color.Black
    Public Property ForeColor As Color
        Get
            Return _ForeColor
        End Get
        Set(value As Color)
            _ForeColor = value
            RequiresRepaint()
        End Set
    End Property
    Private _TextBackColor As Color = Color.Transparent
    Public Property TextBackColor As Color
        Get
            Return _TextBackColor
        End Get
        Set(value As Color)
            _TextBackColor = value
            RequiresRepaint()
        End Set
    End Property
    Private _BackColor As Color = Color.Transparent
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(value As Color)
            _BackColor = value
            RequiresRepaint()
        End Set
    End Property
    Friend _Selected As Boolean
    Public ReadOnly Property Selected As Boolean
        Get
            Return _Selected
        End Get
    End Property
    Friend _Index As Integer
    Public ReadOnly Property Index As Integer
        Get
            Return _Index
        End Get
    End Property
    Public ReadOnly Property Height As Integer
        Get
            Dim ImageHeight As Integer = 0
            If Not IsNothing(Image) And Not ImageScaling Then
                ImageHeight = Image.Height
            End If
            Return If(Separator = SeparatorPosition.None, 0, 1) + Convert.ToInt32({1 + Font.GetHeight + 1, ImageHeight}.Max)
        End Get
    End Property
    Friend _ExpandCollapseBounds As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property ExpandCollapseBounds As Rectangle
        Get
            Return _ExpandCollapseBounds
        End Get
    End Property
    Friend _FavoriteBounds As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property FavoriteBounds As Rectangle
        Get
            Return _FavoriteBounds
        End Get
    End Property
    Friend _CheckBounds As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property CheckBounds As Rectangle
        Get
            Return _CheckBounds
        End Get
    End Property
    Friend _ImageBounds As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property ImageBounds As Rectangle
        Get
            Return _ImageBounds
        End Get
    End Property
    Friend _Bounds As New Rectangle
    Public ReadOnly Property Bounds As Rectangle
        Get
            Return _Bounds
        End Get
    End Property
    Friend _Level As Integer
    Public ReadOnly Property Level As Integer
        Get
            Return _Level
        End Get
    End Property
    Private ReadOnly _Path As New List(Of KeyValuePair(Of Integer, String))
    Public ReadOnly Property NamePath() As String
        Get
            _Path.Clear()
            GetNamePath(Me)
            Return Join(_Path.OrderBy(Function(x) x.Key).Select(Function(y) y.Value).ToArray, "±")
        End Get
    End Property
    Private Sub GetNamePath(_Node As Node)
        _Path.Add(New KeyValuePair(Of Integer, String)(_Node.Level, _Node.Name))
        If Not IsNothing(_Node.Parent) Then GetNamePath(_Node.Parent)
    End Sub
    Public ReadOnly Property TextPath() As String
        Get
            _Path.Clear()
            GetTextPath(Me)
            Return Join(_Path.OrderBy(Function(x) x.Key).Select(Function(y) y.Value).ToArray, "±")
        End Get
    End Property
    Private Sub GetTextPath(_Node As Node)
        _Path.Add(New KeyValuePair(Of Integer, String)(_Node.Level, _Node.Text))
        If Not IsNothing(_Node.Parent) Then GetTextPath(_Node.Parent)
    End Sub
    Public ReadOnly Property Parents As List(Of Node)
        Get
            Dim _Parents As New List(Of Node)
            Dim ParentNode As Node = Parent
            Do While ParentNode IsNot Nothing
                _Parents.Add(ParentNode)
                ParentNode = ParentNode.Parent
            Loop
            Return _Parents
        End Get
    End Property
    Private Sub GetDescendants(_Node As Node)

        For Each Child In _Node.Children
            _Descendants.Add(Child)
            If Child.HasChildren Then GetDescendants(Child)
        Next

    End Sub
    Private ReadOnly _Descendants As New List(Of Node)
    Public ReadOnly Property Descendants As List(Of Node)
        Get
            _Descendants.Clear()
            GetDescendants(Me)
            Return _Descendants
        End Get
    End Property
    Private _Separator As SeparatorPosition = SeparatorPosition.None
    Public Property Separator As SeparatorPosition
        Get
            Return _Separator
        End Get
        Set(value As SeparatorPosition)
            If _Separator <> value Then
                _Separator = value
                RequiresRepaint()
            End If
        End Set
    End Property
    Private _ShowNodeIndex As Boolean = False
    Public Property ShowNodeIndex As Boolean
        Get
            Return _ShowNodeIndex
        End Get
        Set(value As Boolean)
            _ShowNodeIndex = value
            RequiresRepaint()
        End Set
    End Property
    Friend _Visible As Boolean = False
    Public ReadOnly Property Visible As Boolean
        Get
            Return _Visible
        End Get
    End Property
    Friend _VisibleIndex As Integer
    Public ReadOnly Property VisibleIndex As Integer
        Get
            Return _VisibleIndex
        End Get
    End Property
    Public ReadOnly Property Siblings As List(Of Node)
        Get
            Dim BrothersSistersAndMe As New List(Of Node)
            If IsRoot Then
                If Not TreeViewer Is Nothing Then BrothersSistersAndMe.AddRange(TreeViewer.Ancestors)

            Else
                BrothersSistersAndMe.AddRange(Parent.Children)

            End If
            BrothersSistersAndMe.Sort(Function(x, y) x.Index.CompareTo(y.Index))
            Return BrothersSistersAndMe
        End Get
    End Property
    Public ReadOnly Property FirstSibling As Node
        Get
            If Siblings.Any Then
                Return Siblings.First
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property IsFirstSibling As Boolean
        Get
            Return FirstSibling Is Me
        End Get
    End Property
    Public ReadOnly Property NextSibling As Node
        Get
            If Siblings.Any Then
                If Me Is LastSibling Then
                    Return Nothing
                Else
                    Return Siblings(Index + 1)
                End If
            Else
                Return Nothing
            End If

        End Get
    End Property
    Public ReadOnly Property LastSibling As Node
        Get
            If Siblings.Any Then
                Return Siblings.Last
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property IsLastSibling As Boolean
        Get
            Return LastSibling Is Me
        End Get
    End Property
    Public ReadOnly Property DataType As Type
        Get
            Return GetDataType(SortValue)
        End Get
    End Property
    Public Property CanDragDrop As Boolean = True
    Public Property CanFavorite As Boolean = False
    Private Favorite_ As Boolean = False
    Public Property Favorite As Boolean
        Get
            Return Favorite_
        End Get
        Set(value As Boolean)
            If value <> Favorite_ Then
                Favorite_ = value
                If TreeViewer?.FavoritesFirst Then Parent?.Children.SortAscending()
                RequiresRepaint()
            End If
        End Set
    End Property
    Private _SortValue As String = String.Empty
    Public Property SortValue As String
        Get
            If If(_SortValue, String.Empty).Length = 0 Then
                _SortValue = Text
            End If
            Return _SortValue
        End Get
        Set(value As String)
            If value <> _SortValue Then
                _SortValue = value
            End If
        End Set
    End Property

    Friend _Clicked As Boolean
    Private Sub RequiresRepaint()

        If Not IsNothing(TreeViewer) Then
            If _Clicked Then
                _Clicked = False
                TreeViewer.RequiresRepaint()
            Else
                TreeViewer.NodeTimer_Start(Me)
            End If
        End If

    End Sub
    Public Sub Expand()

        If Children.Any Then
            _Expanded = True
            ShowHide(Children, True)
            RequiresRepaint()
        End If

    End Sub
    Public Sub Collapse()

        If Children.Any Then
            _Expanded = False
            ShowHide(Children, False)
            RequiresRepaint()
        End If

    End Sub
    Public Sub Click()
        _Selected = True
        RequiresRepaint()
    End Sub
    Public Sub RemoveMe()
        Try
            Parent?.Children.Remove(Me)
        Catch ex As InvalidOperationException
        End Try
    End Sub
    Private Sub ShowHide(Nodes As List(Of Node), Optional Flag As Boolean = True)

        For Each Node As Node In Nodes
            If Node.Parent Is Nothing Then
                Node._Visible = True

            Else
                If Node.Parent.Expanded Then
                    Node._Visible = Flag

                Else
                    Node._Visible = False

                End If
            End If
            If Node.HasChildren Then ShowHide(Node.Children, Node._Visible)
        Next

    End Sub
    Public ReadOnly Property SortType As Type
        Get
            Return GetDataType((From c In Children Select c.SortValue).ToList)
        End Get
    End Property
    Public Sub SortChildren(Optional SortOrder As SortOrder = SortOrder.Ascending)

        Select Case SortType
            Case GetType(String)
                If SortOrder = SortOrder.Ascending Then Children.Sort(Function(x, y) String.Compare(Convert.ToString(x.SortValue, InvariantCulture), Convert.ToString(y.SortValue, InvariantCulture), StringComparison.Ordinal))
                If SortOrder = SortOrder.Descending Then Children.Sort(Function(y, x) String.Compare(Convert.ToString(x.SortValue, InvariantCulture), Convert.ToString(y.SortValue, InvariantCulture), StringComparison.Ordinal))

            Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                If SortOrder = SortOrder.Ascending Then Children.Sort(Function(x, y) Convert.ToInt64(x.SortValue, InvariantCulture).CompareTo(Convert.ToInt64(y.SortValue, InvariantCulture)))
                If SortOrder = SortOrder.Descending Then Children.Sort(Function(y, x) Convert.ToInt64(x.SortValue, InvariantCulture).CompareTo(Convert.ToInt64(y.SortValue, InvariantCulture)))

            Case GetType(Double), GetType(Decimal)
                If SortOrder = SortOrder.Ascending Then Children.Sort(Function(x, y) Convert.ToDouble(x.SortValue, InvariantCulture).CompareTo(Convert.ToDouble(y.SortValue, InvariantCulture)))
                If SortOrder = SortOrder.Descending Then Children.Sort(Function(y, x) Convert.ToDouble(x.SortValue, InvariantCulture).CompareTo(Convert.ToDouble(y.SortValue, InvariantCulture)))

            Case GetType(Date)
                If SortOrder = SortOrder.Ascending Then Children.Sort(Function(x, y) Convert.ToDateTime(x.SortValue, InvariantCulture).CompareTo(Convert.ToDateTime(y.SortValue, InvariantCulture)))
                If SortOrder = SortOrder.Descending Then Children.Sort(Function(y, x) Convert.ToDateTime(x.SortValue, InvariantCulture).CompareTo(Convert.ToDateTime(y.SortValue, InvariantCulture)))

            Case GetType(Boolean)
                If SortOrder = SortOrder.Ascending Then Children.Sort(Function(x, y) Convert.ToBoolean(x.SortValue, InvariantCulture).CompareTo(Convert.ToBoolean(y.SortValue, InvariantCulture)))
                If SortOrder = SortOrder.Descending Then Children.Sort(Function(y, x) Convert.ToBoolean(x.SortValue, InvariantCulture).CompareTo(Convert.ToBoolean(y.SortValue, InvariantCulture)))

        End Select

    End Sub
    Public Overrides Function ToString() As String
        Return Join({Text, If(Children.Any, "( " & Children.Count & " )", String.Empty)})
    End Function

#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                Me._Font.Dispose()
                Me.Font.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

Public Class HeaderLevels
    Inherits List(Of ColumnHeadCollection)
    Public ReadOnly Property Parent As TreeViewer
    Friend Sub New(tree As TreeViewer)
        Parent = tree
    End Sub
    Public Shadows Function Add(Headers As ColumnHeadCollection) As ColumnHeadCollection

        If Headers IsNot Nothing Then
            Headers.Parent_ = Parent
            MyBase.Add(Headers)
        End If
        Return Headers

    End Function
End Class

REM //////////////// HEADER COLLECTION \\\\\\\\\\\\\\\\
Public Class ColumnHeadCollection
    Inherits List(Of ColumnHead)
    Friend Parent_ As TreeViewer
    Public ReadOnly Property Parent As TreeViewer
        Get
            Return Parent_
        End Get
    End Property
    Public Sub New()
    End Sub
    Public Sub New(columns As String())

        If columns IsNot Nothing Then
            For Each column As String In columns
                Add(column)
            Next
        End If

    End Sub
    Public Property Height As Integer = 32
    Public Shadows Function AddRange(Headers As ColumnHead()) As ColumnHead()

        For Each Header In Headers
            Add(Header)
        Next
        Return Headers

    End Function
    Public Shadows Function AddRange(headerNames As String()) As ColumnHead()

        Dim Headers As New List(Of ColumnHead)
        For Each headerName In headerNames
            Dim newHeader As New ColumnHead(headerName)
            Headers.Add(newHeader)
            Add(newHeader)
        Next
        Return Headers.ToArray

    End Function
    Public Shadows Function Add(ByVal Text As String) As ColumnHead

        Dim addHeader As New ColumnHead(Text) With {
            .Width = 100
        }
        Return Add(addHeader)

    End Function
    Public Shadows Function Add(ByVal Text As String, ByVal Width As Integer) As ColumnHead

        Dim Header As New ColumnHead(Text) With {
            .Text = Text,
            .Width = Width
        }
        Return Add(Header)

    End Function
    Public Shadows Function Add(addHeader As ColumnHead) As ColumnHead

        AddHandler addHeader.Changed, AddressOf OnChanged
        MyBase.Add(addHeader)
        addHeader.Parent_ = Me
        OnChanged(Me, New EventArgs)
        Return addHeader

    End Function
    Public Shadows Sub Insert(ByVal index As Integer, ByVal insertHeader As ColumnHead)

        MyBase.Insert(index, insertHeader)
        OnChanged(Me, New EventArgs)

    End Sub
    Public Shadows Sub Remove(ByVal removeHeader As ColumnHead)

        MyBase.Remove(removeHeader)
        OnChanged(Me, New EventArgs)

    End Sub
    Public Event Changed(ByVal sender As Object, ByVal e As EventArgs)
    Protected Friend Sub OnChanged(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent Changed(sender, e)
    End Sub
    Public Event Clicked(ByVal sender As ColumnHead)
    Protected Friend Sub OnClicked(ByVal sender As ColumnHead)
        RaiseEvent Clicked(sender)
    End Sub
End Class
REM //////////////// HEADER \\\\\\\\\\\\\\\\
<Serializable> Public Class ColumnHead
    Friend Parent_ As ColumnHeadCollection
    Public ReadOnly Property Parent As ColumnHeadCollection
        Get
            Return Parent_
        End Get
    End Property
    Public Sub New(headerText As String)
        Text = headerText
    End Sub
    Property Resizeable As Boolean
    Private WithEvents Style_ As New CellStyle With {
        .BackColor = Color.Transparent,
        .ShadeColor = Color.White,
        .ForeColor = Color.Black,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.None
    }
    Public ReadOnly Property Style As CellStyle
        Get
            Return Style_
        End Get
    End Property
    Friend ReadOnly Property StyleChanged As Boolean
    Private Sub Style_PropertyChanged(sender As Object, e As StyleEventArgs) Handles Style_.PropertyChanged
        Using defaultStyle As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.White, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
            _StyleChanged = Style_ <> defaultStyle
        End Using
        Parent?.Parent?.Invalidate()
    End Sub
    Private Text_ As String = "Header"
    Property Text As String
        Get
            Return Text_
        End Get
        Set(ByVal Value As String)
            If Text_ <> Value Then
                Text_ = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private Width_ As Integer = 100
    Property Width As Integer
        Get
            Return Width_
        End Get
        Set(ByVal Value As Integer)
            If Width_ <> Value Then
                Width_ = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private Image_ As Image = Nothing
    Property Image As Image
        Get
            Return Image_
        End Get
        Set(ByVal Value As Image)
            If Not SameImage(Image_, Value) Then
                Image_ = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Friend Bounds_ As Rectangle
    ReadOnly Property Bounds As Rectangle
        Get
            Return Bounds_
        End Get
    End Property
    Public Event Changed(ByVal sender As Object, ByVal e As EventArgs)
    Protected Friend Sub OnChanged(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent Changed(sender, e)
    End Sub
    Public Event Clicked(ByVal sender As Header)
    Protected Friend Sub OnClicked(ByVal sender As Header)
        RaiseEvent Clicked(sender)
    End Sub
    Public Overrides Function ToString() As String
        Return Join({Text, Width.ToString, Style.Alignment.ToString}, ".")
    End Function
End Class