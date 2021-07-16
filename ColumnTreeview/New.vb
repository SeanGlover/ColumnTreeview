Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.ComponentModel
Imports Controls

Public Enum MouseRegion
    None
    Column
    Expander
    Favorite
    CheckBox
    Image
    Node
    Field
End Enum
Public Class HitRegion
    Implements IEquatable(Of HitRegion)
    Public Property Region As MouseRegion
    Public Property Node As Node
    Public Property Column As ColumnHead
    Public Overrides Function GetHashCode() As Integer
        Return Region.GetHashCode Xor Node.GetHashCode
    End Function
    Public Overloads Function Equals(other As HitRegion) As Boolean Implements IEquatable(Of HitRegion).Equals
        If other Is Nothing Then
            Return Me Is Nothing
        Else
            Return Region = other.Region AndAlso Node Is other.Node
        End If
    End Function
    Public Shared Operator =(value1 As HitRegion, value2 As HitRegion) As Boolean
        If value1 Is Nothing Then
            Return value2 Is Nothing
        Else
            Return value1.Equals(value2)
        End If
    End Operator
    Public Shared Operator <>(value1 As HitRegion, value2 As HitRegion) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(obj As Object) As Boolean
        If TypeOf obj Is HitRegion Then
            Return CType(obj, HitRegion) = Me
        Else
            Return False
        End If
    End Function
End Class
Public Class AlignFormat
    Implements IEquatable(Of AlignFormat)
    Public Enum TypeGroup
        None
        Booleans
        Decimals
        Integers
        Dates
        Times
        Images
        Strings
    End Enum
    Public ReadOnly Property HorizontalAlignment As StringAlignment
    Public ReadOnly Property DataType As Type
    Public ReadOnly Property FormatString As String
    Public ReadOnly Property Group As TypeGroup

    Public Sub New(DataType As Type)

        _DataType = DataType

        Select Case DataType
            Case GetType(Boolean), GetType(Byte), GetType(Short), GetType(Integer), GetType(Long), GetType(Date), GetType(DateAndTime), GetType(Image), GetType(Bitmap), GetType(Icon)
                HorizontalAlignment = StringAlignment.Center

            Case GetType(Decimal), GetType(Double)
                HorizontalAlignment = StringAlignment.Far

            Case Else
                HorizontalAlignment = StringAlignment.Near

        End Select

        Select Case DataType
            Case GetType(Boolean)
                _Group = TypeGroup.Booleans
                _FormatString = String.Empty

            Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                _Group = TypeGroup.Integers
                _FormatString = String.Empty

            Case GetType(Date)
                _Group = TypeGroup.Dates
                _FormatString = Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern

            Case GetType(DateAndTime)
                _Group = TypeGroup.Times
                _FormatString = Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.FullDateTimePattern

            Case GetType(Decimal), GetType(Double)
                _Group = TypeGroup.Decimals
                _FormatString = "C2"

            Case GetType(Image), GetType(Bitmap), GetType(Icon)
                _Group = TypeGroup.Images
                _FormatString = String.Empty

            Case GetType(String)
                _Group = TypeGroup.Strings
                _FormatString = String.Empty

            Case Else
                _Group = TypeGroup.None
                _FormatString = String.Empty

        End Select

    End Sub
    Public Overrides Function GetHashCode() As Integer
        Return HorizontalAlignment.GetHashCode Xor DataType.GetHashCode Xor FormatString.GetHashCode Xor Group.GetHashCode
    End Function
    Public Overloads Function Equals(other As AlignFormat) As Boolean Implements IEquatable(Of AlignFormat).Equals
        Return DataType Is other?.DataType
    End Function
    Public Shared Operator =(Object1 As AlignFormat, Object2 As AlignFormat) As Boolean
        Return Object1.Equals(Object2)
    End Operator
    Public Shared Operator <>(Object1 As AlignFormat, Object2 As AlignFormat) As Boolean
        Return Not Object1 = Object2
    End Operator
    Public Overrides Function Equals(obj As Object) As Boolean
        If TypeOf obj Is AlignFormat Then
            Return CType(obj, AlignFormat) = Me
        Else
            Return False
        End If
    End Function
End Class
Public Class ColumnEventArgs
    Inherits EventArgs
    Public ReadOnly Property Column As ColumnHead
    Public Sub New(Value As ColumnHead)
        Column = Value
    End Sub
End Class

Public Class NodeEventArgs
    Inherits EventArgs
    Public ReadOnly Property Node As Node
    Public ReadOnly Property Hit As HitRegion
    Public ReadOnly Property Nodes As List(Of Node)
    Public ReadOnly Property ProposedText As String = String.Empty
    Friend Sub New(Value As Node)
        Node = Value
    End Sub
    Friend Sub New(Values As List(Of Node))
        Nodes = Values
    End Sub
    Friend Sub New(Value As Node, NewText As String)
        Node = Value
        ProposedText = NewText
    End Sub
    Friend Sub New(region As HitRegion)
        Hit = region
        Node = region.Node
    End Sub
End Class
Public Class TreeViewer
    Inherits Control

    Private WithEvents Karen As New Hooker
    Public WithEvents VScroll As New VScrollBar
    Public WithEvents HScroll As New HScrollBar
    Private WithEvents NodeTimer As New Timer With {.Interval = 200}
    Private WithEvents ScrollTimer As New Timer With {.Interval = 50}
    Private ReadOnly DragData As New DragInfo
    Private _Cursor As Cursor

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

    Private WithEvents TSMIMultiSelect_ As New ToolStripMenuItem With {.Text = "Multi-Select", .Checked = False, .CheckOnClick = True}
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

    Private IgnoreSizeChanged As Boolean = False
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ C O N S T A N T S
    Private Const CheckHeight As Integer = 14
    Private Const VScrollWidth As Integer = 14
    Private Const HScrollHeight As Integer = 12
    Private ExpandHeight As Integer = 10

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ N O D E   I N D E X I N G
    Private NodeIndex As Integer = 0
    Private VisibleIndex As Integer
    Private RollingHeight As Integer
    Private RollingWidth As Integer

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ I M A G E S
    Private ReadOnly FavoriteImage As Image
    Private ExpandImage As Image
    Private CollapseImage As Image

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ M O U S E O V E R
    Private LastMouseNode As Node = Nothing
    Private LastMouseColumn As ColumnHead = Nothing
    Private MousePoint As Point

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ E V E N T S
    Public Event Alert(sender As Object, e As AlertEventArgs)
    Public Event ColumnClicked(sender As Object, e As ColumnEventArgs)
    Public Event NodesChanged(sender As Object, e As NodeEventArgs)
    Public Event NodeBeforeAdded(sender As Object, e As NodeEventArgs)
    Public Event NodeAfterAdded(sender As Object, e As NodeEventArgs)
    Public Event NodeBeforeRemoved(sender As Object, e As NodeEventArgs)
    Public Event NodeAfterRemoved(sender As Object, e As NodeEventArgs)
    Public Event NodeBeforeEdited(sender As Object, e As NodeEventArgs)
    Public Event NodeAfterEdited(sender As Object, e As NodeEventArgs)
    Public Event NodeDragStart(sender As Object, e As NodeEventArgs)
    Public Event NodeDragOver(sender As Object, e As NodeEventArgs)
    Public Event NodeDropped(sender As Object, e As NodeEventArgs)
    Public Event NodeChecked(sender As Object, e As NodeEventArgs)
    Public Event NodeExpanded(sender As Object, e As NodeEventArgs)
    Public Event NodeCollapsed(sender As Object, e As NodeEventArgs)
    Public Event NodeClicked(sender As Object, e As NodeEventArgs)
    Public Event NodeRightClicked(sender As Object, e As NodeEventArgs)
    Public Event NodeEnter(sender As Object, e As NodeEventArgs)
    Public Event NodeLeave(sender As Object, e As NodeEventArgs)
    Public Event NodeFavorited(sender As Object, e As NodeEventArgs)
    Public Event NodeDoubleClicked(sender As Object, e As NodeEventArgs)

    Private Class DragInfo
        Implements IEquatable(Of DragInfo)
        Implements IDisposable
        Friend MousePoints As New List(Of Point)
        Friend IsDragging As Boolean
        Friend DragNode As Node
        Friend DropHighlightNode As Node

        Public Overrides Function GetHashCode() As Integer
            Return MousePoints.GetHashCode Xor IsDragging.GetHashCode Xor DragNode.GetHashCode Xor DropHighlightNode.GetHashCode
        End Function
        Public Overloads Function Equals(other As DragInfo) As Boolean Implements IEquatable(Of DragInfo).Equals
            Return DragNode Is other?.DragNode
        End Function
        Public Shared Operator =(Object1 As DragInfo, Object2 As DragInfo) As Boolean
            Return Object1.Equals(Object2)
        End Operator
        Public Shared Operator <>(Object1 As DragInfo, Object2 As DragInfo) As Boolean
            Return Not Object1 = Object2
        End Operator
        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is DragInfo Then
                Return CType(obj, DragInfo) = Me
            Else
                Return False
            End If
        End Function
#Region "IDisposable Support"
        Private DisposedValue As Boolean ' To detect redundant calls IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not DisposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    DragNode.Dispose()
                    DropHighlightNode.Dispose()
                End If
                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            DisposedValue = True
        End Sub
        ' TODO: override Finalize() only if Dispose( disposing As Boolean) above has code to free unmanaged resources.
        Protected Overrides Sub Finalize()
            ' Do not change this code.  Put cleanup code in Dispose( disposing As Boolean) above.
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
    Public Enum CheckState
        None
        All
        Mixed
    End Enum
    Public Enum CheckStyle
        Slide
        Check
    End Enum
    Public Enum ExpandStyle
        PlusMinus
        Arrow
        Book
        LightBulb
    End Enum

    Public Sub New()

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)

        FavoriteImage = Base64ToImage(StarString)

        BackColor = Color.GhostWhite
#Region " GLOBAL OPTIONS SET-UP "
        With TSDD_Options
#Region " NODE EDITING "
            With TSMI_NodeEditing
                .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                .ImageScaling = ToolStripItemImageScaling.None
                .Image = Base64ToImage(EditString)
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
                IC_NodeEdit.Image = Base64ToImage(EditString)
                IC_NodeEdit.DropDown.SelectionColor = Color.Transparent
                IC_NodeAdd.Image = Base64ToImage(AddString)
                IC_NodeAdd.DropDown.SelectionColor = Color.Transparent
                IC_NodeRemove.Image = Base64ToImage(RemoveString)
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
            TSMIMultiSelect_.DropDownItems.AddRange({TSMI_SelectAll, TSMI_SelectNone})
            .Items.Add(TSMIMultiSelect_)
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
            TSMIMultiSelect_.Text = "Multi-Select".ToString(InvariantCulture)
        Else
            TSMIMultiSelect_.Text = "Single-Select".ToString(InvariantCulture)
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
    Private Sub OptionsClicked(sender As Object, e As EventArgs) Handles B_Arrow.Click, B_Book.Click, B_Default.Click, B_Light.Click, TSMI_ExpandCollapseAll.Click, TSMI_Checkboxes.Click, TSMI_CheckUncheckAll.Click, TSMIMultiSelect_.Click, TSMI_SelectAll.Click, TSMI_SelectNone.Click, TSMI_Sort.Click

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
        If sender Is TSMIMultiSelect_ Then MultiSelect = TSMIMultiSelect_.Checked
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

    Public Property FreezeRoot As Boolean = False
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
    Private WithEvents ColumnHeaders_ As New HeaderLevels(Me)
    Public ReadOnly Property ColumnHeaders As HeaderLevels
        Get
            Return ColumnHeaders_
        End Get
    End Property
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
                HScroll.Value = 0
                VScroll.Value = 0
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

                End If
#End Region
                If Table IsNot Nothing Then
                    With ColumnHeaders
                        .Clear()
                        Dim columnNames As New List(Of String)(From ch In Table.Columns Select DirectCast(ch, DataColumn).ColumnName)
                        If columnNames.Any Then .Add(New ColumnHeadCollection(columnNames.ToArray))
                        If ColumnHeaders.Any Then RecursiveBuild(Ancestors, 0, Table)
                    End With
                    _LoadTime = Now.Subtract(startLoad)
                End If
                Invalidate()
            End If
        End Set
    End Property
    Private Sub RecursiveBuild(nc As NodeCollection, lvl As Byte, tbl As DataTable)

        Dim headers As ColumnHeadCollection = ColumnHeaders(lvl)
        Dim headName As String = headers(0).Text 'Must be the first
        If tbl?.Columns.Contains(headName) Then
            Dim columnTypes As New Dictionary(Of String, Type)
            For Each column As DataColumn In tbl.Columns
                columnTypes.Add(column.ColumnName, column.DataType)
                headers.DataTypes(column.ColumnName) = column.DataType
            Next
            Dim tables = (From r In tbl.AsEnumerable Group r By groupX = r(headName).ToString Into headGroup = Group
                          Select New With {
                                                     .groupBy = groupX,
                                                     .table = headGroup.ToList
                                                     }).ToDictionary(Function(k) k.groupBy, Function(v) v.table)
            '/// Each table will be a subset of Rows
            For Each tableGroup In tables
                Dim firstRow As DataRow = tableGroup.Value.First
                Dim headNode As Node = nc.Add(tableGroup.Key)
                With headNode
                    .HeaderLevel_ = lvl
                    .IsHeader_ = True
                    .ColumnIndex_ = 0
                    .Header_ = headers(.ColumnIndex)
                    .Font = Font
                    .Row_ = firstRow
                    .Value = firstRow(headName)
                End With
                Dim xx As Integer = 0
                tableGroup.Value.ForEach(Sub(row)
                                             Dim subHeaders = headers.Skip(1).ToList
                                             If subHeaders.Any Then
                                                 Dim groupBy As String = subHeaders.First.Text
                                                 Dim nodeText As String = If(0 = 1, xx.ToString & "_", String.Empty) & row(groupBy).ToString
                                                 Dim childNode As Node = headNode.Children.Add(groupBy, nodeText)
                                                 With childNode
                                                     .HeaderLevel_ = lvl
                                                     .IsFieldParent_ = True
                                                     .ColumnIndex_ = 1
                                                     .Header_ = headers(.ColumnIndex)
                                                     .DataType_ = GetType(String)
                                                     .IsField_ = False
                                                     .Font = Font
                                                     .Row_ = row
                                                     .Value = row(groupBy)
                                                 End With
                                                 subHeaders.RemoveAt(0)
                                                 Dim columnIndex As Byte = 2
                                                 subHeaders.ForEach(Sub(column)
                                                                        Dim fieldNode As Node = childNode.Fields.Add(column.Text, row(column.Text).ToString)
                                                                        With fieldNode
                                                                            .HeaderLevel_ = lvl
                                                                            .ColumnIndex_ = columnIndex
                                                                            .Header_ = headers(.ColumnIndex)
                                                                            .IsField_ = True
                                                                            .Font = Font
                                                                            .Value = row(column.Text)
                                                                        End With
                                                                        columnIndex += CByte(1)
                                                                    End Sub)
                                             End If
                                             If lvl + 1 < ColumnHeaders.Count Then RecursiveBuild(headNode.Children, lvl + CByte(1), tableGroup.Value.CopyToDataTable)
                                             xx += 1
                                         End Sub)
            Next
        Else
            'Throw New Exception
        End If

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public ReadOnly Property OptionsOpen As Boolean
    Public Property FavoritesFirst As Boolean = True
    Private CanFavorite_ As Boolean = False
    Public Property CanFavorite As Boolean
        Get
            Return CanFavorite_
        End Get
        Set(value As Boolean)
            If value <> CanFavorite_ Then
                CanFavorite_ = value
                For Each node In Ancestors.All
                    node.CanFavorite = value
                Next
                RequiresRepaint()
            End If
        End Set
    End Property
    Private _ExpanderStyle As ExpandStyle = ExpandStyle.PlusMinus
    Public Property ExpanderStyle As ExpandStyle
        Get
            Return _ExpanderStyle
        End Get
        Set(value As ExpandStyle)
            _ExpanderStyle = value
            UpdateExpandImage()
        End Set
    End Property
    Private CheckboxStyle_ As CheckStyle = CheckStyle.Slide
    Public Property CheckboxStyle As CheckStyle
        Get
            Return CheckboxStyle_
        End Get
        Set(value As CheckStyle)
            CheckboxStyle_ = value
            RefreshNodesBounds_Lines(Ancestors)
        End Set
    End Property
    Private _CheckBoxes As CheckState = CheckState.Mixed
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <Category("Options")>
    <Description("None=No checkboxes, All=Each node will have a checkbox, Mixed=Some with, some without")>
    <RefreshProperties(RefreshProperties.All)>
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
    Public ReadOnly Property DropHighlightNode As Node
        Get
            Return DragData.DropHighlightNode
        End Get
    End Property
    Public Property MouseOverExpandsNode As Boolean = False
    Public Property CanAdd As Boolean = True
    Private MultiSelect_ As Boolean
    Public Property MultiSelect As Boolean
        Get
            Return MultiSelect_
        End Get
        Set(value As Boolean)
            If MultiSelect_ <> value Then
                If Not value Then
                    Ancestors.All.ForEach(Sub(node)
                                              node._Selected = False
                                              node.Fields.ForEach(Sub(field)
                                                                      field._Selected = False
                                                                  End Sub)
                                          End Sub)
                End If
                MultiSelect_ = value
                Invalidate()
            End If
        End Set
    End Property
    Public ReadOnly Property NodeHeight As Integer
        Get
            Return CInt(Ancestors.All.Average(Function(a) a.Height))
        End Get
    End Property
    Public ReadOnly Property SelectedNodes As New List(Of Node)
    Public ReadOnly Property Roots As List(Of Node)
        Get
            Return Ancestors.All.Where(Function(n) n.Level = 0).ToList
        End Get
    End Property
    Public ReadOnly Property Levels(level As Byte) As List(Of Node)
        Get
            Return Ancestors.All.Where(Function(n) n.Level = level).ToList
        End Get
    End Property
    Public ReadOnly Property Ancestors As New NodeCollection(Me)
    Public Property GroupBy As String
        Get
            Return If(Table Is Nothing OrElse Table.Columns.Count = 0, Nothing, Table.Columns(0).ColumnName)
        End Get
        Set(value As String)
            If value IsNot Nothing And Table IsNot Nothing AndAlso Table.Columns.Contains(value) Then
                Dim tempTable = Table.Copy
                DataSource = Nothing
                Ancestors.Clear()
                tempTable.Columns(value).SetOrdinal(0)
                DataSource = tempTable
                ColumnHeaders.ForEach(Sub(headers)
                                          With headers
                                              '.Styles.Theme = Theme.White
                                              '.MouseStyles.Theme = Theme.Purple
                                              .ForEach(Sub(header)
                                                           'header.Image = My.Resources.beeTiny
                                                       End Sub)
                                          End With
                                      End Sub)
            End If
        End Set
    End Property
    Public Property MaxNodes As Integer
    Public Property LineStyle As DashStyle = DashStyle.Dot
    Public Property LineColor As Color = Color.Blue
    Public Property RootLines As Boolean = True
    Public Property DropHighlightColor As Color = Color.Gainsboro
    Public Property MouseOverColor As Color = Color.Gainsboro
    Public Property SelectionColor As Color = Color.Gainsboro
    Public Property Offset As New Point(5, 3)
    Public Overrides Property AutoSize As Boolean = True
    Public Property ShowOptions As Boolean = True
    Public Property StopMe As Boolean
    Friend ReadOnly Property Hit As HitRegion
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

        If e IsNot Nothing Then
            Dim HitRegion As HitRegion = HitTest(e.Location)
            If HitRegion.Column Is Nothing Then
                Dim HitNode As Node = HitRegion.Node
                With DragData
                    .DragNode = HitNode
                    .IsDragging = False
                    .MousePoints = New List(Of Point)
                End With
                If e.Button = MouseButtons.Right Then
                    If ShowOptions Then
                        Dim ShowLocation As Point = Cursor.Position
                        ShowLocation.Offset(10, 0)
                        TSDD_Options.AutoClose = False
                        TSDD_Options.Show(ShowLocation)
                    End If

                Else
                    HideOptions()
                    If HitNode IsNot Nothing Then
                        With HitNode
                            Select Case HitRegion.Region
                                Case MouseRegion.Expander
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

                                Case MouseRegion.Favorite
                                    .Favorite = Not .Favorite
                                    RaiseEvent NodeFavorited(Me, New NodeEventArgs(HitNode))

                                Case MouseRegion.CheckBox
                                    ._Clicked = True
                                    .Checked = Not .Checked
                                    RaiseEvent NodeChecked(Me, New NodeEventArgs(HitNode))

                                Case MouseRegion.Image, MouseRegion.Node, MouseRegion.Field
                                    ._Clicked = True
                                    If Not MultiSelect Then
                                        For Each Node In SelectedNodes.Except({HitNode})
                                            Node._Selected = False
                                        Next
                                        SelectedNodes.Clear()
                                    End If
                                    HitNode._Selected = Not HitNode.Selected
                                    If HitNode.Selected Then
                                        SelectedNodes.Add(HitNode)
                                    Else
                                        SelectedNodes.Remove(HitNode)
                                    End If
                                    If e.Button = MouseButtons.Left Then
                                        RaiseEvent NodeClicked(Me, New NodeEventArgs(HitRegion))

                                    ElseIf e.Button = MouseButtons.Right Then
                                        RaiseEvent NodeRightClicked(Me, New NodeEventArgs(HitRegion))

                                    End If

                            End Select
                        End With
                        Invalidate()
                    End If
                End If
            Else
                Dim hitColumn As ColumnHead = HitRegion.Column
                Dim fieldName As String = hitColumn.Name
                Dim hitGroup = hitColumn.GridFormat.Group
                hitColumn.SortOrder = If(hitColumn.SortOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
                If hitColumn.Index = 0 Then
                    SortCollection(Ancestors, hitColumn)
                Else
                    Ancestors.ForEach(Sub(root)
                                          SortCollection(root.Children, hitColumn)
                                      End Sub)
                End If
                RaiseEvent ColumnClicked(Me, New ColumnEventArgs(HitRegion.Column))
            End If
            RequiresRepaint()
        End If
        MyBase.OnMouseDown(e)

    End Sub
    Private Sub SortCollection(collection As NodeCollection, hitColumn As ColumnHead)

        Dim fieldName As String = hitColumn.Name
        Dim hitGroup = hitColumn.GridFormat.Group
        If hitColumn.SortOrder = SortOrder.Ascending Then
            collection.Sort(Function(x, y)
                                Dim xValue As Object = If(hitColumn.Index < 2, x.Value, x.Fields.Item(fieldName).Value)
                                Dim yValue As Object = If(hitColumn.Index < 2, y.Value, y.Fields.Item(fieldName).Value)
                                Select Case hitGroup
                                    Case AlignFormat.TypeGroup.Decimals, AlignFormat.TypeGroup.Integers
                                        Return CDbl(xValue).CompareTo(CDbl(yValue))

                                    Case AlignFormat.TypeGroup.Dates, AlignFormat.TypeGroup.Times
                                        Return CDate(xValue).CompareTo(CDate(yValue))

                                    Case AlignFormat.TypeGroup.Booleans
                                        Return CBool(xValue).CompareTo(CBool(yValue))

                                    Case AlignFormat.TypeGroup.Strings
                                        Return xValue.ToString.CompareTo(yValue.ToString)

                                    Case Else
                                        'AlignFormat.TypeGroup.Images, AlignFormat.TypeGroup.None
                                        Return 0

                                End Select
                            End Function)
        Else
            collection.Sort(Function(y, x)
                                Dim xValue As Object = If(hitColumn.Index < 2, x.Value, x.Fields.Item(fieldName).Value)
                                Dim yValue As Object = If(hitColumn.Index < 2, y.Value, y.Fields.Item(fieldName).Value)
                                Select Case hitGroup
                                    Case AlignFormat.TypeGroup.Decimals, AlignFormat.TypeGroup.Integers
                                        Return CDbl(xValue).CompareTo(CDbl(yValue))

                                    Case AlignFormat.TypeGroup.Dates, AlignFormat.TypeGroup.Times
                                        Return CDate(xValue).CompareTo(CDate(yValue))

                                    Case AlignFormat.TypeGroup.Booleans
                                        Return CBool(xValue).CompareTo(CBool(yValue))

                                    Case AlignFormat.TypeGroup.Strings
                                        Return xValue.ToString.CompareTo(yValue.ToString)

                                    Case Else
                                        'AlignFormat.TypeGroup.Images, AlignFormat.TypeGroup.None
                                        Return 0

                                End Select
                            End Function)
        End If

    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            MousePoint = e.Location
            If e.Button = MouseButtons.None Then
                Dim redrawMe As Boolean = False
                Dim hitInfo = HitTest(e.Location)
                _Hit = hitInfo
                If Hit.Column IsNot LastMouseColumn Then
                    LastMouseColumn = Hit.Column
                    redrawMe = True
                End If
                If hitInfo.Node IsNot LastMouseNode Then
                    If hitInfo.Node Is Nothing Then
                        RaiseEvent NodeLeave(Me, New NodeEventArgs(LastMouseNode))
                    Else
                        RaiseEvent NodeEnter(Me, New NodeEventArgs(Hit.Node))
                    End If
                    If MouseOverExpandsNode Then
                        If hitInfo.Node Is Nothing Then
                            LastMouseNode.Collapse()
                        Else
                            hitInfo.Node.Expand()
                        End If
                    End If
                    LastMouseNode = hitInfo.Node
                    redrawMe = True
                End If
                If redrawMe Then Invalidate()

            ElseIf e.Button = MouseButtons.Left Then
                With DragData
                    If Not .MousePoints.Contains(e.Location) Then .MousePoints.Add(e.Location)
                    .IsDragging = .DragNode IsNot Nothing AndAlso (.MousePoints.Count >= 5 And Not .DragNode.Bounds_Text.Contains(.MousePoints.Last))
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
        With DragData
            .IsDragging = False
            .MousePoints.Clear()
        End With
        ScrollTimer.Stop()
        MyBase.OnMouseUp(e)

    End Sub
    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)

        If e IsNot Nothing Then
            Dim HitRegion As HitRegion = HitTest(e.Location)
            Dim HitNode As Node = HitRegion.Node
            If Not IsNothing(HitNode) Then
                RaiseEvent NodeDoubleClicked(Me, New NodeEventArgs(HitRegion))
            End If
        End If
        MyBase.OnMouseDoubleClick(e)

    End Sub

#Region " KEYPRESS EVENTS "
    Private Sub On_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs)

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

        If e IsNot Nothing AndAlso e.Modifiers = Keys.Control AndAlso e.KeyCode = Keys.C And SelectedNodes.Any Then
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
                _Cursor = Cursors.No
            End If
        End With

    End Sub
    Protected Overrides Sub OnDragLeave(e As EventArgs)

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
    Protected Overrides Sub OnDragEnter(e As DragEventArgs)

        ScrollTimer.Stop()
        Invalidate()
        MyBase.OnDragEnter(e)

    End Sub
    Protected Overrides Sub OnDragOver(e As DragEventArgs)

        If e IsNot Nothing Then
            e.Effect = DragDropEffects.All
            Dim Location As Point = PointToClient(New Point(e.X, e.Y))
            Dim HitRegion As HitRegion = HitTest(Location)
            Dim HitNode As Node = HitRegion?.Node
            If HitNode IsNot DragData.DropHighlightNode Then
                DragData.DropHighlightNode = HitNode
                Invalidate()
                If Not IsNothing(HitNode) Then
                    e.Data.SetData(GetType(Object), HitNode)
                    If HitRegion.Region = MouseRegion.Expander And HitNode.HasChildren Then
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
    Protected Overrides Sub OnDragDrop(e As DragEventArgs)

        If e IsNot Nothing Then
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
    Protected Overrides Sub OnGiveFeedback(e As GiveFeedbackEventArgs)

        If e IsNot Nothing Then
            e.UseDefaultCursors = False
            Cursor.Current = _Cursor
        End If
        MyBase.OnGiveFeedback(e)

    End Sub
    Private Sub DragScroll() Handles ScrollTimer.Tick

        If MouseButtons.HasFlag(MouseButtons.Left) And Not IsNothing(ScrollTimer.Tag) Then
            DrawNodes_Set()
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
    Protected Overrides Sub OnPaint(e As PaintEventArgs)

        If e IsNot Nothing Then
            With e.Graphics
                .TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
                .FillRectangle(New SolidBrush(BackColor), ClientRectangle)
                If BackgroundImage IsNot Nothing Then
                    Dim xOffset As Integer = {ClientRectangle.Width, Math.Abs(Convert.ToInt32((ClientRectangle.Width - BackgroundImage.Width) / 2))}.Min
                    Dim yOffset As Integer = {ClientRectangle.Height, Math.Abs(Convert.ToInt32((ClientRectangle.Height - BackgroundImage.Height) / 2))}.Min
                    Dim Bounds_Image As New Rectangle(xOffset, yOffset, {ClientRectangle.Width, BackgroundImage.Width}.Min, {ClientRectangle.Height, BackgroundImage.Height}.Min)
                    .DrawImage(BackgroundImage, Bounds_Image)
                End If
                Columns_SetBounds()
                If Ancestors.Any Then
                    Dim drawRootLines As Boolean = RootLines And Ancestors.Count > 1 'Doesn't make sense if only one root, when User wants RootLines
                    Dim firstRootNode As Node = Ancestors.First
                    Dim verticalRootLine_x As Integer = CInt({
                                            firstRootNode.Bounds_Check.Left,
                                            firstRootNode.Bounds_Favorite.Left,
                                            firstRootNode.Bounds_Image.Left,
                                            firstRootNode.Bounds_ShowHide.Left
                                                                          }.Min / 2)
                    Using linePen As New Pen(LineColor) With {.DashStyle = LineStyle}
                        For Each Node As Node In Ancestors.Draw
                            With Node
                                Dim mouseInTip As Boolean = False
                                If .Separator = Node.SeparatorPosition.Above And Node IsNot firstRootNode Then
                                    Using Pen As New Pen(Color.Blue, 1)
                                        Pen.DashStyle = DashStyle.DashDot
                                        e.Graphics.DrawLine(Pen, New Point(0, .Bounds_Text.Top), New Point(ClientRectangle.Right, .Bounds_Text.Top))
                                    End Using

                                ElseIf .Separator = Node.SeparatorPosition.Below And Node IsNot Ancestors.Last Then
                                    Using Pen As New Pen(Color.Blue, 1)
                                        Pen.DashStyle = DashStyle.DashDot
                                        e.Graphics.DrawLine(Pen, New Point(0, .Bounds_Text.Bottom), New Point(ClientRectangle.Right, .Bounds_Text.Bottom))
                                    End Using

                                End If
                                If .HasChildren Then e.Graphics.DrawImage(If(.Expanded, CollapseImage, ExpandImage), .Bounds_ShowHide)
                                If .CanFavorite Then e.Graphics.DrawImage(If(.Favorite, My.Resources.star, My.Resources.starEmpty), .Bounds_Favorite)

                                If .TipText IsNot Nothing Then
                                    Dim triangleHeight As Single = 8
                                    Dim trianglePoints As New List(Of PointF) From {
                                            New PointF(.Bounds_Text.Left, .Bounds_Text.Top),
                                            New PointF(.Bounds_Text.Left + triangleHeight, .Bounds_Text.Top),
                                            New PointF(.Bounds_Text.Left, .Bounds_Text.Top + triangleHeight)
                                    }
                                    e.Graphics.FillPolygon(Brushes.DarkOrange, trianglePoints.ToArray)
                                    mouseInTip = InTriangle(MousePoint, trianglePoints.ToArray)
                                End If
                                If .Image IsNot Nothing Then
                                    e.Graphics.DrawImage(.Image, .Bounds_Image)
                                    If Hit?.Node Is Node And Hit?.Region = MouseRegion.Image Then
                                        Using imageBrush As New SolidBrush(Color.FromArgb(96, MouseOverColor))
                                            e.Graphics.FillRectangle(imageBrush, .Bounds_Image)
                                        End Using
                                    End If
                                End If
                                Dim boundsColumn As Rectangle = .Header.Bounds
                                Dim boundsNode As Rectangle = If(Node.HasChildren,
                                    .Bounds_Text,
                                    New Rectangle(boundsColumn.Left, .Bounds_Text.Top, boundsColumn.Width, .Bounds_Text.Height)
                                    )
                                Using textBrush As New SolidBrush(.ForeColor)
                                    Using sf As New StringFormat With {
                                            .Alignment = If(Node.Header Is Nothing, StringAlignment.Near, Node.Header.GridFormat.HorizontalAlignment),
                                            .LineAlignment = StringAlignment.Center
                                        }
                                        e.Graphics.DrawString(If(mouseInTip, .TipText, .Text),
                                                                  .Font,
                                                                  textBrush,
                                                                  boundsNode,
                                                                  sf)
                                    End Using
                                End Using
                                Using Brush As New SolidBrush(If(DragData.DropHighlightNode Is Node, DropHighlightColor, .BackColor))
                                    'boundsNode.Inflate(-1, -1)
                                    e.Graphics.FillRectangle(Brush, boundsNode)
                                End Using
                                If Hit?.Node Is Node And .TipText Is Nothing Then
                                    Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, MouseOverColor))
                                        e.Graphics.FillRectangle(SemiTransparentBrush, boundsNode)
                                    End Using
                                End If
                                If Not .HeaderLevel = 255 Then
                                    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ H E A D   N O D E
                                    Using dottedPen As New Pen(Brushes.Silver, 1) With {
                                            .DashStyle = DashStyle.Dot
                                        }
                                        boundsColumn = .Header.Bounds
                                        Dim boundsNodeHead As New Rectangle(boundsColumn.Left, .Bounds_Text.Top, boundsColumn.Width - 1, .Bounds_Text.Height)
                                        e.Graphics.DrawRectangle(dottedPen, boundsNodeHead)
                                        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ F I E L D   N O D E S
                                        For Each fieldNode In .Fields
                                            With fieldNode
                                                boundsColumn = .Header.Bounds
                                                ._Bounds_Text = New Rectangle(boundsColumn.Left, Node.Bounds_Text.Top, boundsColumn.Width - 1, Node.Bounds_Text.Height)
                                                If .Bounds_Text.Right > 0 And .Bounds_Text.Left < Width Then
                                                    Using textBrush As New SolidBrush(.ForeColor)
                                                        Using sf As New StringFormat With {
                                                                .Alignment = fieldNode.Header.GridFormat.HorizontalAlignment,
                                                                .LineAlignment = StringAlignment.Center
                                                                                          }
                                                            e.Graphics.DrawString(If(mouseInTip, .TipText, .Text),
                                                                                      .Font,
                                                                                      textBrush,
                                                                                      .Bounds_Text,
                                                                                      sf)
                                                        End Using
                                                    End Using
                                                    Using Brush As New SolidBrush(If(DragData.DropHighlightNode Is fieldNode, DropHighlightColor, .BackColor))
                                                        'boundsNode.Inflate(-1, -1)
                                                        e.Graphics.FillRectangle(Brush, boundsNode)
                                                    End Using
                                                    e.Graphics.DrawRectangle(dottedPen, .Bounds_Text)
                                                    If Hit?.Node Is fieldNode And .TipText Is Nothing Then
                                                        Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, MouseOverColor))
                                                            e.Graphics.FillRectangle(SemiTransparentBrush, .Bounds_Text)
                                                        End Using
                                                    End If
                                                    If .Selected Then
                                                        Dim selectionBounds As Rectangle = .Bounds_Text
                                                        Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, SelectionColor))
                                                            e.Graphics.FillRectangle(SemiTransparentBrush, .Bounds_Text)
                                                        End Using
                                                        e.Graphics.DrawRectangle(Pens.Black, selectionBounds)
                                                    End If
                                                End If
                                            End With
                                        Next
                                    End Using
                                End If
                                If .Selected Then
                                    Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, SelectionColor))
                                        e.Graphics.FillRectangle(SemiTransparentBrush, boundsNode)
                                    End Using
                                    e.Graphics.DrawRectangle(Pens.Black, boundsNode)
                                End If
                                If .CheckBox Then
                                    If FreezeRoot Then
                                        Dim boundsCheck As Rectangle = .Bounds_Check
                                        boundsCheck.Inflate(3, 3)
                                        e.Graphics.FillEllipse(Brushes.Gainsboro, boundsCheck)
                                    End If
                                    If CheckboxStyle = CheckStyle.Check Then
                                        '/// Check background as White or Gray
                                        Using checkBrush As New SolidBrush(If(.PartialChecked, Color.FromArgb(192, Color.LightGray), Color.White))
                                            e.Graphics.FillRectangle(checkBrush, .Bounds_Check)
                                        End Using
                                        ''/// Draw the checkmark ( only if .Checked or .PartialChecked )
                                        If .Checked Or .PartialChecked Then
                                            Using CheckFont As New Font("Marlett", 10)
                                                TextRenderer.DrawText(e.Graphics, "a".ToString(InvariantCulture), CheckFont, .Bounds_Check, If(.Checked, Color.Blue, Color.DarkGray), TextFormatFlags.NoPadding Or TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
                                            End Using
                                        End If
                                        '/// Draw the surrounding Check square
                                        Using Pen As New Pen(Color.Blue, 1)
                                            e.Graphics.DrawRectangle(Pen, .Bounds_Check)
                                        End Using

                                    ElseIf CheckboxStyle = CheckStyle.Slide Then
                                        e.Graphics.DrawImage(If(.Checked, My.Resources.slideStateOn, If(.PartialChecked, My.Resources.slideStateMixed, My.Resources.slideStateOff)), .Bounds_Check)
                                        If Hit?.Node Is Node And Hit?.Region = MouseRegion.CheckBox Then
                                            Using checkBrush As New SolidBrush(Color.FromArgb(96, MouseOverColor))
                                                e.Graphics.FillRectangle(checkBrush, .Bounds_Check)
                                            End Using
                                        End If

                                    End If
                                End If
                                Dim objectBounds As Rectangle = .Bounds_ShowHide
                                Dim objectCenter As Integer = CInt(objectBounds.Height / 2)

                                '/// Vertical line between this node and child nodes
                                If .HasChildren And .Expanded Then
                                    Dim childLast_bounds As Rectangle = Node_LeftMostBounds(.Children.Last)
                                    Dim verticalNodeLineTop_xy As New Point(objectBounds.Left + objectCenter, objectBounds.Bottom)
                                    Dim verticalNodeLineBottom_xy As New Point(objectBounds.Left + objectCenter, {childLast_bounds.Top + objectCenter, ClientRectangle.Height}.Min)
                                    If FreezeRoot Then
                                        e.Graphics.FillRectangle(Brushes.Gainsboro,
                                                                 New Rectangle(verticalNodeLineTop_xy.X - 4,
                                                                               verticalNodeLineTop_xy.Y,
                                                                               8,
                                                                               verticalNodeLineBottom_xy.Y - verticalNodeLineTop_xy.Y)
                                                                               )
                                    End If
                                    e.Graphics.DrawLine(linePen, verticalNodeLineTop_xy, verticalNodeLineBottom_xy)
                                End If

                                Dim horizontalNodeLine_x As Integer = {
                                                .Bounds_Check.Left,
                                                .Bounds_Favorite.Left,
                                                .Bounds_ShowHide.Left
                                                                              }.Min
                                Dim VerticalCenter As Integer = objectBounds.Top + objectCenter

                                If IsNothing(.Parent) Then
                                    If drawRootLines Then e.Graphics.DrawLine(linePen, New Point(verticalRootLine_x, VerticalCenter), New Point(horizontalNodeLine_x, VerticalCenter))
                                Else
                                    REM /// HORIZONTAL LINES LEFT OF EXPAND/COLLAPSE
                                    Dim NodeHorizontalLeftPoint As New Point(.Parent.Bounds_ShowHide.Left + objectCenter + 1, VerticalCenter)
                                    Dim NodeHorizontalRightPoint As New Point(horizontalNodeLine_x, VerticalCenter)
                                    If FreezeRoot Then
                                        e.Graphics.FillRectangle(Brushes.Gainsboro,
                                                                 New Rectangle(NodeHorizontalLeftPoint.X,
                                                                               NodeHorizontalLeftPoint.Y - 4,
                                                                               NodeHorizontalRightPoint.X - NodeHorizontalLeftPoint.X,
                                                                               8)
                                                                               )
                                    End If
                                    e.Graphics.DrawLine(linePen, NodeHorizontalLeftPoint, NodeHorizontalRightPoint)
                                End If

                            End With
                        Next
                        '/// Vertical root line between top ( first ) node and bottom ( last ) node ... but don't draw if the top IS the bottom too ( 1 node only )
                        If drawRootLines Then
                            Dim lastNode As Node = Ancestors.Last
                            Dim LineTop As Integer = 2 + firstRootNode.Bounds_ShowHide.Top + CInt(firstRootNode.Bounds_ShowHide.Height / 2)
                            Dim TopPoint As New Point(verticalRootLine_x, {0, LineTop}.Max)
                            Dim LineBottom As Integer = lastNode.Bounds_Text.Top + Convert.ToInt32(lastNode.Height / 2)
                            Dim BottomPoint As New Point(verticalRootLine_x, {LineBottom, Height}.Min)
                            e.Graphics.DrawLine(linePen, TopPoint, BottomPoint)
                        End If
                    End Using
                Else
                    HScroll.Hide()
                    VScroll.Hide()
                End If
                ColumnHeaders.Draw.ForEach(Sub(header)
                                               With header
                                                   Dim drawStyle As CellStyle = If(header Is Hit?.Column, .MouseStyle, .Style)
                                                   If drawStyle.Theme = Theme.None Then
                                                       If drawStyle.BackImage Is Nothing Then
                                                           Using LinearBrush As New LinearGradientBrush(.Bounds, drawStyle.BackColor, drawStyle.ShadeColor, LinearGradientMode.Vertical)
                                                               e.Graphics.FillRectangle(LinearBrush, .Bounds)
                                                           End Using
                                                       Else
                                                           e.Graphics.DrawImage(drawStyle.BackImage, .Bounds)
                                                       End If
                                                   Else
                                                       Dim currentTheme = If(header Is Hit?.Column, .MouseStyle.Theme, drawStyle.Theme)
                                                       e.Graphics.DrawImage(GlossyImages(currentTheme), .Bounds)
                                                   End If
                                                   If .Image IsNot Nothing Then
                                                       Dim boundsImage As Rectangle = .Bounds_Image
                                                       boundsImage.Inflate(-3, 0)
                                                       boundsImage.Offset(2, 0)
                                                       e.Graphics.DrawImage(.Image, boundsImage)
                                                   End If
                                                   If .SortIcon IsNot Nothing Then
                                                       Dim boundsSort As Rectangle = .Bounds_Sort
                                                       boundsSort.Inflate(-4, 0)
                                                       boundsSort.Offset(-4, 0)
                                                       e.Graphics.DrawImage(.SortIcon, boundsSort)
                                                   End If
                                                   Using headerTextBrush As New SolidBrush(drawStyle.ForeColor)
                                                       e.Graphics.DrawString(
                                                      .Text,
                                                      drawStyle.Font,
                                                      headerTextBrush,
                                                      .Bounds_Text,
                                                      drawStyle.Alignment
                                                  )
                                                   End Using
                                                   ControlPaint.DrawBorder3D(e.Graphics, .Bounds, Border3DStyle.Raised)
                                               End With
                                           End Sub)
            End With
            ControlPaint.DrawBorder3D(e.Graphics, ClientRectangle, Border3DStyle.Sunken)
        End If

    End Sub
    Private Function Node_LeftMostBounds(branch As Node) As Rectangle

        If branch Is Nothing Then
            Return Nothing
        Else
            Dim nodeBounds As New List(Of Rectangle) From {branch.Bounds_Text, branch.Bounds_Check, branch.Bounds_ShowHide, branch.Bounds_Favorite, branch.Bounds_Image}
            nodeBounds.Sort(Function(a, b)
                                Dim Level1 = a.Left.CompareTo(b.Left)
                                If Level1 <> 0 Then
                                    Return Level1
                                Else
                                    Dim Level2 = a.Width.CompareTo(b.Width)
                                    Return Level2
                                End If
                            End Function)
            Return nodeBounds.First
        End If

    End Function
    Public ReadOnly Property UnRestrictedSize As Size
        Get
            Dim headersWidth As Integer = 0
            For Each headerLevel As ColumnHeadCollection In ColumnHeaders
                For Each header As ColumnHead In headerLevel
                    If header.Visible Then headersWidth += header.Bounds.Width
                Next
            Next
            'Dim HeadersWidth As Integer = ColumnHeaders.Sum(Function(h) h.Sum(Function(header)
            '                                                                      Return header.Bounds.Width
            '                                                                  End Function))
            Return New Size({RollingWidth + Offset.X, HeadersWidth}.Max, RollingHeight + Offset.Y)
        End Get
    End Property
    Friend Sub RequiresRepaint()

        IgnoreSizeChanged = True
        REM /// RESET INDEX / HEIGHT
        NodeIndex = 0
        VisibleIndex = 0
        RollingWidth = Offset.X
        RollingHeight = Offset.Y + ColumnHeaders.Sum(Function(h) h.Height)

        REM /// ITERATE ALL NODES CHANGING BOUNDS
        RefreshNodesBounds_Lines(Ancestors)
        Columns_SetBounds() 'Must be done here to ensure initial view of H/V Scrolls is accurate

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
        Dim sizes As New List(Of Size) From {maxScreenSize, maxParentSize, maxUserSize}
        Dim nonZeroWidths As New List(Of Integer)(From s In sizes Where s.Width > 0 Select s.Width)
        Dim nonZeroHeights As New List(Of Integer)(From s In sizes Where s.Height > 0 Select s.Height)
        Dim maxWidth As Integer = nonZeroWidths.Min
        Dim maxHeight As Integer = nonZeroHeights.Min

        Dim hScrollVisible As Boolean = unboundedSize.Width > maxWidth
        Dim vscrollVisible As Boolean = unboundedSize.Height > maxHeight

        If AutoSize Then 'Can resize
            Dim proposedWidth = {unboundedSize.Width, maxWidth}.Min
            Dim proposedHeight = {unboundedSize.Height, maxHeight}.Min
            If hScrollVisible Then proposedHeight = {proposedHeight + HScrollHeight, maxHeight}.Min
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
            .Maximum = {0, unboundedSize.Width + If(vscrollVisible, VScroll.Width, 0)}.Max
            .Visible = hScrollVisible
            .Left = 0
            .Width = maxWidth
            .Top = Height - .Height
            .LargeChange = maxWidth
            If .Visible Then
                .Show()
            Else
                .Value = 0
                .Hide()
            End If
        End With
        With VScroll
            Dim maxVscrollHeight As Integer = maxHeight - If(hScrollVisible, HScroll.Height, 0)
            .Minimum = 0
            .Maximum = {0, unboundedSize.Height + If(hScrollVisible, HScroll.Height, 0)}.Max
            .Visible = vscrollVisible
            .Left = maxWidth - VScrollWidth
            .Height = maxVscrollHeight
            .Top = 0
            .LargeChange = maxVscrollHeight
            If .Visible Then
                .Show()
            Else
                .Value = 0
                .Hide()
            End If
        End With
#End Region
        DrawNodes_Set()

        REM /// FINALLY- REPAINT
        Invalidate()
        IgnoreSizeChanged = False

    End Sub
    Private Sub Columns_SetBounds()

        ColumnHeaders.Draw.Clear()
        ColumnHeaders.ForEach(Sub(headers)
                                  Dim rollingLeft As Integer = -HScroll.Value
                                  For Each header As ColumnHead In headers.Where(Function(h) h.Visible)
                                      With header
                                          Dim freezeColumn As Boolean = FreezeRoot And .Index = 0
                                          Dim offsetFreeze As Integer = If(freezeColumn, HScroll.Value, 0)
                                          .Bounds_Image = If(.Image Is Nothing,
             New Rectangle(rollingLeft + offsetFreeze, 0, 0, headers.Height),
             New Rectangle(rollingLeft + offsetFreeze, CInt((headers.Height - .Image.Height) / 2), 2 + .Image.Width + 1, .Image.Height))
                                          Dim widthProposed As Integer = .Bounds_Image.Width
                                          If .SortIcon Is Nothing Then
                                              widthProposed += { .ContentWidth, .TextSize.Width}.Max
                                              .Bounds_Sort = New Rectangle(rollingLeft + offsetFreeze + widthProposed, 0, 0, 0)
                                          Else
                                              widthProposed += If(.ContentWidth - .TextSize.Width > .SortIcon.Width, .ContentWidth, .TextSize.Width + 2 + .SortIcon.Width + 2)
                                              .Bounds_Sort = New Rectangle(rollingLeft + offsetFreeze + widthProposed - .SortIcon.Width, CInt((headers.Height - .SortIcon.Height) / 2), .SortIcon.Width, .SortIcon.Height)
                                          End If
                                          .Bounds_Text = New Rectangle(.Bounds_Image.Right, 0, .Bounds_Sort.Left - .Bounds_Image.Right, headers.Height)
                                          .Bounds = New Rectangle(rollingLeft + offsetFreeze, 0, widthProposed, headers.Height)
                                          If freezeColumn Or .Bounds.Right > 0 And rollingLeft < Width Then ColumnHeaders.Draw.Add(header)
                                          rollingLeft += .Bounds.Width
                                      End With
                                  Next
                              End Sub)
        ColumnHeaders.Draw.Sort(Function(x, y)
                                    Return y.Index.CompareTo(x.Index)
                                End Function)

    End Sub
    Private Sub RefreshNodesBounds_Lines(Nodes As NodeCollection)

        NodeIndex = 0
        Nodes.ForEach(Sub(node)
                          With node
                              ._Index = NodeIndex
                              ._Visible = .Parent Is Nothing OrElse .Parent.Expanded
                              If .Visible Then
                                  Node_SetBounds(node)
                                  ._VisibleIndex = VisibleIndex
                                  VisibleIndex += 1
                                  If .Bounds_Text.Right > RollingWidth Then RollingWidth = .Bounds_Text.Right
                                  RollingHeight += .Height
                                  If .HasChildren Then RefreshNodesBounds_Lines(.Children)
                              End If
                              NodeIndex += 1
                          End With
                          '/// Test
                          If FavoritesFirst And node.CanFavorite Then node.Children.SortAscending(False) 'Do not let the Sort require repaint as it cycles back here to an infinate loop
                      End Sub)

    End Sub
    Private Sub Node_SetBounds(node As Node)

        Dim y As Integer = RollingHeight - VScroll.Value
        Const HorizontalSpacing As Integer = 3

        With node
#Region " S E T   B O U N D S "
            Dim leftMost As Integer = Offset.X + HorizontalSpacing + If(.Parent Is Nothing, If(RootLines, 6, 0), .Parent.Bounds_ShowHide.Right) - If(FreezeRoot, 0, HScroll.Value)

            If ExpandBeforeText Then
                '■■■■■■■■■■■■■ P r e f e r
#Region " +- Icon precedes Text "
                'Bounds.X cascades from [1] Favorite, [2] Checkbox, [3] Image, [4] ShowHide, [5] Text
                REM FAVORITE
                ._Bounds_Favorite.X = leftMost
                ._Bounds_Favorite.Y = y + CInt((.Height - FavoriteImage.Height) / 2)
                ._Bounds_Favorite.Width = If(.CanFavorite, FavoriteImage.Width, 0)
                ._Bounds_Favorite.Height = If(.CanFavorite, FavoriteImage.Height, .Height)

                REM CHECKBOX
                ._Bounds_Check.X = ._Bounds_Favorite.Right + If(._Bounds_Favorite.Width = 0, 0, HorizontalSpacing)
                If CheckboxStyle = CheckStyle.Slide Then
                    ._Bounds_Check.Width = If(.CheckBox, My.Resources.slideStateOff.Width, 0)
                    ._Bounds_Check.Height = My.Resources.slideStateOff.Height
                    ._Bounds_Check.Y = y + CInt((.Height - ._Bounds_Check.Height) / 2)
                Else
                    ._Bounds_Check.Width = If(.CheckBox, CheckHeight, 0)
                    ._Bounds_Check.Height = CheckHeight
                    ._Bounds_Check.Y = y + CInt((.Height - CheckHeight) / 2)
                End If

                REM IMAGE
                ._Bounds_Image.X = ._Bounds_Check.Right + If(._Bounds_Check.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Image.Height = If(IsNothing(.Image), 0, If(.ImageScaling, .Height, .Image.Height))
                'MAKE IMAGE SQUARE IF SCALING
                ._Bounds_Image.Width = If(IsNothing(.Image), 0, If(.ImageScaling, ._Bounds_Image.Height, .Image.Width))
                ._Bounds_Image.Y = y + CInt((.Height - ._Bounds_Image.Height) / 2)

                REM EXPAND/COLLAPSE
                ._Bounds_ShowHide.X = ._Bounds_Image.Right + HorizontalSpacing '+ If(IsNothing(.Parent), If(RootLines, 6, 0), .Parent._Bounds_Image.Right + HorizontalSpacing)
                ._Bounds_ShowHide.Y = y + CInt((.Height - ExpandHeight) / 2)
                ._Bounds_ShowHide.Width = If(.HasChildren, ExpandHeight, 0)
                ._Bounds_ShowHide.Height = ExpandHeight

                REM TEXT
                ._Bounds_Text.X = ._Bounds_ShowHide.Right + If(._Bounds_ShowHide.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Text.Y = y
                ._Bounds_Text.Height = .Height
                '._Bounds_Text.Width = .TextWidth 'Not necessary- set when the Text changes at the Node level
#End Region
            Else
#Region " +- Icon follows Text "
                REM EXPAND/COLLAPSE
                ._Bounds_ShowHide.X = leftMost
                ._Bounds_ShowHide.Y = y + CInt((.Height - ExpandHeight) / 2)
                ._Bounds_ShowHide.Width = If(.HasChildren, ExpandHeight, 0)
                ._Bounds_ShowHide.Height = ExpandHeight

                REM FAVORITE
                ._Bounds_Favorite.X = ._Bounds_ShowHide.Right + If(._Bounds_ShowHide.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Favorite.Y = y + CInt((.Height - FavoriteImage.Height) / 2)
                ._Bounds_Favorite.Width = If(.CanFavorite, FavoriteImage.Width, 0)
                ._Bounds_Favorite.Height = If(.CanFavorite, FavoriteImage.Height, .Height)

                REM CHECKBOX
                ._Bounds_Check.X = ._Bounds_Favorite.Right + If(._Bounds_Favorite.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Check.Width = If(.CheckBox, CheckHeight, 0)
                ._Bounds_Check.Height = CheckHeight
                ._Bounds_Check.Y = y + CInt((.Height - ._Bounds_Check.Height) / 2)

                REM IMAGE
                ._Bounds_Image.X = ._Bounds_Check.Right + If(._Bounds_Check.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Image.Height = If(IsNothing(.Image), 0, If(.ImageScaling, .Height, .Image.Height))
                'MAKE IMAGE SQUARE IF SCALING
                ._Bounds_Image.Width = If(IsNothing(.Image), 0, If(.ImageScaling, ._Bounds_Image.Height, .Image.Width))
                ._Bounds_Image.Y = y + CInt((.Height - ._Bounds_Image.Height) / 2)

                REM TEXT
                ._Bounds_Text.X = ._Bounds_Image.Right + If(._Bounds_Image.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Text.Y = y
                ._Bounds_Text.Height = .Height
#End Region
            End If

            If .Header IsNot Nothing Then
                .Header.ContentWidth = If(.ColumnIndex = 0, .Bounds_Text.Right, .Bounds_Text.Width)
                .Fields.ForEach(Sub(field)
                                    With field
                                        .Header.ContentWidth = .Bounds_Text.Width
                                        ._Bounds_Text.Y = node.Bounds_Text.Y
                                        ._Bounds_Text.Height = node.Bounds_Text.Height
                                    End With
                                End Sub)
            End If
#End Region
        End With

    End Sub

    '//////////////////////////////////// SCROLLING, SCROLLING, SCROLLING
    Private Sub OnScrolled(sender As Object, e As ScrollEventArgs) Handles HScroll.Scroll, VScroll.Scroll

        If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
            HScrollLeftRight(e.OldValue - e.NewValue)

        ElseIf e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            VScrollUpDown(e.OldValue - e.NewValue)

        End If
        DrawNodes_Set()

    End Sub
    Private Sub HScrollLeftRight(X_Change As Integer)

        If X_Change = 0 Then Exit Sub
        Ancestors.Visible.ForEach(Sub(node)
                                      If Not (FreezeRoot And {0, 1}.Contains(node.ColumnIndex)) Then
                                          node._Bounds_ShowHide.X += X_Change
                                          node._Bounds_Favorite.X += X_Change
                                          node._Bounds_Check.X += X_Change
                                          node._Bounds_Image.X += X_Change
                                          node._Bounds_Text.X += X_Change
                                      End If
                                  End Sub)
        Invalidate()

    End Sub
    Private Sub VScrollUpDown(Y_Change As Integer)

        If Y_Change = 0 Then Exit Sub
        Ancestors.Visible.ForEach(Sub(node)
                                      node._Bounds_ShowHide.Y += Y_Change
                                      node._Bounds_Favorite.Y += Y_Change
                                      node._Bounds_Check.Y += Y_Change
                                      node._Bounds_Image.Y += Y_Change
                                      node._Bounds_Text.Y += Y_Change
                                  End Sub)
        Invalidate()

    End Sub
    '//////////////////////////////////// SCROLLING, SCROLLING, SCROLLING
    Protected Overrides Sub OnParentChanged(e As EventArgs)
        RequiresRepaint()
        MyBase.OnParentChanged(e)
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

        If Not IgnoreSizeChanged Then RequiresRepaint()
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
            RequiresRepaint()
        End With

    End Sub
#End Region

    Public Function HitTest(Location As Point) As HitRegion

        Dim hotSpot As New HitRegion
        With hotSpot
            .Region = MouseRegion.None
            Dim columnsDraw As New List(Of ColumnHead)(ColumnHeaders.Draw)
            columnsDraw.ForEach(Sub(column)
                                    If column.Bounds.Contains(Location) Then
                                        .Region = MouseRegion.Column
                                        .Node = Nothing
                                        .Column = column
                                        Exit Sub
                                    End If
                                End Sub)
            If .Region = MouseRegion.None Then
                Dim ancestorsDraw As New List(Of Node)(Ancestors.Draw)
                ancestorsDraw.ForEach(Sub(node)
                                          If node.Bounds_Text.Contains(Location) Then
                                              .Region = MouseRegion.Node
                                              .Node = node

                                          ElseIf node.Bounds_Check.Contains(Location) Then
                                              .Region = MouseRegion.CheckBox
                                              .Node = node

                                          ElseIf node.Bounds_ShowHide.Contains(Location) Then
                                              .Region = MouseRegion.Expander
                                              .Node = node

                                          ElseIf node.Bounds_Favorite.Contains(Location) Then
                                              .Region = MouseRegion.Favorite
                                              .Node = node

                                          ElseIf node.Bounds_Image.Contains(Location) Then
                                              .Region = MouseRegion.Image
                                              .Node = node

                                          Else
                                              Dim boundsNode As New Rectangle(node.Header.Bounds.X, node.Bounds_Text.Y, node.Header.Bounds.Width, node.Bounds_Text.Height)
                                              If boundsNode.Contains(Location) Then
                                                  .Region = MouseRegion.Node
                                                  .Node = node
                                              Else
                                                  node.Fields.ForEach(Sub(field)
                                                                          If field.Bounds_Text.Contains(Location) Then
                                                                              .Region = MouseRegion.Field
                                                                              .Node = field
                                                                          End If
                                                                      End Sub)
                                              End If
                                          End If
                                          If Not .Region = MouseRegion.None Then Exit Sub
                                      End Sub)
            End If
        End With
        Return hotSpot

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
    Private Sub ExpandCollapseNodes(Nodes As NodeCollection, State As Boolean)

        For Each Node As Node In Nodes
            If State Then
                Node.Expand()
            Else
                Node.Collapse()
            End If
            If Node.HasChildren Then ExpandCollapseNodes(Node.Children, State)
        Next

    End Sub
    Private Sub DrawNodes_Set()

        Ancestors.Client.Clear()
        For Each Node In Ancestors.Visible
            If ClientRectangle.Contains(New Point(1, Node.Bounds_Text.Top)) Or ClientRectangle.Contains(New Point(1, Node.Bounds_Text.Bottom)) Then
                Ancestors.Client.Add(Node)
            End If
        Next

        Ancestors.Draw.Clear()
        Ancestors.Draw.AddRange(Ancestors.Client)
        Dim NodesDraw = (From N In Ancestors.Visible Group N By N.Parent Into Group Select Parent, Children =
                ((From C In Group Where C.Bounds_Text.Bottom < 0 Order By C.Bounds_Text.Y Descending).Take(1)).Union _
                ((From C In Group Where C.Bounds_Text.Top > ClientRectangle.Height Order By C.Bounds_Text.Y Ascending).Take(1)))
        For Each ParentGroup In NodesDraw
            Ancestors.Draw.AddRange(ParentGroup.Children)
        Next

    End Sub
End Class

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ NODE COLLECTION
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
        _Tree = TreeViewer
    End Sub
    Private _Tree As TreeViewer
    Public ReadOnly Property Tree As TreeViewer
        Get
            Return _Tree
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
            Dim topNode As Node
            For Each topNode In Me
                Queue.Enqueue(topNode)
            Next
            While Queue.Any
                topNode = Queue.Dequeue
                Nodes.Add(topNode)
                For Each childNode In topNode.Children
                    Queue.Enqueue(childNode)
                Next
            End While
            Nodes.Sort(Function(x, y) x.VisibleIndex.CompareTo(y.VisibleIndex))
            Return Nodes
        End Get
    End Property
    Public ReadOnly Property Levels(level As Byte) As List(Of Node)
        Get
            Return All.Where(Function(n) n.Level = level).ToList
        End Get
    End Property
    Public ReadOnly Property Selected As List(Of Node)
        Get
            Return All.Where(Function(n) n.Selected).ToList
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
    Public ReadOnly Property Table As DataTable
        Get
            Dim rows As New List(Of DataRow)(From n In Me Select n.Row)
            Return rows.CopyToDataTable
        End Get
    End Property
    Public ReadOnly Property Client As New List(Of Node)
    Public ReadOnly Property Draw As New List(Of Node)
    Public ReadOnly Property DataType As Type
        Get
            Dim Types As New List(Of Type)([Select](Function(n) n.DataType))
            Return GetDataType(Types)
        End Get
    End Property

    Public Overloads Function Contains(Name As String) As Boolean

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
    Public Overloads Function Add(AddNode As Node) As Node

        If AddNode IsNot Nothing Then
            With AddNode
                ._Index = Count
                If Parent Is Nothing Then   ' *** ROOT NODE
                    'mTreeViewer was set when TreeViewer was created with New NodeCollection
                    ._Tree = Tree
                    ._Visible = True
                    ._Level = 0

                Else                        ' *** CHILD NODE
                    REM /// Get TreeViewer value from Parent. A Node Collection shares some Node properties
                    _Tree = Parent.Tree
                    ._Tree = Parent.Tree
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
                If Tree IsNot Nothing Then
                    .CanFavorite = Tree.CanFavorite
                    .Font = Tree.Font
                    Tree.NodeTimer_Start(AddNode)
                End If
            End With
            MyBase.Add(AddNode)
        End If
        Return AddNode

    End Function
    Public Overloads Function AddRange(Nodes As List(Of Node)) As List(Of Node)

        If Nodes IsNot Nothing Then
            For Each Node As Node In Nodes
                Add(Node)
            Next
            Tree?.RequiresRepaint()
        End If
        Return Nodes

    End Function
    Public Overloads Function AddRange(Nodes As Node()) As Node()

        If Nodes IsNot Nothing Then
            For Each Node As Node In Nodes
                Add(Node)
            Next
            Tree?.RequiresRepaint()
        End If
        Return Nodes

    End Function
    Public Overloads Function AddRange(Nodes As String()) As Node()

        Dim NewNodes As New List(Of Node)
        If Nodes IsNot Nothing Then
            For Each NewNode As String In Nodes
                Add(NewNode)
                NewNodes.Add(New Node With {.Text = NewNode})
            Next
            Tree?.RequiresRepaint()
        End If
        Return NewNodes.ToArray

    End Function
    Public Overloads Function Clear(Nodes As NodeCollection) As NodeCollection

        Clear()
        Tree?.RequiresRepaint()
        Return Nodes

    End Function
    Public Overloads Function Insert(Index As Integer, InsertNode As Node) As Node

        If InsertNode IsNot Nothing Then
            If IsNothing(Tree) Then
                MyBase.Insert(Index, InsertNode)
            Else
                If Tree.Ancestors.All.Contains(InsertNode) Then
                    'Throw New ArgumentException("This node already exists in the Treeviewer. Try Removing the Node")

                Else
                    MyBase.Insert(Index, InsertNode)
                    InsertNode._Tree = Tree
                    InsertNode._Parent = Parent
                    Tree.NodeTimer_Start(InsertNode)
                End If
            End If
        End If
        Return InsertNode

    End Function
    Public Overloads Function Remove(RemoveNode As Node) As Node

        If RemoveNode IsNot Nothing Then
            MyBase.Remove(RemoveNode)
            Tree.NodeTimer_Start(RemoveNode)
        End If
        Return RemoveNode

    End Function
    Public Shadows Function Item(Name As String) As Node

        Dim Nodes As New List(Of Node)((From N In Me Where N.Name = Name).ToArray)
        Return If(Nodes.Any, Nodes.First, Nothing)

    End Function
    Public Shadows Function ItemByTag(TagObject As Object) As Node

        Dim Nodes As New List(Of Node)((From N In All Where N.Tag Is TagObject).ToArray)
        Return If(Nodes.Any, Nodes.First, Nothing)

    End Function
    Friend Sub SortAscending(Optional repaint As Boolean = True)

        If Tree?.FavoritesFirst Then
            Select Case DataType
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
            Select Case DataType
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
        If repaint Then
            Tree?.RequiresRepaint()
        Else
            Tree?.Invalidate()
        End If

    End Sub
    Friend Sub SortDescending(Optional repaint As Boolean = True)

        If Tree?.FavoritesFirst Then
            Select Case DataType
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
            Select Case DataType
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
        If repaint Then Tree?.RequiresRepaint()

    End Sub
End Class

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ NODE
Public Class Node
    Implements IDisposable
    Public Sub New()
        Children = New NodeCollection(_Tree) With {._Parent = Me}
        Fields = New NodeCollection(_Tree) With {._Parent = Me}
    End Sub
    Public Sub New(nodeText As String)
        Text = nodeText
        Children = New NodeCollection(_Tree) With {._Parent = Me}
        Fields = New NodeCollection(_Tree) With {._Parent = Me}
    End Sub
    Public Enum SeparatorPosition
        None
        Above
        Below
    End Enum

    Friend _Tree As TreeViewer
    Public ReadOnly Property Tree As TreeViewer
        Get
            Return _Tree
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
    Public ReadOnly Property HeaderLevel As Byte
        Get
            Return HeaderLevel_
        End Get
    End Property
    Friend ColumnIndex_ As Byte = 255
    Public ReadOnly Property ColumnIndex As Byte
        Get
            Return ColumnIndex_
        End Get
    End Property
    Friend IsHeader_ As Boolean = False
    Public ReadOnly Property IsHeader As Boolean
        Get
            Return IsHeader_
        End Get
    End Property
    Friend IsFieldParent_ As Boolean = False
    Public ReadOnly Property IsFieldParent As Boolean
        Get
            Return IsFieldParent_
        End Get
    End Property
    Friend IsField_ As Boolean = False
    Public ReadOnly Property IsField As Boolean
        Get
            Return IsField_
        End Get
    End Property
    Friend Row_ As DataRow
    Public ReadOnly Property Row As DataRow
        Get
            Return Row_
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
            If Tree Is Nothing Then
                Return _CheckBox
            Else
                If Tree.CheckBoxes = TreeViewer.CheckState.All Then
                    _CheckBox = True

                ElseIf Tree.CheckBoxes = TreeViewer.CheckState.None Then
                    _CheckBox = False

                ElseIf Tree.CheckBoxes = TreeViewer.CheckState.Mixed Then
                    If Descendants.Any Then _CheckBox = Descendants.Where(Function(d) d.CheckBox).Count > 0

                End If
                Return _CheckBox
            End If
        End Get
        Set(value As Boolean)
            If value <> _CheckBox Then
                _CheckBox = value
                If Tree IsNot Nothing And value And Tree.CheckBoxes = TreeViewer.CheckState.None Then Tree.CheckBoxes = TreeViewer.CheckState.Mixed
                RequiresRepaint()
            End If
        End Set
    End Property
    Private _Checked As Boolean = False
    Public Property Checked As Boolean
        Get
            Dim checkboxes = Descendants.Where(Function(d) d.CheckBox)
            If checkboxes.Any Then
                Dim countChecked As Integer = checkboxes.Where(Function(c) c.Checked).Count
                Dim allChecked = checkboxes.Count = countChecked
                _Checked = allChecked
            End If
            Return _Checked
        End Get
        Set(value As Boolean)
            If value <> _Checked Then
                _Checked = value
                'If value Then CheckBox = True
                For Each Child In Descendants
                    Child.Checked = value
                Next
                RequiresRepaint()
            End If
        End Set
    End Property
    Public ReadOnly Property PartialChecked As Boolean
        Get
            Dim checks As New List(Of Node)(Descendants.Where(Function(d) d.CheckBox))
            Dim checkCount As Integer = checks.Where(Function(c) c.Checked).Count
            Return checkCount > 0 And checkCount < checks.Count
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
    Friend Value_ As Object
    Friend Property Value As Object
        Get
            Return Value_
        End Get
        Set(value As Object)
            Value_ = If(IsDBNull(value), Nothing, value)
            Text = Format(Value_, Header.GridFormat.FormatString)
        End Set
    End Property
    Friend ReadOnly Property TextWidth As Integer
    Private Sub TextWidth_Set()

        _TextWidth = MeasureText(Text, Font).Width   'TextRenderer.MeasureText(If(Text, String.Empty), Font).Width
        _Bounds_Text.Width = TextWidth
        RequiresRepaint()

    End Sub 'Only changes to Font & Text affect the width
    Private _Text As String
    Public Property Text As String
        Get
            Return _Text & If(ShowNodeIndex, " [" & Index.ToString(InvariantCulture) & "]", String.Empty)
        End Get
        Set(value As String)
            If Not _Text = value Then
                _Text = Replace(value, "&", "&&")
                TextWidth_Set()
            End If
        End Set
    End Property
    Public Property TipText As String
    Public ReadOnly Property Options As New List(Of Object)
    Public ReadOnly Property ChildOptions As New List(Of Object)
    Private Font_ As Font
    Public Property Font As Font
        Get
            If Font_ Is Nothing Then
                If Tree Is Nothing Then
                    Return New Font("Calibri", 9)
                Else
                    Font_ = Tree.Font
                    Return Font_
                End If
            Else
                Return Font_
            End If
        End Get
        Set(value As Font)
            If value IsNot Nothing Then
                If Font_ Is Nothing OrElse value.Name <> Font_.Name And value.Style <> Font_.Style And value.Size <> Font_.Size Then
                    Font_ = value
                    TextWidth_Set()
                End If
            End If
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
    Friend _Bounds_ShowHide As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_ShowHide As Rectangle
        Get
            Return _Bounds_ShowHide
        End Get
    End Property
    Friend _Bounds_Favorite As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_Favorite As Rectangle
        Get
            Return _Bounds_Favorite
        End Get
    End Property
    Friend _Bounds_Check As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_Check As Rectangle
        Get
            Return _Bounds_Check
        End Get
    End Property
    Friend _Bounds_Image As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_Image As Rectangle
        Get
            Return _Bounds_Image
        End Get
    End Property
    Friend _Bounds_Text As New Rectangle
    Public ReadOnly Property Bounds_Text As Rectangle
        Get
            Return _Bounds_Text
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
    Public ReadOnly Property TextPath As String
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
    Public ReadOnly Property Siblings As List(Of Node)
        Get
            Dim BrothersSistersAndMe As New List(Of Node)(If(IsRoot, Tree?.Ancestors, Parent.Children))
            BrothersSistersAndMe.Sort(Function(x, y) x.Index.CompareTo(y.Index))
            Return BrothersSistersAndMe
        End Get
    End Property
    Public ReadOnly Property FirstSibling As Node
        Get
            Return Siblings.First 'Always Siblings since Me is included
        End Get
    End Property
    Public ReadOnly Property LastSibling As Node
        Get
            Return Siblings.Last 'Always Siblings since Me is included
        End Get
    End Property
    Public ReadOnly Property NextSibling As Node
        Get
            If Me Is LastSibling Then
                Return Nothing
            Else
                Return Siblings(Index + 1)
            End If
        End Get
    End Property
    Friend Header_ As ColumnHead
    Public ReadOnly Property Header As ColumnHead
        Get
            Return Header_
        End Get
    End Property
    Friend DataType_ As Type = Nothing
    Public ReadOnly Property DataType As Type
        Get
            If HeaderLevel = 255 Or ColumnIndex = 255 Then
                'Not created using DataTable
                DataType_ = If(DataType_, GetDataType(SortValue))

            Else
                'Created using DataTable - HeaderLevel and ColumnIndex both set to a value
                DataType_ = Header.DataType

            End If
            Return DataType_
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
                If Tree?.FavoritesFirst Then Parent?.Children.SortAscending()
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

        If Not IsNothing(Tree) Then
            If _Clicked Then
                _Clicked = False
                Tree.RequiresRepaint()
            Else
                Tree.NodeTimer_Start(Me)
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

        Dim typeString As String = If(DataType Is Nothing, String.Empty, ", Type=" & Replace(DataType.ToString, "System.", String.Empty))
        Dim textString As String = If(Text, "( empty )")
        Dim childCount As String = If(Children.Any, ", Children ( " & Children.Count & " )", String.Empty)

        If Name Is Nothing Then
            Return Join({textString, typeString, childCount}, String.Empty)
        Else
            Return Join({"Key=", Name, ", Text=", textString, typeString, childCount}, String.Empty)
        End If

    End Function

#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                Font_.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose( disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose( disposing As Boolean) above.
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

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ HEADER LEVELS
Public Class HeaderLevels
    Inherits List(Of ColumnHeadCollection)
    Friend Event Changed()
    Public ReadOnly Property Parent As TreeViewer
    Friend Sub New(tree As TreeViewer)
        Parent = tree
    End Sub
    Public ReadOnly Property Draw As New List(Of ColumnHead)
    Private WithEvents Styles_ As New CellStyle With {
        .BackColor = Color.Transparent,
        .ShadeColor = Color.White,
        .ForeColor = Color.Black,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.Black
    }
    Public Property Styles As CellStyle
        Get
            Return Styles_
        End Get
        Set(value As CellStyle)
            If value <> Styles_ Then
                Styles_ = value
                ForEach(Sub(headers)
                            headers.Styles = value
                        End Sub)
                Parent?.RequiresRepaint()
            End If
        End Set
    End Property
    Private WithEvents MouseStyles_ As New CellStyle With {
        .BackColor = Color.Transparent,
        .ShadeColor = Color.White,
        .ForeColor = Color.Black,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.Yellow
    }
    Public Property MouseStyles As CellStyle
        Get
            Return MouseStyles_
        End Get
        Set(value As CellStyle)
            If value <> MouseStyles_ Then
                MouseStyles_ = value
                ForEach(Sub(headers)
                            headers.MouseStyles = value
                        End Sub)
                Parent?.RequiresRepaint()
            End If
        End Set
    End Property
    Private Sub MouseStyle_Changed(sender As Object, e As StyleEventArgs) Handles Styles_.PropertyChanged, MouseStyles_.PropertyChanged

        Dim senderStyle As CellStyle = DirectCast(sender, CellStyle)
        If sender Is Styles Then
            ForEach(Sub(heads)
                        With heads.Styles
                            If e.ChangedProperty = CellStyle.Properties.Alignment Then .Alignment = senderStyle.Alignment
                            If e.ChangedProperty = CellStyle.Properties.BackColor Then .BackColor = senderStyle.BackColor
                            If e.ChangedProperty = CellStyle.Properties.Font Then .Font = senderStyle.Font
                            If e.ChangedProperty = CellStyle.Properties.ForeColor Then .ForeColor = senderStyle.ForeColor
                            If e.ChangedProperty = CellStyle.Properties.Height Then .Height = senderStyle.Height
                            If e.ChangedProperty = CellStyle.Properties.ImageScaling Then .ImageScaling = senderStyle.ImageScaling
                            If e.ChangedProperty = CellStyle.Properties.Padding Then .Padding = senderStyle.Padding
                            If e.ChangedProperty = CellStyle.Properties.ShadeColor Then .ShadeColor = senderStyle.ShadeColor
                            If e.ChangedProperty = CellStyle.Properties.Theme Then .Theme = senderStyle.Theme
                            If e.ChangedProperty = CellStyle.Properties.Image Then .BackImage = senderStyle.BackImage
                        End With
                    End Sub)
        Else
            ForEach(Sub(heads)
                        With heads.MouseStyles
                            If e.ChangedProperty = CellStyle.Properties.Alignment Then .Alignment = senderStyle.Alignment
                            If e.ChangedProperty = CellStyle.Properties.BackColor Then .BackColor = senderStyle.BackColor
                            If e.ChangedProperty = CellStyle.Properties.Font Then .Font = senderStyle.Font
                            If e.ChangedProperty = CellStyle.Properties.ForeColor Then .ForeColor = senderStyle.ForeColor
                            If e.ChangedProperty = CellStyle.Properties.Height Then .Height = senderStyle.Height
                            If e.ChangedProperty = CellStyle.Properties.ImageScaling Then .ImageScaling = senderStyle.ImageScaling
                            If e.ChangedProperty = CellStyle.Properties.Padding Then .Padding = senderStyle.Padding
                            If e.ChangedProperty = CellStyle.Properties.ShadeColor Then .ShadeColor = senderStyle.ShadeColor
                            If e.ChangedProperty = CellStyle.Properties.Theme Then .Theme = senderStyle.Theme
                            If e.ChangedProperty = CellStyle.Properties.Image Then .BackImage = senderStyle.BackImage
                        End With
                    End Sub)
        End If
        Parent?.RequiresRepaint()

    End Sub
    Public Shadows Sub Clear()

        ForEach(Sub(headers)
                    RemoveHandler headers.Changed, AddressOf OnChanged
                End Sub)
        MyBase.Clear()
        RaiseEvent Changed()

    End Sub
    Public Shadows Function Add(addHeaders As ColumnHeadCollection) As ColumnHeadCollection

        If addHeaders IsNot Nothing Then
            With addHeaders
                .Parent_ = Parent
                .Styles = Styles
                .MouseStyles = MouseStyles
                .ForEach(Sub(header)
                             header.Font = Parent.Font
                         End Sub)
                AddHandler .Changed, AddressOf OnChanged
            End With
            MyBase.Add(addHeaders)
            RaiseEvent Changed()
        End If
        Return addHeaders

    End Function
    Public Shadows Function Remove(dropHeaders As ColumnHeadCollection) As ColumnHeadCollection

        If dropHeaders IsNot Nothing Then
            With dropHeaders
                .Parent_ = Nothing
                .Styles.Dispose()
                .MouseStyles.Dispose()
                RemoveHandler .Changed, AddressOf OnChanged
            End With
            MyBase.Remove(dropHeaders)
            RaiseEvent Changed()
        End If
        Return dropHeaders

    End Function
    Private Sub OnChanged()
        RaiseEvent Changed()
    End Sub
End Class

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ HEADER COLLECTION
Public Class ColumnHeadCollection
    Inherits List(Of ColumnHead)
    Implements IDisposable

    Friend Event Changed()
    Friend Parent_ As TreeViewer
    Public ReadOnly Property Parent As TreeViewer
        Get
            Return Parent_
        End Get
    End Property

    Public WithEvents DataTypes As New SpecialDictionary(Of String, Type)
    Public Sub New()
    End Sub
    Public Sub New(columns As String())

        If columns IsNot Nothing Then
            For Each column As String In columns
                Add(column)
                DataTypes.Add(column, GetType(String)) 'Default
            Next
            AddHandler DataTypes.PropertyChanged, AddressOf DataType_Set
        End If

    End Sub
    Private Sub DataType_Set(sender As Object, e As DictionaryEventArgs)

        Dim columnName As String = e.Key.ToString
        Dim lastType As Type = CType(e.LastValue, Type)
        Dim newType As Type = CType(e.Value, Type)
        Dim header As ColumnHead = Item(columnName)
        header.GridFormat = New AlignFormat(newType)

    End Sub

    Private WithEvents Styles_ As New CellStyle With {
        .BackColor = Color.Transparent,
        .ShadeColor = Color.White,
        .ForeColor = Color.Black,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.Black
    }
    Public Property Styles As CellStyle
        Get
            Return Styles_
        End Get
        Set(value As CellStyle)
            If value <> Styles_ Then
                Styles_ = value
                ForEach(Sub(header)
                            header.Style_ = value
                        End Sub)
                Parent?.RequiresRepaint()
            End If
        End Set
    End Property

    Private WithEvents MouseStyles_ As New CellStyle With {
        .BackColor = Color.Transparent,
        .ShadeColor = Color.White,
        .ForeColor = Color.Black,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.Yellow
    }
    Public Property MouseStyles As CellStyle
        Get
            Return MouseStyles_
        End Get
        Set(value As CellStyle)
            If value <> MouseStyles_ Then
                MouseStyles_ = value
                ForEach(Sub(header)
                            header.MouseStyle_ = value
                        End Sub)
                Parent?.RequiresRepaint()
            End If
        End Set
    End Property
    Private Sub Style_Changed(sender As Object, e As StyleEventArgs) Handles Styles_.PropertyChanged, MouseStyles_.PropertyChanged

        Dim senderStyle As CellStyle = DirectCast(sender, CellStyle)
        If sender Is Styles Then
            ForEach(Sub(header)
                        With header.Style
                            If e.ChangedProperty = CellStyle.Properties.Alignment Then .Alignment = senderStyle.Alignment
                            If e.ChangedProperty = CellStyle.Properties.BackColor Then .BackColor = senderStyle.BackColor
                            If e.ChangedProperty = CellStyle.Properties.Font Then .Font = senderStyle.Font
                            If e.ChangedProperty = CellStyle.Properties.ForeColor Then .ForeColor = senderStyle.ForeColor
                            If e.ChangedProperty = CellStyle.Properties.Height Then .Height = senderStyle.Height
                            If e.ChangedProperty = CellStyle.Properties.ImageScaling Then .ImageScaling = senderStyle.ImageScaling
                            If e.ChangedProperty = CellStyle.Properties.Padding Then .Padding = senderStyle.Padding
                            If e.ChangedProperty = CellStyle.Properties.ShadeColor Then .ShadeColor = senderStyle.ShadeColor
                            If e.ChangedProperty = CellStyle.Properties.Theme Then .Theme = senderStyle.Theme
                            If e.ChangedProperty = CellStyle.Properties.Image Then .BackImage = senderStyle.BackImage
                        End With
                    End Sub)
        Else
            ForEach(Sub(header)
                        With header.MouseStyle
                            If e.ChangedProperty = CellStyle.Properties.Alignment Then .Alignment = senderStyle.Alignment
                            If e.ChangedProperty = CellStyle.Properties.BackColor Then .BackColor = senderStyle.BackColor
                            If e.ChangedProperty = CellStyle.Properties.Font Then .Font = senderStyle.Font
                            If e.ChangedProperty = CellStyle.Properties.ForeColor Then .ForeColor = senderStyle.ForeColor
                            If e.ChangedProperty = CellStyle.Properties.Height Then .Height = senderStyle.Height
                            If e.ChangedProperty = CellStyle.Properties.ImageScaling Then .ImageScaling = senderStyle.ImageScaling
                            If e.ChangedProperty = CellStyle.Properties.Padding Then .Padding = senderStyle.Padding
                            If e.ChangedProperty = CellStyle.Properties.ShadeColor Then .ShadeColor = senderStyle.ShadeColor
                            If e.ChangedProperty = CellStyle.Properties.Theme Then .Theme = senderStyle.Theme
                            If e.ChangedProperty = CellStyle.Properties.Image Then .BackImage = senderStyle.BackImage
                        End With
                    End Sub)
        End If
        Parent?.RequiresRepaint()

    End Sub
    Private Height_ As Integer = 32
    Public Property Height As Integer
        Get
            Return Height_
        End Get
        Set(value As Integer)
            ForEach(Sub(header)
                        If header.Height > Height_ Then Height_ = header.Height
                    End Sub)
        End Set
    End Property
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
    Public Shadows Function Add(Text As String) As ColumnHead

        Dim addHeader As New ColumnHead(Text)
        Return Add(addHeader)

    End Function
    Public Shadows Function Add(addHeader As ColumnHead) As ColumnHead

        If addHeader Is Nothing Then
            Return Nothing
        Else
            With addHeader
                .Parent_ = Me
                .Style_ = Styles
                .MouseStyle_ = MouseStyles
                .Index_ = Count
            End With
            MyBase.Add(addHeader)
            RaiseEvent Changed()
        End If
        Return addHeader

    End Function
    Public Shadows Function Item(headerName As String) As ColumnHead

        Dim names As New List(Of ColumnHead)
        names.AddRange(Where(Function(h) h.Name = headerName))
        Return If(names.Any, names.First, Nothing)

    End Function
    Public Shadows Sub Insert(index As Integer, insertHeader As ColumnHead)

        MyBase.Insert(index, insertHeader)
        RaiseEvent Changed()

    End Sub
    Public Shadows Sub Remove(removeHeader As ColumnHead)

        MyBase.Remove(removeHeader)
        RaiseEvent Changed()

    End Sub

#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                Styles_.Dispose()
                MouseStyles_.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose( disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose( disposing As Boolean) above.
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

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ COLUMN HEAD
<Serializable> Public Class ColumnHead
    Implements IDisposable

    Friend Event Changed(sender As Object, e As EventArgs)
    Friend Parent_ As ColumnHeadCollection
    Public ReadOnly Property Parent As ColumnHeadCollection
        Get
            Return Parent_
        End Get
    End Property
    Public Sub New(headerName As String)
        Name_ = headerName 'Can NOT be changed
        Text_ = headerName 'Can be changed
    End Sub
    Friend Index_ As Integer = -1
    Public ReadOnly Property Index As Integer
        Get
            Return Index_
        End Get
    End Property
    Private SortOrder_ As New SortOrder
    Public Property SortOrder As SortOrder
        Get
            Return SortOrder_
        End Get
        Set(value As SortOrder)
            If value <> SortOrder_ Then
                Dim sortUp As String = "iVBORw0KGgoAAAANSUhEUgAAABEAAAALCAYAAACZIGYHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVChTY2SAgAYoTQ4A6z0PxP8pwPep4hLG//9BhlEGyHFJAzaLSQ2T+yBDkDFVXEL3MMEaFjBAbJhghAUMU8ElDAwAvNhdwMSXsO4AAAAASUVORK5CYII="
                Dim sortDown As String = "iVBORw0KGgoAAAANSUhEUgAAABEAAAALCAYAAACZIGYHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABLSURBVChTY2SAgAYojQ80/P//H8rEBOeBGCRLCN8HGYINU8UljLgkSAGkuAQGsLqI2DCBYYywoYpLBixM0AFYL6lhgo7vU8ElDA0AaFFdwFj1ubQAAAAASUVORK5CYII="
                _SortIcon = If(value = SortOrder.None, Nothing, Base64ToImage(If(value = SortOrder.Ascending, sortDown, sortUp)))
                Parent?.ForEach(Sub(column)
                                    If column IsNot Me Then column.SortOrder = SortOrder.None
                                End Sub)
                SortOrder_ = value
                Parent?.Parent?.RequiresRepaint()
            End If
        End Set
    End Property
    Friend ReadOnly Property SortIcon As Image
    Friend WithEvents Style_ As New CellStyle With {
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
    Friend WithEvents MouseStyle_ As New CellStyle With {
        .BackColor = Color.Transparent,
        .ShadeColor = Color.White,
        .ForeColor = Color.Black,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.None
    }
    Public ReadOnly Property MouseStyle As CellStyle
        Get
            Return MouseStyle_
        End Get
    End Property
    Friend ReadOnly Property StyleChanged As Boolean
    Private Sub Style_PropertyChanged(sender As Object, e As StyleEventArgs) Handles Style_.PropertyChanged
        Using defaultStyle As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.White, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
            _StyleChanged = Style_ <> defaultStyle
        End Using
        Parent?.Parent?.Invalidate()
    End Sub
    Friend Name_ As String = "Header"
    Public ReadOnly Property Name As String
        Get
            Return Name_
        End Get
    End Property
    Private Font_ As Font
    Public Property Font As Font
        Get
            Return Font_
        End Get
        Set(value As Font)
            If value IsNot Nothing Then
                If Font_ Is Nothing OrElse value.Name <> Font_.Name And value.Style <> Font_.Style And value.Size <> Font_.Size Then
                    Font_ = value
                    TextSize_Set()
                End If
            End If
        End Set
    End Property
    Friend ReadOnly Property TextSize As Size
    Private Sub TextSize_Set()
        'Nodes use Functions.MeasureText ( Tighter ), ColumnHead is TextRendererMeasureText ( a little wider )
        '"Contact_Name" Functions.MeasureText.Width = 85 ( IBM Plex Mono, 8 ) but TextRendererMeasureText.Width = 91! DrawString wraps at 85, but not at 91
        _TextSize = TextRenderer.MeasureText(If(Text, String.Empty), Font) '(If(Text, String.Empty), Font)
    End Sub
    Private Text_ As String = "Header"
    Public Property Text As String
        Get
            Return Text_
        End Get
        Set(value As String)
            If Text_ <> value Then
                Text_ = value
                TextSize_Set()
                RaiseEvent Changed(Me, Nothing)
            End If
        End Set
    End Property
    Private Image_ As Image = Nothing
    Property Image As Image
        Get
            Return Image_
        End Get
        Set(Value As Image)
            If Not SameImage(Image_, Value) Then
                Image_ = Value
                RaiseEvent Changed(Me, Nothing)
            End If
        End Set
    End Property
    Public ReadOnly Property Height As Integer
        Get
            Return 2 + {If(Image Is Nothing, 0, Image.Height), TextSize.Height}.Max + 2
        End Get
    End Property
    Private ContentWidth_ As Integer
    Friend Property ContentWidth As Integer
        Get
            Return ContentWidth_
        End Get
        Set(value As Integer)
            ContentWidth_ = {ContentWidth_, value}.Max
        End Set
    End Property
    Public Property Visible As Boolean = True
    Friend Property Bounds_Image As New Rectangle
    Friend Property Bounds_Text As New Rectangle
    Friend Property Bounds_Sort As New Rectangle
    Friend Property Bounds As New Rectangle
    ReadOnly Property DataType As Type
        Get
            If GridFormat Is Nothing Then
                Return Parent?.DataTypes(Name)
            Else
                Return GridFormat.DataType
            End If
        End Get
    End Property
    Public Property GridFormat As AlignFormat
    Public Overrides Function ToString() As String

        Dim typeString As String = If(DataType Is Nothing, String.Empty, ", Type=" & Replace(DataType.ToString, "System.", String.Empty))
        Return $"Index={Index}, {Text}, Width={Bounds.Width}, {typeString}"

    End Function

#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                Font_.Dispose()
                Style_.Dispose()
                MouseStyle_.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose( disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose( disposing As Boolean) above.
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