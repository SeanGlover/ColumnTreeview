Option Explicit On
Option Strict On
Imports System.ComponentModel
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Drawing2D

Public Structure HitInfo
    Public Region As String
    Public Branch As Branch
    Public Item As Item
End Structure
Public Structure DropHighlight
    Public Off As Boolean
    Public Item As Object
End Structure
Public Class HeaderTreeView
    Inherits Headers
#Region " Images "
    Public ReadOnly Property ExpandedIcon As Image
    Public ReadOnly Property CollapsedIcon As Image
    Public ReadOnly Property StarIcon As Image
    Public ReadOnly Property DataTableIcon As Image
    Public ReadOnly Property DragImage As Bitmap
    Private Const StarString As String = "iVBORw0KGgoAAAANSUhEUgAAAAwAAAAMCAIAAADZF8uwAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAXtJREFUKFNtkU1LAlEUhkcQbFt/xT/gylX+AJcRXDDCRS7CCCwCFzXgJsiLiYQwOWUZVx2/onGERO9cJ7OoRWbgJ+aIEgWGKF01AqGXszjvy3M4cI6m3W6rqqrVajUaDTOv8Xis0y0sLS0yGONyuVyv15vzarVatVoNYzmfx0wul6tWq43/RCdlWc5ms0wmk8HKg/z4Oqt3tacUe382mpQkKcOkUhIMlZyowuF+5K7T//x+fmleK90g6frkijd8k06nGf9FwmjjDDYuFH+jxGg0GgyGjUbnHD25OLLBomhUZLxngh5APfCbtkmCNIbDYfPjy+q7N1gIDfXAwSH0Cxnt4tZJ4Vaph3FVrrWOkmXgJEYg6gHLocQEMgK4yoqb3oLrtLR2rOwGlH1/EUBi3qMQ5NB0nWkHAihaIHFCAqdFG2oBKy7PoMBlxGxjgSO4DsOsG3k8AoTowI2oBfbgivUwkpImF+f5K54XeEEICUIsFhemDbU0RChOH/EDcKchcY4euAgAAAAASUVORK5CYII="
    Private Const LightOn As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAUtJREFUKFNj/P//PwMK+P/r6ytGRmZWLhFUcQYGoFI4+PryxMfbk388X/rj2ZIPt6d8e3MRWRah9MurC9+er/375fi/T3v/ftr99+vJr49Xfnt3C64aofTtrSV/Ph76/bju9/Npv59P+v245s+no29vLsWi9M3VGT/uVPy4lvTzQdfPe80/Lkf+vFvz+soMLEpfnu/4diHy683Wn6/W/3y96euNpm8Xol+c68ai9NnF+Z9Oer/cYfJ6n8Orfc6vdpp+Oun//PIyLEp///hyb1fS620Gd1ca3Ftl8Gab4f29eX9+/8SiFCj0/fPrh9tib6/2vLnU6N6mkF/fP2EPLIjoo2OTrq+JuDLf5NHRCcjqQDGFxr91fPn2/oC1zTa3Tm8koPTIkSN7du8qKS4+deoUAaWXL19uamoqKSm5ffs2AaVA6b9ggKYOyAUAkObu3QMxkwMAAAAASUVORK5CYII="
    Private Const LightOff As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAARZJREFUKFOFkM1qg1AQhdOdT5uF6QM0O9MHaHYppFtXulJBgxupeiVWEfzb+IMUBK1YQRHtoKCmCJ7FZZj57rlzz1Pf97tHNU3TdR2GYf/6O0AnpWmqKIqu64ZhqKoahuFyOqN5ngNUVdXvoLIsPc+L43iiZxSciqIA459BUMBlTdNWULBMkiSKou9BUARBgBCCvUd6dgUD3/cdx4F3Qa7r2rYtSdKKq2VZ9/sdbOA0TVOWZag1hFbQtm1FURQEgaZpiqJ4nuc4Dn65gkKrrmuAPq7X89uZJMksy9bDGrvyp/x+uRAEwTDMknv41jj40vXjy3G/38MmGyhC6k24PR8OLMtuoJDX6+mE4zjEvIHCGDKfYl/Sf9M5/Uxpz2tBAAAAAElFTkSuQmCC"
    Private Const BookClosed As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAALDgAACw4BQL7hQQAAAddJREFUKFNtkd9P2lAcxYsPvvnv6njBYIAhEKVhxkHQoREFtOWHcYEZQKoIpSkt0dKglh+9GAwphlbBUSzsEhyyZOfp/vice3LPVzcej5E5jUZjRVF1OmRpaXFhQTd/hUB0puG7JtTkZArEYkKZl7rd3yNo/atPVNNGQlWOnwMsLJz9bMYTtdRFg2XbqqpN4Q8UcvWGEjmtneBVPAwIQk6nnn8l2jguRqL3HCdBwwSd5FZlyIUj9dBxPRGXotFW6Bj4/VWvl9/YYEzmPGgqyCS3Jp8nAXzvwC9gWDMYaPh8925PGUVLViv1ZSWztpZqtWSk03kLYRXf3u3Oj7sd7517m99ycc7NksPBWCyUXp8xmZNZku313xBRlLe+k55dat2RW7dR1q+UxUyajKRxNafXE1Zb+qZye3RSBEBGRPDs2SXyDB/AaLMFZuWMxvyqIQs5myPNcDeQc7kKktRD+v3B1fXDvr+QpfnLAv9tmzQYrlaWCTvkylwQKzqd5NPT60dZw6HWbr8eBpjwGctWhCOM3kQvqBIbwot2+/XLy+CfXuFGVd9zefEwyNDcA+g8YjEWRckZ9zmCqQ8W1xC7mUshgDFuDz3N/c9gZ0e93qD5qMB/zHNw/QdY3clc1dADtgAAAABJRU5ErkJggg=="
    Private Const BookOpen As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAALDgAACw4BQL7hQQAAAeJJREFUKFN10V9v0lAYBnCWJV5544fxM3jntYlGs5joxbIlOhNHNEsVNS4zkexPtiUsm8ZZNlmxIRtsrJSuDih1qXQUgrUtpS3dpCtDkI5hwRYSuJg+d+c5vzfnTc5Au912dKJVGlcuXxocHOge/xGLWuGkX/t0KZuvNM7NbnMxDqtS1foSmNHKxhf6+IdctRrjrElli0GcRZJib8am6bQ25Sb1spHmNEY4zXA6tENHknm9Ykx7v9VqzfPOUzalqKOXr/d1vc5JJT/KfsZ4SdXOGk3r6tUi5t8QcOyoXm/alPyqAC7MogVFW92iITTXp/NBglTmZ9Nq8bdNiaQ8AaBdurSRWN9levTFXIA8KE6/TRWVmk3xGD/+NHTSoQsg5t2me/TZjM+iU5MHslR1tFqtIEKNjPkSRMGic6sRMJQSlZJyfBqn2GVfNE5Iz4GYKFYs2mZy+VGnB3CF0T125kN4PXwYJRj3SugdhB/mJAwXnOOoIJTtBf6YJisow48X3bN7bzybd50LgBvywgksnksxcmArO/pgk+NPbNoNX1DHJlau33JdvTZ047bn5tDynXvv7w+Dj5zwwydrvPizT03T/M7La3DkI7QD+ndBGPHCyKdANIiQMTJTrRl9+r+v7/V/AT5wyfCHirK9AAAAAElFTkSuQmCC"
    Private Const ExpandedString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAO9JREFUOE+lk9sKAVEUhnkoD+E5vJJzuHEshzs1DiGUQ0RRigsKJceRcWZ+tppxmjXKTK2bqe/791p7Lz0AnaaPCbSUDI8mK3iiBbgjebjCOThCGdiDKVh9SZi9nFylWvue9wyVBZ5YEaIIXK4ijqcrtvsLeOGMGX/EeH7AYLJDdyjAYDQpC9yRwk+43d/QAnZstWQGN3prWuC890wdW4IrHZ4W2AJpuWfW52cxuNha0gKLP/E1sNdkBmebC1pg9nFv01aCU/W5ukC6KgrmqjN1AbtnNThentIC9sKUhvf5j3yJ/+6DpkV6bPK/yRJ3A/PE7e2oP8DgAAAAAElFTkSuQmCCAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="
    Private Const CollapsedString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAARpJREFUOE+lk9tKAmEUhfWhfIiew1eyUkkvPCaNdZdHbPCACIqCklZERQZCZhNZeZylS5jRhvlHcAb2zcD61tqLfzsBOGx9BNgZXdwffCKYLCEgFXF2IcN3XoA3nsNJJAtPOK1Ptd5Z+21NdUDwsgxVBRZLFdPZEj9/CyjjOd6VKd6GEzwPfnH/OobryG0OCEilvWK58SIGMLbRmW6ac+fpG1fynRjgX+9sjE0AY1PcfPiCVLAAnMby+s4UGqfWVZDI98QJjqMZ08LoTHG5PUI82xUDPJH0v7YZmyk08U3rA9HMHsBuYbvOFOcaQ4RTt9YJWFix2Ueq8rgpjDszNp0pDl1bAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="
    Private Const ExpandedArrowString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAM9JREFUOE/N0kkKg0AQBdCf+x/KlaALxXkeF04o4g0qXU1sDPQqBhLhI72o17+gH0SEWx8Dd3JrWLa/c/t3Ac/ziOM4Dtm2LWNZFpmmKWMYhsq1tVphmiYw0Pc9tW1LVVVRnueUpilFUURBEEjAdd23tdVhWRZckaZpqCxLiSRJIocFwpfogW3bwMg4juDqXdfRifCwQCCawPd9PbDvO9Z1VQgPMcJ/0QRZloGRMAz1wHEcOJF5njEMA14I6rpGURQSieNYD3z6Hv7oIf1shSf3G9UMQ+Vu/QAAAABJRU5ErkJggg=="
    Private Const CollapsedArrowString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAANFJREFUOE+l00kKhDAQheHq+x/KlaALxXkeF04o4g1ep0J3Y0OEiILoIvnyK9QLAD26GHhyKzc7joNpmmhZFtq2jfZ9p+M4lGsvgTOyrqtEVKWXQN/3YGQcR1nCiDZg2zbatsUZmef5HlBVFZqmQdd1smQYBn3AsizkeY6yLP8Q7U8wTRNpmv4QLhAl+gUMRFGEJEnA76KE6rrWBwzDQBAE4KdAKMsyKoriHvBBSJTQF9H+B7zZdV3yPI9836cwDCmOY/2CO7PxaJDkJN85TbX2Db5d1YfJcQ3TAAAAAElFTkSuQmCC"
    Private Const DataTableString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAmJJREFUOE99Ul1PGkEU3T60Puiv8iPR+OqrH0++m0ir0TQxJjWR4EdrCBgBG4wQSyEiQmhgXUAqKwtbwGRZ2Fl2MaAoghWILCxMh5AaNU3v070z98w55955AyHE/gbP85Ikoardbv9+qLx914NupXq9r7ev1YLtNuzt7cHQ0VOwLPuUcwCgHPUxLF+twrs7mM9Dkoy9ACSTyVcAVLIpvlbrAK6uYChEY7IMRbFB01W/v3R8nLNaBbM5ZTQm9HrSbA4TRIJNgVpNLpXkQkEOhaLY5aWcycjRaP30tOL1llyugt2es9nEvT1ma+tselqr1pgZBsTjgKYBQQSxTKbJcS3E4PPdOByi1crr9b+USu/8vGVy8vPg4AfDrv3hAd7ewlwOBoMUxvMNBAiHKy5X3mRKoYfVanJl5cfc3LeJic2Bgfe7Xx3lMry+htksDAQoTBAaglCnqNLRkWAwxLVaSqUilpYcCoV5aurL6OhH1ZoxFgORCDg/BzgewACQugxeb8FmQ47T29uR1dWOpC6DZvswn2+KYhOAps93jjw8ptPSPyWNj28MDy+oNYeiCDkOMgzEcRLjuEdBaFFU0e3OmkyMTkevr/uXl51dSUNDcxub39EO0YYuLqDHQyJJtf+YRlNaVeoIgsTxkMdzRlE0YqgJgkSSRYulox6ZVirxxUXbzMweYhgZWTw5oZ9/HzTWKsvKkUjF7y/a7Zf7+yyCra35FhasyHR/vyIaTb8AyHKbpks4fuN05p7vYXbWNDb2aWfH+by787W79f19PZG4JgjO4biwWKiDg59uN1UuV151o/IPPuNL2ItzNKQAAAAASUVORK5CYII="
    Private Sub LoadImages()
        _StarIcon = FromBase64String(StarString)
        _DataTableIcon = FromBase64String(DataTableString)
        Select Case mExpandStyle
            Case TreeExpanderStyle.Arrow
                _ExpandedIcon = FromBase64String(ExpandedArrowString)
                _CollapsedIcon = FromBase64String(CollapsedArrowString)
            Case TreeExpanderStyle.PlusMinus
                _ExpandedIcon = FromBase64String(ExpandedString)
                _CollapsedIcon = FromBase64String(CollapsedString)
            Case TreeExpanderStyle.Book
                _ExpandedIcon = FromBase64String(BookOpen)
                _CollapsedIcon = FromBase64String(BookClosed)
            Case TreeExpanderStyle.LightBulb
                _ExpandedIcon = FromBase64String(LightOn)
                _CollapsedIcon = FromBase64String(LightOff)
        End Select
    End Sub
    Private Function FromBase64String( ImageString As String) As Image
        Dim b() As Byte = Convert.FromBase64String(ImageString)
        Dim Image As Image
        Try
            Using MemoryStream As New System.IO.MemoryStream()
                MemoryStream.Position = 0
                MemoryStream.Write(b, 0, b.Length)
                Image = Image.FromStream(MemoryStream)
                Dim Bmp As New Bitmap(Image)
                Bmp.MakeTransparent(Bmp.GetPixel(0, 0))
                Image = CType(Bmp, Image)
                MemoryStream.Close()
            End Using
            Return Image
        Finally

        End Try
    End Function
#End Region

    Public Sub New()
        Size = New Size(400, 300)
        LoadImages()
        AddHandler Branches_.Changed, AddressOf OnBranchesChanged
        AddHandler SelectedBranches_.Changed, AddressOf OnSelectedBranchesChanged
        AddHandler SelectedItems_.Changed, AddressOf OnSelectedItemsChanged
        AddHandler PreviewKeyDown, AddressOf OnPreviewKeyDown
        MyBase.OnFontChanged(New EventArgs)
    End Sub

#Region " Properties & Fields "
    Public DropHighlight As New DropHighlight
    Private ClickedItem As Object
    Private ItemClicked As Boolean = False
    Private Dragging As Boolean = False
    Private RowIndex As Integer = 0
    Private HeaderIndex As Integer = 0
    <Browsable(False)> Public MultiSelecting As Boolean = False     'SHOULD BE PRIVATE
    Private mCursor As Boolean = True
    Private mDataSet As DataSet
    Private mDefaultContextMenuStrip As Boolean = False
    Public Property DefaultContextMenuStrip As Boolean
        Get
            Return mDefaultContextMenuStrip
        End Get
        Set( Value As Boolean)
            If Value <> mDefaultContextMenuStrip Then
                mDefaultContextMenuStrip = Value
                If Value Then
                    Dim CMS As New ContextMenuStrip
                    Dim Expand As New ToolStripMenuItem("Expand All", CollapsedIcon, AddressOf OnCMSClick)
                    Dim Collapse As New ToolStripMenuItem("Collapse All", ExpandedIcon, AddressOf OnCMSClick)
                    CMS.Items.AddRange({Expand, Collapse})
                    ContextMenuStrip = CMS
                End If
            End If
        End Set
    End Property
    ReadOnly Property RowHeight As Integer
        Get
            RowHeight = Convert.ToInt32(Font.GetHeight) + 2
            Dim Q_Image As IEnumerable(Of Int32) = (From A In (From N In Branches Select CType(N, Branch)) Where Not IsNothing(A.Image) Select A.Image.Height)
            If Not Q_Image.Count = 0 Then
                If Q_Image.Max + 2 > RowHeight Then RowHeight = Q_Image.Max + 2
            End If
            Return RowHeight
        End Get
    End Property
    Protected WithEvents Branches_ As New BranchCollection(Me)
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> ReadOnly Property Branches() As BranchCollection
        Get
            Return Branches_
        End Get
    End Property
    Private mBranchLineStyle As DashStyle = DashStyle.Dash
    <DescriptionAttribute("Tree Header BranchLineStyle")> Public Property BranchLineStyle() As System.Drawing.Drawing2D.DashStyle
        Get
            Return mBranchLineStyle
        End Get
        Set( Value As System.Drawing.Drawing2D.DashStyle)
            If Value <> mBranchLineStyle Then
                mBranchLineStyle = Value
                Invalidate()
            End If
        End Set
    End Property
    Private mBranchLineColor As Color = Color.Black
    <DescriptionAttribute("Tree Header BranchLineColor")> Public Property BranchLineColor() As Color
        Get
            Return mBranchLineColor
        End Get
        Set( Value As Color)
            mBranchLineColor = Value
            Invalidate()
        End Set
    End Property
    Private mRootLines As Boolean = True
    Public Property RootLines() As Boolean
        Get
            Return mRootLines
        End Get
        Set( Value As Boolean)
            mRootLines = Value
            Invalidate()
        End Set
    End Property
    Private mGridLines As Boolean = True
    <DescriptionAttribute("Tree Header GridLines")> Public Property GridLines() As Boolean
        Get
            Return mGridLines
        End Get
        Set( Value As Boolean)
            mGridLines = Value
            Invalidate()
        End Set
    End Property
    Private mExpandStyle As TreeExpanderStyle
    <DescriptionAttribute("PlusMinus, Arrow.")> Public Property ExpanderStyle() As TreeExpanderStyle
        Get
            Return mExpandStyle
        End Get
        Set( Value As TreeExpanderStyle)
            If Value <> mExpandStyle Then
                mExpandStyle = Value
                LoadImages()
            End If
        End Set
    End Property
    Private mIndent As Integer = 0
    Property Indent() As Integer
        Get
            Return mIndent
        End Get
        Set( Value As Integer)
            mIndent = Value
        End Set
    End Property
    Public ReadOnly Property LoadTime As TimeSpan
    Public ReadOnly Property ImageList As ImageList
    Private SortType_ As SortType = SortType.NoSort
    Property SortType() As SortType
        Get
            Return SortType_
        End Get
        Set( Value As SortType)
            SortType_ = Value
            Branches_.SortType = SortType_
            Invalidate()
        End Set
    End Property
    Private mHeaderStyle As Border3DStyle = Border3DStyle.Etched
    Property HeaderStyle As Border3DStyle
        Get
            Return mHeaderStyle
        End Get
        Set( Value As Border3DStyle)
            mHeaderStyle = Value
        End Set
    End Property
    Public Property FullRowSelect As Boolean
    Public Property DropHighlightColor As Color
    Property MultiSelect As Boolean
    Private mSelectStyle As SelectStyle = SelectStyle.Box
    Property SelectStyle As SelectStyle
        Get
            Return mSelectStyle
        End Get
        Set( value As SelectStyle)
            mSelectStyle = value
        End Set
    End Property
    Private mBackImageAlpha As Integer = 0
    Property BackImageAlpha As Integer
        Get
            Return mBackImageAlpha
        End Get
        Set( value As Integer)
            If value >= 0 AndAlso value <= 255 Then
                mBackImageAlpha = value
            Else : mBackImageAlpha = 0
            End If
        End Set
    End Property
    ReadOnly Property BranchCollection() As List(Of Branch)
        Get
            Dim Branches As New List(Of Branch)
            Dim Queue As New Queue(Of Branch)
            Dim TopBranch As Branch, Branch As Branch
            For Each TopBranch In Branches
                Queue.Enqueue(TopBranch)
            Next
            While (Queue.Count > 0)
                TopBranch = Queue.Dequeue
                Branches.Add(TopBranch)
                For Each Branch In TopBranch.Branches
                    Queue.Enqueue(Branch)
                Next
            End While
            Return Branches
        End Get
    End Property
    Private WithEvents SelectedBranches_ As New BranchCollection(Me)
    <Browsable(False)> ReadOnly Property SelectedBranches() As BranchCollection
        Get
            Return SelectedBranches_
        End Get
    End Property
    Private WithEvents SelectedItems_ As New ItemCollection(Me)
    <Browsable(False)> ReadOnly Property SelectedItems() As ItemCollection
        Get
            Return SelectedItems_
        End Get
    End Property
    Private WithEvents Selection_ As New Dictionary(Of Object, Object)
    <Browsable(False)> ReadOnly Property Selection() As Dictionary(Of Object, Object)
        Get
            Return Selection_
        End Get
    End Property
    Private WithEvents VisibleBranches_ As New List(Of Branch)
    Private ReadOnly Property VisibleBranches As List(Of Branch)
        Get
            Dim Q_Visible As New List(Of Branch)(From N In Branches.All Where N.Visible AndAlso N.Bounds.Bottom < ClientRectangle.Bottom Order By N.Bounds.Top Ascending Select N)
            Return Q_Visible
        End Get
    End Property
#End Region
#Region " Drawing "
    Protected Overrides Sub DrawItems( Graphics As Graphics,  Rect As Rectangle)

        If Headers.Any() Then
            DrawBranches(Graphics)
            If mGridLines Then DrawGridLines(Graphics)
        Else

        End If

    End Sub
    Private Sub DrawBranches( Graphics As Graphics)

        For Each Header As Header In Headers
            If IsNothing(BackgroundImage) Then
                Using Brush As Brush = New SolidBrush(BackColor)
                    Graphics.FillRectangle(Brush, New Rectangle(Header.Bounds.Left, 0, Header.Bounds.Width, ClientSize.Height))
                End Using
            Else
                Dim Section As New Rectangle(Header.Bounds.Left, 0, Header.Bounds.Width, BackgroundImage.Height)
                Graphics.DrawImage(DrawHeaderImage(CType(BackgroundImage, Bitmap), Section), New Point(Header.Bounds.Left, 0))
                Using AlphaBrush As New SolidBrush(Color.FromArgb(BackImageAlpha, BackColor))
                    Graphics.FillRectangle(AlphaBrush, New Rectangle(Header.Bounds.Left, 0, Header.Bounds.Width, ClientSize.Height))
                End Using
            End If
            For Each Branch As Branch In (From N In Branches.All Where DirectCast(N, Branch).Visible = True AndAlso DirectCast(N, Branch).Bounds.Bottom < ClientRectangle.Bottom Select DirectCast(N, Branch)).ToList
                Dim BranchBounds As Rectangle = Branch.Bounds
                If Not BranchBounds.Bottom < HeaderHeight Then
                    If BranchBounds.Top > ClientSize.Height Then Exit For
                    If Headers.IndexOf(Header) = 0 Then
                        Dim BranchImageWidth As Integer = 0, BranchImageHeight As Integer = 0
                        If Branch.Selected Then DrawSelection(Graphics, BranchBounds, SelectedColor)
                        If Not DropHighlight.Off AndAlso Not IsNothing(DropHighlight.Item) AndAlso DropHighlight.Item.GetType Is GetType(Branch) AndAlso Branch Is DropHighlight.Item Then
                            DrawSelection(Graphics, BranchBounds, DropHighlightColor)
                        End If
                        If Not IsNothing(ImageList) Then
                            BranchImageWidth = StarIcon.Width
                            BranchImageHeight = StarIcon.Height
                            If IsNothing(Branch.Image) Then Branch.Image = StarIcon 'Default
                            BranchImageWidth = Branch.Image.Width
                            BranchImageHeight = Branch.Image.Height
                            Graphics.DrawImage(Branch.Image, New Rectangle(2 + BranchBounds.X, BranchBounds.Y + Convert.ToInt32((BranchBounds.Height - BranchImageHeight) / 2),
                                                                         BranchImageWidth, BranchImageHeight))
                        End If
                        Dim TextHeight As Integer = TextRenderer.MeasureText(Branch.Text, Font).Height
                        TextRenderer.DrawText(Graphics, Branch.Text, Font, New Point(BranchBounds.Left + 0 + BranchImageWidth, Convert.ToInt32((RowHeight - TextHeight) / 2) + BranchBounds.Top), Branch.ForeColor)
                        If Branch.HasChildren Then
                            Dim ExpanderRect As New Rectangle(Branch.BoxBounds.X + Convert.ToInt32((Branch.BoxBounds.Width - CollapsedIcon.Width) / 2), BranchBounds.Y + Convert.ToInt32((BranchBounds.Height - CollapsedIcon.Height) / 2), CollapsedIcon.Width, CollapsedIcon.Height)
                            If Branch.State = Branchestate.Collapsed Then
                                Graphics.DrawImage(CollapsedIcon, ExpanderRect)
                            Else
                                Graphics.DrawImage(ExpandedIcon, ExpanderRect)
                            End If
                        End If
                        DrawBranchLines(Graphics, Branch)
                    Else
                        If Not IsNothing(Branch.Items) AndAlso Branch.Items.Count >= Headers.IndexOf(Header) Then
                            Dim Item As Item = Branch.Items(Headers.IndexOf(Header) - 1)
                            Dim ItemImageWidth As Integer = 0, ItemImageHeight As Integer = 0
                            Dim ItemBounds As Rectangle = Item.Bounds
                            If Item.Selected And Not FullRowSelect = True OrElse FullRowSelect AndAlso Item.Parent.Selected Then DrawSelection(Graphics, ItemBounds, SelectedColor)
                            If Not DropHighlight.Off AndAlso Not IsNothing(DropHighlight.Item) Then
                                If FullRowSelect AndAlso Item.Parent Is DropHighlight.Item Then DrawSelection(Graphics, ItemBounds, DropHighlightColor)
                                If Not FullRowSelect AndAlso Item Is DropHighlight.Item Then DrawSelection(Graphics, ItemBounds, DropHighlightColor)
                            End If
                            If Not IsNothing(Item.Image) Then
                                ItemImageWidth = Item.Image.Width
                                ItemImageHeight = Item.Image.Height
                                Graphics.DrawImage(Item.Image, New Rectangle(2 + ItemBounds.X, ItemBounds.Y + Convert.ToInt32((BranchBounds.Height - ItemImageHeight) / 2),
                                                                             Item.Image.Width, Item.Image.Height))
                            End If
                            Dim TextSize As Size = TextRenderer.MeasureText(Item.Text, Font)
                            Dim TextVertOffset As Integer = Convert.ToInt32((RowHeight - TextSize.Height) / 2)
                            Dim TextHoriOffset As Integer = 0
                            If Headers(1 + Branch.Items.IndexOf(Item)).Alignment = HorizontalAlignment.Center Then
                                TextHoriOffset = Convert.ToInt32((ItemBounds.Width - ItemImageWidth - TextSize.Width) / 2)
                            ElseIf Headers(1 + Branch.Items.IndexOf(Item)).Alignment = HorizontalAlignment.Right Then
                                TextHoriOffset = Convert.ToInt32(ItemBounds.Width - TextSize.Width - ItemImageWidth)
                            End If
                            TextRenderer.DrawText(Graphics, Item.Text, Font, New Point(ItemBounds.Left + TextHoriOffset + ItemImageWidth, TextVertOffset + ItemBounds.Top), Item.ForeColor)
                        End If
                    End If
                End If
            Next
        Next
        TotalHeight = If(VisibleBranches.Any, VisibleBranches.Count * RowHeight, 0) + 2

    End Sub
    Private Sub DrawSelection( Graphics As Graphics,  Rect As Rectangle,  Color As Color)
        Using SelectedBrush As New SolidBrush(Color)
            If SelectStyle = SelectStyle.Box Then
                Graphics.FillRectangle(SelectedBrush, Rect)
                Graphics.DrawRectangle(SystemPens.Highlight, Rect)
            ElseIf SelectStyle = SelectStyle.Highlight Then
                Using Brush As New Drawing2D.LinearGradientBrush(Rect, BackColor, Color, Drawing2D.LinearGradientMode.Vertical)
                    Graphics.FillRectangle(Brush, Rect)
                End Using
            End If
        End Using
    End Sub
    Private Sub DrawBranchLines( g As Graphics,  Branch As Branch)

        Using Pen As New Pen(mBranchLineColor, 1)
            Pen.DashStyle = mBranchLineStyle
            Do While Not IsNothing(Branch.Parent)
                Dim BranchBounds As Rectangle = Branch.Bounds
                Dim ParentExpandBounds As Rectangle = Branch.Parent.BoxBounds, ExpandBounds As Rectangle = Branch.BoxBounds
                Dim ParentBoundsLeft As Integer = ParentExpandBounds.Left + Convert.ToInt32(ParentExpandBounds.Width / 2)
                g.DrawLine(Pen, ParentBoundsLeft, BranchBounds.Top + Convert.ToInt32(BranchBounds.Height / 2), If(Branch.HasChildren, ExpandBounds.Left, BranchBounds.Left), BranchBounds.Top + Convert.ToInt32(BranchBounds.Height / 2))  '///Horizontal
                g.DrawLine(Pen, ParentBoundsLeft, ParentExpandBounds.Bottom, ParentBoundsLeft, BranchBounds.Top + Convert.ToInt32(BranchBounds.Height / 2))  '///Vertical
                Branch = Branch.Parent
            Loop
            If RootLines Then
                Dim RootIndent As Integer = 6
                Dim TopBounds As Rectangle = Branches(0).Bounds, BottomBounds As Rectangle = (From N In Branches Select DirectCast(N, Branch).Bounds).Last
                g.DrawLine(Pen, RootIndent, TopBounds.Top + Convert.ToInt32(TopBounds.Height / 2), RootIndent, BottomBounds.Top + Convert.ToInt32(BottomBounds.Height / 2))   '///Vertical
                For Each RootBranch As Branch In Branches
                    Pen.DashStyle = mBranchLineStyle
                    g.DrawLine(Pen, RootIndent, RootBranch.Bounds.Top + Convert.ToInt32(RootBranch.Bounds.Height / 2), RootIndent * 2, RootBranch.Bounds.Top + Convert.ToInt32(RootBranch.Bounds.Height / 2))  '///Horizontal
                Next
            End If
        End Using

    End Sub
    Private Sub DrawGridLines( g As Graphics)
        For Each Header As Header In Headers
            Using Pen As New Pen(Color.Gainsboro, 1)
                Dim Dashes As Single() = {2, 2, 2, 2}
                Pen.DashPattern = Dashes
                g.DrawLine(Pen, New Point(Header.Bounds.Left - 1, 0), New Point(Header.Bounds.Left - 1, Bounds.Bottom))
            End Using
        Next
    End Sub
    Private Function DrawHeaderImage( Source As Bitmap,  Section As Rectangle) As Bitmap
        Dim Bitmap As New Bitmap(Section.Width, Section.Height)
        Using G As Graphics = Graphics.FromImage(Bitmap)
            G.DrawImage(Source, 0, 0, Section, GraphicsUnit.Pixel)
        End Using
        Return Bitmap
    End Function
#End Region
#Region " Helper Methods "
    Private Sub UpdateDragImage()
        Dim ItemImage As Image = Nothing, ImageWidth As Integer = 0, Text As String = String.Empty, Font As Font = Nothing
        If TypeOf (ClickedItem) Is Branch Then
            ItemImage = DirectCast(ClickedItem, Branch).Image
            Text = DirectCast(ClickedItem, Branch).Text
        ElseIf TypeOf (ClickedItem) Is Item Then
            ItemImage = DirectCast(ClickedItem, Item).Image
            Text = DirectCast(ClickedItem, Item).Text
        End If
        Using Graphics As Graphics = CreateGraphics()
            Dim TextSize As SizeF = Graphics.MeasureString(Text, Font)
            Dim RectHeight As Integer = Convert.ToInt32(TextSize.Height)
            If Not IsNothing(ItemImage) Then
                ImageWidth = ItemImage.Width
                If ItemImage.Height > TextSize.Height Then RectHeight = ItemImage.Height
            End If
            Dim DragRect As New Rectangle(0, 0, ImageWidth + Convert.ToInt32(TextSize.Width), RectHeight)
            Using bmp As New Bitmap(DragRect.Width + 8, DragRect.Height, Drawing.Imaging.PixelFormat.Format24bppRgb)
                Using G_Image As Graphics = Graphics.FromImage(bmp)
                    TextRenderer.DrawText(G_Image, Text, Font, New Point(ImageWidth + 1, 0), Color.White)
                    Dim X As Integer
                    Dim Y As Integer
                    Dim Red As Integer
                    Dim Green As Integer
                    Dim Blue As Integer
                    For X = 0 To bmp.Width - 1
                        For Y = 0 To bmp.Height - 1
                            Red = 255 - bmp.GetPixel(X, Y).R
                            Green = 255 - bmp.GetPixel(X, Y).G
                            Blue = 255 - bmp.GetPixel(X, Y).B
                            bmp.SetPixel(X, Y, Color.FromArgb(Red, Green, Blue))
                        Next Y
                    Next X
                    If Not IsNothing(ItemImage) Then G_Image.DrawImage(ItemImage, New Point(0, 0))
                End Using
                bmp.MakeTransparent(Color.White)
                _DragImage = DirectCast(bmp.Clone, Bitmap)
            End Using
        End Using
    End Sub
    Public Function DrawRoundedRectangle( Rect As Rectangle, Optional  Corner As Integer = 10) As GraphicsPath
        Dim Graphix As New System.Drawing.Drawing2D.GraphicsPath
        Dim ArcRect As New RectangleF(Rect.Location, New SizeF(Corner, Corner))
        Graphix.AddArc(ArcRect, 180, 90)
        Graphix.AddLine(Rect.X + CInt(Corner / 2), Rect.Y, Rect.X + Rect.Width - CInt(Corner / 2), Rect.Y)
        ArcRect.X = Rect.Right - Corner
        Graphix.AddArc(ArcRect, 270, 90)
        Graphix.AddLine(Rect.X + Rect.Width, Rect.Y + CInt(Corner / 2), Rect.X + Rect.Width, Rect.Y + Rect.Height - CInt(Corner / 2))
        ArcRect.Y = Rect.Bottom - Corner
        Graphix.AddArc(ArcRect, 0, 90)
        Graphix.AddLine(Rect.X + CInt(Corner / 2), Rect.Y + Rect.Height, Rect.X + Rect.Width - CInt(Corner / 2), Rect.Y + Rect.Height)
        ArcRect.X = Rect.Left
        Graphix.AddArc(ArcRect, 90, 90)
        Graphix.AddLine(Rect.X, Rect.Y + CInt(Corner / 2), Rect.X, Rect.Y + Rect.Height - CInt(Corner / 2))
        Return Graphix
    End Function
#End Region
    Private WithEvents BindingSource As New BindingSource
    Public ReadOnly Property Groups As New List(Of String)
    Private Table_ As DataTable
    Public ReadOnly Property Table As DataTable
        Get
            Return Table_
        End Get
    End Property
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
                Clear()
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
                    Headers.Clear()
                    Dim columnNames As New List(Of String)(From c In Table.Columns Select DirectCast(c, DataColumn).ColumnName)
                    Dim branchGroups As New List(Of String)(columnNames.Intersect(Groups))
                    If branchGroups.Any Then
                        'Groups ... CustomerNbr|CustomerName|Adresss, x|y|z
                        'Groups Headers is delimited by <|>, Children by next item
                        Dim levels As New Dictionary(Of String, String)
                        For Each group In Groups
                            Dim columns As String() = Split(group, "|")
                            Dim firstColumn As String = columns.First
                            Dim columnGroup = From row In Table.AsEnumerable Group row By xGroup = row(firstColumn).ToString Into rowGroup = Group
                            For Each column In columnGroup
                                Branches.Add(column.xGroup)
                            Next
                        Next

                    Else

                    End If
                    _LoadTime = Now.Subtract(startLoad)
                End If
#End Region
            End If
        End Set
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Sub Load( DataSet As DataSet)

        mDataSet = DataSet
        For Each Table As DataTable In DataSet.Tables
            If Not IsNothing(ContextMenuStrip) Then ContextMenuStrip.Items.Add(New ToolStripMenuItem(Table.TableName, DataTableIcon, AddressOf OnCMSClick))
            If (DataSet.Tables.IndexOf(Table)) = 0 Then
                For Each Header As DataColumn In Table.Columns
                    If Table.Columns.IndexOf(Header) = 0 Then
                        Headers.Add(Header.ColumnName, 250)
                    Else
                        Headers.Add(Header.ColumnName)
                    End If
                Next
            End If
            For Each DataRow As DataRow In Table.Rows
                Dim DR As DataRow = DataRow
                Dim Key As String = DataRow.Table.PrimaryKey(0).ColumnName
                Dim Parent = (From A In DataSet.Relations Where DR.GetParentRows(CType(A, DataRelation).RelationName).Count > 0 Select DR.GetParentRow(A.ToString))
                If Parent.Count = 0 Then
                    Dim Branch As Branch = Branches.Add(DataRow(Key).ToString, DataRow(Key).ToString)
                    For I = 1 To DataRow.ItemArray.Count - 1
                        Branch.Items.Add(DataRow.ItemArray(I).ToString, Color.Black)
                    Next
                Else
                    Dim ParentRow As DataRow = Parent.First
                    Dim ParentKey As String = ParentRow.Table.PrimaryKey(0).ColumnName
                    Dim Branch As Branch = Branches.Find(ParentRow(ParentKey).ToString)(0).Branches.Add(DataRow(Key).ToString, DataRow(Key).ToString)
                    For I = 1 To Headers.Count - 1
                        Branch.Items.Add(DataRow.ItemArray(I).ToString, Color.Black)
                    Next
                End If
            Next
        Next

    End Sub
    Public Sub Clear()

        VisibleBranches_.Clear()
        SelectedBranches_.Clear()
        SelectedItems_.Clear()
        Branches_.Clear()
        Invalidate()

    End Sub
    Protected Friend Function HitTest( e As MouseEventArgs) As HitInfo
        Dim Hit As HitInfo = Nothing
        Dim Q_RowHit As IEnumerable(Of Branch) = (From A In Branches.All Where e.Y >= A.Bounds.Top AndAlso e.Y <= A.Bounds.Bottom Select A)
        If Not Q_RowHit.Count = 0 Then
            Hit.Branch = Q_RowHit.First
            If e.X <= Hit.Branch.Bounds.Left Then
                Hit.Region = "±"
            ElseIf e.X <= Hit.Branch.Bounds.Right Then
                Hit.Region = "Branch"
            Else
                Hit.Region = "Item"
                Dim Q_ItemHit As IEnumerable(Of Item) = (From A In (From I In Hit.Branch.Items Select DirectCast(I, Item)) Where
                                e.X >= A.Bounds.X AndAlso e.X <= A.Bounds.Right AndAlso e.Y >= A.Bounds.Y AndAlso e.Y <= A.Bounds.Bottom Select A)
                If Q_ItemHit.Count = 0 Then
                    Hit.Branch = Nothing
                    Hit.Item = Nothing
                Else
                    Hit.Item = Q_ItemHit.First
                End If
            End If
        End If
        Return Hit
    End Function
    Protected Friend Function HitTest( X As Integer,  Y As Integer) As HitInfo
        Dim Hit As HitInfo = Nothing
        Dim Q_RowHit As IEnumerable(Of Branch) = (From A In Branches.All Where Y >= A.Bounds.Top AndAlso Y <= A.Bounds.Bottom Select A)
        If Not Q_RowHit.Count = 0 Then
            Hit.Branch = Q_RowHit.First
            If X <= Hit.Branch.Bounds.Left Then
                Hit.Region = "±"
            ElseIf X <= Hit.Branch.Bounds.Right Then
                Hit.Region = "Branch"
            Else
                Hit.Region = "Item"
                Dim Q_ItemHit As IEnumerable(Of Item) = (From A In (From I In Hit.Branch.Items Select DirectCast(I, Item)) Where
                                X >= A.Bounds.X AndAlso X <= A.Bounds.Right AndAlso Y >= A.Bounds.Y AndAlso Y <= A.Bounds.Bottom Select A)
                If Q_ItemHit.Count = 0 Then
                    Hit.Branch = Nothing
                    Hit.Item = Nothing
                Else
                    Hit.Item = Q_ItemHit.First
                End If
            End If
        End If
        Return Hit
    End Function

#Region " Overrides "
    Protected Shadows Sub OnPreviewKeyDown( sender As Object,  e As PreviewKeyDownEventArgs)
        Select Case (e.KeyCode)
            Case Keys.Down, Keys.Up, Keys.Left, Keys.Right, Keys.Enter
                e.IsInputKey = True
        End Select
    End Sub
    Protected Overrides Sub OnKeyUp( e As KeyEventArgs)
        If e.KeyCode = Keys.ShiftKey Then MultiSelecting = False
    End Sub
    Protected Overrides Sub OnKeyDown( e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        Select Case (e.KeyCode)
            Case Keys.Down, Keys.Up, Keys.Left, Keys.Right, Keys.Enter, Keys.ShiftKey
                Dim Q_Last = (From A In Selection Select A.Key)
                If Not Q_Last.Count = 0 Then
                    Dim Q_Branch As Branch = If(TryCast(Q_Last.Last, Branch), TryCast(Q_Last.Last, Item).Parent)
                    If e.KeyCode = Keys.Enter Then
                        If Q_Branch.State = Branchestate.Collapsed Then
                            Q_Branch.Expand()
                            RaiseEvent AfterExpand(Q_Branch, e)
                        ElseIf Q_Branch.State = Branchestate.Expanded Then
                            Q_Branch.Collapse()
                            RaiseEvent AfterCollapse(Q_Branch, e)
                        End If
                    Else
                        Invalidate()
                        If e.KeyCode = Keys.ShiftKey Then MultiSelecting = True
                        If e.KeyCode = Keys.Left AndAlso Not HeaderIndex = 0 Then HeaderIndex -= 1
                        If e.KeyCode = Keys.Right AndAlso Not HeaderIndex = Q_Branch.Items.Count Then HeaderIndex += 1
                        If e.KeyCode = Keys.Down AndAlso Not RowIndex = VisibleBranches.Count - 1 Then RowIndex += 1
                        If e.KeyCode = Keys.Up AndAlso Not RowIndex = 0 Then RowIndex -= 1
                    End If
                    If Not e.KeyCode = Keys.ShiftKey Then
                        If HeaderIndex = 0 Then
                            VisibleBranches(RowIndex).Selected = True
                        Else
                            If VisibleBranches(RowIndex).Items.Count >= HeaderIndex Then VisibleBranches(RowIndex).Items(HeaderIndex - 1).Selected = True
                        End If
                    End If
                End If
        End Select
    End Sub
    Protected Overrides Sub OnMouseDown( e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        Dragging = False
        Dim Hit As HitInfo = HitTest(e)
        If Not IsNothing(Hit.Branch) Then
            If Hit.Region = "±" Then
                ItemClicked = False
                If Hit.Branch.State = Branchestate.Expanded Then
                    RaiseEvent BeforeCollapse(DirectCast(ClickedItem, Branch), New CancelEventArgs)
                    Hit.Branch.Collapse()
                    RaiseEvent AfterCollapse(DirectCast(ClickedItem, Branch), New EventArgs)
                ElseIf Hit.Branch.State = Branchestate.Collapsed Then
                    RaiseEvent BeforeExpand(DirectCast(ClickedItem, Branch), New CancelEventArgs)
                    Hit.Branch.Expand()
                    RaiseEvent AfterExpand(DirectCast(ClickedItem, Branch), New EventArgs)
                End If
                Invalidate()
            ElseIf Hit.Region = "Branch" Then
                ItemClicked = True
                ClickedItem = Hit.Branch
                HeaderIndex = 0
                RowIndex = VisibleBranches.IndexOf(Hit.Branch)
            ElseIf Hit.Region = "Item" AndAlso Not IsNothing(Hit.Item) Then
                ItemClicked = True
                ClickedItem = Hit.Item
                HeaderIndex = Hit.Branch.Items.IndexOf(Hit.Item) + 1
                RowIndex = VisibleBranches.IndexOf(Hit.Branch)
            End If
            Invalidate()
        End If
    End Sub
    Protected Overrides Sub OnMouseMove( e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If ItemClicked AndAlso e.Button = MouseButtons.Left Then
            Refresh()
            If Not Dragging AndAlso Not IsNothing(ClickedItem) Then    '/// START DRAG ///
                Dragging = True
                mCursor = True
                Dim Items As New List(Of Object)
                Items.AddRange(From Branch In SelectedBranches Select Branch)
                Items.AddRange(From Item In SelectedItems Select Item)
                If Items.Count = 0 Then Items.Add(ClickedItem) 'FullRowSelect, Must be Branch, Not Item!
                Dim Data As New DataObject
                Data.SetData(GetType(Object), Items)
                RaiseEvent DragStart(Items, New DragEventArgs(Nothing, 0, e.X, e.Y, DragDropEffects.Copy Or DragDropEffects.Move, DragDropEffects.All))
            ElseIf Dragging Then    '/// DRAG OVER ///
                Dim Data As New DataObject
                Data.SetData(GetType(Object), Nothing)
                MyBase.OnDragOver(New DragEventArgs(Data, 0, e.X, e.Y, DragDropEffects.Copy Or DragDropEffects.Move, DragDropEffects.All))
                DoDragDrop(Data, DragDropEffects.Copy Or DragDropEffects.Move)
            End If
        Else
            ClickedItem = Nothing
        End If
    End Sub
    Protected Overrides Sub OnMouseUp( e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        Dim Hit As HitInfo = HitTest(e)
        If Not IsNothing(Hit.Branch) AndAlso Not Hit.Region = "±" AndAlso Not Dragging Then
            If IsNothing(Hit.Item) Then
                Hit.Branch.Selected = Not (Hit.Branch.Selected)
            Else
                Hit.Item.Selected = Not (Hit.Item.Selected)
            End If
            RaiseEvent ItemClick(ClickedItem, e)
        End If
        Dragging = False
        ItemClicked = False
        ClickedItem = Nothing
        Refresh()
    End Sub
    Protected Overrides Sub OnGiveFeedback( e As GiveFeedbackEventArgs)
        MyBase.OnGiveFeedback(e)
        e.UseDefaultCursors = False
        If mCursor Then
            UpdateDragImage()
            mCursor = False
            Cursor.Current = New Cursor(DragImage.GetHicon())
        End If
    End Sub
    Protected Overrides Sub OnDragEnter( e As DragEventArgs)
        MyBase.OnDragEnter(e)
    End Sub
    Protected Overrides Sub OnDragLeave( e As System.EventArgs)
        MyBase.OnDragLeave(e)
        DropHighlight = Nothing
        Refresh()
    End Sub
    Protected Overrides Sub OnDragOver( e As DragEventArgs)
        e.Effect = DragDropEffects.Copy
        Dim X As Integer = PointToClient(New Point(e.X, e.Y)).X, Y As Integer = PointToClient(New Point(e.X, e.Y)).Y
        Dim Hit As HitInfo = HitTest(X, Y)
        Dim HitObject As Object = Nothing
        If IsNothing(Hit.Branch) Then
            DropHighlight = Nothing
        Else
            If IsNothing(Hit.Item) OrElse Not IsNothing(Hit.Item) AndAlso FullRowSelect Then
                HitObject = Hit.Branch
            Else
                HitObject = Hit.Item
            End If
        End If
        DropHighlight.Item = HitObject
        e.Data.SetData(GetType(Object), HitObject)
        Refresh()
    End Sub
    Protected Overrides Sub OnDragDrop( e As DragEventArgs)
        Dim X As Integer = PointToClient(New Point(e.X, e.Y)).X, Y As Integer = PointToClient(New Point(e.X, e.Y)).Y
        Dim Hit As HitInfo = HitTest(X, Y)
        Dim HitObject As Object = Nothing
        If Not IsNothing(Hit.Branch) Then
            If IsNothing(Hit.Item) Then
                HitObject = Hit.Branch
            Else
                HitObject = Hit.Item
            End If
        End If
        e.Data.SetData(GetType(Object), HitObject)
        DropHighlight.Item = Nothing
        Refresh()
        MyBase.OnDragDrop(e)
    End Sub
#End Region
#Region " Events "
    Protected Friend Sub OnCMSClick( sender As Object,  e As EventArgs)
        Dim TSMI As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If TSMI.Text.Contains("Expand") OrElse TSMI.Text.Contains("Collapse") Then
            For Each Branch As Branch In BranchCollection
                If TSMI.Text.Contains("Expand") Then
                    Branch.Expand()
                    RaiseEvent AfterExpand(Branch, e)
                ElseIf TSMI.Text.Contains("Collapse") Then
                    Branch.Collapse()
                    RaiseEvent AfterCollapse(Branch, e)
                End If
            Next Branch
        Else
            Dim DS As New DataSet
            For TableIndex As Integer = mDataSet.Tables.IndexOf(TSMI.Text) To mDataSet.Tables.Count - 1
                Dim DataTable As New DataTable
                DataTable = mDataSet.Tables(TableIndex).Copy
                Dim Q_PrimaryKey As IEnumerable(Of DataColumn) = (From A In mDataSet.Tables(TableIndex).PrimaryKey Join B In DataTable.Columns On DirectCast(A, DataColumn).ColumnName Equals DirectCast(B, DataColumn).ColumnName Select DirectCast(B, DataColumn))
                DataTable.PrimaryKey = Q_PrimaryKey.ToArray
                DS.Tables.Add(DataTable)
            Next
            Dim Relations As DataRelationCollection = mDataSet.Relations
            For Each Relation As DataRelation In Relations
                If Not IsNothing(DS.Tables(Relation.ParentTable.TableName)) Then
                    Dim TableName As String = Relation.ParentTable.TableName
                    Dim Q_Parent As IEnumerable(Of DataColumn) = (From A In Relation.ParentColumns Join B In DS.Tables(Relation.ParentTable.TableName).Columns On DirectCast(A, DataColumn).ColumnName Equals DirectCast(B, DataColumn).ColumnName Select DirectCast(B, DataColumn))
                    Dim Q_Child As IEnumerable(Of DataColumn) = (From A In Relation.ChildColumns Join B In DS.Tables(Relation.ChildTable.TableName).Columns On DirectCast(A, DataColumn).ColumnName Equals DirectCast(B, DataColumn).ColumnName Select DirectCast(B, DataColumn))
                    DS.Relations.Add(New DataRelation(Relation.RelationName, Q_Parent(0), Q_Child(0)))
                End If
            Next
            Clear()
            Headers.Clear()
            Load(DS)
        End If
    End Sub
    Public Event BranchesChanged( sender As Object,  e As EventArgs)
    Protected Friend Sub OnBranchesChanged( sender As Object,  e As EventArgs)
        Invalidate()
        RaiseEvent BranchesChanged(sender, e)
    End Sub
    Public Event SelectedBranchesChanged( sender As Object,  e As EventArgs)
    Protected Friend Sub OnSelectedBranchesChanged( sender As Object,  e As EventArgs)
        Invalidate()
        RaiseEvent SelectedBranchesChanged(sender, e)
    End Sub
    Public Event SelectedItemsChanged( sender As Object,  e As EventArgs)
    Protected Friend Sub OnSelectedItemsChanged( sender As Object,  e As EventArgs)
        RaiseEvent SelectedItemsChanged(sender, e)
    End Sub
    Public Event BeforeExpand( Branch As Branch,  e As CancelEventArgs)
    Public Event AfterExpand( Branch As Branch,  e As EventArgs)
    Public Event BeforeCollapse( Branch As Branch,  e As CancelEventArgs)
    Public Event AfterCollapse( Branch As Branch,  e As EventArgs)
    Public Event DragStart( sender As Object,  e As DragEventArgs)
    Public Event ItemClick( sender As Object,  e As EventArgs)
    Protected Friend Sub OnItemClick( sender As Object,  e As EventArgs)
        Invalidate()
        RaiseEvent ItemClick(sender, e)
    End Sub
#End Region
End Class
'///////////////////////////////////////////////////Branch COLLECTION///////////////////////////////////////////////////
#Region " Branch Classes "
<Serializable()> Public Class BranchCollection
    Inherits List(Of Branch)
    Public ReadOnly Property Tree As HeaderTreeView
    Public ReadOnly Property Parent As Branch

    Public Sub New( Parent As HeaderTreeView, Optional  Branch As Branch = Nothing)
        Tree = Parent
        _Parent = Branch
    End Sub

    Public ReadOnly Property Headers As New HeaderCollection
    Public ReadOnly Property All() As List(Of Branch)
        Get
            Dim Branches As New List(Of Branch)
            Dim Queue As New Queue(Of Branch)
            Dim TopBranch As Branch, Branch As Branch
            For Each TopBranch In Me
                Queue.Enqueue(TopBranch)
            Next
            While (Queue.Count > 0)
                TopBranch = Queue.Dequeue
                Branches.Add(TopBranch)
                For Each Branch In TopBranch.Branches
                    Queue.Enqueue(Branch)
                Next
            End While
            Return Branches
        End Get
    End Property
    Private SortType_ As SortType = SortType.NoSort
    Property SortType() As SortType
        Get
            Return SortType_
        End Get
        Set( Value As SortType)
            SortType_ = Value
            Sort()
            For Each Branch As Branch In Me
                Branch.Branches.SortType = SortType_
            Next
        End Set
    End Property
    Public Overloads Function Add(Name As String, Text As String, Image As Image) As Branch
        Return Add(New Branch With {.Name = Name, .Text = Text, .Image = Image})
    End Function
    Public Overloads Function Add(Name As String, Text As String) As Branch
        Return Add(New Branch With {.Name = Name, .Text = Text})
    End Function
    Public Overloads Function Add(Text As String, Image As Image) As Branch
        Return Add(New Branch With {.Text = Text, .Image = Image})
    End Function
    Public Overloads Function Add(Text As String) As Branch
        Return Add(New Branch With {.Text = Text})
    End Function
    Public Overloads Function Add( AddBranch As Branch) As Branch

        If AddBranch IsNot Nothing Then
            With AddBranch
                .Index_ = Count
                If Parent Is Nothing Then   ' *** ROOT Branch
                    'mTreeViewer was set when TreeViewer was created with New BranchCollection
                    .Tree_ = Tree
                    .Visible_ = True
                    .Level_ = 0

                Else                        ' *** CHILD Branch
                    REM /// Get TreeViewer value from Parent. A Branch Collection shares some Branch properties
                    _Tree = Parent.Tree
                    .Tree_ = Parent.Tree
                    .Parent_ = Parent
                    .Level_ = Parent.Level + 1
                    'If Parent.Expanded Then
                    '    .Visible_ = True
                    '    If Count = 0 Then
                    '        ._VisibleIndex = 0
                    '    Else
                    '        ._VisibleIndex = Last.VisibleIndex + 1
                    '    End If
                    'Else
                    '    ._VisibleIndex = -1
                    'End If

                End If
                If Tree IsNot Nothing Then
                    '.Font = Tree.Font
                    'Tree.BranchTimer_Start(AddBranch)
                End If
            End With
            MyBase.Add(AddBranch)
            OnChanged(Me, New EventArgs)
            OnBoundsChanged()
        End If
        Return AddBranch

    End Function
    Public Shadows Function AddRange( newBranches As Branch()) As List(Of Branch)

        If newBranches Is Nothing Then
            Return Nothing
        Else
            MyBase.AddRange(newBranches)
            Return newBranches.ToList
        End If

    End Function
    Public Shadows Function Find( Name As String) As List(Of Branch)

        Dim findBranches As New List(Of Branch)(From A In All Where Name = A.Name Select A)
        Return findBranches

    End Function
    Public Shadows Sub Insert(index As Integer,  Value As Branch)

        MyBase.Insert(index, Value)
        OnChanged(Me, New EventArgs)
        OnBoundsChanged()

    End Sub
    Public Shadows Sub Remove(removeValue As Branch)

        removeValue.Selected = False
        MyBase.Remove(removeValue)
        OnChanged(Me, New EventArgs)
        OnBoundsChanged()
        Tree.Invalidate()

    End Sub
    Public Shadows Sub Clear()

        MyBase.Clear()
        OnChanged(Me, New EventArgs)
        OnBoundsChanged()
        Tree.Invalidate()

    End Sub

    Private RelativeIndex As Integer = 0
    Private Sub GetRowIndex( Branches As BranchCollection)
        For Each N As Branch In Branches
            If N.Visible Then
                N.RowIndex = RelativeIndex
                RelativeIndex += 1
            End If
            If N.HasChildren Then GetRowIndex(N.Branches)
        Next
    End Sub
#Region " Events "
    Public Event Changed( sender As Object,  e As EventArgs)
    Protected Friend Sub OnChanged( sender As Object,  e As EventArgs)
        Sort()
        RaiseEvent Changed(sender, e)
    End Sub
    Private Sub OnBoundsChanged()
        RelativeIndex = 0
        GetRowIndex(Tree.Branches)
    End Sub
#End Region

    Public Shadows Sub Sort()

    End Sub

End Class
'///////////////////////////////////////////////////Branches///////////////////////////////////////////////////
<Serializable> Public Class Branch
    Public Sub New(Name As String, Text As String, Optional Image As Image = Nothing)

        mName = Name
        Text_ = Text
        mImage = Image
        Branches_ = New BranchCollection(Tree, Me)
        Items_ = New ItemCollection(Tree, Me)

    End Sub
    Public Sub New()
    End Sub
#Region " Properties & Fields "
    Friend Index_ As Integer
    Public ReadOnly Property Index As Integer
        Get
            Return Index_
        End Get
    End Property
    Private RelativeIndex As Integer = 0
    ReadOnly Property RootIndent As Integer
        Get
            If Tree.RootLines Then
                Return 10
            Else
                Return 4
            End If
        End Get
    End Property
    Private mRowIndex As Integer
    Public Property RowIndex As Integer
        Get
            Return mRowIndex
        End Get
        Set(value As Integer)
            mRowIndex = value
        End Set
    End Property
    Friend Tree_ As HeaderTreeView
    Public ReadOnly Property Tree As HeaderTreeView
        Get
            Return Tree_
        End Get
    End Property
    Private WithEvents Branches_ As New BranchCollection(Tree)
    Public ReadOnly Property Branches As BranchCollection
        Get
            Return Branches_
        End Get
    End Property
    Friend Parent_ As Branch
    Public ReadOnly Property Parent As Branch
        Get
            Return Parent_
        End Get
    End Property
    ReadOnly Property NextBranch() As Branch
        Get
            Dim ParentCollection As BranchCollection
            If IsNothing(Parent) Then
                ParentCollection = Tree.Branches
            Else
                ParentCollection = Parent.Branches
            End If
            If ParentCollection.Count = 1 + ParentCollection.IndexOf(Me) Then
                Return Nothing
            Else
                Return ParentCollection(1 + ParentCollection.IndexOf(Me))
            End If
        End Get
    End Property
    Private mName As String
    Property Name() As String
        Get
            Return mName
        End Get
        Set( Value As String)
            mName = Value
        End Set
    End Property
    Private Text_ As String
    Property Text() As String
        Get
            Return Text_
        End Get
        Set( Value As String)
            Text_ = Value
        End Set
    End Property
    Private mForeColor As Color = Color.Black
    Property ForeColor() As Color
        Get
            Return mForeColor
        End Get
        Set( Value As Color)
            mForeColor = Value
        End Set
    End Property
    Private mTag As Object
    Property Tag() As Object
        Get
            Return mTag
        End Get
        Set( Value As Object)
            mTag = Value
        End Set
    End Property
    Private mImage As Image
    Property Image() As Image
        Get
            Return mImage
        End Get
        Set( Value As Image)
            mImage = Value
        End Set
    End Property
    Friend Level_ As Integer
    ReadOnly Property Level As Integer
        Get
            Return Level_
        End Get
    End Property
    Private mSelected As Boolean = False
    Property Selected() As Boolean
        Get
            Return mSelected
        End Get
        Set( Value As Boolean)
            If mSelected <> Value Then
                Dim BranchBounds As Rectangle = Bounds
                If BranchBounds.Bottom >= Tree.Bounds.Bottom - If(Tree.HScroll.Visible, Tree.HScroll.Height, 0) Then Tree.VScroll.Value += 2 + BranchBounds.Height
                If BranchBounds.Top <= Tree.HeaderHeight Then
                    If RowIndex = 0 Then
                        Tree.VScroll.Value = 0
                    ElseIf Not Tree.VScroll.Value = 0 Then
                        If Tree.VScroll.Value >= 2 + BranchBounds.Height Then Tree.VScroll.Value -= 2 + BranchBounds.Height
                    End If
                End If
                mSelected = Value
                If mSelected Then
                    If Not Tree.MultiSelect OrElse Tree.MultiSelect AndAlso Not Tree.MultiSelecting Then
                        For Each Item As Object In (From A In Tree.Selection Select A.Key).ToList
                            If TypeOf (Item) Is Branch Then DirectCast(Item, Branch).Selected = False
                            If TypeOf (Item) Is Item Then DirectCast(Item, Item).Selected = False
                        Next
                    End If
                    Tree.SelectedBranches.Add(Me)
                    Tree.Selection.Add(Me, Me)
                Else
                    Tree.SelectedBranches.Remove(Me)
                    Tree.Selection.Remove(Me)
                End If
            End If
        End Set
    End Property
    Friend Visible_ As Boolean = False
    ReadOnly Property Visible As Boolean
        Get
            Dim VisibleFlag As Boolean
            If IsNothing(Parent) Then
                VisibleFlag = True
                REM /// Root Branches are visible provided they fit in the ClientRectangle
            Else
                Dim Parent As Branch = Me
                Dim HiddenFlag As Boolean = False
                Do While Not IsNothing(Parent.Parent)
                    Parent = Parent.Parent
                    If Parent.State = Branchestate.Collapsed Then
                        HiddenFlag = True
                        REM /// If any one of the Branch's parents are collapsed then the Branch is not visible
                        Exit Do
                    End If
                Loop
                VisibleFlag = Not HiddenFlag
            End If
            Return VisibleFlag
            Return Visible_
            REM /// Test both cases to see if in ClientRectangle by looking at Bounds
        End Get
    End Property
    Private mState As Branchestate = Branchestate.Collapsed
    ReadOnly Property State() As Branchestate
        Get
            If HasChildren Then
                Return mState
            Else
                Return Branchestate.Collapsed
            End If
        End Get
    End Property
    ReadOnly Property Bounds() As Rectangle
        Get
            If Visible Then
                Dim X As Integer = BoxBounds.Right - If(IsNothing(Parent) AndAlso Not HasChildren, Tree.ExpandedIcon.Width, 0)
                Dim Y As Integer = Tree.HeaderHeight + Tree.RowHeight * RowIndex - Tree.VScroll.Value
                Dim W As Integer = Tree.Headers(0).Width - X
                Dim H As Integer = Tree.RowHeight
                Return New Rectangle(X, Y, W, H)
            Else
                Return Nothing
            End If
        End Get
    End Property
    ReadOnly Property BoxBounds() As Rectangle
        Get
            If Visible Then
                Dim X As Integer = RootIndent + Level * (Tree.Indent + 8) - Tree.HScroll.Value      'Tree.ExpandedIcon.Width
                Dim Y As Integer = Tree.HeaderHeight + Tree.RowHeight * RowIndex - Tree.VScroll.Value
                Dim W As Integer = Tree.ExpandedIcon.Width
                Dim H As Integer = Tree.ExpandedIcon.Height
                Return New Rectangle(X, Y, W, H)
            Else
                Return Nothing
            End If
        End Get
    End Property
    Private WithEvents Items_ As ItemCollection
    ReadOnly Property Items() As ItemCollection
        Get
            Return Items_
        End Get
    End Property
    ReadOnly Property HasChildren() As Boolean
        Get
            Return Branches.Any
        End Get
    End Property
#End Region
#Region " Methods "
    Public Sub Expand()
        mState = Branchestate.Expanded
        RelativeIndex = 0
        GetRowIndex(Tree.Branches)
        Tree.Invalidate()
    End Sub
    Public Sub Collapse()
        mState = Branchestate.Collapsed
        RelativeIndex = 0
        GetRowIndex(Tree.Branches)
        Tree.Invalidate()
    End Sub
    Public Sub Remove()
        Parent_.Branches.Remove(Me)
    End Sub
#End Region
    Private Sub GetRowIndex( Branches As BranchCollection)
        For Each N As Branch In Branches
            If N.Visible Then
                N.RowIndex = RelativeIndex
                RelativeIndex += 1
            End If
            If N.HasChildren Then GetRowIndex(N.Branches)
        Next
    End Sub
#Region " Events "
    Public Event Changed( sender As Object,  e As EventArgs)
    Protected Friend Sub OnChanged( sender As Object,  e As EventArgs)
        RaiseEvent Changed(sender, e)
    End Sub
#End Region
End Class
#End Region
'///////////////////////////////////////////////////ITEMS///////////////////////////////////////////////////
#Region " Item Classes "
<Serializable()> Public Class ItemCollection
    Inherits CollectionBase
#Region " Constructor "
    Public Sub New( Tree As HeaderTreeView, Optional  Branch As Branch = Nothing, Optional  ImageList As ImageList = Nothing)
        Tree_ = Tree
        Owner_ = Branch
        ImageList_ = ImageList
    End Sub
#End Region
#Region " Properties & Fields "
    Private ReadOnly Tree_ As HeaderTreeView
    ReadOnly Property Tree() As HeaderTreeView
        Get
            Return Tree_
        End Get
    End Property
    Private ReadOnly Owner_ As Branch
    ReadOnly Property Owner() As Branch
        Get
            Return Owner_
        End Get
    End Property
    Private ReadOnly ImageList_ As ImageList
    ReadOnly Property ImageList() As ImageList
        Get
            Return ImageList_
        End Get
    End Property
    Default Property Item( Index As Integer) As Item
        Get
            Return CType(List(Index), Item)
        End Get
        Set( Value As Item)
            List(Index) = Value
        End Set
    End Property
#End Region
#Region " Methods "
    Public Function Add( Value As Item) As Integer
        Dim Result As Integer
        Result = List.Add(Value)
        'AddHandler Value.MouseDown, AddressOf OnMouseDown
        OnChanged(Me, New EventArgs)
        Return Result
    End Function
    Public Function Add( Text As String, Optional  ForeColor As Color = Nothing, Optional  Image As Image = Nothing) As Item
        Dim Result As New Item(Owner) With {
            .Text = Text,
            .ForeColor = ForeColor,
            .Image = Image
        }
        Add(Result)
        Return Result
    End Function
    Public Function IndexOf( Value As Item) As Integer
        Return List.IndexOf(Value)
    End Function
    Public Sub Insert( index As Integer,  Value As Item)
        List.Insert(index, Value)
        OnChanged(Me, New EventArgs)
    End Sub
    Public Sub Remove( Value As Item)
        List.Remove(Value)
        OnChanged(Me, New EventArgs)
    End Sub
    Public Function Contains( Value As Item) As Boolean
        Return List.Contains(Value)
    End Function
    Public Shadows Sub Clear()
        MyBase.Clear()
        OnChanged(Me, New EventArgs)
    End Sub
#End Region
#Region " Events "
    Public Event Changed( sender As Object,  e As EventArgs)
    Protected Friend Sub OnChanged( sender As Object,  e As EventArgs)
        RaiseEvent Changed(sender, e)
    End Sub
#End Region
End Class
Public Class Item
#Region " Constructor "
    Public Sub New(Optional Branch As Branch = Nothing)
        Tree_ = Branch.Tree
        Parent_ = Branch
    End Sub
#End Region
#Region " Properties & Fields "
    Private ReadOnly Tree_ As HeaderTreeView
    ReadOnly Property Tree() As HeaderTreeView
        Get
            Return Tree_
        End Get
    End Property
    Private ReadOnly Parent_ As Branch
    ReadOnly Property Parent() As Branch
        Get
            Return Parent_
        End Get
    End Property
    Private mImage As Image
    Property Image() As Image
        Get
            Return mImage
        End Get
        Set( Value As Image)
            mImage = Value
        End Set
    End Property
    Private Text_ As String
    Property Text() As String
        Get
            Return Text_
        End Get
        Set( Value As String)
            Text_ = Value
        End Set
    End Property
    Private mAlignment As HorizontalAlignment
    Property Alignment() As HorizontalAlignment
        Get
            Return (mAlignment)
        End Get
        Set( Value As HorizontalAlignment)
            mAlignment = Value
        End Set
    End Property
    Private mForeColor As Color = Color.Black
    Property ForeColor() As Color
        Get
            Return mForeColor
        End Get
        Set( Value As Color)
            mForeColor = Value
        End Set
    End Property
    Private mTag As Object
    Property Tag() As Object
        Get
            Return mTag
        End Get
        Set( Value As Object)
            mTag = Value
        End Set
    End Property
    Private mSelected As Boolean = False
    Property Selected() As Boolean
        Get
            Return mSelected
        End Get
        Set( Value As Boolean)
            If Tree.FullRowSelect Then
                Parent.Selected = Not (Parent.Selected)
            Else
                If mSelected <> Value Then
                    Dim ItemBounds As Rectangle = Bounds
                    If ItemBounds.Bottom >= Tree.Bounds.Bottom - If(Tree.HScroll.Visible, Tree.HScroll.Height, 0) Then Tree.VScroll.Value += 2 + ItemBounds.Height
                    If ItemBounds.Top <= Tree.HeaderHeight Then
                        If Parent.RowIndex = 0 Then
                            Tree.VScroll.Value = 0
                        Else
                            Tree.VScroll.Value -= 2 + ItemBounds.Height
                        End If
                    End If
                    mSelected = Value
                    If mSelected Then
                        If Not Tree.MultiSelect OrElse Tree.MultiSelect AndAlso Not Tree.MultiSelecting Then
                            For Each Item As Object In (From A In Tree.Selection Select A.Key).ToList
                                If TypeOf (Item) Is Branch Then DirectCast(Item, Branch).Selected = False
                                If TypeOf (Item) Is Item Then DirectCast(Item, Item).Selected = False
                            Next
                        End If
                        Tree.SelectedItems.Add(Me)
                        Tree.Selection.Add(Me, Me)
                    Else
                        Tree.SelectedItems.Remove(Me)
                        Tree.Selection.Remove(Me)
                    End If
                End If
            End If
        End Set
    End Property
    ReadOnly Property Bounds() As Rectangle
        Get
            If Parent.Visible Then
                Dim X As Integer = Tree.Headers(1 + Parent.Items.IndexOf(Me)).Bounds.Left
                Dim Y As Integer = Parent.Bounds.Y
                Dim W As Integer = Tree.Headers(1 + Parent.Items.IndexOf(Me)).Width
                Dim H As Integer = Parent.Bounds.Height
                Return New Rectangle(X, Y, W, H)
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region
#Region " Events "

#End Region
End Class
#End Region
#Region " Enumerations "
Public Enum Branchestate
    Expanded
    Collapsed
End Enum
Public Enum SortType
    TextAscending
    TextDescending
    IndexAscending
    IndexDescending
    NoSort
End Enum
Public Enum TreeExpanderStyle
    PlusMinus = 0
    Arrow = 1
    Book = 2
    LightBulb = 4
End Enum
Public Enum SelectStyle
    Box = 0
    Highlight = 1
End Enum
#End Region

Public MustInherit Class Headers
    Inherits Control
#Region " Constructor "
    Friend WithEvents VScroll As New VScrollBar
    Friend WithEvents HScroll As New HScrollBar
    Public Sub New()
        VScroll.Minimum = 0
        VScroll.SmallChange = 1
        VScroll.Parent = Me
        VScroll.Hide()
        HScroll.Minimum = 0
        HScroll.SmallChange = 1
        HScroll.Parent = Me
        HScroll.Hide()
        AddHandler VScroll.ValueChanged, AddressOf OnScrollChange
        AddHandler HScroll.ValueChanged, AddressOf OnScrollChange
        AddHandler Headers_.Changed, AddressOf OnHeadersChanged
        AddHandler Headers_.Clicked, AddressOf OnHeadersClicked
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)
        BackColor = SystemColors.Window
    End Sub
#End Region
#Region " Properties & Fields "
    Private AllowPaint As Boolean = True
    Private Property HeadersVisible_ As Boolean = True
    Property HeadersVisible() As Boolean
        Get
            Return HeadersVisible_
        End Get
        Set( Value As Boolean)
            HeadersVisible_ = Value
        End Set
    End Property
    Public ReadOnly Property HeaderHeight As Integer
        Get
            HeaderHeight = Convert.ToInt32(Font.GetHeight) + 6
            Dim Q_Image As IEnumerable(Of Int32) = (From A In (From I In Headers Select CType(I, Header)) Where Not IsNothing(A.Image) Select A.Image.Height)
            If Not Q_Image.Count = 0 Then
                If Q_Image.Max + 6 > HeaderHeight Then HeaderHeight = Q_Image.Max + 6
            End If
            Return If(HeadersVisible_, HeaderHeight, 1)
        End Get
    End Property
    Private mSelectedColor As Color = Color.WhiteSmoke
    Property SelectedColor() As Color
        Get
            Return mSelectedColor
        End Get
        Set( Value As Color)
            If (Not (mSelectedColor.Equals(Value))) Then
                mSelectedColor = Value
                Invalidate()
            End If
        End Set
    End Property
    Private mBackColor As Color = SystemColors.Window
    <DefaultValue(GetType(Color), "Window")> Shadows Property BackColor() As Color
        Get
            Return mBackColor
        End Get
        Set( Value As Color)
            If (Not (mBackColor.Equals(Value))) Then
                mBackColor = Value
                Invalidate()
            End If
        End Set
    End Property
    Private mHeaderForeColor As Color = Color.Black
    Property HeaderForeColor() As Color
        Get
            Return mHeaderForeColor
        End Get
        Set( Value As Color)
            If (Not (mHeaderForeColor.Equals(Value))) Then
                mHeaderForeColor = Value
                Invalidate()
            End If
        End Set
    End Property
    Private mHeaderBackColor As Color = SystemColors.Control
    Property HeaderBackColor() As Color
        Get
            Return mHeaderBackColor
        End Get
        Set( Value As Color)
            If (Not (mHeaderBackColor.Equals(Value))) Then
                mHeaderBackColor = Value
                Invalidate()
            End If
        End Set
    End Property
    Private mHeaderHatchColor As Color = SystemColors.Control
    Property HeaderHatchColor() As Color
        Get
            Return mHeaderHatchColor
        End Get
        Set( Value As Color)
            If (Not (mHeaderHatchColor.Equals(Value))) Then
                mHeaderHatchColor = Value
                Invalidate()
            End If
        End Set
    End Property
    Private mHeaderHatchStyle As HatchStyle = HatchStyle.ZigZag
    Property HeaderHatchStyle() As HatchStyle
        Get
            Return mHeaderHatchStyle
        End Get
        Set( Value As HatchStyle)
            If (Not (mHeaderHatchStyle.Equals(Value))) Then
                mHeaderHatchStyle = Value
                Invalidate()
            End If
        End Set
    End Property
    Private WithEvents Headers_ As New HeaderCollection
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> ReadOnly Property Headers() As HeaderCollection
        Get
            Return Headers_
        End Get
    End Property
    Private ReadOnly Property TotalWidth() As Integer
        Get
            TotalWidth = 2
            For Each Header As Header In Headers_
                TotalWidth += Header.Width
            Next
            If VScroll.Visible Then TotalWidth += VScroll.Width
            Return TotalWidth
        End Get
    End Property
    Private mTotalHeight As Integer
    Protected Property TotalHeight() As Integer
        Get
            Return mTotalHeight
        End Get
        Set( Value As Integer)
            mTotalHeight = Value
        End Set
    End Property
#End Region
#Region " Events "
    Public Event ScrollChange( sender As Object,  e As EventArgs)
    Protected Friend Sub OnScrollChange( sender As Object,  e As EventArgs)
        Invalidate()
        RaiseEvent ScrollChange(sender, e)
        OnHeadersChanged(sender, e)
    End Sub
    Public Event HeadersChanged( sender As Object,  e As EventArgs)
    Protected Friend Sub OnHeadersChanged( sender As Object,  e As EventArgs)
        Invalidate()
        RaiseEvent HeadersChanged(sender, e)
        Dim PrevColWidths As Integer = 1
        For Each Header As Header In Headers
            Header.Bounds = New Rectangle(PrevColWidths - HScroll.Value, 1, Header.Width, HeaderHeight)
            PrevColWidths += Header.Width
        Next
    End Sub
    Public Event HeadersClicked( sender As Header)
    Protected Friend Sub OnHeadersClicked( sender As Header)
        Dim CTvw As HeaderTreeView = CType(Me, HeaderTreeView)
        If CTvw.SortType = SortType.IndexAscending Then
            CTvw.SortType = SortType.IndexDescending
        ElseIf CTvw.SortType = SortType.IndexDescending Then
            CTvw.SortType = SortType.IndexAscending
        ElseIf CTvw.SortType = SortType.TextDescending Then
            CTvw.SortType = SortType.TextAscending
        ElseIf CTvw.SortType = SortType.TextAscending Then
            CTvw.SortType = SortType.TextDescending
        End If
        Invalidate()
        RaiseEvent HeadersClicked(sender)
    End Sub
#End Region
#Region " Drawing "
    Public Sub BeginUpdate()
        AllowPaint = False
    End Sub
    Public Sub EndUpdate()
        AllowPaint = True
        Invalidate()
    End Sub
    Protected Overrides Sub OnPaint( e As PaintEventArgs)
        If Not AllowPaint Then Exit Sub
        Dim Graphics As Graphics = e.Graphics
        Dim Rectangle As Rectangle = ClientRectangle
        DrawBackground(Graphics, Rectangle)
        DrawItems(Graphics, Rectangle)
        SetupScrolls(Rectangle, mTotalHeight + HeaderHeight + 2)
        DrawHeaders(Graphics, Rectangle)
        DrawBorder(Graphics, Rectangle)
    End Sub
    Private Sub SetupScrolls( r As Rectangle,  TotalHeight As Integer)
        Dim HVis As Boolean = HScroll.Visible
        Dim VVis As Boolean = VScroll.Visible
        If HScroll.Visible Then
            TotalHeight += HScroll.Height
        End If
        If TotalHeight > r.Height Then
            VScroll.Top = r.Top + 2
            VScroll.Left = r.Right - 2 - VScroll.Width
            VScroll.Maximum = TotalHeight
            VScroll.LargeChange = r.Height
            If HScroll.Visible Then
                VScroll.Height = r.Height - 4 - HScroll.Height
            Else
                VScroll.Height = r.Height - 4
            End If
            If VScroll.Value > (TotalHeight - r.Height) Then
                VScroll.Value = (TotalHeight - r.Height)
            End If
            VScroll.Show()
        Else
            VScroll.Hide()
            VScroll.Value = 0
        End If
        If TotalWidth > r.Width Then
            HScroll.Top = r.Bottom - 2 - HScroll.Height
            HScroll.Left = r.Left + 2
            HScroll.Maximum = TotalWidth
            HScroll.LargeChange = r.Width
            If VScroll.Visible Then
                HScroll.Width = r.Width - 4 - VScroll.Width
            Else
                HScroll.Width = r.Width - 4
            End If
            If HScroll.Value > (TotalWidth - r.Width) Then
                HScroll.Value = (TotalWidth - r.Width)
            End If
            HScroll.Show()
        Else
            HScroll.Hide()
            HScroll.Value = 0
        End If
        If (HVis <> HScroll.Visible) OrElse (VVis <> VScroll.Visible) Then
            Invalidate()
        End If
    End Sub
    Private Sub DrawBackground( g As Graphics,  r As Rectangle)
        g.FillRectangle(New SolidBrush(mBackColor), r)
    End Sub
    Protected MustOverride Sub DrawItems( g As Graphics,  r As Rectangle)
    Private Sub DrawHeaders( g As Graphics,  r As Rectangle)
        If HeadersVisible Then
            Dim CTvw As HeaderTreeView = DirectCast(Me, HeaderTreeView)
            Dim Ascending As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAM9JREFUOE/N0kkKg0AQBdCf+x/KlaALxXkeF04o4g0qXU1sDPQqBhLhI72o17+gH0SEWx8Dd3JrWLa/c/t3Ac/ziOM4Dtm2LWNZFpmmKWMYhsq1tVphmiYw0Pc9tW1LVVVRnueUpilFUURBEEjAdd23tdVhWRZckaZpqCxLiSRJIocFwpfogW3bwMg4juDqXdfRifCwQCCawPd9PbDvO9Z1VQgPMcJ/0QRZloGRMAz1wHEcOJF5njEMA14I6rpGURQSieNYD3z6Hv7oIf1shSf3G9UMQ+Vu/QAAAABJRU5ErkJggg=="
            Dim Descending As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAHtJREFUOE/FkU0KgCAUBu3+h3IV2EIQEUUUf05iL40i08BFOOsZ/PAtKSU0BARDoCH7mDMpwBj3Xm5MAjuE0GvqgBACtnPOGNNsHgFjLMZYbK21lPLd3EGxvffWWrCVUkIIznnVnAHYAKV0y8CwNQN24fqDWXf4OP//k3bpJLwlth796QAAAABJRU5ErkJggg=="
            Dim Sort As String = Ascending
            Dim AllHeadersWidth As Integer = 1
            For Each Header As Header In Headers
                Dim HeaderImageWidth As Integer = 0, HeaderImageHeight As Integer = 0
                Using HatchBrush As New HatchBrush(mHeaderHatchStyle, mHeaderHatchColor, mHeaderBackColor)
                    g.FillRectangle(HatchBrush, Header.Bounds)
                End Using
                If Not IsNothing(Header.Image) Then
                    HeaderImageWidth = Header.Image.Width
                    HeaderImageHeight = Header.Image.Height
                    g.DrawImage(Header.Image, New Rectangle(Header.Bounds.Left + 3, Header.Bounds.Top + 2, HeaderImageWidth, HeaderImageHeight))
                End If
                TextRenderer.DrawText(g, Header.Text, Font, New Point(Header.Bounds.Left + 3 + HeaderImageWidth, Header.Bounds.Top + Convert.ToInt32((HeaderHeight - Font.GetHeight) / 2) - 1), mHeaderForeColor)
                ControlPaint.DrawBorder3D(g, Header.Bounds, CTvw.HeaderStyle)
                If Not CTvw.SortType = SortType.NoSort OrElse Not HeadersVisible Then
                    If CTvw.SortType.ToString.Contains("Descending") Then Sort = Descending
                    Using ms As New MemoryStream(Convert.FromBase64String(Sort))
                        Using SortArrow As Bitmap = DirectCast(Bitmap.FromStream(ms), Bitmap)
                            SortArrow.MakeTransparent(Color.White)
                            g.DrawImageUnscaled(SortArrow, New Rectangle(Header.Bounds.Right - SortArrow.Width - 6, Header.Bounds.Top + 2, SortArrow.Width, SortArrow.Height))
                        End Using
                    End Using
                End If
                AllHeadersWidth += Header.Width
            Next
            If AllHeadersWidth < r.Width Then
                Dim PadColRect As Rectangle
                PadColRect = New Rectangle(AllHeadersWidth, 1, (r.Width - AllHeadersWidth) - 1, HeaderHeight)
                Using HatchBrush As New HatchBrush(mHeaderHatchStyle, mHeaderHatchColor, mHeaderBackColor)
                    g.FillRectangle(HatchBrush, PadColRect)
                End Using
                ControlPaint.DrawBorder3D(g, PadColRect, CTvw.HeaderStyle)
            End If
        End If
    End Sub
    Private Sub DrawBorder( g As Graphics,  r As Rectangle)
        ControlPaint.DrawBorder3D(g, r, Border3DStyle.Sunken)
        If VScroll.Visible AndAlso HScroll.Visible Then g.FillRectangle(SystemBrushes.Control, VScroll.Left, HScroll.Top, VScroll.Width, HScroll.Height)
    End Sub
#End Region
#Region " Helper Methods "

#End Region
#Region " Overrides "
    Private TheHeader As Header
    Private ColSizeHover As Boolean = False
    Private ColStartDrag As Integer
    Private OrigColWidth As Integer
    Private ReadOnly ToolTip As New ToolTip
    Protected Overrides Sub OnMouseMove( e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If e.Button = MouseButtons.Left AndAlso ColSizeHover AndAlso Not IsNothing(TheHeader) Then
            If e.X < TheHeader.Bounds.Left + 4 Then Exit Sub
            TheHeader.Width = OrigColWidth + (e.X - ColStartDrag)
            ToolTip.SetToolTip(Me, TheHeader.Text & " Width=" & TheHeader.Width.ToString)
            MyBase.Refresh()
            Invalidate()
        ElseIf HeadersVisible Then
            For Each Header As Header In Headers
                If ((e.X >= Header.Bounds.Right - 4) AndAlso (e.X <= Header.Bounds.Right + 4)) AndAlso (e.Y <= HeaderHeight) AndAlso Header.Resizeable Then
                    Cursor = Cursors.VSplit
                    If Header.Resizeable Then
                        ColSizeHover = True
                    End If
                    TheHeader = Header
                    Exit For
                Else
                    Cursor = Cursors.Default
                    ColSizeHover = False
                End If
            Next
        End If
    End Sub
    Protected Overrides Sub OnMouseDown( e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If HeadersVisible Then
            If e.Y <= HeaderHeight Then
                If ColSizeHover Then
                    ColStartDrag = e.X
                    OrigColWidth = TheHeader.Width
                Else
                    Dim ClickedHeader As IEnumerable(Of Header) = (From A In Headers Where e.X >= DirectCast(A, Header).Bounds.Left _
                                                   And e.X <= DirectCast(A, Header).Bounds.Right Select DirectCast(A, Header))
                    If ClickedHeader.Count = 0 Then
                        OnHeadersClicked(Nothing)
                    Else
                        OnHeadersClicked(ClickedHeader.First)
                    End If
                End If
            End If
        End If
    End Sub
    Protected Overrides Sub OnMouseUp( e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        If e.Button = MouseButtons.Left Then
            ' ColSizing = e.X
            ToolTip.SetToolTip(Me, String.Empty)
            ToolTip.Hide(Me)
        End If
    End Sub
    Protected Overrides Sub OnFontChanged( e As System.EventArgs)
        MyBase.OnFontChanged(e)
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseWheel( e As MouseEventArgs)
        MyBase.OnMouseWheel(e)
        If VScroll.Visible Then
            If ((VScroll.Value + -e.Delta) <= (VScroll.Maximum - VScroll.LargeChange)) AndAlso
            (VScroll.Value + -e.Delta >= VScroll.Minimum) Then
                VScroll.Value += -e.Delta
                Invalidate()
            Else
                If -e.Delta < 0 Then
                    VScroll.Value = VScroll.Minimum
                    Invalidate()
                Else
                    VScroll.Value = (VScroll.Maximum - VScroll.LargeChange)
                    Invalidate()
                End If
            End If
        End If
    End Sub
#End Region
End Class
'///////////////////////////////////////////////////Header COLLECTION///////////////////////////////////////////////////
#Region " Header Classes "
<Serializable()> Public Class HeaderCollection
    Inherits List(Of Header)
    Public Shadows Function AddRange( Headers As Header()) As Header()

        For Each Header In Headers
            Add(Header)
        Next
        Return Headers

    End Function
    Public Shadows Function Add(addHeader As Header) As Header

        AddHandler addHeader.Changed, AddressOf OnChanged
        MyBase.Add(addHeader)
        OnChanged(Me, New EventArgs)
        Return addHeader

    End Function
    Public Shadows Function Add( Text As String) As Header

        Dim addHeader As New Header(Text) With {
            .Width = 100
        }
        MyBase.Add(addHeader)
        Return addHeader

    End Function
    Public Shadows Function Add( Text As String,  Width As Integer) As Header

        Dim Header As New Header(Text) With {
            .Text = Text,
            .Width = Width
        }
        MyBase.Add(Header)
        Return Header

    End Function
    Public Shadows Sub Insert( index As Integer,  insertHeader As Header)

        MyBase.Insert(index, insertHeader)
        OnChanged(Me, New EventArgs)

    End Sub
    Public Shadows Sub Remove( removeHeader As Header)

        MyBase.Remove(removeHeader)
        OnChanged(Me, New EventArgs)

    End Sub
    Public Event Changed( sender As Object,  e As EventArgs)
    Protected Friend Sub OnChanged( sender As Object,  e As EventArgs)
        RaiseEvent Changed(sender, e)
    End Sub
    Public Event Clicked( sender As Header)
    Protected Friend Sub OnClicked( sender As Header)
        RaiseEvent Clicked(sender)
    End Sub
End Class
'///////////////////////////////////////////////////HeaderS///////////////////////////////////////////////////
<Serializable> Public Class Header
    Public Sub New(headerText As String)
        Text = headerText
    End Sub
    Private mResizeable As Boolean = True
    Property Resizeable As Boolean
        Get
            Return mResizeable
        End Get
        Set( Value As Boolean)
            mResizeable = Value
        End Set
    End Property
    Private ForeColor_ As Color = Color.Black
    Property ForeColor() As Color
        Get
            Return ForeColor_
        End Get
        Set( Value As Color)
            If ForeColor_ <> Value Then
                ForeColor_ = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private Text_ As String = "Header"
    Property Text() As String
        Get
            Return Text_
        End Get
        Set( Value As String)
            If Text_ <> Value Then
                Text_ = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private mAlignment As HorizontalAlignment = HorizontalAlignment.Left
    Property Alignment As HorizontalAlignment
        Get
            Return mAlignment
        End Get
        Set( Value As HorizontalAlignment)
            If mAlignment <> Value Then
                mAlignment = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private mWidth As Integer = 100
    Property Width() As Integer
        Get
            Return mWidth
        End Get
        Set( Value As Integer)
            If mWidth <> Value Then
                mWidth = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private mImage As Image = Nothing
    Property Image() As Image
        Get
            Return mImage
        End Get
        Set( Value As Image)
            If mImage IsNot Value Then
                mImage = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private mBounds As Rectangle
    Property Bounds() As Rectangle
        Get
            Return mBounds
        End Get
        Set( Value As Rectangle)
            If mBounds <> Value Then
                mBounds = Value
                OnChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Public Event Changed( sender As Object,  e As EventArgs)
    Protected Friend Sub OnChanged( sender As Object,  e As EventArgs)
        RaiseEvent Changed(sender, e)
    End Sub
    Public Event Clicked( sender As Header)
    Protected Friend Sub OnClicked( sender As Header)
        RaiseEvent Clicked(sender)
    End Sub
End Class
#End Region