﻿
Imports System.Collections.ObjectModel

Public Class Tile
    Inherits Border

    Private mousedownpoint As Point?
    Private adorner As DragAdorner
    Private window As Window

    Sub New()
        MyBase.New()
        AllowDrop = True
    End Sub

#Region "Events"

    Private Sub Tile_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.PreviewMouseLeftButtonDown
        mousedownpoint = e.GetPosition(sender)
    End Sub

    Private Sub Tile_PreviewMouseMove(sender As Object, e As MouseEventArgs) Handles Me.PreviewMouseMove
        Dim mousemovepoint = e.GetPosition(sender)
        If Not mousedownpoint.HasValue OrElse e.LeftButton = MouseButtonState.Released OrElse
        (Point.Subtract(mousedownpoint.Value, mousemovepoint).Length < SystemParameters.MinimumHorizontalDragDistance And
         Point.Subtract(mousedownpoint.Value, mousemovepoint).Length < SystemParameters.MinimumVerticalDragDistance) Then Exit Sub

        DoDragDrop(mousemovepoint)
    End Sub

    Private Sub Window_PreviewDragOver(sender As Object, e As DragEventArgs)
        If adorner Is Nothing OrElse window Is Nothing Then Exit Sub
        adorner.LeftOffset = e.GetPosition(window).X
        adorner.TopOffset = e.GetPosition(window).Y
    End Sub

    Private Sub Tile_PreviewDragOver(sender As Object, e As DragEventArgs) Handles Me.PreviewDragOver
        If Not e.Data.GetDataPresent(Me.GetType) Then Exit Sub
        Dim sourceelement As Tile = e.Data.GetData(Me.GetType) : If sourceelement Is Nothing Then Exit Sub
        Dim targetelement As Tile = sender
        If sourceelement Is targetelement Then Exit Sub
        Dim sourcedata As Object = sourceelement.DataContext : If sourcedata Is Nothing Then Exit Sub
        Dim targetdata As Object = targetelement.DataContext : If targetdata Is Nothing Then Exit Sub
        Dim panel = FindVisualParent(Of Panel)(Me) : If panel Is Nothing Then Exit Sub
        Dim itemscontrol = FindVisualParent(Of ItemsControl)(panel) : If itemscontrol Is Nothing Then Exit Sub

        Dim itemssource = itemscontrol.ItemsSource
        If itemssource IsNot Nothing Then
            Dim sourcetype As Type = itemssource.GetType
            If Not sourcetype.IsGenericType OrElse Not (sourcetype.GetGenericTypeDefinition = GetType(ObservableCollection(Of))) Then Exit Sub
            Dim collection As Object = itemssource ' ObservableCollection of sometype
            Dim sourcedataindex As Integer = collection.IndexOf(sourcedata) : If sourcedataindex = -1 Then Exit Sub
            Dim targetdataindex As Integer = collection.IndexOf(targetdata) : If targetdataindex = -1 Then Exit Sub
            collection.Move(sourcedataindex, targetdataindex)
        End If
    End Sub

    Private Sub Tile_PreviewDrop(sender As Object, e As DragEventArgs) Handles Me.PreviewDrop
        If Not e.Data.GetDataPresent(Me.GetType) Then Exit Sub
        Dim sourceelement As Tile = e.Data.GetData(Me.GetType) : If sourceelement Is Nothing Then Exit Sub
        Dim targetelement As Tile = sender
        If sourceelement Is targetelement Then Exit Sub
        Dim sourcedata As Object = sourceelement.DataContext : If sourcedata Is Nothing Then Exit Sub
        Dim targetdata As Object = targetelement.DataContext : If targetdata Is Nothing Then Exit Sub
        Dim panel = FindVisualParent(Of Panel)(Me) : If panel Is Nothing Then Exit Sub
        Dim itemscontrol = FindVisualParent(Of ItemsControl)(panel) : If itemscontrol Is Nothing Then Exit Sub

        Dim items = itemscontrol.Items
        If items IsNot Nothing Then
            Dim sourcedataindex As Integer = items.IndexOf(sourcedata) : If sourcedataindex = -1 Then Exit Sub
            Dim targetdataindex As Integer = items.IndexOf(targetdata) : If targetdataindex = -1 Then Exit Sub
            items.RemoveAt(sourcedataindex)
            items.Insert(targetdataindex, sourcedata)
            itemscontrol.UpdateLayout()
        End If
    End Sub

#End Region

#Region "Subs"

    Private Sub DoDragDrop(startpoint As Point)
        window = Window.GetWindow(Me) : If window Is Nothing Then Exit Sub
        Dim lastallowdrop As Boolean = window.AllowDrop : window.AllowDrop = True
        AddHandler window.PreviewDragOver, AddressOf Window_PreviewDragOver
        adorner = New DragAdorner(Me, Me, startpoint) : AdornerLayer.GetAdornerLayer(window.Content).Add(adorner)
        Dim lastopacity = Opacity : Opacity = 0.1
        Dim dragdata As New DataObject(Me) : DragDrop.DoDragDrop(Me, dragdata, DragDropEffects.Move)
        Opacity = lastopacity
        AdornerLayer.GetAdornerLayer(window.Content).Remove(adorner) : adorner = Nothing
        RemoveHandler window.PreviewDragOver, AddressOf Window_PreviewDragOver
        window.AllowDrop = lastallowdrop
    End Sub

    Public Function FindVisualParent(Of T As DependencyObject)(ByVal child As Object, Optional until As DependencyObject = Nothing) As T
        Dim parent As DependencyObject = If(child.Parent IsNot Nothing, child.Parent, VisualTreeHelper.GetParent(child))

        If parent IsNot Nothing Then
            If TypeOf parent Is T Then
                Return parent
            ElseIf parent Is until Then
                Return Nothing
            Else
                Return FindVisualParent(Of T)(parent, until)
            End If
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region "DragAdorner"
    Private NotInheritable Class DragAdorner
        Inherits Adorner

        Private _child As Rectangle
        Private _leftoffset As Double
        Private _topoffset As Double
        Private _startpoint As Point

        Public Sub New(ByVal adornedElement As UIElement, ByVal content As UIElement, Optional startpoint As Point = Nothing)
            MyBase.New(adornedElement)

            _startpoint = startpoint
            _child = New Rectangle()
            _child.Width = content.RenderSize.Width
            _child.Height = content.RenderSize.Height
            _child.Fill = New ImageBrush(BitmapFrame.Create(RenderToBitmap(content)))
        End Sub

        Protected Overrides Function MeasureOverride(ByVal constraint As System.Windows.Size) As System.Windows.Size
            _child.Measure(constraint)
            Return _child.DesiredSize
        End Function

        Protected Overrides Function ArrangeOverride(ByVal finalSize As System.Windows.Size) As System.Windows.Size
            _child.Arrange(New Rect(finalSize))
            Return finalSize
        End Function

        Protected Overrides Function GetVisualChild(ByVal index As Integer) As System.Windows.Media.Visual
            Return _child
        End Function

        Protected Overrides ReadOnly Property VisualChildrenCount() As Integer
            Get
                Return 1
            End Get
        End Property

        Public Property LeftOffset() As Double
            Get
                Return _leftoffset
            End Get
            Set(ByVal value As Double)
                _leftoffset = value
                UpdatePosition()
            End Set
        End Property

        Public Property TopOffset() As Double
            Get
                Return _topoffset
            End Get
            Set(ByVal value As Double)
                _topoffset = value
                UpdatePosition()
            End Set
        End Property

        Public Sub UpdatePosition()
            Dim adornerLayer As AdornerLayer = Me.Parent
            If Not adornerLayer Is Nothing Then
                adornerLayer.Update(AdornedElement)
            End If
        End Sub

        Public Overrides Function GetDesiredTransform(ByVal transform As System.Windows.Media.GeneralTransform) As System.Windows.Media.GeneralTransform
            Dim result As GeneralTransformGroup = New GeneralTransformGroup()
            result.Children.Add(New TranslateTransform(_leftoffset - _startpoint.X, _topoffset - _startpoint.Y))
            Return result
        End Function

        Public Function RenderToBitmap(ByVal element As FrameworkElement) As RenderTargetBitmap
            Dim topLeft As Double = 0
            Dim topRight As Double = 0
            Dim width As Integer = CInt(element.ActualWidth)
            Dim height As Integer = CInt(element.ActualHeight)
            Dim dpiX As Double = 96
            Dim dpiY As Double = 96
            Dim pixelFormat As PixelFormat = PixelFormats.[Default]
            Dim elementBrush As VisualBrush = New VisualBrush(element)
            Dim visual As DrawingVisual = New DrawingVisual()
            Dim dc As DrawingContext = visual.RenderOpen()
            dc.DrawRectangle(elementBrush, Nothing, New Rect(topLeft, topRight, width, height))
            dc.Close()
            Dim bitmap As RenderTargetBitmap = New RenderTargetBitmap(width, height, dpiX, dpiY, pixelFormat)
            bitmap.Render(visual)
            Return bitmap
        End Function
    End Class

#End Region

End Class