Imports System.Collections.ObjectModel
Imports System.Windows.Media.Animation

Public Class DragDropReorderer

    Private Shared elements As New List(Of UIElement)
    Private Shared mousedownpoint As Point?
    Private Shared adorner As DragAdorner
    Private Shared window As Window

#Region "AttachedProperties"

    Public Shared ReadOnly AllowDragProperty As DependencyProperty = DependencyProperty.RegisterAttached(
        "AllowDrag", GetType(Boolean), GetType(DragDropReorderer), New PropertyMetadata(False,
            Sub(uielement As UIElement, e As DependencyPropertyChangedEventArgs)
                If e.NewValue IsNot Nothing AndAlso e.NewValue = True Then
                    AddHandler uielement.PreviewMouseLeftButtonDown, AddressOf UIElement_PreviewMouseLeftButtonDown
                    AddHandler uielement.PreviewMouseMove, AddressOf UIElement_PreviewMouseMove
                    AddHandler uielement.PreviewDragOver, AddressOf UIElement_PreviewDragOver
                    AddHandler uielement.PreviewDrop, AddressOf UIElement_PreviewDrop
                    uielement.AllowDrop = True
                    If Not elements.Contains(uielement) Then elements.Add(uielement)
                End If
                If e.NewValue IsNot Nothing AndAlso e.NewValue = False Then
                    RemoveHandler uielement.PreviewMouseLeftButtonDown, AddressOf UIElement_PreviewMouseLeftButtonDown
                    RemoveHandler uielement.PreviewMouseMove, AddressOf UIElement_PreviewMouseMove
                    RemoveHandler uielement.PreviewDragOver, AddressOf UIElement_PreviewDragOver
                    RemoveHandler uielement.PreviewDrop, AddressOf UIElement_PreviewDrop
                    If elements.Contains(uielement) Then elements.Remove(uielement)
                End If
            End Sub))

    Public Shared Function GetAllowDrag(element As UIElement) As Boolean
        If element Is Nothing Then Throw New ArgumentException
        Return element.GetValue(AllowDragProperty)
    End Function

    Public Shared Sub SetAllowDrag(element As UIElement, value As Boolean)
        element.SetValue(AllowDragProperty, value)
    End Sub

    Public Shared ReadOnly DraggedOpacityProperty As DependencyProperty = DependencyProperty.RegisterAttached("DraggedOpacity", GetType(Double), GetType(DragDropReorderer), New PropertyMetadata(0.0))

    Public Shared Function GetDraggedOpacity(element As UIElement) As Double
        If element Is Nothing Then Throw New ArgumentException
        Return element.GetValue(DraggedOpacityProperty)
    End Function

    Public Shared Sub SetDraggedOpacity(element As UIElement, value As Double)
        element.SetValue(DraggedOpacityProperty, value)
    End Sub

    Public Shared ReadOnly DraggedAnimationProperty As DependencyProperty = DependencyProperty.RegisterAttached("DraggedAnimation", GetType(Boolean), GetType(DragDropReorderer), New PropertyMetadata(True))

    Public Shared Function GetDraggedAnimation(element As UIElement)
        If element Is Nothing Then Throw New ArgumentException
        Return element.GetValue(DraggedAnimationProperty)
    End Function

    Public Shared Sub SetDraggedAnimation(element As UIElement, value As Boolean)
        element.SetValue(DraggedAnimationProperty, value)
    End Sub

#End Region

#Region "Events"

    Private Shared Sub UIElement_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        mousedownpoint = e.GetPosition(sender)
    End Sub

    Private Shared Sub UIElement_PreviewMouseMove(sender As Object, e As MouseEventArgs)
        Dim mousemovepoint = e.GetPosition(sender)
        If Not mousedownpoint.HasValue OrElse e.LeftButton = MouseButtonState.Released OrElse
        (Point.Subtract(mousedownpoint.Value, mousemovepoint).Length < SystemParameters.MinimumHorizontalDragDistance And
         Point.Subtract(mousedownpoint.Value, mousemovepoint).Length < SystemParameters.MinimumVerticalDragDistance) Then Exit Sub

        DoDragDrop(mousemovepoint, sender)
    End Sub

    Private Shared Sub Window_PreviewDragOver(sender As Object, e As DragEventArgs)
        If adorner Is Nothing OrElse window Is Nothing Then Exit Sub
        adorner.LeftOffset = e.GetPosition(window).X
        adorner.TopOffset = e.GetPosition(window).Y
    End Sub

    Private Shared Sub UIElement_PreviewDragOver(sender As Object, e As DragEventArgs)
        If Not e.Data.GetDataPresent(sender.GetType) Then Exit Sub
        Dim sourceelement As Object = e.Data.GetData(sender.GetType) : If sourceelement Is Nothing Then Exit Sub
        Dim targetelement As Object = sender
        If sourceelement Is targetelement Then Exit Sub
        Dim sourcedata As Object = sourceelement.DataContext : If sourcedata Is Nothing Then Exit Sub
        Dim targetdata As Object = targetelement.DataContext : If targetdata Is Nothing Then Exit Sub
        Dim panel = FindVisualParent(Of Panel)(sender) : If panel Is Nothing Then Exit Sub
        Dim itemscontrol = FindVisualParent(Of ItemsControl)(panel) : If itemscontrol Is Nothing Then Exit Sub
        Dim itemssource = itemscontrol.ItemsSource : If itemssource Is Nothing Then Exit Sub
        Dim sourcetype As Type = itemssource.GetType
        If Not sourcetype.IsGenericType OrElse Not (sourcetype.GetGenericTypeDefinition = GetType(ObservableCollection(Of))) Then Exit Sub
        Dim collection As Object = itemssource ' ObservableCollection of sometype
        Dim sourcedataindex As Integer = collection.IndexOf(sourcedata) : If sourcedataindex = -1 Then Exit Sub
        Dim targetdataindex As Integer = collection.IndexOf(targetdata) : If targetdataindex = -1 Then Exit Sub
        collection.Move(sourcedataindex, targetdataindex)
        If GetDraggedAnimation(sender) Then FadeIn(itemscontrol.ItemContainerGenerator.ContainerFromIndex(sourcedataindex))
    End Sub

    Private Shared Sub UIElement_PreviewDrop(sender As Object, e As DragEventArgs)
        If Not e.Data.GetDataPresent(sender.GetType) Then Exit Sub
        Dim sourceelement As Object = e.Data.GetData(sender.GetType) : If sourceelement Is Nothing Then Exit Sub
        Dim targetelement As Object = sender
        If sourceelement Is targetelement Then Exit Sub
        Dim sourcedata As Object = sourceelement.DataContext : If sourcedata Is Nothing Then Exit Sub
        Dim targetdata As Object = targetelement.DataContext : If targetdata Is Nothing Then Exit Sub
        Dim panel = FindVisualParent(Of Panel)(sender) : If panel Is Nothing Then Exit Sub
        Dim itemscontrol = FindVisualParent(Of ItemsControl)(panel) : If itemscontrol Is Nothing Then Exit Sub
        If itemscontrol.ItemsSource IsNot Nothing Then Exit Sub
        Dim items = itemscontrol.Items : If items Is Nothing Then Exit Sub
        Dim sourcedataindex As Integer = items.IndexOf(sourcedata) : If sourcedataindex = -1 Then Exit Sub
        Dim targetdataindex As Integer = items.IndexOf(targetdata) : If targetdataindex = -1 Then Exit Sub
        items.RemoveAt(sourcedataindex)
        items.Insert(targetdataindex, sourcedata)
        itemscontrol.UpdateLayout()
    End Sub

#End Region

#Region "Subs"

    Private Shared Sub DoDragDrop(startpoint As Point, uielement As UIElement)
        window = Window.GetWindow(uielement) : If window Is Nothing Then Exit Sub
        Dim lastallowdrop As Boolean = window.AllowDrop : window.AllowDrop = True
        AddHandler window.PreviewDragOver, AddressOf Window_PreviewDragOver
        adorner = New DragAdorner(uielement, uielement, startpoint) : AdornerLayer.GetAdornerLayer(window.Content).Add(adorner)
        Dim lastopacity = uielement.Opacity : uielement.Opacity = GetDraggedOpacity(uielement)
        Dim dragdata As New DataObject(uielement) : DragDrop.DoDragDrop(uielement, dragdata, DragDropEffects.Move)
        uielement.Opacity = lastopacity
        AdornerLayer.GetAdornerLayer(window.Content).Remove(adorner) : adorner = Nothing
        RemoveHandler window.PreviewDragOver, AddressOf Window_PreviewDragOver
        window.AllowDrop = lastallowdrop
    End Sub

    Public Shared Function FindVisualParent(Of T As DependencyObject)(ByVal child As Object, Optional until As DependencyObject = Nothing) As T
        Dim parent As DependencyObject = If(child.Parent, VisualTreeHelper.GetParent(child))

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

    Private Shared Sub AnimateOpacity(ByVal target As DependencyObject, ByVal from As Double, ByVal [to] As Double)
        Dim opacityAnimation = New DoubleAnimation With {
            .From = from,
            .[To] = [to],
            .Duration = TimeSpan.FromMilliseconds(500)
        }
        Storyboard.SetTarget(opacityAnimation, target)
        Storyboard.SetTargetProperty(opacityAnimation, New PropertyPath("Opacity"))
        Dim _storyboard = New Storyboard()
        _storyboard.Children.Add(opacityAnimation)
        _storyboard.Begin()
    End Sub

    Public Shared Sub FadeIn(ByVal target As DependencyObject)
        AnimateOpacity(target, 0, 1)
    End Sub

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