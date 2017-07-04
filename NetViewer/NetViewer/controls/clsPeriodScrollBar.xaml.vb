Imports System.Windows.Controls.Primitives

Public Class clsPeriodScrollBar

    Public Shared ReadOnly ValueFromProperty As DependencyProperty = DependencyProperty.Register("ValueFrom",
                                                            GetType(Date),
                                                            GetType(clsPeriodScrollBar),
                                                            New FrameworkPropertyMetadata(Today.AddDays(-10), AddressOf ValueFromPropertyChanged))
    Public Shared ReadOnly ValueToProperty As DependencyProperty = DependencyProperty.Register("ValueTo",
                                                            GetType(Date),
                                                            GetType(clsPeriodScrollBar),
                                                            New FrameworkPropertyMetadata(Today.AddDays(-20), AddressOf ValueToPropertyChanged))
    Public Shared ReadOnly MinimumProperty As DependencyProperty = DependencyProperty.Register("Minimum",
                                                            GetType(Date),
                                                            GetType(clsPeriodScrollBar),
                                                            New FrameworkPropertyMetadata(Today.AddDays(-30), AddressOf MinimumPropertyChanged))
    Public Shared ReadOnly MaximumProperty As DependencyProperty = DependencyProperty.Register("Maximum",
                                                            GetType(Date),
                                                            GetType(clsPeriodScrollBar),
                                                            New FrameworkPropertyMetadata(Today, AddressOf MaximumPropertyChanged))

    Public Event ValueFromChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Date))
    Public Event ValueToChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Date))

    Private Shared Sub ValueFromPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As clsPeriodScrollBar = CType(d, clsPeriodScrollBar)
        With instance
            Dim newvalue As Date = CType(e.NewValue, Date)
            If newvalue < .GetValue(MinimumProperty) Then newvalue = .GetValue(MinimumProperty)

            .SetValue(e.Property, newvalue)
            '.UpdateSlider()
        End With
    End Sub

    Private Shared Sub ValueToPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As clsPeriodScrollBar = CType(d, clsPeriodScrollBar)
        With instance
            Dim newvalue As Date = CType(e.NewValue, Date)
            If newvalue > .GetValue(MaximumProperty) Then newvalue = .GetValue(MaximumProperty)

            .SetValue(e.Property, newvalue)
            '.UpdateSlider()
        End With
    End Sub

    Private Shared Sub MinimumPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As clsPeriodScrollBar = CType(d, clsPeriodScrollBar)
        With instance
            Dim newvalue As Date = CType(e.NewValue, Date)
            .SetValue(ValueFromProperty, newvalue)
            .SetValue(MinimumProperty, newvalue)
            '.UpdateSlider()
        End With
    End Sub

    Private Shared Sub MaximumPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim instance As clsPeriodScrollBar = CType(d, clsPeriodScrollBar)
        With instance
            Dim newvalue As Date = CType(e.NewValue, Date)
            .SetValue(ValueToProperty, newvalue)
            .SetValue(MaximumProperty, newvalue)
            '.UpdateSlider()
        End With
    End Sub

    Public Property ValueFrom() As Date
        Get
            Return GetValue(ValueFromProperty)
        End Get
        Set(ByVal value As Date)
            SetValue(ValueFromProperty, value)
        End Set
    End Property

    Public Property ValueTo() As Date
        Get
            Return GetValue(ValueToProperty)
        End Get
        Set(ByVal value As Date)
            SetValue(ValueToProperty, value)
        End Set
    End Property

    Public Property Minimum() As Date
        Get
            Return GetValue(MinimumProperty)
        End Get
        Set(ByVal value As Date)
            SetValue(MinimumProperty, value)
        End Set
    End Property

    Public Property Maximum() As Date
        Get
            Return GetValue(MaximumProperty)
        End Get
        Set(ByVal value As Date)
            SetValue(MaximumProperty, value)
        End Set
    End Property

    Sub New()
        InitializeComponent()
    End Sub

    Public Sub UpdateSlider()
        'for stars
        gSlider.ColumnDefinitions(0).Width = New GridLength(Math.Abs((GetValue(ValueFromProperty) - GetValue(MinimumProperty)).TotalSeconds / (GetValue(MaximumProperty) - GetValue(MinimumProperty)).TotalSeconds), GridUnitType.Star)
        gSlider.ColumnDefinitions(1).Width = New GridLength(Math.Abs((GetValue(ValueToProperty) - GetValue(ValueFromProperty)).TotalSeconds / (GetValue(MaximumProperty) - GetValue(MinimumProperty)).TotalSeconds), GridUnitType.Star)
        gSlider.ColumnDefinitions(2).Width = New GridLength(Math.Abs((GetValue(MaximumProperty) - GetValue(ValueToProperty)).TotalSeconds / (GetValue(MaximumProperty) - GetValue(MinimumProperty)).TotalSeconds), GridUnitType.Star)
    End Sub

    Public Sub UpdateValues()
        Dim _oldfrom As Date = GetValue(ValueFromProperty)
        Dim _oldto As Date = GetValue(ValueToProperty)

        SetValue(ValueFromProperty, GetValue(MinimumProperty).AddSeconds((GetValue(MaximumProperty) - GetValue(MinimumProperty)).TotalSeconds * gSlider.ColumnDefinitions(0).ActualWidth / gSlider.ActualWidth))
        SetValue(ValueToProperty, GetValue(MinimumProperty).AddSeconds((GetValue(MaximumProperty) - GetValue(MinimumProperty)).TotalSeconds * (gSlider.ColumnDefinitions(0).ActualWidth + gSlider.ColumnDefinitions(1).ActualWidth) / gSlider.ActualWidth))

        RaiseEvent ValueFromChanged(Me, New RoutedPropertyChangedEventArgs(Of Date)(_oldfrom, GetValue(ValueFromProperty)))
        RaiseEvent ValueToChanged(Me, New RoutedPropertyChangedEventArgs(Of Date)(_oldto, GetValue(ValueToProperty)))
    End Sub

    Private Sub gsLeft_DragDelta(sender As Object, e As DragDeltaEventArgs) Handles gsLeft.DragDelta
        UpdateValues()
        e.Handled = True
    End Sub

    Private Sub gsCenter_DragDelta(sender As Object, e As DragDeltaEventArgs) Handles gsCenter.DragDelta
        UpdateValues()
        e.Handled = True
    End Sub

    Private Sub gsRight_DragDelta(sender As Object, e As DragDeltaEventArgs) Handles gsRight.DragDelta
        UpdateValues()
        e.Handled = True
    End Sub

    Private Sub clsPeriodScrollBar_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        'UpdateSlider()
    End Sub

End Class
