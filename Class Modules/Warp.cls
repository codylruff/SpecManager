VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Warp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Classes")
'====================================
'CLASS DESCRIPTION:
' Represents a specific warp order.
'====================================
'Factory Module Constructor function:
'Function CreateWarp(Specification As WarpingSpecification, NumberOfBobbins As Integer, _
'                     PackageWeightlbs As Double, WarpLengthYds As Double) As Warp
'----------------
' CONSTANTS
'----------------
Const MinimumPackageLengthyds = 10936
'----------------
'CLASS VARIABLES
'----------------
Private Specification_          As Specification
Private NumberOfBobbins_        As Double
Private PackageWeightlbs_       As Double
Private WarpLengthYds_          As Double
'----------------
' AUTO-PROPERTIES
'----------------
' WarpLengthYds
Public Property Let WarpLengthYds(value As Double): WarpLengthYds_ = value: End Property
Public Property Get WarpLengthYds() As Double: WarpLengthYds = WarpLengthYds_: End Property
' PackageWeightlbs
Public Property Let PackageWeightlbs(value As Double): PackageWeightlbs_ = value: End Property
Public Property Get PackageWeightlbs() As Double: PackageWeightlbs = PackageWeightlbs_: End Property
' NumberOfBobbins
Public Property Let NumberOfBobbins(value As Double): NumberOfBobbins_ = value: End Property
Public Property Get NumberOfBobbins() As Double: NumberOfBobbins = NumberOfBobbins_: End Property
' Specification
Public Property Set Specification(value As Specification): Set Specification_ = value: End Property
Public Property Get Specification() As Specification: Set Specification = Specification_: End Property
'----------------
' PROPERTIES
'----------------
' Is Minimum Package Length
Public Property Get IsMinimumPackageLength() As Boolean
    ' Allows for less changes in weaving
    IsMinimumPackageLength = IIf(ResidualLengthYds > MinimumPackageLengthyds, True, False)
End Property
' Residual Length yards
Public Property Get ResidualLengthYds() As Double
    ResidualLengthYds = PackageLengthYds - WarpLengthYds
End Property
' PackageLengthYds
Public Property Get PackageLengthYds() As Double
    With Me.Specification.Properties
        ' 454 g / lb and 0.9144 m / yd
        PackageLengthYds = PackageWeightlbs / (.Item("Dtex") / 10000 / 454 * 0.9144)
    End With
End Property
' NumberOfSections
Public Property Get NumberOfSections() As Double
    With Me.Specification.Properties
        NumberOfSections = .Item("EndsPerInch") * .Item("FinalWidthCm") / 2.54 / NumberOfBobbins
    End With
End Property
