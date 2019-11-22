VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BallisticPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Represents a ballistic package for testing
'====================================
Implements ITestingPlan
'Factory Module Constructor function:
'Function CreateBallisticPackage() As BallisticPackage
'----------------
'CLASS VARIABLES
'----------------
Private GSM_TO_PSF As Double
Private IN_IN_TO_YARDS As Double
Private SQIN_TO_SQFT As Double
Private m_ActualPsf As Double
Private m_TargetPsf As Double
Private m_ConditionedWeight As Double
Private m_FabricWidthInches As Double
Private m_PackageLengthInches As Double
Private m_NumberOfLayers As Integer
Private m_Template As SpecificationTemplate
' Properties dictionary
Public Properties   As Object
'------------
' PROPERTIES
'------------
Public Property Set Template(obj As SpecificationTemplate): Set m_Template = obj: End Property
Public Property Get Template() As SpecificationTemplate: Set Template = m_Template: End Property
Public Property Let PackageLengthInches(value As Double): m_PackageLengthInches = value: End Property
Public Property Get PackageLengthInches() As Double: PackageLengthInches = m_PackageLengthInches: End Property
Public Property Let FabricWidthInches(value As Double): m_FabricWidthInches = value: End Property
Public Property Get FabricWidthInches() As Double: FabricWidthInches = m_FabricWidthInches: End Property
Public Property Let ConditionedWeight(value As Double): m_ConditionedWeight = value: End Property
Public Property Get ConditionedWeight() As Double: ConditionedWeight = m_ConditionedWeight: End Property
Public Property Let TargetPsf(value As Double): m_TargetPsf = value: End Property
Public Property Get TargetPsf() As Double: TargetPsf = m_TargetPsf: End Property
' Returns the number of layers required to build a package having a min of m_TargetPsf.
Public Property Let NumberOfLayers(value As Integer)
    m_NumberOfLayers = value
End Property
Public Property Get NumberOfLayers() As Integer
    If m_NumberOfLayers = 0 Then
        NumberOfLayers = Application.WorksheetFunction.RoundUp(m_TargetPsf / (m_ConditionedWeight * GSM_TO_PSF), 0)
    Else
        NumberOfLayers = m_NumberOfLayers
    End If
End Property
'----------
' METHODS
'----------
Public Sub CalculatePackageProperties()
' Perform calculations to determine package properties
    Dim props_temp As Object
    Set props_temp = Me.Template.Properties
    props_temp("Ballistic Sample Length") = MinRequiredSampleLength
    props_temp("Number of Layers") = NumberOfLayers
    props_temp("Package Areal Density [psf]") = ActualPsf
    props_temp("Package Weight [lbm]") = PackageMass
    Set Me.Properties = props_temp
End Sub

Function ActualPsf() As Double
' Returns the actual psf as calculated from the min number of layers, this number will always be greater than the m_TargetPsf
    ActualPsf = NumberOfLayers * ConditionedWeight * GSM_TO_PSF
End Function

Function MinRequiredSampleLength() As Double
' Returns the minimum number of yards of fabric required to build the package.
    MinRequiredSampleLength = Application.WorksheetFunction.Ceiling(NumberOfLayers / (Application.WorksheetFunction.RoundDown(FabricWidthInches / PackageLengthInches, 0)) * IN_IN_TO_YARDS, 0.5)
End Function

Function PackageMass() As Double
' Returns the estimated package mass based on the actual psf
    PackageMass = ActualPsf * (PackageLengthInches * PackageLengthInches * SQIN_TO_SQFT)
End Function

' ITestingPlan Methods
Function ITestingPlan_GetTemplate() As SpecificationTemplate
    Set Template = App.templates("Ballistic Testing Requirements")
    Set ITestingPlan_GetTemplate = Template
End Function

Function ITestingPlan_GetPropertiesJson() As String
    CalculatePackageProperties
    ITestingPlan_GetPropertiesJson = JsonVBA.ConvertToJson(Me.Properties)
End Function

' Constructor
Private Sub Class_Initialize()
    GSM_TO_PSF = 0.000204861
    IN_IN_TO_YARDS = 0.459459459
    SQIN_TO_SQFT = 0.006944444
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
' Deconstructor
Private Sub Class_Terminate()
    Set Template = Nothing
    Set Properties = Nothing
End Sub
