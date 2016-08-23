VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_PredeclaredId = False
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Attribute VB_GlobalNameSpace = False
Attribute VB_Name = "INyDIA_SourceTube"
    Public DetectedVolume As Long
    Private internal_RemainingVolume As Long
    	Public Property Get RemainingVolume
			RemainingVolume = DetectedVolume - NoOfDispensedQuotas * ReqQuotaVolume
		End Property
    Public ReqQuotaVolume As Long
    Public UseIncompleteQuota As Boolean
    Public PreviousTubesQuotas As Integer
    Private internal_NoOfQuotas As Integer
    	Public Property Get NoOfQuotas
			NoOfQuotas = IIf(UseIncompleteQuota And DetectedVolume > 0 And LastQuotaVolume > MinLastQuota, Int(DetectedVolume/ReqQuotaVolume) + 1, Int(DetectedVolume/ReqQuotaVolume))
		End Property
	Private internal_LastQuotaVolume As Long
		Public Property Get LastQuotaVolume
			LastQuotaVolume = Abs((DetectedVolume Mod ReqQuotaVolume) * UseIncompleteQuota)
		End Property
	Private internal_IsTubeEmpty As Boolean
		Public Property Get IsTubeEmpty
				IsTubeEmpty = RemainingVolume <= MinLastQuota
		End Property
	Public TubeOrderNo As Integer
	Public NoOfDispensedQuotas As Integer
	Public MinLastQuota As Long

Private Sub Class_Initialize()
    'Debug.Print "Class_Initialize"
	UseIncompleteQuota = True
	MinLastQuota = 100
End Sub
