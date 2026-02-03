Attribute VB_Name = "Module06Reports"
Sub StandardReport()
'
' Report - Standard
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\New Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_NewReport.Show
End Sub

Sub A3Report()
'
' Report - A3
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\A3 Report.dotm" _
        , NewTemplate:=False, DocumentType:=0

End Sub

Sub ProofReport()
'
' Proof Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\Proof Report.dotm" _
        , NewTemplate:=False, DocumentType:=0

End Sub

Sub PCR()
'
' PCR - Westminster Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\PCR report - Westminster.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_PCR_Westminster.Show
End Sub

Sub NIACamden()
'
' NIA - Camden Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\PCR report - Camden.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_PCR_Camden.Show
End Sub

Sub NIARBKC()
'
' NIA - Camden Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\PCR report - RBKC.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_PCR_RBKC.Show
End Sub

Sub NIACoL()
'
' NIA - Camden Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\PCR report - City of London.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_PCR_London.Show
End Sub

Sub NIALA()
'
' NIA - Camden Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\NIA report - Local Authority.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_NIALocalAuthorityReport.Show
End Sub

Sub NPPFEBF()
'
' NPPF & EBF Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\NPPF & EBF report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_NPPF_EBF.Show
End Sub


Sub NPPF6472EBF()
'
' NPPF, BS6472 & EBF Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\NPPF,6472 & EBF report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_NPPF_6472_EBF.Show
End Sub

Sub ClassMAReport()
'
' NPPF, BS6472 & EBF Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\Class MA report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_MA.Show
End Sub

Sub BS4142Report()
'
' BS4142 Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\BS4142 Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_BS4142Report.Show
End Sub


Sub EalingReport()
'
' Ealing Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\Ealing report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_Ealing_Report.Show
End Sub

Sub CMPCamden()
'
' Camden CMP Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\CMP Camden Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_CMPCamdenReport.Show
End Sub

Sub CMPReport()
'
' CMP Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\CMP Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_CMPReport.Show
End Sub




Sub CMPRBKC()
'
' RBKC CMP Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\CMP RBKC Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_CMPRBKCReport.Show
End Sub


Sub PCTWall()
'
' PCT Wall Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\ADE2003 PCT SI Wall Test Report.dotm" _
        , NewTemplate:=False, DocumentType:=0

End Sub

Sub PCTFloor()
'
' PCT Floor Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\ADE2003 PCT SI Floor Test Report.dotm" _
        , NewTemplate:=False, DocumentType:=0

End Sub

Sub PCTWallandFloor()
'
' PCT Floor Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\ADE2003 PCT SI Wall and Floor Test Report.dotm" _
        , NewTemplate:=False, DocumentType:=0

End Sub

Sub ADE()
'
' ADE/SI Design Review Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\ADE Review.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_ADEReport.Show

End Sub


Sub NMPReport()
'
' Gym Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\NMP Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_NMPReport.Show
End Sub

Sub NAW()
'
' NAW Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\NAW Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_NAWReport.Show
End Sub

Sub NIHL()
'
' NIHL Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\NIHL.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_NIHLReport.Show
End Sub

Sub Gym()
'
' Gym Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\Gym Report.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_GymReport.Show
End Sub

Sub Odour()
'
' Odour Assessment Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\Odour Assessment.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    Form5_Odour.Show
End Sub


Sub GenSpec()
'
' General Specification Report
'
'
    Documents.Add Template:= _
        AddinFolder & "\5. Reports\GenSpec.dotm" _
        , NewTemplate:=False, DocumentType:=0
    
    form5_GenSpec.Show
End Sub

Sub equip_table()

'
' Insert Equipment Table text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\Equipment List.docx", Range:="Equipment", Link:=False
    
        Selection.WholeStory
    Selection.LanguageID = wdEnglishUK
    Selection.NoProofing = False
    Application.CheckLanguage = False
        
End Sub
Sub LA90Leqtable()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="LA90Leqtable", Link:=False
    
End Sub
Sub LA90Leqtable2pos()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="LA90Leqtable2pos", Link:=False
    
End Sub

Sub ReportIntro()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="intro", Link:=False
    
End Sub

Sub ReportNPPF()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="NPPF", Link:=False
    
End Sub

Sub Report8233()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS82332014full", Link:=False
    
End Sub

Sub ReportWHO()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="WHOfull", Link:=False
    
End Sub

Sub AVO()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="AVO", Link:=False
    
End Sub

Sub ProPGCriteria()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ProPGCriteria", Link:=False
    
End Sub

Sub Report41421997()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS41421997method", Link:=False
    
End Sub

Sub Report41422014()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS41422014method", Link:=False
    
End Sub

Sub Report41422014long()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS41422014methodlong", Link:=False
    
End Sub

Sub ClassAA()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ClassAA", Link:=False
    
End Sub
Sub ClassAB()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ClassAB", Link:=False
    
End Sub
Sub ClassMA()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ClassMA", Link:=False
    
End Sub
Sub ClassO()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ClassO", Link:=False
    
End Sub
Sub ClassQ()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ClassQ", Link:=False
    
End Sub

Sub IOAMusic()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="IOAMusic", Link:=False
    
End Sub

Sub NANR()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="NANR", Link:=False
    
End Sub

Sub Report6472()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS64721", Link:=False
    
End Sub

Sub Report5228()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS52282009", Link:=False
    
End Sub

Sub ReportCRTN()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="CRTN", Link:=False
    
End Sub

Sub ReportMPS2()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="MPS2", Link:=False
    
End Sub

Sub ReportDMRB()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="DMRB", Link:=False
    
End Sub

Sub ReportIEMA()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="IEMA", Link:=False
    
End Sub

Sub ReportNAW()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="NAW", Link:=False
    
End Sub

Sub Defraodourintro()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Defraodourintro", Link:=False
    
End Sub

Sub Defraodour()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Defraodour", Link:=False
    
End Sub

Sub ReportHW17()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="HW17", Link:=False
    
End Sub

Sub ReportP13()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="P13", Link:=False
    
End Sub
Sub Report2008Pol8()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="P82008", Link:=False
    
End Sub
Sub Report2008Hea13Court()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H132008Courts", Link:=False
    
End Sub
Sub Report2008Hea13Edu()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H132006Education", Link:=False
    
End Sub
Sub Report2008Hea13Health()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H132006Health", Link:=False
    
End Sub
Sub Report2008Hea13Industrial()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H132006Industrial", Link:=False
    
End Sub
Sub Report2008Hea13Offices()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H132006Office", Link:=False
    
End Sub
Sub Report2008Hea13Prisons()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H132006Prisons", Link:=False
    
End Sub
Sub Report2008Hea13Retail()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H132006Retail", Link:=False
    
End Sub
Sub Report2011Pol05()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="P52011", Link:=False
    
End Sub
Sub Report2011Hea5Schools()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52011Schools", Link:=False
    
End Sub
Sub Report2011Hea5HE()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52011HE", Link:=False
    
End Sub
Sub Report2011Hea5Health()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52011Health", Link:=False
    
End Sub
Sub Report2011Hea5Resi()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52011Resi", Link:=False
    
End Sub
Sub Report2011Hea5Other()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52011Otheri", Link:=False
    
End Sub
Sub Report2014Pol05()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="P52014", Link:=False
    
End Sub
Sub Report2014Hea05Education()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52014Education", Link:=False
    
End Sub
Sub Report2014Hea05Health()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52014Health", Link:=False
    
End Sub
Sub Report2014Hea05Office()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52014Office", Link:=False
    
End Sub
Sub Report2014Hea05Law()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52014Law", Link:=False
    
End Sub
Sub Report2014Hea05Industrial()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52014Industrial", Link:=False
    
End Sub
Sub Report2014Hea05Resi()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="H52014Resi", Link:=False
    
End Sub
Sub ReportEco2006()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Eco2006", Link:=False
    
End Sub
Sub ReportCfSH2008()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="CfSH2008", Link:=False
    
End Sub
Sub ReportCfSH2010()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="CfSH2010", Link:=False
    
End Sub
Sub ReportBCO2005()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BCO2005", Link:=False
    
End Sub
Sub ReportBCO2009()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BCO2009", Link:=False
    
End Sub
Sub ReportBCO2011()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BCO2011", Link:=False
    
End Sub

Sub Basis()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Basis", Link:=False
    
End Sub

Sub Islingtonmusic()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Islingtonmusic", Link:=False
    
End Sub

Sub ReportENS()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ENS", Link:=False
    
End Sub

Sub ReportNIA()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="NIA", Link:=False
    
End Sub

Sub Spectraladaptation()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Spectraladaptation", Link:=False
    
End Sub

Sub Noisemapping()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="noisemapping", Link:=False
    
End Sub

Sub ReportEBF()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="EBF", Link:=False
    
End Sub

Sub Carmovements()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Carmovements", Link:=False
    
End Sub

Sub Partywalljunction()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Partywalljunction", Link:=False
    
End Sub
Sub Glazedjunction()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Glazedjunction", Link:=False
    
End Sub

Sub ReportLift()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Lift", Link:=False
    
End Sub
Sub ReportSVP()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="SVP", Link:=False
    
End Sub
Sub ReportTenancy()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Tenancy", Link:=False
    
End Sub

Sub carpet()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="carpet", Link:=False
    
End Sub

Sub carpetlong()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="carpetlong", Link:=False
    
End Sub

Sub Section7()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Section7", Link:=False
    
End Sub

Sub SICommtoResi()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="SICommtoResi", Link:=False
    
End Sub

Sub OfficeSIPrivacy()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="OfficeSIPrivacy", Link:=False
    
End Sub

Sub BeerGarden()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BeerGarden", Link:=False
    
End Sub

Sub Pubnoise()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Pubnoise", Link:=False
    
End Sub


Sub Hybridapproach()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Hybridapproach", Link:=False
    
End Sub


Sub groundborne()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="groundborne", Link:=False
    
End Sub

Sub ProPGRiskAssessment()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ProPGRiskAssessment", Link:=False
    
End Sub


Sub IntoOut()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="IntoOut", Link:=False
    
End Sub

Sub ADO()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="ADO", Link:=False
    
End Sub

Sub BS7445assess()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS7445assess", Link:=False
    
End Sub


Sub BS4142assess()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS4142assess", Link:=False
    
End Sub


Sub BS4142context()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BS4142context", Link:=False
    
End Sub

Sub Throughwindow()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Throughwindow", Link:=False
    
End Sub


Sub Overheating()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Overheating", Link:=False
    
End Sub

Sub AVOGRisk()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="AVOGRisk", Link:=False
    
End Sub

Sub RTReview()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="RTReview", Link:=False
    
End Sub

Sub NewBuild()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="NewBuild", Link:=False
    
End Sub

Sub Conversion()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Conversion", Link:=False
    
End Sub


'Sub Hea05CritOffice()
'
' Insert autotext

'Selection.Collapse Direction:=wdCollapseEnd
'    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Hea05CritOffice", Link:=False
    
'End Sub
'Sub Hea05DesOffice()
''
'' Insert autotext
'
'Selection.Collapse Direction:=wdCollapseEnd
'    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Hea05DesOffice", Link:=False
'
'End Sub

Sub Pol05report()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Pol05report", Link:=False
    
End Sub
Sub S3resi()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="S3resi", Link:=False
    
End Sub

Sub S4resi()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="S4resi", Link:=False
    
End Sub
Sub Hea05CritOffice()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Hea05CritOffice", Link:=False
    
End Sub

Sub Hea05CritOffice1()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Hea05CritOffice", Link:=False
    
End Sub

Sub Hea05DesOffice()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Hea05DesOffice", Link:=False
    
End Sub

Sub S3Mixed()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="S3Mixed", Link:=False
    
End Sub

Sub S4Mixed()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="S4Mixed", Link:=False
    
End Sub

Sub BB93new()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BB93new", Link:=False
    
End Sub

Sub BB93refurb()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="BB93refurb", Link:=False
    
End Sub

Sub L90eqdescription()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="L90eqdescription", Link:=False
    
End Sub

Sub Nuisance()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Nuisance", Link:=False
    
End Sub

Sub Minerals()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Minerals", Link:=False
    
End Sub

Sub Officepartitions()
'
' Insert autotext

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\5. Reports\report builder.docx", Range:="Officepartitions", Link:=False
    
End Sub
