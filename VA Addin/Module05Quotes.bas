Attribute VB_Name = "Module05Quotes"

Sub NewQuote()
'
' New Quote
'
'
    Application.ScreenUpdating = False
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\New quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_NewQuote.Show
    Application.ScreenUpdating = True
End Sub


Sub FullHotel()
'
' Full Brief - Hotel
'
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Full Brief - Hotel.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
        Form4_FullBriefHotelQuote.Show

End Sub

Sub FullOffice()

' Full Brief - Office


    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Full Brief - Office.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
        Form4_FullBriefOfficeQuote.Show

End Sub

Sub FullResi()

' Full Brief - Office


    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Full Brief - Resi.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
        Form4_FullBriefResiQuote.Show

End Sub

'Sub FullMixed()
''
'' Full Brief - Hotel
''
''
'    Documents.Add Template:= _
'        AddinFolder & "\4. Quotes\Full Brief - Mixed.dotm" _
'        , NewTemplate:=False, DocumentType:=0
'
'End Sub


Sub PCTQuote()
'
' PCT Quote Macro
'
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\PCT Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_PCTQuote.Show
    
End Sub

Sub PCRQuote()
'
' PCR Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\PCR Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_PCRQuote.Show
    
End Sub

Sub LANIAQuote()
'
' PCR Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\PCR non-Westminster quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_PCRNonWstMnstrQuote.Show
    
End Sub


Sub BS4142Quote()
'
' BS4142 Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\BS4142 Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_BS4142Quote.Show
    
End Sub

Sub NPPFQuote()
'
' NPPF Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\NPPF Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_NPPFQuote.Show
    
End Sub

Sub NIAEBFQuote()
'
' NIA (NPPF) & EBF Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\NIA&EBF Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_NiaEbfQuote.Show
    
End Sub

Sub OfficetoResiQuote()
'
' Office to Resi Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Office to Resi Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_OfficetoResiQuote.Show
    
End Sub

Sub RBKCCMPQuote()
'
' RBKC CMP Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\RBKC CMP Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_CMPQuote.Show
    
End Sub

Sub LongTermQuote()
'
' RBKC CMP Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Long Term Monitoring Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_LongTermQuote.Show
    
End Sub

Sub GymQuote()
'
' Gym Quote
'
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Gym quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
        Form4_Gymquote.Show
        

End Sub

Sub Licensing()
'
' Licensing Application
'
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Licensing Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
        Form4_LicensingQuote.Show

End Sub

Sub A1toA3()
'
' Change of use A1 to A3
'
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\A1toA3 Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
        Form4_A1toA3.Show

End Sub


Sub ETSUQuote()
'
' NAW Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\Windfarm Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_NewQuote.Show
    
End Sub

Sub NAWQuote()
'
' NAW Quote Macro
'
    Documents.Add Template:= _
        AddinFolder & "\4. Quotes\NAW Quote.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form4_NAWQuote.Show
    
End Sub

Sub Hourlyrates()
'
' Insert Hourly Rates

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Hourlyrates", Link:=False
        
End Sub

Sub Qintroduction()
'
' Insert Introduction

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="intro", Link:=False
        
End Sub

Sub Licensedpremises()
'
' Insert Licensed Premises quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Licensed", Link:=False
        
End Sub

Sub ENSsection()
'
' Insert ENS quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="ENS", Link:=False
        
End Sub

Sub BS4142section()
'
' Insert BS4142 quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="BS4142", Link:=False
        
End Sub

Sub BS4142resi()
'
' Insert BS4142 quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="BS4142resi", Link:=False
        
End Sub

            
Sub NPPFsection()
'
' Insert NPPF quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="NPPF", Link:=False
        
End Sub


Sub ProPG()
'
' Insert ProPG quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="ProPG", Link:=False
        
End Sub

Sub INC()
'
' Insert INC quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="INC", Link:=False
        
End Sub

Sub EBFsection()
'
' Insert EBF quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="EBF", Link:=False
        
End Sub


Sub Glazingappraisal()

'
' Insert Glazing Appraisal quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Glazing", Link:=False
        
End Sub

Sub Slam()

'
' Insert Car door Slam quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Slam", Link:=False
        
End Sub

Sub RoomAcoustics()

'
' Insert Room ACoustics quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="RoomAcoustics", Link:=False
        
End Sub


Sub BuildingServices()

'
' Insert Buiding Services quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="BuildingServices", Link:=False
        
End Sub

Sub PCRsection()
'
' Insert Planning Compliance quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="PC", Link:=False
        
End Sub

Sub Odourquote()
'
' Insert Planning Compliance quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Odour", Link:=False
        
End Sub

Sub Vibsection()
'
' Insert Vibration/BS6472 quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Vibration", Link:=False
        
End Sub

Sub ADEsection()
'
' Insert ADE Design Review quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="ADE", Link:=False
        
End Sub

Sub SITsection()
'
' Insert SIT quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="SIT", Link:=False
        
End Sub

Sub SITADEsection()
'
' Insert SIT & ADE quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="SITADE", Link:=False
        
End Sub

Sub PCTsection()
'
' Insert PCT quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="PCT", Link:=False
        
End Sub

Sub OfficeSI()
'
' Insert Office SI Tests quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="OfficeSI", Link:=False
        
End Sub

Sub RT()
'
' Insert Reverberation Time Tests quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="RT", Link:=False
        
End Sub

Sub Mappingsection()
'
' Insert Noise Mapping quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Mapping", Link:=False
        
End Sub

Sub Auralisationsection()
'
' Insert Noise Mapping quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Auralisation", Link:=False
        
End Sub

Sub EventSimulation()
'
' Insert Music Event Simulation quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="EventSimulation", Link:=False
        
End Sub

Sub NMP()
'
' Insert Noise Management Plan quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="NMP", Link:=False
        
End Sub

Sub Trafficshortsection()
'
' Insert Traffic Noise (short) quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Trafficshort", Link:=False
        
End Sub

Sub Trafficlongsection()
'
' Insert Traffic Noise (long) quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Trafficlong", Link:=False
        
End Sub

Sub Constructionsection()
'
' Insert Construction Noise quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Construction", Link:=False
        
End Sub

Sub CMP()
'
' Insert CMP quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="CMP", Link:=False
        
End Sub

Sub S61section()
'
' Insert Section 61 Assessment quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="S61", Link:=False
        
End Sub

Sub PostNoiseSurvey()

'
' Insert Point Installation Plant Noise Survey quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="PostNoiseSurvey", Link:=False
        
End Sub

Sub ESsection()
'
' Insert ES Chapter quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="ES", Link:=False
        
End Sub

Sub Reportingsection()
'
' Insert Reporting quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Reporting", Link:=False
        
End Sub

Sub Shortform()
'
' Insert ENS quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="ShortForm", Link:=False
        
End Sub

Sub novation()
'
' Insert novation quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="novation", Link:=False
        
End Sub

Sub SITerms()
'
' Insert novation quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="SITerms", Link:=False
        
End Sub


Sub JD()
'
' Opens JD folder

Dim sPath As String

sPath = "X:\CVs\CURRENT CVs\JD"

retVal = Shell("explorer.exe " & sPath, vbNormalFocus)

End Sub


Sub SL()
'
' Opens SL folder

Dim sPath As String

sPath = "X:\CVs\CURRENT CVs\SL"

retVal = Shell("explorer.exe " & sPath, vbNormalFocus)

End Sub


Sub Postsurvey()
'
' Insert novation quote text

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\4. Quotes\Quotes Source.docx", Range:="Postsurvey", Link:=False
        
End Sub
