Attribute VB_Name = "basRibbonCallbacks"

'################################################################
'#                                                              #
'#      Created with / Erstellt mit:                            #
'#      IDBE Ribbon Creator 2013                                #
'#      Version 1.1000                                          #
'#                                                              #
'#      (c) 2007-2013 IDBE Avenius                              #
'#                                                              #
'#      http://www.ribboncreator2013.com                        #
'#      http://www.ribboncreator.com                            #
'#      http://www.accessribon.com                              #
'#      http://www.avenius.com                                  #
'#                                                              #
'#      You may send change requests or report errors to:       #
'#      Aenderungswuensche oder Fehler bitte an:                #
'#                                                              #
'#      mailto://info@ribboncreator2010.com                     #
'#                                                              #
'################################################################


' Globals

Public gobjRibbon As IRibbonUI

Public bolEnabled As Boolean    ' Used in Callback "getEnabled"
                                ' Further informations in Callback "getEnabled"
                                ' Für Callback "getEnabled"
                                ' Genauere Informationen in Callback "getEnabled".
                               
Public bolVisible As Boolean    ' Used in Callback "getVisible"
                                ' More information in Callback "getVisible
                                ' Für Callback "getVisible"
                                ' Further informations in Callback "getVisible

' For Sample Callback "GetContent"
' Fuer Beispiel Callback "GetContent"
Public Type ItemsVal
    id As String
    label As String
    imageMso As String
End Type


' Callbacks

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
'Callbackname in XML File "onLoad"

    Set gobjRibbon = ribbon
End Sub

Public Sub OnActionButton(control As IRibbonControl)
'Callback in XML File "onAction"

    ' Callback for event button click
    ' Callback für Button Click
    
    Select Case control.id
        'Case "btnInfo"
        
        '
' control.id information comes from IDBE Ribbon Creator 2010
' menu key is stored in AddinFolder\Button references for addin.xslx
'
' Layout for control is:
    '   Brief description
    '   Case "button number"
    '       Macro name

        
' VA Letterhead
        Case "btn1"
            Letterhead
' VA Report Background
        Case "btn2"
            Reportbackground
' VA Logo
        Case "btn3"
            Logo
' VA Logo Large
        Case "btn217"
            Logolarge
            '
            '

' Report Custom Table
        Case "btn4"
            CustomTable
' Report Octave Band 63-8k
        Case "btn5"
            OctTable
' Report Octave Band 125-4k
        Case "btn6"
            ShortOctTable
' Report LA90 & LAeq Table
        Case "btn349"
            LA90Leqtable
' Report LA90 & LAeq Table - 2 positions
        Case "btn350"
            LA90Leqtable2pos
' Report Octave Band 125-4k
        Case "btn6"
            ShortOctTable
' ' Letter Custom Table
'        Case "btn7"
'            CustomTable
'' Letter Octave Band 63-8k
'        Case "btn8"
'            OctTableL
'' Letter Octave Band 125-4k
'        Case "btn9"
'            ShortOctTableL
            '
            '
' Quote Intro
        Case "btn10"
            Qintro
' Quote Section
        Case "btn11"
            QSectionheading
'' Quote Section Text Frontpage
'        Case "btn12"
'            QSectionfrontpage
' Quote Section Text
        Case "btn13"
            QSection
' Quote Section Subheading
        Case "btn14"
            QSubheading
' Quote Subheading
        Case "btn15"
            QSectionsubheading
' Quote Indented Section Text
        Case "btn16"
            QIndentedsection
' Table Heading
        Case "btn205"
                Tableheading
' Table Text
        Case "btn206"
                Tabletext
' Quote Table Title
        Case "btn207"
            Qtabletitle
' Quote Figure Number
        Case "btn214"
            Qfigure
'
' Report Chapter Title
        Case "btn17"
            RChapter
' Report Section Heading
        Case "btn18"
            RSectionheading
' Report Section Text
        Case "btn19"
            RSection
' Report Italic Subheading
        Case "btn20"
            RSubsection
' Report Subheading
        Case "btn317"
            RHeadingL4
' Report Bullet Point
        Case "btn21"
            RBullet
' Report Table Heading
        Case "btn22"
            Tableheading
' Report Table Text
        Case "btn23"
            Tabletext
' Report Table Title
        Case "btn24"
            RTabletitle
' Report Figure Number
        Case "btn215"
            Rfigure
'
' NIHIL Chapter Title
        Case "btn25"
            ExpertChapter
' NIHIL Section Text
        Case "btn26"
            ExpertSection
' NIHIL Section Text Indented
        Case "btn27"
            ExpertSectionIndent
' Table Heading
        Case "btn208"
                Tableheading
' Table Text
        Case "btn209"
                Tabletext
' Expert Table Title
        Case "btn210"
            Experttabletitle
' Expert Figure Number
        Case "btn216"
            Expertfigure

            '
            '
' 1 Page Letter
        Case "btn29"
            New1Page
' 2 Page Letter
        Case "btn30"
            New2Page
' Instruction Letter
        Case "btn31"
            Instruction
'' Warranty Letter
'        Case "btn32"
'            Warranty
' PCT Site Checklist
        Case "btn33"
            Checklist
' Debt Letter 1
        Case "btn267"
            DebtLetter1
' Debt Letter 2
        Case "btn268"
            DebtLetter2
' Debt Letter 3
        Case "btn269"
            DebtLetter3
 ' Debt Letter 4
        Case "btn270"
            DebtLetter4
            '
            '
            '
            
' New Quote
        Case "btn34"
            NewQuote
' Full Brief (Hotel) Quote
        Case "btn35"
            FullHotel
' Full Brief (Office) Quote
        Case "btn36"
            FullOffice
' Full Brief Residential Quote
        Case "btn37"
            FullResi
' PCT Quote
        Case "btn38"
            PCTQuote
' PCR Quote
        Case "btn39"
            PCRQuote
' NIA Local Authority Quote
        Case "btn40"
            LANIAQuote
' BS4142 Quote
        Case "btn41"
            BS4142Quote
' NPPF Quote
        Case "btn42"
            NPPFQuote
' NIA (NPPF) & EBF Quote
        Case "btn43"
            NIAEBFQuote
' Office to Resi Quote
        Case "btn294"
            OfficetoResiQuote
' Gym Quote
        Case "btn301"
            GymQuote
' Licencing/Leisure Quote
        Case "btn286"
            Licensing
' Change of Use: A1 to A3 Quote
        Case "btn318"
            A1toA3
' RBKC CMP Quote
        Case "btn262"
            RBKCCMPQuote
' Long Term Con/Dem Monitoring
        Case "btn354"
            LongTermQuote
' Windfarm Quote
        Case "btn44"
            ETSUQuote
' NAW Quote
        Case "btn45"
            NAWQuote
' Quote Introduction
        Case "btn211"
            Qintroduction
' Hourly Rates
        Case "btn46"
            Hourlyrates
' Licensed Premises
        Case "btn47"
            Licensedpremises
' ENS Scope
        Case "btn48"
            ENSsection
' BS4142 Scope (Commercial)
        Case "btn49"
            BS4142section
' BS4142 Scope (Resi)
        Case "btn366"
            BS4142resi
' NPPF Scope
        Case "btn50"
            NPPFsection
' EBF Scope
        Case "btn51"
            EBFsection
' Glazing Appraisal
        Case "btn52"
            Glazingappraisal
' Car Door Slam
        Case "btn323"
            Slam
' Room Acoustics
        Case "btn296"
            RoomAcoustics
' Building Services
        Case "btn302"
            BuildingServices
' PCR Scope
        Case "btn53"
            PCRsection
' Odour Assessment
        Case "btn278"
            Odourquote
' BS6472 Scope
        Case "btn54"
            Vibsection
' ADE Design Review Scope
        Case "btn55"
            ADEsection
' Investigative SIT Scope
        Case "btn56"
            SITsection
' Investigative SIT & ADE Design Review Scope
        Case "btn57"
            SITADEsection
' PCT Scope
        Case "btn58"
            PCTsection
' Office SI Tests
        Case "btn297"
            OfficeSI
' Reverbertation Time Tests
        Case "btn319"
            RT
' Noise Modelling Scope
        Case "btn59"
            Mappingsection
' Auralisation Scope
        Case "btn60"
            Auralisationsection
' Music Event Simulation
        Case "btn298"
            EventSimulation
' Noise Management Plan
        Case "btn299"
            NMP
' Traffic Noise (short)
        Case "btn61"
            Trafficshortsection
' Traffic Noise (long)
        Case "btn62"
            Trafficlongsection
' Construction Noise
        Case "btn63"
            Constructionsection
' CMP
        Case "btn64"
            CMP
' Section 61 Assessment
        Case "btn65"
            S61section
' Post Installation Survey (Plant)
        Case "btn371"
            PostNoiseSurvey
' Post works internal noise survey
        Case "btn_372"
            Postsurvey
' ES Chapter
        Case "btn66"
            ESsection
' Report Scope
        Case "btn67"
            Reportingsection
' Short Form Fees & Terms
        Case "btn68"
            Shortform
' SI/PCT Term and Conditions
        Case "btn295"
            SITerms
' Novation
        Case "btn263"
            novation
' ProPG
        Case "btn334"
            ProPG
' Internal Noise Criteria
        Case "btn343"
            INC
' EBF
        Case "btn344"
            EBFsection
' Room Acoustics
        Case "btn2966"
            RoomAcoustics
' Post Installation Noise Survey (Plant)
        Case "btn371"
            PostNoiseSurvey
            '
            '
'CVs
            '
' JD
        Case "btn69"
            JD
' SL
        Case "btn70"
            SL
            '
            '
            '
            '
' Standard Report
        Case "btn71"
            StandardReport
'Introduction
        Case "btn75"
            ReportIntro
'NPPF
        Case "btn79"
            ReportNPPF
'BS8233:2014
        Case "btn80"
            Report8233
'WHO:1999
        Case "btn81"
            ReportWHO
'ADO
        Case "btn_373"
            ADO
'AVO Guidance
        Case "btn347"
            AVO
'ProPG Criteria
        Case "btn348"
            ProPGCriteria
'BS4142:1997
        Case "btn82"
            Report41421997
'BS4142:2014
        Case "btn83"
            Report41422014
'BS4142:2014 long
        Case "btn212"
            Report41422014long
'Class AA
        Case "btn363"
            ClassAA
'Class AB
        Case "btn364"
            ClassAB
'Class MA
        Case "btn365"
            ClassMA
'Class O
        Case "btn303"
            ClassO
'Class Q
        Case "btn367"
            ClassQ
'IOA Pub/Music Noise
        Case "btn324"
            IOAMusic
'NANR45
        Case "btn361"
            NANR
'BS6472-1:2008
        Case "btn84"
            Report6472
'BS5228:2009
        Case "btn85"
            Report5228
'CRTN
        Case "btn86"
            ReportCRTN
'MPS2
        Case "btn87"
            ReportMPS2
'DMRB
        Case "btn88"
            ReportDMRB
'IEMA
        Case "btn89"
            ReportIEMA
'NAW
        Case "btn90"
            ReportNAW
            '
'Defra Odour Guidance (intro)
        Case "btn284"
            Defraodourintro
            '
'Defra Odour Guidance
        Case "btn285"
            Defraodour
            '
'BREEAM 2006 HW17
        Case "btn91"
            ReportHW17
'BREEAM 2006 P13
        Case "btn92"
            ReportP13
'BREEAM 2008 Pol 8
        Case "btn93"
            Report2008Pol8
'BREEAM 2008 Hea 13 Courts
        Case "btn94"
            Report2008Hea13Court
'BREEAM 2008 Hea 13 Education
        Case "btn95"
            Report2008Hea13Edu
'BREEAM 2008 Hea 13 Healthcare
        Case "btn96"
            Report2008Hea13Health
'BREEAM 2008 Hea 13 Industrial
        Case "btn97"
            Report2008Hea13Industrial
'BREEAM 2008 Hea 13 Offices
        Case "btn98"
            Report2008Hea13Offices
'BREEAM 2008 Hea 13 Prisons
        Case "btn99"
            Report2008Hea13Prisons
'BREEAM 2008 Hea 13 Retail
        Case "btn100"
            Report2008Hea13Retail
'BREEAM 2011 Pol 05
        Case "btn101"
            Report2011Pol05
'BREEAM 2011 Hea 05 Schools
        Case "btn102"
            Report2011Hea5Schools
'BREEAM 2011 Hea 05 Further Education
        Case "btn103"
            Report2011Hea5HE
'BREEAM 2011 Hea 05 Healthcare
        Case "btn104"
            Report2011Hea5Health
'BREEAM 2011 Hea 05 Residential
        Case "btn105"
            Report2011Hea5Resi
'BREEAM 2011 Hea 05 Other
        Case "btn106"
            Report2011Hea5Other
'BREEAM 2014 Pol 05
        Case "btn107"
            Report2014Pol05
'BREEAM 2014 Hea 05 Education
        Case "btn108"
            Report2014Hea05Education
'BREEAM 2014 Hea 05 Healthcare
        Case "btn109"
            Report2014Hea05Health
'BREEAM 2014 Hea 05 Office
        Case "btn110"
            Report2014Hea05Office
'BREEAM 2014 Hea 05 Law Courts
        Case "btn111"
            Report2014Hea05Law
'BREEAM 2014 Hea 05 Industrial
        Case "btn112"
            Report2014Hea05Industrial
'BREEAM 2014 Hea 05 Residential
        Case "btn113"
            Report2014Hea05Resi
'EcoHomes 2006
        Case "btn114"
            ReportEco2006
'CfSH 2008
        Case "btn115"
            ReportCfSH2008
'CfSH 2010
        Case "btn116"
            ReportCfSH2010
'BCO 2005
        Case "btn117"
            ReportBCO2005
'BCO 2009
        Case "btn118"
            ReportBCO2009
'BCO 2011
        Case "btn119"
            ReportBCO2011
'Islington Music Noise
        Case "btn360"
            Islingtonmusic
'ENS
        Case "btn76"
            ReportENS
'NIA
        Case "btn77"
            ReportNIA
'Description of background
        Case "btn_394"
            L90eqdescription
'Spectral Adaption for Plant
        Case "btn374"
            Spectraladaptation
'BS7445 Assessment
        Case "btn_377"
            BS7445assess
'BS4142 Assessment
        Case "btn_378"
            BS4142assess
'BS4142 Context Discussion
        Case "btn_379"
            BS4142context
'Commercial Noise Assessment - Internal
        Case "btn_380"
            Throughwindow
'Nuisance
        Case "btn_395"
            Nuisance
'Minerals
        Case "btn_396"
            Minerals
'Noise Mapping
        Case "btn368"
            Noisemapping
' Basis of Criteria
        Case "btn213"
            Basis
'EBF
        Case "btn78"
            ReportEBF
'Ventilation/Overheating
        Case "btn_384"
            Overheating
'AVOG Risk Assessment
        Case "btn_385"
            AVOGRisk
'Car Door Slams & Movements
        Case "btn327"
            Carmovements
'Party Wall Junction
        Case "btn321"
            Partywalljunction
'Glazed Partition Junction
        Case "btn322"
            Glazedjunction
'BREEAM Hea 05 Criteria Report - Office
        Case "btn_383"
            Hea05CritOffice
'BREEAM Hea 05 Design Stage Report - Office
        Case "btn_384000"
            Hea05DesOffice
'BREEAM Pol 05 Report
        Case "btn_385000"
            Pol05report
'Stage 3 Report - Residential
        Case "btn_386"
            S3resi
'Stage 4 Report - Residential
        Case "btn_387"
            S4resi
'Stage 3 Report - Office
        Case "btn_388000"
            Hea05CritOffice1
'Stage 4 Report - Office
        Case "btn_389000"
            Hea05DesOffice
'Stage 3 Report - Mixed Use
        Case "btn_390000"
            S3Mixed
'Stage 4 Report - MixedUse
        Case "btn_391"
            S4Mixed
'Stage 4 report - BB93, New build
        Case "btn_392"
            BB93new
'Stage 4 report - BB93, Refurbishment
        Case "btn_393"
            BB93refurb
'Lift Noise
        Case "btn120"
            ReportLift
'SVP
        Case "btn121"
            ReportSVP
'Tenancy Agreements
        Case "btn122"
            ReportTenancy
'Hard Floor vs Carpet
        Case "btn264"
            carpet
'Hard Floor vs Carpet (long)
        Case "btn265"
            carpetlong
'Section 7: Reverberation in common parts
        Case "btn328"
            Section7
'Reverberation Time Review
        Case "btn_388"
            RTReview
'SI Review (Commercial to Resi)
        Case "btn372"
            SICommtoResi
'SOffice Review (Privacy)
        Case "btn373"
            OfficeSIPrivacy
'Office 'Standard' Constructions
        Case "btn_397"
            Officepartitions
'ADE New Build
        Case "btn_389"
            NewBuild
'ADE Conversion
        Case "btn_390"
            Conversion
' ProPG Risk Assessment
        Case "btn369"
            ProPGRiskAssessment
' Beer Garden
        Case "btn370"
            BeerGarden
' In to Out/Breakout
        Case "btn_369"
            IntoOut
' Night-time Pub Noise
        Case "btn351"
            Pubnoise
' Glazing - hyrbrid approach (Class MA)
        Case "btn352"
            Hybridapproach
' Ground Noise Predictions
        Case "btn353"
            groundborne
' PCR - Westminster
        Case "btn123"
            PCR
' NIA - Camden
        Case "btn124"
            NIACamden
' NIA - RBKC
        Case "btn125"
            NIARBKC
' NIA - City of London
        Case "btn126"
            NIACoL
' NIA - Local Authority
        Case "btn127"
            NIALA
' NPPF & EBF Report
        Case "btn128"
            NPPFEBF
' NPPF, BS6472 & EBF Report
        Case "btn129"
            NPPF6472EBF
' BS4142 Report
        Case "btn130"
            BS4142Report
' Class MA Caonversion
        Case "btn345"
            ClassMAReport
' Ealing Report
        Case "btn358"
            EalingReport
' CMP Report
        Case "btn355"
            CMPReport
' CMP (Camden) Report
        Case "btn313"
            CMPCamden
' CMP (RBKC) Report
        Case "btn346"
            CMPRBKC
' PCT Wall Report
        Case "btn131"
            PCTWall
' PCT Floor Report
        Case "btn132"
            PCTFloor
' PCT Wall & Floor Report
        Case "btn133"
            PCTWallandFloor
' ADE SI Review
        Case "btn316"
            ADE
' NAW Report
        Case "btn290"
            NAW
' NIHL Report
        Case "btn72"
            NIHL
' Gym Report
        Case "btn311"
            Gym
' NMP/SI Report
        Case "btn312"
            NMPReport
' Odour Report
        Case "btn_257"
            Odour
' General Specification Report
        Case "btn73"
            GenSpec
' Insert Equipment Table
        Case "btn74"
            equip_table
            '
            '
' Plant Noise Schedule
        Case "btn134"
            PNS
' Anti-vibration Mount Schedule
        Case "btn135"
            AVM
' Fan Coil Unit Schedule
        Case "btn136"
            FCU
' Roomside Silencer Schedule
        Case "btn137"
            RSS
' Atmospheric Silencer Scedule
        Case "btn138"
            ASS
' Plantoom Structural Schedule
        Case "btn139"
            PRS
' Lift Specification
        Case "btn141"
            Lifts
' WHO Tables
        Case "btn142"
            NewWHO
' BB93 Tables
        Case "btn143"
            Newbb93
' A3 Figure Landcape
        Case "btn145"
            A3landscape
' Landscape Figure
        Case "btn146"
            Landscape
' Portrait Figure
        Case "btn147"
            Portrait
' Drawing Issue Sheet
        Case "btn148"
            Drawingissue
' Appendix A
        Case "btn150"
            AppendixA
' Appendix Facer
        Case "btn289"
            AppendixFacer
' New Appendix
        Case "btn151"
            NewAppendix
' Survey Sheet
        Case "btn152"
            Surveysheet
            '
            '
            '
' Environmental Noise Survey RAMS
        Case "btn304"
            ENSRAMS
' Noise at Work RAMS
        Case "btn305"
            NAWRAMS
' Sound Insulation Test RAMS
        Case "btn306"
            SITRAMS
' Environmental Policy
        Case "btn308"
            EnvPol
' Health and Safety Policy
        Case "btn309"
            HASPol
' Quality Policy
        Case "btn310"
            QalPol
' GDPR Policy
        Case "btn307"
            GDPR
            '
            '
            '
' Inserts "Further to..."
        Case "btn153"
            Further
' Inserts "We trust..."
        Case "btn154"
            Trust
' Insert ADE 1992 text
        Case "btn155"
            ADE1992
' Inserts ADE 2003 text
        Case "btn156"
            ADE2003
' Inserts ANC Noise from Building Services text
        Case "btn157"
            ANCBS
' Inserts BB93 text
        Case "btn158"
            BB93
' Inserts BS8233 2014 text
        Case "btn159"
           BS82332014
' Inserts BS8233 2014 table
        Case "btn160"
           BS82332014Table
' Inserts BS8233 1999 text
        Case "btn161"
            BS8233
' Inserts BS8233 Table
        Case "btn162"
            BS8233Table
' Inserts HTM 08-01 rebadged
        Case "btn163"
            ATDM4032
' Inserts HTM 08-01
        Case "btn164"
            HTM0801
' Inserts HTM2045
        Case "btn165"
            HTM2045
' Inserts BREEAM 2011
        Case "btn166"
            BREEAMNC
' Inserts BREEAM Education
        Case "btn167"
            BREEAMEducation
' Inserts BREEAM Healthcare
        Case "btn168"
            BREEAMHealth
' Inserts BREEAM Offices
        Case "btn169"
            BREEAMOffice
' Inserts EPA
        Case "btn170"
            EPA
' Inserts BS7445-1
        Case "btn171"
            BS7445P1
' Insert BS7445-2
        Case "btn172"
            BS7445P2
' Inserts BS4142 2014
        Case "btn173"
            BS41422014
' Inserts BS4142 1997
        Case "btn174"
            BS4142
' Inserts BS5228-1
        Case "btn175"
            BS5228P1
' Inserts BS5228-2
        Case "btn176"
            BS5228P2
' Inserts BS5228-4
        Case "btn177"
            BS5228P4
' Inserts CRTN
        Case "btn178"
            CRTN
' Inserts DMRB (1994)
        Case "btn179"
            DMRB1994
' Inserts DMRB (2011)
        Case "btn180"
            DMRB2011
' Inserts ISO 10140-2
        Case "btn181"
            ISO10140P2
' Inserts ISO 10140-3
        Case "btn182"
            ISO10140P3
' Inserts ISO 10848-2
        Case "btn183"
            ISO10848
' Inserts ISO 140-3
        Case "btn184"
            ISO140P3
' Inserts ISO 140-4
        Case "btn185"
            ISO140P4
' Inserts ISO 140-5
        Case "btn186"
            ISO140P5
' Inserts ISO 140-7
        Case "btn187"
            ISO140P7
' Inserts ISO 140-10
        Case "btn188"
            ISO140P10
' Inserts ISO 717-1
        Case "btn189"
            ISO717P1
' Inserts ISO 717-2
        Case "btn190"
            ISO717P2
' Inserts ISO 16283-1
        Case "btn191"
            ISO16283
' Inserts ISO 16283-2
        Case "btn192"
            ISO162832
' Inserts NPPF
        Case "btn193"
            NPPF
' Inserts PPG24
        Case "btn194"
            PPG24
' Inserts WHO 1980
        Case "btn195"
            WHO1980
' Inserts WHO 1999
        Case "btn196"
            WHO1999
' Inserts BS6472-1
        Case "btn197"
            BS6472P1
' Inserts BS6742-2
        Case "btn198"
            BS6472P2
' Inserts BS6841
        Case "btn199"
            BS6841
' Inserts BS7385
        Case "btn200"
            BS7385
' Inserts CoP Concerts 1995
        Case "btn201"
            CoP
' Inserts Draft CoP Pubs 1998
        Case "btn202"
            DCop1998
' Inserts CoP Pubs 2003
        Case "btn203"
            GPGPubs2003
' Inserts Clay Target Shooting
        Case "btn204"
            CTS
            
' Selects all and updates language
        Case "btn362"
            UpdateLanguage
        
        
        Case Else
            MsgBox "Button """ & control.id & """ clicked" & vbCrLf & _
                           "Es wurde auf Button """ & control.id & """ in Ribbon geklickt", _
                           vbInformation
    End Select
End Sub

'Command Button

Sub OnActionButtonHelp(control As IRibbonControl, ByRef CancelDefault)
    ' Callbackname in XML File Command "onAction"

    ' Callback for command event button click
    ' Callback fuer Command Button Click

    MsgBox "Button ""Help"" clicked" & vbCrLf & _
                           "Es wurde auf Button ""Hilfe"" geklickt", _
                           vbInformation
    CancelDefault = True

End Sub

Sub OnActionCheckBox(control As IRibbonControl, _
                               pressed As Boolean)
    ' Callbackname in XML File "OnActionCheckBox"
    
    ' Callback for event checkbox click
    ' Callback für Checkbox Click

    Select Case control.id
        'Case "chkMyCheckbox"
        '    If pressed = True Then
        '
        '    Else
        '
        '    End If
        '
        Case Else
            MsgBox "The Value of the Checkbox """ & control.id & """ is: " & pressed & vbCrLf & _
                   "Der Wert der Checkbox """ & control.id & """ ist: " & pressed, _
                   vbInformation
    End Select

End Sub

Sub GetPressedCheckBox(control As IRibbonControl, _
                       ByRef bolReturn)
    
    ' Callbackname in XML File "GetPressedCheckBox"
    
    ' Callback for checkbox
    ' indicates how the control is displayed
    ' Callback für Checkbox wie das Control
    ' angezeigt werden soll

    Select Case control.id
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                bolReturn = True
            Else
                bolReturn = False
            End If
    End Select

End Sub


Sub OnActionTglButton(control As IRibbonControl, _
                       pressed As Boolean)
                              
    ' Callbackname in XML File "onAction"
    
    ' Callback für einen Toggle Button Klick
    ' Callback for a Toggle Buttons click event

    Select Case control.id
        '    If pressed = True Then
        '
        '    Else
        '
        '    End If
        
   ' Toggles Bookmarks on and off
        Case "tgb1"
            If pressed = True Then
                Bookmarkon
            Else
                Bookmarkoff
            End If
            
            
        Case Else
            MsgBox "The Value of the Toggle Button """ & control.id & """ is: " & pressed & vbCrLf & _
                   "Der Wert der Toggle Button """ & control.id & """ ist: " & pressed, _
                   vbInformation
    End Select

End Sub

Sub GetPressedTglButton(control As IRibbonControl, _
                       ByRef pressed)
' Callbackname in XML File "getPressed"

' Callback für ein Access ToogleButton Control wie dieser Angezeigt werden soll
' Callback for an Access ToogleButton Control. Indicates how the control is displayed

    Select Case control.id
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = True
            Else
                pressed = False
            End If
    End Select
End Sub

Public Sub GetEnabled(control As IRibbonControl, ByRef enabled)
    ' Callbackname in XML File "getEnabled"
    
    ' To set the property "enabled" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Enabled Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        'Case "ID_XMLRibbControl"
        '    enabled = bolEnabled
        Case Else
            enabled = True
    End Select
End Sub

Public Sub GetVisible(control As IRibbonControl, ByRef visible)
    ' Callbackname in XML File "getVisible"
    
    ' To set the property "visible" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Visible Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        'Case "ID_XMLRibbControl"
        '    visible = bolVisible
        Case Else
            visible = True
    End Select
End Sub

Sub GetLabel(control As IRibbonControl, ByRef label)
    ' Callbackname in XML File "getLabel"
    ' To set the property "label" to a Ribbon Control

    Select Case control.id
        ''GetLabel''
        Case Else
            label = "*getLabel*"

    End Select

End Sub

Sub GetScreentip(control As IRibbonControl, ByRef screentip)
    ' Callbackname in XML File "getScreentip"
    ' To set the property "screentip" to a Ribbon Control

    Select Case control.id
        ''GetScreentip''
        Case Else
            screentip = "*getScreentip*"

    End Select

End Sub

Sub GetSupertip(control As IRibbonControl, ByRef supertip)
    ' Callbackname in XML File "getSupertip"
    ' To set the property "supertip" to a Ribbon Control

    Select Case control.id
        ''GetSupertip''
        Case Else
            supertip = "*getSupertip*"

    End Select

End Sub

Sub GetDescription(control As IRibbonControl, ByRef description)
    ' Callbackname in XML File "getDescription"
    ' To set the property "description" to a Ribbon Control

    Select Case control.id
        ''GetDescription''
        Case Else
            description = "*getDescription*"

    End Select

End Sub

Sub GetTitle(control As IRibbonControl, ByRef Title)
    ' Callbackname in XML File "getTitle"
    ' To set the property "title" to a Ribbon MenuSeparator Control

    Select Case control.id
        ''GetTitle''
        Case Else
            Title = "*getTitle*"

    End Select

End Sub

'EditBox

Sub GetTextEditBox(control As IRibbonControl, _
                             ByRef strText)
    ' Callbackname in XML File "GetTextEditBox"
    
    ' Callback für EditBox welcher Wert in der
    ' EditBox eingetragen werden soll.
    ' Callback for an EditBox Control
    ' Indicates which value is to set to the control

    Select Case control.id
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select
    
End Sub

Sub OnChangeEditBox(control As IRibbonControl, _
                              strText As String)
    ' Callbackname in XML File "OnChangeEditBox"
    
    ' Callback Editbox: Rückgabewert der Editbox
    ' Callback Editbox: Return value of the Editbox

    Select Case control.id
        'Case "MyEbx"
            'If strText = "Password" Then
            '
            'End If
        Case Else
            MsgBox "The Value of the EditBox """ & control.id & """ is: " & strText & vbCrLf & _
                   "Der Wert der EditBox """ & control.id & """ ist: " & strText, _
                   vbInformation
    End Select

End Sub

'DropDown

Sub OnActionDropDown(control As IRibbonControl, _
                             selectedId As String, _
                             selectedIndex As Integer)
    ' Callbackname in XML File "OnActionDropDown"
    
    ' Callback onAction (DropDown)
    
    Select Case selectedId
        'Case "MyItemID"
        '
        Case "ddc_405Item0"
        Allmarkup
        Case "ddc_405Item1"
        Nomarkup
        
        Case Else
            MsgBox "The selected ItemID of DropDown-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                   "Die selektierte ItemID des DropDown-Control """ & control.id & """ ist : """ & selectedId & """", _
                   vbInformation
    End Select

End Sub

Sub GetSelectedItemIndexDropDown(control As IRibbonControl, _
                                   ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexDropDown"
    
    ' Callback getSelectedItemIndex (DropDown)
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id
            Case Else
                index = varIndex
        End Select
    End If

End Sub

'Gallery

Sub OnActionGallery(control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionGallery"
    
    ' Callback onAction (Gallery)
    
    Select Case control.id
        'Case "MyGalleryID"
        '   Select Case selectedId
        '      Case "MyGalleryItemID"
        '
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of Gallery-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des Gallery-Control """ & control.id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub

Sub GetSelectedItemIndexGallery(control As IRibbonControl, _
                                   ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexGallery"
    
    ' Callback getSelectedItemIndex (Gallery)
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id

            Case Else
                index = varIndex

        End Select

    End If

End Sub

'Combobox

Sub GetTextComboBox(control As IRibbonControl, _
                      ByRef strText)

    ' Callbackname im XML File "GetTextComboBox"
    
    ' Callback getText (Combobox)
                           
    Select Case control.id
        
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select

End Sub


Sub OnChangeComboBox(control As IRibbonControl, _
                               strText As String)
                           
    ' Callbackname im XML File "OnChangeCombobox"
    
    ' Callback onChange (Combobox)
   
    Select Case control.id
        
        Case Else
            MsgBox "The selected Item of Combobox-Control """ & control.id & """ is : """ & strText & """" & vbCrLf & _
                   "Das selektierte Item des Combobox-Control """ & control.id & """ ist : """ & strText & """", _
                   vbInformation
    End Select

End Sub


' DynamicMenu

Sub GetContent(control As IRibbonControl, _
               ByRef XMLString)

    ' Sample for a Ribbon XML "getContent" Callback
    ' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    ' Beispiel fuer einen Ribbon XML - "getContent" Callback
    ' Siehe auch: http://www.accessribbon.de/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '       und : http://www.accessribbon.de/?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    Select Case control.id

        Case Else
            XMLString = getXMLForDynamicMenu()
    End Select
 
End Sub


' Helper Function
' Hilfsfunktionen

Public Function getXMLForDynamicMenu() As String
    
    ' Creates a XML String for DynamicMenu CallBack - getContent
    
    ' Erstellt den Inhalt fuer das DynamicMenu im Callback getContent
    
    Dim lngDummy    As Long
    Dim strDummy    As String
    Dim strContent  As String
    
    Dim Items(4) As ItemsVal
    Items(0).id = "btnDy1"
    Items(0).label = "Item 1"
    Items(0).imageMso = "_1"
    Items(1).id = "btnDy2"
    Items(1).label = "Item 2"
    Items(1).imageMso = "_2"
    Items(2).id = "btnDy3"
    Items(2).label = "Item 3"
    Items(2).imageMso = "_3"
    Items(3).id = "btnDy4"
    Items(3).label = "Item 4"
    Items(3).imageMso = "_4"
    Items(4).id = "btnDy5"
    Items(4).label = "Item 5"
    Items(4).imageMso = "_5"
    
    strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    
        For lngDummy = LBound(Items) To UBound(Items)
            strContent = strContent & _
                "<button id=""" & Items(lngDummy).id & """" & _
                " label=""" & Items(lngDummy).label & """" & _
                " imageMso=""" & Items(lngDummy).imageMso & """" & _
                " onAction=""OnActionButton""/>" & vbCrLf
        Next
 

    strDummy = strDummy & strContent & "</menu>"
    getXMLForDynamicMenu = strDummy

End Function

Public Function getTheValue(strTag As String, strValue As String) As String
   ' *************************************************************
   ' Erstellt von     : Avenius
   ' Parameter        : Input String, SuchValue String
   ' Erstellungsdatum : 05.01.2008
   ' Bemerkungen      :
   ' Änderungen       :
   '
   ' Beispiel
   ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
   ' Return           : "Test"
   ' *************************************************************
      
   On Error Resume Next
      
   Dim workTb()     As String
   Dim Ele()        As String
   Dim myVariabs()  As String
   Dim i            As Integer

      workTb = Split(strTag, ";")
      
      ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
      For i = LBound(workTb) To UBound(workTb)
         Ele = Split(workTb(i), ":=")
         myVariabs(i, 0) = Ele(0)
         If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
         End If
      Next
      
      For i = LBound(myVariabs) To UBound(myVariabs)
         If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
         End If
      Next
      
End Function







