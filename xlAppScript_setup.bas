Attribute VB_Name = "xlAppScript_setup"
'/\______________________________________________________________________________________________________________________
'//
'//     xlAppScript Setup
'/\_____________________________________________________________________________________________________________________________
'//
'//     License Information:
'//
'//     Copyright (C) 2022-present, Autokit Technology.
'//
'//     Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'//
'//     1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'//
'//     2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'//
'//     3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
'//
'//     THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
'//     THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'//     (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
'//     HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'//
'/\_____________________________________________________________________________________________________________________________
'//
'//     Latest Revision: 4/22/2022
'/\_____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re
'//     Contact: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
'//     Web: xlappscript.org | autokit.tech
'/\_____________________________________________________________________________________________________________________________

Public Function connectWb()

'/\__________________________________________________________________________________________
'//
'//     A function for setting up a runtime environment (workbook) to interact w/ xlAppScript
'/\__________________________________________________________________________________________

On Error GoTo ErrMsg

nl = vbNewLine

'//Label script memory addresses (Very important!)

'//Multi-memory block specific addresses
Range("MAA1").Name = "xlasKinLabelMod"
Range("MAB1").Name = "xlasKinValueMod"
Range("MAC1").Name = "xlasKinLabel"
Range("MAD1").Name = "xlasKinValue"
Range("MAE1").Name = "xlasState"
Range("MAF1").Name = "xlasArticle"
Range("MAG1").Name = "xlasGroup"
Range("MAH1").Name = "xlasList" '//(3/15/2022)
Range("MAL1").Name = "xlasLib"
'//Single-memory block specific addresses
Range("MAS1").Name = "xlasAppLoad" '//Autokit applications(2/9/2022)
Range("MAS2").Name = "xlasEnvironment"
Range("MAS3").Name = "xlasBlock" '//(2/28/2022)
Range("MAS4").Name = "xlasGoto"
Range("MAS5").Name = "xlasInvert" '//ctrl box+
Range("MAS6").Name = "xlasKeyCtrl" '//Autokit applications
Range("MAS7").Name = "xlasRemember" '//ctrl box+
Range("MAS8").Name = "xlasConsoleType" '//ctrl box+
Range("MAS9").Name = "xlasAMemory" '//ctrl box+
Range("MAS10").Name = "xlasSaveFile" '//ctrl box+
Range("MAS11").Name = "xlasSilent" '//Autokit applications
Range("MAS12").Name = "xlasCtrlBoxFColor" '//ctrl box+
Range("MAS13").Name = "xlasCtrlBoxBColor" '//ctrl box+
Range("MAS14").Name = "xlasGlobalControl" '//4/18/2022
Range("MAS15").Name = "xlasLocalContain" '//(3/17/2022)
Range("MAS16").Name = "xlasLocalStatic" '//(3/13/2022)
Range("MAS17").Name = "xlasUpdateEnable" '//(2/24/2022)
Range("MAS18").Name = "xlasWinForm" '//Autokit applications
Range("MAS19").Name = "xlasWinFormLast" '//Autokit applications
Range("MAS20").Name = "xlasLibErrLvl" '//4/22/2022
Range("MAS21").Name = "xlasErrRef"
Range("MAS22").Name = "xlasEnd"
Range("MAS23").Name = "xlasLink": Range("xlasLink").Value = 1
Range("MAS79").Name = "xlasBlkAddr79" '//3/17/2022
Range("MAS80").Name = "xlasBlkAddr80" '//3/17/2022
Range("MAS81").Name = "xlasBlkAddr81" '//3/17/2022
Range("MAS82").Name = "xlasBlkAddr82" '//3/17/2022
Range("MAS83").Name = "xlasBlkAddr83" '//3/17/2022
Range("MAS84").Name = "xlasBlkAddr84" '//3/17/2022
Range("MAS85").Name = "xlasBlkAddr85" '//3/17/2022
Range("MAS86").Name = "xlasBlkAddr86" '//3/17/2022
Range("MAS87").Name = "xlasBlkAddr87" '//3/17/2022
Range("MAS88").Name = "xlasBlkAddr88" '//3/17/2022
Range("MAS89").Name = "xlasBlkAddr89" '//3/17/2022
Range("MAS90").Name = "xlasBlkAddr90" '//3/17/2022
Range("MAS91").Name = "xlasBlkAddr91" '//3/17/2022
Range("MAS92").Name = "xlasBlkAddr92" '//3/17/2022
Range("MAS93").Name = "xlasBlkAddr93" '//3/17/2022
Range("MAS94").Name = "xlasBlkAddr94" '//3/17/2022
Range("MAS95").Name = "xlasBlkAddr95" '//3/17/2022
Range("MAS96").Name = "xlasBlkAddr96" '//3/17/2022
Range("MAS97").Name = "xlasBlkAddr97" '//3/17/2022
Range("MAS98").Name = "xlasBlkAddr98" '//3/17/2022
Range("MAS99").Name = "xlasBlkAddr99" '//3/17/2022
Range("MAS100").Name = "xlasBlkAddr100" '//4/21/2022
Range("MAS101").Name = "xlasBlkAddr101" '//4/21/2022
Range("MAS102").Name = "xlasBlkAddr102" '//4/21/2022
Range("MAS103").Name = "xlasBlkAddr103" '//4/21/2022
Range("MAS104").Name = "xlasBlkAddr104" '//4/21/2022
Range("MAS105").Name = "xlasBlkAddr105" '//4/21/2022
Range("MAS106").Name = "xlasBlkAddr106" '//4/21/2022
Range("MAS107").Name = "xlasBlkAddr107" '//4/21/2022
Range("MAS108").Name = "xlasBlkAddr108" '//4/21/2022
Range("MAS109").Name = "xlasBlkAddr109" '//4/21/2022
Range("MAS110").Name = "xlasBlkAddr110" '//4/21/2022
Range("MAS111").Name = "xlasBlkAddr111" '//4/21/2022
Range("MAS112").Name = "xlasBlkAddr112" '//4/21/2022
Range("MAS113").Name = "xlasBlkAddr113" '//4/21/2022
Range("MAS114").Name = "xlasBlkAddr114" '//4/21/2022
Range("MAS115").Name = "xlasBlkAddr115" '//4/21/2022
Range("MAS116").Name = "xlasBlkAddr116" '//4/21/2022
Range("MAS117").Name = "xlasBlkAddr117" '//4/21/2022
Range("MAS118").Name = "xlasBlkAddr118" '//4/21/2022
Range("MAS119").Name = "xlasBlkAddr119" '//4/21/2022
Range("MAS120").Name = "xlasBlkAddr120" '//4/21/2022
Range("MAS121").Name = "xlasBlkAddr121" '//4/21/2022
Range("MAS122").Name = "xlasBlkAddr122" '//4/21/2022
Range("MAS123").Name = "xlasBlkAddr123" '//4/21/2022
Range("MAS124").Name = "xlasBlkAddr124" '//4/21/2022
Range("MAS125").Name = "xlasBlkAddr125" '//4/21/2022
Range("MAS126").Name = "xlasBlkAddr126" '//4/21/2022
Range("MAS127").Name = "xlasBlkAddr127" '//4/21/2022
Range("MAS128").Name = "xlasBlkAddr128" '//4/21/2022
Range("MAS129").Name = "xlasBlkAddr129" '//4/21/2022
Range("MAS130").Name = "xlasBlkAddr130" '//4/21/2022
Range("MAS131").Name = "xlasBlkAddr131" '//4/21/2022
Range("MAS132").Name = "xlasBlkAddr132" '//4/21/2022
Range("MAS133").Name = "xlasBlkAddr133" '//4/21/2022
Range("MAS134").Name = "xlasBlkAddr134" '//4/21/2022
Range("MAS135").Name = "xlasBlkAddr135" '//4/21/2022
Range("MAS136").Name = "xlasBlkAddr136" '//4/21/2022
Range("MAS137").Name = "xlasBlkAddr137" '//4/21/2022
Range("MAS138").Name = "xlasBlkAddr138" '//4/21/2022
Range("MAS139").Name = "xlasBlkAddr139" '//4/21/2022
Range("MAS140").Name = "xlasBlkAddr140" '//4/21/2022
Range("MAS141").Name = "xlasBlkAddr141" '//4/21/2022
Range("MAS142").Name = "xlasBlkAddr142" '//4/21/2022
Range("MAS143").Name = "xlasBlkAddr143" '//4/21/2022
Range("MAS144").Name = "xlasBlkAddr144" '//4/21/2022
Range("MAS145").Name = "xlasBlkAddr145" '//4/21/2022
Range("MAS146").Name = "xlasBlkAddr146" '//4/21/2022
Range("MAS147").Name = "xlasBlkAddr147" '//4/21/2022
Range("MAS148").Name = "xlasBlkAddr148" '//4/21/2022
Range("MAS149").Name = "xlasBlkAddr149" '//4/21/2022
Range("MAS150").Name = "xlasBlkAddr150" '//4/21/2022
Range("MAS151").Name = "xlasBlkAddr151" '//4/21/2022
Range("MAS152").Name = "xlasBlkAddr152" '//4/21/2022
Range("MAS153").Name = "xlasBlkAddr153" '//4/21/2022
Range("MAS154").Name = "xlasBlkAddr154" '//4/21/2022
Range("MAS155").Name = "xlasBlkAddr155" '//4/21/2022
Range("MAS156").Name = "xlasBlkAddr156" '//4/21/2022
Range("MAS157").Name = "xlasBlkAddr157" '//4/21/2022
Range("MAS158").Name = "xlasBlkAddr158" '//4/21/2022
Range("MAS159").Name = "xlasBlkAddr159" '//4/21/2022
Range("MAS160").Name = "xlasBlkAddr160" '//4/21/2022
Range("MAS161").Name = "xlasBlkAddr161" '//4/21/2022
Range("MAS162").Name = "xlasBlkAddr162" '//4/21/2022
Range("MAS163").Name = "xlasBlkAddr163" '//4/21/2022
Range("MAS164").Name = "xlasBlkAddr164" '//4/21/2022
Range("MAS165").Name = "xlasBlkAddr165" '//4/21/2022
Range("MAS166").Name = "xlasBlkAddr166" '//4/21/2022
Range("MAS167").Name = "xlasBlkAddr167" '//4/21/2022
Range("MAS168").Name = "xlasBlkAddr168" '//4/21/2022
Range("MAS169").Name = "xlasBlkAddr169" '//4/21/2022
Range("MAS170").Name = "xlasBlkAddr170" '//4/21/2022
Range("MAS171").Name = "xlasBlkAddr171" '//4/21/2022
Range("MAS172").Name = "xlasBlkAddr172" '//4/21/2022
Range("MAS173").Name = "xlasBlkAddr173" '//4/21/2022
Range("MAS174").Name = "xlasBlkAddr174" '//4/21/2022
Range("MAS175").Name = "xlasBlkAddr175" '//4/21/2022
Range("MAS176").Name = "xlasBlkAddr176" '//4/21/2022
Range("MAS177").Name = "xlasBlkAddr177" '//4/21/2022
Range("MAS178").Name = "xlasBlkAddr178" '//4/21/2022
Range("MAS179").Name = "xlasBlkAddr179" '//4/21/2022
Range("MAS180").Name = "xlasBlkAddr180" '//4/21/2022
Range("MAS181").Name = "xlasBlkAddr181" '//4/21/2022
Range("MAS182").Name = "xlasBlkAddr182" '//4/21/2022
Range("MAS183").Name = "xlasBlkAddr183" '//4/21/2022
Range("MAS184").Name = "xlasBlkAddr184" '//4/21/2022
Range("MAS185").Name = "xlasBlkAddr185" '//4/21/2022
Range("MAS186").Name = "xlasBlkAddr186" '//4/21/2022
Range("MAS187").Name = "xlasBlkAddr187" '//4/21/2022
Range("MAS188").Name = "xlasBlkAddr188" '//4/21/2022
Range("MAS189").Name = "xlasBlkAddr189" '//4/21/2022
Range("MAS190").Name = "xlasBlkAddr190" '//4/21/2022
Range("MAS191").Name = "xlasBlkAddr191" '//4/21/2022
Range("MAS192").Name = "xlasBlkAddr192" '//4/21/2022
Range("MAS193").Name = "xlasBlkAddr193" '//4/21/2022
Range("MAS194").Name = "xlasBlkAddr194" '//4/21/2022
Range("MAS195").Name = "xlasBlkAddr195" '//4/21/2022
Range("MAS196").Name = "xlasBlkAddr196" '//4/21/2022
Range("MAS197").Name = "xlasBlkAddr197" '//4/21/2022
Range("MAS198").Name = "xlasBlkAddr198" '//4/21/2022
Range("MAS199").Name = "xlasBlkAddr199" '//4/21/2022
Range("MAS200").Name = "xlasBlkAddr200" '//4/21/2022
Range("MAS201").Name = "xlasBlkAddr201" '//4/21/2022
Range("MAS202").Name = "xlasBlkAddr202" '//4/21/2022
Range("MAS203").Name = "xlasBlkAddr203" '//4/21/2022
Range("MAS204").Name = "xlasBlkAddr204" '//4/21/2022
Range("MAS205").Name = "xlasBlkAddr205" '//4/21/2022
Range("MAS206").Name = "xlasBlkAddr206" '//4/21/2022
Range("MAS207").Name = "xlasBlkAddr207" '//4/21/2022
Range("MAS208").Name = "xlasBlkAddr208" '//4/21/2022
Range("MAS209").Name = "xlasBlkAddr209" '//4/21/2022
Range("MAS210").Name = "xlasBlkAddr210" '//4/21/2022
Range("MAS211").Name = "xlasBlkAddr211" '//4/21/2022
Range("MAS212").Name = "xlasBlkAddr212" '//4/21/2022
Range("MAS213").Name = "xlasBlkAddr213" '//4/21/2022
Range("MAS214").Name = "xlasBlkAddr214" '//4/21/2022
Range("MAS215").Name = "xlasBlkAddr215" '//4/21/2022
Range("MAS216").Name = "xlasBlkAddr216" '//4/21/2022
Range("MAS217").Name = "xlasBlkAddr217" '//4/21/2022
Range("MAS218").Name = "xlasBlkAddr218" '//4/21/2022
Range("MAS219").Name = "xlasBlkAddr219" '//4/21/2022
Range("MAS220").Name = "xlasBlkAddr220" '//4/21/2022
Range("MAS221").Name = "xlasBlkAddr221" '//4/21/2022
Range("MAS222").Name = "xlasBlkAddr222" '//4/21/2022
Range("MAS223").Name = "xlasBlkAddr223" '//4/21/2022
Range("MAS224").Name = "xlasBlkAddr224" '//4/21/2022
Range("MAS225").Name = "xlasBlkAddr225" '//4/21/2022
Range("MAS226").Name = "xlasBlkAddr226" '//4/21/2022
Range("MAS227").Name = "xlasBlkAddr227" '//4/21/2022
Range("MAS228").Name = "xlasBlkAddr228" '//4/21/2022
Range("MAS229").Name = "xlasBlkAddr229" '//4/21/2022
Range("MAS230").Name = "xlasBlkAddr230" '//4/21/2022
Range("MAS231").Name = "xlasBlkAddr231" '//4/21/2022
Range("MAS232").Name = "xlasBlkAddr232" '//4/21/2022
Range("MAS233").Name = "xlasBlkAddr233" '//4/21/2022
Range("MAS234").Name = "xlasBlkAddr234" '//4/21/2022
Range("MAS235").Name = "xlasBlkAddr235" '//4/21/2022
Range("MAS236").Name = "xlasBlkAddr236" '//4/21/2022
Range("MAS237").Name = "xlasBlkAddr237" '//4/21/2022
Range("MAS238").Name = "xlasBlkAddr238" '//4/21/2022
Range("MAS239").Name = "xlasBlkAddr239" '//4/21/2022
Range("MAS240").Name = "xlasBlkAddr240" '//4/21/2022
Range("MAS241").Name = "xlasBlkAddr241" '//4/21/2022
Range("MAS242").Name = "xlasBlkAddr242" '//4/21/2022
Range("MAS243").Name = "xlasBlkAddr243" '//4/21/2022
Range("MAS244").Name = "xlasBlkAddr244" '//4/21/2022
Range("MAS245").Name = "xlasBlkAddr245" '//4/21/2022
Range("MAS246").Name = "xlasBlkAddr246" '//4/21/2022
Range("MAS247").Name = "xlasBlkAddr247" '//4/21/2022
Range("MAS248").Name = "xlasBlkAddr248" '//4/21/2022
Range("MAS249").Name = "xlasBlkAddr249" '//4/21/2022
Range("MAS250").Name = "xlasBlkAddr250" '//4/21/2022
Range("MAS251").Name = "xlasBlkAddr251" '//4/21/2022
Range("MAS252").Name = "xlasBlkAddr252" '//4/21/2022
Range("MAS253").Name = "xlasBlkAddr253" '//4/21/2022
Range("MAS254").Name = "xlasBlkAddr254" '//4/21/2022
Range("MAS255").Name = "xlasBlkAddr255" '//4/21/2022
Range("MAS256").Name = "xlasBlkAddr256" '//4/21/2022
Range("MAS257").Name = "xlasBlkAddr257" '//4/21/2022
Range("MAS258").Name = "xlasBlkAddr258" '//4/21/2022
Range("MAS259").Name = "xlasBlkAddr259" '//4/21/2022
Range("MAS260").Name = "xlasBlkAddr260" '//4/21/2022
Range("MAS261").Name = "xlasBlkAddr261" '//4/21/2022
Range("MAS262").Name = "xlasBlkAddr262" '//4/21/2022
Range("MAS263").Name = "xlasBlkAddr263" '//4/21/2022
Range("MAS264").Name = "xlasBlkAddr264" '//4/21/2022
Range("MAS265").Name = "xlasBlkAddr265" '//4/21/2022
Range("MAS266").Name = "xlasBlkAddr266" '//4/21/2022
Range("MAS267").Name = "xlasBlkAddr267" '//4/21/2022
Range("MAS268").Name = "xlasBlkAddr268" '//4/21/2022
Range("MAS269").Name = "xlasBlkAddr269" '//4/21/2022
Range("MAS270").Name = "xlasBlkAddr270" '//4/21/2022
Range("MAS271").Name = "xlasBlkAddr271" '//4/21/2022
Range("MAS272").Name = "xlasBlkAddr272" '//4/21/2022
Range("MAS273").Name = "xlasBlkAddr273" '//4/21/2022
Range("MAS274").Name = "xlasBlkAddr274" '//4/21/2022
Range("MAS275").Name = "xlasBlkAddr275" '//4/21/2022
Range("MAS276").Name = "xlasBlkAddr276" '//4/21/2022
Range("MAS277").Name = "xlasBlkAddr277" '//4/21/2022
Range("MAS278").Name = "xlasBlkAddr278" '//4/21/2022
Range("MAS279").Name = "xlasBlkAddr279" '//4/21/2022

'//Create target script locations
If Dir(drv & envHome & "\.z7", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7")
If Dir(drv & envHome & "\.z7\utility", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility")
If Dir(drv & envHome & "\.z7\utility\debug", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\debug")
If Dir(drv & envHome & "\.z7\utility\temp", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\temp")
If Dir(drv & envHome & "\.z7\utility\miss", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\miss")
If Dir(drv & envHome & "\.z7\utility\miss\colors", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\miss\colors")

MsgBox ("xlAppScript runtime environment connection is complete." & nl & nl & _
"Current environment: " & ThisWorkbook.Name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbInformation

Exit Function

ErrMsg:
MsgBox ("There was an issue trying to connect this runtime environment." & nl & nl & _
"Current environment: " & ThisWorkbook.Name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbCritical

End Function
Public Function disconnectWb()

'/\_____________________________________________________________________________________
'//
'//     A function for removing an environment from interacting w/ xlAppScript
'/\_____________________________________________________________________________________


On Error GoTo ErrMsg

nl = vbNewLine

'//Remove script addresses

'//Multi-memory block specific addresses
ActiveWorkbook.Names("xlasKinLabelMod").Delete
ActiveWorkbook.Names("xlasKinValueMod").Delete
ActiveWorkbook.Names("xlasKinLabel").Delete
ActiveWorkbook.Names("xlasKinValue").Delete
ActiveWorkbook.Names("xlasState").Delete
ActiveWorkbook.Names("xlasArticle").Delete
ActiveWorkbook.Names("xlasGroup").Delete
ActiveWorkbook.Names("xlasList").Delete '//(3/15/2022)
ActiveWorkbook.Names("xlasLib").Delete
'//Single-memory block specific addresses
ActiveWorkbook.Names("xlasAppLoad").Delete '//Autokit applications(2/9/2022)
ActiveWorkbook.Names("xlasEnvironment").Delete
ActiveWorkbook.Names("xlasBlock").Delete '//(2/28/2022)
ActiveWorkbook.Names("xlasGoto").Delete
ActiveWorkbook.Names("xlasInvert").Delete '//ctrl box+
ActiveWorkbook.Names("xlasKeyCtrl").Delete '//Autokit applications
ActiveWorkbook.Names("xlasSilent").Delete '//Autokit applications
ActiveWorkbook.Names("xlasRemember").Delete '//ctrl box+
ActiveWorkbook.Names("xlasConsoleType").Delete '//ctrl box+
ActiveWorkbook.Names("xlasAMemory").Delete '//ctrl box+
ActiveWorkbook.Names("xlasSaveFile").Delete '//ctrl box+
ActiveWorkbook.Names("xlasCtrlBoxBColor").Delete '//ctrl box+
ActiveWorkbook.Names("xlasCtrlBoxFColor").Delete '//ctrl box+
ActiveWorkbook.Names("xlasGlobalControl").Delete '//4/18/2022
ActiveWorkbook.Names("xlasLocalContain").Delete '//(3/17/2022)
ActiveWorkbook.Names("xlasLocalStatic").Delete '//(3/13/2022)
ActiveWorkbook.Names("xlasUpdateEnable").Delete '//(2/24/2022)
ActiveWorkbook.Names("xlasWinForm").Delete '//Autokit applications
ActiveWorkbook.Names("xlasWinFormLast").Delete '//Autokit applications
ActiveWorkbook.Names("xlasLibErrLvl").Delete '//4/22/2022
ActiveWorkbook.Names("xlasErrRef").Delete
ActiveWorkbook.Names("xlasEnd").Delete
Range("xlasLink").Value = vbNullString: ActiveWorkbook.Names("xlasLink").Delete
ActiveWorkbook.Names("xlasBlkAddr79").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr80").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr81").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr82").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr83").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr84").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr85").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr86").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr87").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr88").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr89").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr90").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr91").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr92").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr93").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr94").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr95").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr96").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr97").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr98").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr99").Delete '//3/17/2022
ActiveWorkbook.Names("xlasBlkAddr100").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr101").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr102").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr103").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr104").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr105").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr106").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr107").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr108").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr109").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr110").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr111").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr112").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr113").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr114").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr115").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr116").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr117").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr118").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr119").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr120").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr121").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr122").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr123").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr124").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr125").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr126").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr127").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr128").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr129").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr130").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr131").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr132").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr133").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr134").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr135").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr136").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr137").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr138").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr139").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr140").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr141").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr142").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr143").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr144").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr145").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr146").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr147").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr148").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr149").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr150").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr151").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr152").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr153").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr154").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr155").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr156").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr157").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr158").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr159").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr160").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr161").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr162").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr163").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr164").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr165").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr166").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr167").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr168").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr169").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr170").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr171").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr172").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr173").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr174").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr175").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr176").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr177").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr178").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr179").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr180").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr181").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr182").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr183").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr184").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr185").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr186").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr187").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr188").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr189").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr190").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr191").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr192").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr193").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr194").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr195").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr196").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr197").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr198").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr199").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr200").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr201").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr202").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr203").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr204").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr205").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr206").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr207").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr208").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr209").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr210").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr211").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr212").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr213").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr214").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr215").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr216").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr217").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr218").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr219").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr220").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr221").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr222").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr223").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr224").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr225").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr226").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr227").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr228").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr229").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr230").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr231").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr232").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr233").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr234").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr235").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr236").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr237").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr238").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr239").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr240").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr241").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr242").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr243").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr244").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr245").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr246").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr247").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr248").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr249").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr250").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr251").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr252").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr253").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr254").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr255").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr256").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr257").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr258").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr259").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr260").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr261").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr262").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr263").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr264").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr265").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr266").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr267").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr268").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr269").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr270").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr271").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr272").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr273").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr274").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr275").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr276").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr277").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr278").Delete '//4/21/2022
ActiveWorkbook.Names("xlasBlkAddr279").Delete '//4/21/2022

MsgBox ("xlAppScript runtime environment disconnection is complete." & nl & nl & _
"Current environment: " & ThisWorkbook.Name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbInformation

Exit Function

ErrMsg:
MsgBox ("There was an issue trying to disconnect this runtime environment." & nl & nl & _
"Current environment: " & ThisWorkbook.Name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbCritical

End Function

