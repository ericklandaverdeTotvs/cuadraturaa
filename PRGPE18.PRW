#Include "Protheus.ch"

//ordenar por periodo
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PRGPE18   ºAutor  ³ A Mendez           ºFecha ³  08/24/17   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Generación de consulta de empleados con el saldo de la caja ±±
±±º          ³ de ahorros                                                 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/


User Function PRGPE18()      

   Processa({|| GenTela()},"Generando Información")

Return 



Static Function GenTela()
Local aButtons   := {} 
Local aTempReg   := {}
Local aRegistros := {}
Local aMatriculas:= {}
Local aObjCoords := {}
Local ACOLSIZE   := {}
Local cConsecut  := "00000"
Local cNomPerg   := "PRGPE18"  
Local nX         := 0
Local nY         := 0

GenPerg(cNomPerg)

Aadd(aButtons,{"Excel",{|| u_PRGPE13(aRegistros) }, "Impresión Excel"    ,"Excel"})  
 
If !Pergunte(cNomPerg,.T.)
	Return
EndIf
    
aMatriculas := GenMatriculas()  

ProcRegua(Len(aMatriculas))
    
If Len(aMatriculas) > 0 
    
	For nX := 1 To Len(aMatriculas)
		aTempReg := GenSaldoCA(aMatricula[nX,1])
         
         If Len(aTempReg) > 0 
			For nY:= 1 To Len(aTempReg)
   				AADD(aRegistros,{aTempReg[nY,1],aTempReg[nY,2],aTempReg[nY,3],aTempReg[nY,4],aTempReg[nY,5],;
   								aTempReg[nY,6],aTempReg[nY,7],aTempReg[nY,8],aTempReg[nY,9],aTempReg[nY,10],;
           						aTempReg[nY,11]} ) 
           	Next nY
		Endif    
        
		IncProc("Empleado: " + aMatricula[nX,1] + " - "  + aMatricula[nX,2])
	Next nX
        
Endif

If Len(aRegistros) ==  0
	AADD(aRegistros,{"","","","","","","","","",0,0})    
Endif

aAdvSize := MsAdvSize(.T.,.F.,450)
	
aInfoAdvSize:= { aAdvSize[1] , aAdvSize[2] , aAdvSize[3] , aAdvSize[4] , 5 , 5 }
aAdd( aObjCoords , { 000 , 040 , .T. , .F. } )
aAdd( aObjCoords , { 000 , 000 , .T. , .T. } )
aAdd( aObjCoords , { 000 , 040 , .T. , .F. } )
aObjSize := MsObjSize( aInfoAdvSize , aObjCoords )
	
DEFINE MSDIALOG _oDlg TITLE "Consulta Saldos Caja de Ahorro" FROM aAdvSize[7],aAdvSize[1] TO aAdvSize[6],aAdvSize[5] PIXEL

aTFolder := { OemToAnsi('Empleados') }
oGroup1 := TFolder():New( aObjSize[1,1],aObjSize[1,2],aTFolder,,_oDlg,1,,,.T.,,aObjSize[1,4],aObjSize[1,3] )

aHeader := {"Matrícula","Nombre","Situación","Proceso","Periodo","Semana","Nómina","Concepto","Descripción","Unidades","Importe"}
lHScroll := .T.
lVScroll := .T.
lPixel := .T.
oBrowse := TWBrowse():New(aObjSize[2,1],aObjSize[2,2],aObjSize[2,4],aObjSize[2,3]-40,;
							/*bFields*/,aHeader,aColSize, _oDlg,/*cField*/,,,/*bChange*/,;
	   	                         /*bLDblClick*/,/*bRClick*/,/*oFont*/,,,,,.F.,,.T.,,.F.,,.T.,.T.)
	                                 
oBrowse:SetArray(aRegistros)
oBrowse:bLine := {||{aRegistros[oBrowse:nAt,01],;
					 aRegistros[oBrowse:nAt,02],;
					 aRegistros[oBrowse:nAt,11],;
					 aRegistros[oBrowse:nAt,03],;
					 aRegistros[oBrowse:nAt,04],;
					 aRegistros[oBrowse:nAt,05],;
					 aRegistros[oBrowse:nAt,06],;
					 aRegistros[oBrowse:nAt,07],;
					 aRegistros[oBrowse:nAt,08],;
					 aRegistros[oBrowse:nAt,09],;
					 Transform(aRegistros[oBrowse:nAt,10],'@E 9,999,999.99');
					};
				}
oBrowse:Refresh()

ACTIVATE MSDIALOG _oDlg CENTERED ON INIT (EnchoiceBar(_oDlg, {|| _oDlg:End() },{|| nOpc:=0, _oDlg:End()},.F.,@aButtons))
Return

Static Function GenSaldoCA(cMatricula)  
Local cQuery     := ""             
Local nVSaldo    := 0    
Local cAliasSRA  := "SRATMP"         
Local aSaldosCA  := {} 

cQuery :=  " SELECT RD_MAT,RA_NOMECMP,RA_SITFOLH,RD_HORAS HORAS, RD_PD,RV_DESC,RD_VALOR VALOR,RV_TIPOCOD, RD_PERIODO PERIODO,RD_PROCES PROCES,RD_SEMANA SEMANA,RD_ROTEIR ROTEIR "
cQuery +=  " FROM " +InitSqlName("SRD") + " SRD "
cQuery +=  " INNER JOIN " + InitSqlName("SRV") + " SRV ON RV_COD = RD_PD AND SRV.D_E_L_E_T_  = ' ' AND RV_FILIAL = '"+xFilial("SRV")+"' "
cQuery +=  " INNER JOIN " + InitSqlName("SRA") + " SRA ON RA_MAT = RD_MAT AND SRA.D_E_L_E_T_  = ' ' AND RA_FILIAL = RD_FILIAL "
cQuery +=  " WHERE RD_FILIAL = '"+xFilial("SRD")+"' AND RD_MAT = '" + cMatricula + "'"
cQuery +=  " AND (RD_PERIODO >= " + mv_par01 + " And " + "RD_PERIODO <= " + mv_par02 +  ")"  
cQuery +=  " AND RD_PD in ('464','469','580','581','582','583','468','584','585') "          
cQuery +=  " AND SRD.d_e_l_e_t_ = ' '"

DbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSRA,.T.,.T.)

If !(cAliasSRA)->(EOF()) 

	If mv_par06 == 1
		AADD(aSaldosCA,{(cAliasSRA)->RD_MAT,(cAliasSRA)->RA_NOMECMP,(cAliasSRA)->PROCES,'','','','','Saldo ',0,nVSaldo,(cAliasSRA)->RA_SITFOLH})
       	nVSaldo:= 0 
    Endif   
       
    While !(cAliasSRA)->(EOF())
     
	    If (cAliasSRA)->RD_PD $ '464,469,580,581'
	       If mv_par06 == 1
       	     	nVSaldo += (cAliasSRA)->VALOR
       	     Else 
       	        AADD(aSaldosCA,{(cAliasSRA)->RD_MAT ,(cAliasSRA)->RA_NOMECMP,(cAliasSRA)->PROCES,;
       	                        (cAliasSRA)->PERIODO,(cAliasSRA)->SEMANA    ,(cAliasSRA)->ROTEIR    ,(cAliasSRA)->RD_PD ,;
       	                        (cAliasSRA)->RV_DESC,(cAliasSRA)->HORAS      ,(cAliasSRA)->VALOR,(cAliasSRA)->RA_SITFOLH})  
       	   Endif  	
	    ElseIf (cAliasSRA)->RD_PD $'582,583,468,584,585' 
	       If mv_par06 == 1
       	     	nVSaldo -= (cAliasSRA)->VALOR
       	     Else 
       	        AADD(aSaldosCA,{(cAliasSRA)->RD_MAT ,(cAliasSRA)->RA_NOMECMP,(cAliasSRA)->PROCES,;
       	                        (cAliasSRA)->PERIODO,(cAliasSRA)->SEMANA    ,(cAliasSRA)->ROTEIR    ,(cAliasSRA)->RD_PD ,;
       	                        (cAliasSRA)->RV_DESC,(cAliasSRA)->HORAS     ,(cAliasSRA)->VALOR*-1,(cAliasSRA)->RA_SITFOLH})  
       	   Endif  	
	    EndIf
	 
	    (cAliasSRA)->(DbSkip())
     
    EndDo
  
    If mv_par06 == 1
         aSaldosCA[1,10] := nVSaldo 
    Endif 
     
Endif
(cAliasSRA)->(DbCloseArea()) 
Return(aSaldosCA) 

       
Static Function GenMatricu()  

Local cQuery      := ""             
Local nVSaldo     := 0   
Local cAliasSRD   := "TMPSRD"
Local aMatriculas := {}          

cQuery :=  " SELECT RA_MAT, RA_NOMECMP "
cQuery +=  " FROM " +InitSqlName("SRA") + " SRA "
cQuery +=  " WHERE RA_FILIAL = '"+xFilial("SRA")+"' "   
cQuery +=  " And (RA_PROCES >= " + mv_par03 + " Or " + "RA_PROCES <= " + mv_par04 + ")"  
Do Case
	Case mv_par05 == 1
		cQuery +=  " And RA_SITFOLH IN  (' ')"   
   	Case mv_par05 == 2
       	cQuery +=  " And RA_SITFOLH IN  ('D')"   
   	Case mv_par05 == 3
       	cQuery +=  " And RA_SITFOLH IN  (' ','A','D','F','T' )"   
EndCase
cQuery +=  " And SRA.D_E_L_E_T_ = ' '"
DbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSRD,.T.,.T.)

If !(cAliasSRD)->(EOF()) 
	While !(cAliasSRD)->(EOF())
	AADD(aMatriculas,{(cAliasSRD)->RA_MAT,(cAliasSRD)->RA_NOMECMP})
	(cAliasSRD)->(DbSkip())  
  EndDo  
Endif
(cAliasSRD)->(DbCloseArea())
Return(aMatriculas)


Static Function GenPerg(cNomPerg)

Local _sAlias := Alias()
Local i := 0
Local j := 0
Local cPreg := PADR(cNomPerg,10)
Local aRegs := {}
Local aHelpSpa := {}
Local cKey := ""

dbSelectArea("SX1")
dbSetOrder(1)
aAdd(aRegs,{cPreg,"01",	"","Del Periodo"  ,"",	"MV_CHA",	"C",06,0,	0,	"G","NaoVazio()","MV_PAR01","","","","",			"","","","","",		"","","","","",		"","","","","",		"","","","","",		"RCH11","","","",""})
aAdd(aRegs,{cPreg,"02",	"","Al Periodo"  ,"",	"MV_CHB",	"C",06,0,	0,	"G","NaoVazio()","MV_PAR02","","","","",			"","","","","",		"","","","","",		"","","","","",		"","","","","",		"RCH11","","","",""})
aAdd(aRegs,{cPreg,"03",	"","De Proceso"  ,"",	"MV_CHC",	"C",05,0,	0,	"G","NaoVazio()","MV_PAR03","","","","",			"","","","","",		"","","","","",		"","","","","",		"","","","","",		"RCJ","","","",""})		
aAdd(aRegs,{cPreg,"04",	"","A Proceso"   ,"",	"MV_CHD",	"C",05,0,	0,	"G","NaoVazio()","MV_PAR04","","","","",			"","","","","",		"","","","","",		"","","","","",		"","","","","",		"RCJ","","","",""})		
aAdd(aRegs,{cPreg,"05",	"","Situación"   ,"",	"MV_CHE",	"C",01,0,	0,	"C",""          ,"MV_PAR06","Activos","Activos","Activos","",			"","Inactivos","Inactivos","Activos","",		"","Todos","Todos","Todos","",		"","","","","",		"","","","","",		"","","","",""})
aAdd(aRegs,{cPreg,"06",	"","Tipo Rep"    ,"",	"MV_CHF",	"C",01,0,	0,	"C",""          ,"MV_PAR07","Saldos","Saldos","Saldos","",			"","Movimientos","Movimientos","Movimientos","",		"","","","","",		"","","","","",		"","","","","",		"","","","",""})


For i:=1 to Len(aRegs)
	If !dbSeek(cPreg+aRegs[i,2])
		RecLock("SX1",.T.)
		For j:=1 to FCount()
			If j <= Len(aRegs[i])
				FieldPut(j,aRegs[i,j])
			Endif
		Next
		MsUnlock()
		ElseIF  I== 6 .oR.  I== 12 //Cree siempre el registro de matricula
		RecLock("SX1",.F.)
        For j:=1 to FCount()
        	If j <= Len(aRegs[i])
				FieldPut(j,aRegs[i,j])
           Endif
	       Next
	       MsUnlock()
	Endif
Next

dbSelectArea(_sAlias)
Return