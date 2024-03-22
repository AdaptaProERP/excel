// Programa   : BRDEFFIELDCOLXLS
// Fecha/Hora : 13/06/2017 15:04:37
// Propósito  : Definición de Campos con Columnas de Excel
// Creado Por : Automáticamente por BRWMAKER
// Llamado por: <DPXBASE>
// Aplicación : Gerencia 
// Tabla      : <TABLA>

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cTable,cTitle,cFileIXL,aRef,bPostSave,cRunIxl,cFileXls,cCodigo)
   LOCAL aData,aDbf:={},aHead:={},cMax:="A"
   LOCAL nMaxCol:=26+26

   IF Type("oDefXlsCol")="O" .AND. oDefXlsCol:oWnd:hWnd>0
      EJECUTAR("BRRUNNEW",oDefXlsCol,GetScript())
      RETURN oDefXlsCol
   ENDIF

   DEFAULT  cTable :="DPCLIENTES",;
            aRef   :={},;
            cTitle :="CAMPOS DE LAS TABLAS",;
            cFileXls:=""      

   DEFAULT cCodigo:=SQLGET("DPIMPRXLS","IXL_CODIGO","IXL_TABLA"+GetWhere("=",cTable))

   AADD(aDbf,ARRAY(	nMaxCol))
   cMax:="A"

   AEVAL(aDbf[1],{|a,n,c| AADD(aHead,cMax),cMax:=NEXTCOL(cMax)})
  
   cTitle:="Definición para Importar desde Archivos EXCEL " +IF(Empty(cTitle),"",cTitle)

   oDp:oFrm:=NIL

   aData:=EJECUTAR("IXLLOADDATA",cTable,cTitle,cFileIXL,aRef,bPostSave,cRunIxl,cFileXls,cCodigo)

   AEVAL(aData,{|a,n|aData[n,3]:=SAYOPTIONS("DPCAMPOS","CAM_TYPE",a[3]),;
                     aData[n,6]:=IF(Empty(a[6]),SPACE(2),a[6]) })

   ViewData(aData,cTitle)

   oDp:oFrm:=oDefXlsCol
            
RETURN .T. 

FUNCTION ViewData(aData,cTitle,cWhere_)
   LOCAL oBrw,oCol,aTotal:=ATOTALES(aData)
   LOCAL oFont,oFontB
   LOCAL aPeriodos:=ACLONE(oDp:aPeriodos)
   LOCAL aCoors:=GetCoors( GetDesktopWindow() )
   LOCAL oTable

   oTable:=OpenTable("SELECT * FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)
   oTable:End()

   DEFINE FONT oFont  NAME "Tahoma" SIZE 0, -12 
   DEFINE FONT oFontB NAME "Tahoma" SIZE 0, -12 

   DpMdi(cTitle,"oDefXlsCol","BRDEFFIELDCOLXLS.EDT")

   oDefXlsCol:SetTable(oTable)

   oDefXlsCol:Windows(0,0,aCoors[3]-170,MIN(900,aCoors[4]-10),.T.) // Maximizado

   oDefXlsCol:lMsgBar  :=.F.
   oDefXlsCol:cNombre  :=""
   oDefXlsCol:cWhereQry:=""
   oDefXlsCol:oWhere   :=TWHERE():New(oDefXlsCol)
   oDefXlsCol:lWhen    :=.T.
   oDefXlsCol:cBrwCod  :=""
   oDefXlsCol:lTmdi    :=.T.
   oDefXlsCol:cWhereCli:=""
   oDefXlsCol:cTitleCli:=NIL
   oDefXlsCol:cCodigo  :=cCodigo
   oDefXlsCol:cMemo    :=""
   oDefXlsCol:cCondic  :="Requerimientos: "+CRLF+"Archivo Excel no debe tener Columnas dinámicas ni campos"+CRLF+"memos ni multi-lineal."+CRLF+"Dejará el sistema sin Memoria y no podrá funcionar."

   oDefXlsCol:oBrwD    :=NIL
   oDefXlsCol:oMemo    :=NIL
   oDefXlsCol:oMeterR  :=NIL
   oDefXlsCol:oSay     :=NIL
   oDefXlsCol:nRecord  :=0
   oDefXlsCol:lBarDef  :=.T.
   oDefXlsCol:nCantid  :=10
   oDefXlsCol:aHead    :=ACLONE(aHead)
   oDefXlsCol:cFileIXL :="FORMS\"+ALLTRIM(cCodigo)+".IXL"
   oDefXlsCol:bPostSave:={||NIL}
   oDefXlsCol:nMaxCol  :=nMaxCol // MAXCOL

   oDefXlsCol:nClrText1:=CLR_HBLUE

   oDefXlsCol:oBrwA:=TXBrowse():New( IF(oDefXlsCol:lTmdi,oDefXlsCol:oWnd,oDefXlsCol:oDlg ))

   oDefXlsCol:oBrwA:aData:=ACLONE(aData)

   oDefXlsCol:oBrwA:SetArray( aData, .F. )
   oDefXlsCol:oBrwA:SetFont(oFont)

   oDefXlsCol:oBrwA:lFooter     := .F.
   oDefXlsCol:oBrwA:lHScroll    := .T.
   oDefXlsCol:oBrwA:nHeaderLines:= 2
   oDefXlsCol:oBrwA:nDataLines  := 1
   oDefXlsCol:oBrwA:nFooterLines:= 1

   oDefXlsCol:aData    :=ACLONE(aData)
   oDefXlsCol:nClrText :=0
   oDefXlsCol:nClrPane1:=oDp:nClrPane1
   oDefXlsCol:nClrPane2:=oDp:nClrPane2

   AEVAL(oDefXlsCol:oBrwA:aCols,{|oCol|oCol:oHeaderFont:=oFont})

   oCol:=oDefXlsCol:oBrwA:aCols[1]
   oCol:cHeader      :='Campo'
   oCol:bLClickHeader:={|r,c,f,o| SortArray( o, oDefXlsCol:oBrwA:aArrayData ) } 
   oCol:nWidth       := 100+40+40

   oCol:=oDefXlsCol:oBrwA:aCols[2]
   oCol:cHeader      :='Descripción'
   oCol:bLClickHeader:= {|r,c,f,o| SortArray( o, oDefXlsCol:oBrwA:aArrayData ) } 
   oCol:nWidth       := 280

   oCol:=oDefXlsCol:oBrwA:aCols[3]
   oCol:cHeader      :='Tipo'
   oCol:bLClickHeader:= {|r,c,f,o| SortArray( o, oDefXlsCol:oBrwA:aArrayData ) } 
   oCol:nWidth       := 70-10

   oCol:bClrStd      := {|nClrText,uValue|uValue:=oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,4],;
                          nClrText:=COLOR_OPTIONS("DPCAMPOS","CAM_TYPE",uValue),;
                         {nClrText,iif( oDefXlsCol:oBrwA:nArrayAt%2=0, oDefXlsCol:nClrPane1, oDefXlsCol:nClrPane2 ) } } 

   oCol:=oDefXlsCol:oBrwA:aCols[4]
   oCol:cHeader      :='Len'
   oCol:bLClickHeader:= {|r,c,f,o| SortArray( o, oDefXlsCol:oBrwA:aArrayData ) } 
   oCol:nWidth       := 30

   oCol:=oDefXlsCol:oBrwA:aCols[5]
   oCol:cHeader      :='Valor'+CRLF+"Por Defecto"
   oCol:bLClickHeader:= {|r,c,f,o| SortArray( o, oDefXlsCol:oBrwA:aArrayData ) } 
   oCol:nWidth       := 180
   oCol:nEditType    :=1
   oCol:bOnPostEdit  :={|oCol,uValue|oDefXlsCol:PUTDEFAULT(oCol,uValue,5)}


   oCol:=oDefXlsCol:oBrwA:aCols[6]
   oCol:cHeader      :='Col'+CRLF+"XLS"
   oCol:bLClickHeader:= {|r,c,f,o| SortArray( o, oDefXlsCol:oBrwA:aArrayData ) } 
   oCol:bLClickHeader:= {|r,c,f,o| oDefXlsCol:SORTXLSCOL(r,c,f,o) } 
   oCol:nWidth       := 40
   oCol:nEditType  :=1
   oCol:bOnPostEdit:={|oCol,uValue|oDefXlsCol:PUTCOLUMNA(oCol,uValue,6)}


   oDefXlsCol:oBrwA:bClrStd               := {|oBrw,nClrText,aLine|oBrw    :=oDefXlsCol:oBrwA,aLine:=oBrw:aArrayData[oBrw:nArrayAt],;
                                                                   nClrText:=oDefXlsCol:nClrText,;
                                                                   nClrText:=IF(Empty(aLine[6]),nClrText,oDefXlsCol:nClrText1),;
                                            {nClrText,iif( oBrw:nArrayAt%2=0, oDefXlsCol:nClrPane1, oDefXlsCol:nClrPane2 ) } }


   oDefXlsCol:oBrwA:bClrHeader            := {|| { oDp:nLbxClrHeaderText, oDp:nLbxClrHeaderPane}}
   oDefXlsCol:oBrwA:bClrFooter            := {|| { oDp:nLbxClrHeaderText, oDp:nLbxClrHeaderPane}}

   oDefXlsCol:oBrwA:bLDblClick:={|oBrw|oDefXlsCol:RUNCLICK() }

   oDefXlsCol:oBrwA:bChange:={||oDefXlsCol:BRWCHANGE()}
   oDefXlsCol:oBrwA:CreateFromCode()

   oDefXlsCol:oBrwD:=TXBrowse():New( IF(oDefXlsCol:lTmdi,oDefXlsCol:oWnd,oDefXlsCol:oDlg ))
   oDefXlsCol:oBrwD:SetArray( aDbf, .F. )
   oDefXlsCol:oBrwD:SetFont(oFont)

   oDefXlsCol:oBrwD:lFooter     := .F.
   oDefXlsCol:oBrwD:lHScroll    := .T.
   oDefXlsCol:oBrwD:nHeaderLines:= 2
   oDefXlsCol:oBrwD:nDataLines  := 1
   oDefXlsCol:oBrwD:nFooterLines:= 1

   oCol:=oDefXlsCol:oBrwD:aCols[1]
   oCol:cHeader      :='A'
   oCol:bLClickHeader:= {|r,c,f,o| SortArray( o, oDefXlsCol:oBrwD:aArrayData ) } 
   oCol:nWidth       := 40

   AEVAL(oDefXlsCol:oBrwD:aCols,{|oCol,n|oCol:oHeaderFont:=oFont,oCol:cHeader:=aHead[n],oCol:nWidth:=40})

   oDefXlsCol:oBrwD:bClrStd               := {|oBrwD,nClrText,aLine|oBrwD:=oDefXlsCol:oBrwD,aLine:=oBrwD:aArrayData[oBrwD:nArrayAt],;
                                              nClrText:=oDefXlsCol:nClrText,;
                                            {nClrText,iif( oBrwD:nArrayAt%2=0, oDefXlsCol:nClrPane1, oDefXlsCol:nClrPane2 ) } }

   oDefXlsCol:oBrwD:bClrHeader            := {|| { oDp:nLbxClrHeaderText, oDp:nLbxClrHeaderPane}}
   oDefXlsCol:oBrwD:bClrFooter            := {|| { oDp:nLbxClrHeaderText, oDp:nLbxClrHeaderPane}}

//   oDefXlsCol:oBrwD:bLDblClick:={|oBrwD|oDefXlsCol:RUNCLICK() }

   oDefXlsCol:oBrwD:bChange   :={||oDefXlsCol:BRWCHANGE()}
   oDefXlsCol:oBrwD:bLDblClick:={||oDefXlsCol:RUNCLICKD()}

   oDefXlsCol:oBrwD:CreateFromCode()

   oDefXlsCol:bValid   :={||.T. }  // 17/07/2023 genera incidencia {|| EJECUTAR("BRWSAVEPAR",oDefXlsCol)}

   @ 0,0 SPLITTER oDefXlsCol:oHSplit ;
         HORIZONTAL;
         PREVIOUS CONTROLS oDefXlsCol:oBrwA ;
         HINDS CONTROLS oDefXlsCol:oBrwD;
         TOP MARGIN 40 ;
         BOTTOM MARGIN 40 ;
         SIZE 300, 4  PIXEL ;
         OF oDefXlsCol:oWnd ;
         _3DLOOK

  oDefXlsCol:oWnd:oClient := oDefXlsCol:oHSplit

  oDefXlsCol:Activate({||oDefXlsCol:ViewDatBar()})

RETURN .T.

/*
// Barra de Botones
*/
FUNCTION ViewDatBar()
   LOCAL oCursor,oBar,oBtn,oFont,oCol
   LOCAL oDlg:=IF(oDefXlsCol:lTmdi,oDefXlsCol:oWnd,oDefXlsCol:oDlg)
   LOCAL nLin:=0
   LOCAL nWidth:=oDefXlsCol:oBrwA:nWidth()
   LOCAL nAltoBrw:=150+80
   LOCAL nClrText:=0

   /*   
   //  Ubicamos el Area del Primer Objeto o Browse.
   */

   oDefXlsCol:oBrwA:Move(032,0,800,nAltoBrw,.T.)

   oDefXlsCol:oHSplit:Move(oDefXlsCol:oBrwA:nHeight()+oDefXlsCol:oBrwA:nTop(),0)
   oDefXlsCol:oBrwD:Move(oDefXlsCol:oBrwA:nHeight()+oDefXlsCol:oBrwA:nTop()+5,0,800,400,.T.)

   oDefXlsCol:oHSplit:AdjLeft()
   oDefXlsCol:oHSplit:AdjRight()

   oDefXlsCol:oBrwA:GoBottom(.T.)
   oDefXlsCol:oBrwA:Refresh(.T.)

   DEFINE CURSOR oCursor HAND
   IF !oDp:lBtnText 
     DEFINE BUTTONBAR oBar SIZE 52-15,60-15 OF oDlg 3D CURSOR oCursor
   ELSE 
     DEFINE BUTTONBAR oBar SIZE oDp:nBtnWidth+5,oDp:nBarnHeight+6 OF oDlg 3D CURSOR oCursor 
   ENDIF 

   DEFINE FONT oFont  NAME "Tahoma"   SIZE 0, -11 BOLD

 // Emanager no Incluye consulta de Vinculos

   oDefXlsCol:oFontBtn   :=oFont    
   oDefXlsCol:nClrPaneBar:=oDp:nGris
   oDefXlsCol:oBrw:oLbx  :=oDefXlsCol

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XSAVE.BMP";
          TOP PROMPT "Grabar"; 
          ACTION  (CursorWait(),oDefXlsCol:SAVELBX())

   oBtn:cToolTip:="Guardar Definiciones"


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\RUN.BMP";
          TOP PROMPT "Ejecutar"; 
          ACTION  (CursorWait(),oDefXlsCol:XLSLOADDBF())

   oBtn:cToolTip:="Lectura XLS"
/*
   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\PROGRAMA.BMP";
          TOP PROMPT "Programa"; 
          ACTION  EJECUTAR("DPIMPRXLS_PRG",oDefXlsCol:cCodigo)

   oBtn:cToolTip:="Editar Programa Fuente"
*/

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          MENU EJECUTAR("BRBTNMENU",{"Etiqueta de Referencias"},;
                                     "oDefXlsCol");
          FILENAME "BITMAPS\PROGRAMA.BMP";
          TOP PROMPT "Programa";
          ACTION EJECUTAR("DPIMPRXLS_PRG",oDefXlsCol:cCodigo)

   oBtn:cToolTip:="Opciones de Ejecucion"

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XEDIT.BMP";
          TOP PROMPT "Editar"; 
          ACTION  oDefXlsCol:XLSMODCOL()

   oBtn:cToolTip:="Editar"


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XDELETE.BMP";
          TOP PROMPT "Eliminar"; 
          ACTION  oDefXlsCol:XLSDELCOL()

   oBtn:cToolTip:="Remover ID de Columna"

   IF ISSQLFIND("DPBRWLNKCONCAT","BRC_CODIGO"+GetWhere("=","PLANTILLADOC"))

       DEFINE BUTTON oBtn;
       OF oBar;
       NOBORDER;
       FONT oFont;
       FILENAME "BITMAPS\XBROWSE.BMP";
       TOP PROMPT "Detalles"; 
       ACTION  EJECUTAR("BRWRUNBRWLINK",oDefXlsCol:oBrwA,"PLANTILLADOC",oDefXlsCol:cSql,oDefXlsCol:nPeriodo,oDefXlsCol:dDesde,oDefXlsCol:dHasta,oDefXlsCol)

       oBtn:cToolTip:="Ejecutar Browse Vinculado(s)"
       oDefXlsCol:oBtnRun:=oBtn

       oDefXlsCol:oBrwA:bLDblClick:={||EVAL(oDefXlsCol:oBtnRun:bAction) }

   ENDIF


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\EXCEL.BMP";
          TOP PROMPT "Excel"; 
          ACTION  oDefXlsCol:RUNFILEXLS()

   oBtn:cToolTip:="Abrir Archivo XLS "+ALLTRIM(oDefXlsCol:IXL_FILE)

   oDefXlsCol:oBtnXls:=oBtn

/*
   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\AUDITORIA.BMP";
          ACTION oDefXlsCol:VERAUDITORIA()

   oBtn:cToolTip:="Explorar en Pistas de Auditoria"
*/

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\ZOOM.BMP";
          TOP PROMPT "Zoom"; 
          ACTION IF(oDefXlsCol:oWnd:IsZoomed(),oDefXlsCol:oWnd:Restore(),oDefXlsCol:oWnd:Maximize())

   oBtn:cToolTip:="Maximizar"


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XFIND.BMP";
          TOP PROMPT "Buscar"; 
          ACTION  EJECUTAR("BRWSETFIND",oDefXlsCol:oBrw)

   oBtn:cToolTip:="Buscar"

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\FILTRAR.BMP";
          TOP PROMPT "Filtrar"; 
          ACTION  EJECUTAR("BRWSETFILTER",oDefXlsCol:oBrw)

   oBtn:cToolTip:="Filtrar Registros"

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\OPTIONS.BMP",NIL,"BITMAPS\OPTIONSG.BMP";
          TOP PROMPT "Opciones"; 
          ACTION  EJECUTAR("BRWSETOPTIONS",oDefXlsCol:oBrw);
          WHEN LEN(oDefXlsCol:oBrwA:aArrayData)>1

   oBtn:cToolTip:="Filtrar según Valores Comunes"


   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\html.BMP";
          TOP PROMPT "Html"; 
          ACTION  (EJECUTAR("BRWTOHTML",oDefXlsCol:oBrw))

   oBtn:cToolTip:="Generar Archivo html"

   oDefXlsCol:oBtnHtml:=oBtn
/*
   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\PREVIEW.BMP";
          TOP PROMPT "Preview"; 
          ACTION  (EJECUTAR("BRWPREVIEW",oDefXlsCol:oBrw))

   oBtn:cToolTip:="Previsualización"

   oDefXlsCol:oBtnPreview:=oBtn
*/

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\xTOP.BMP";
          TOP PROMPT "Primero"; 
          ACTION  (oDefXlsCol:oBrwA:GoTop(),oDefXlsCol:oBrwA:Setfocus())

  DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\xFIN.BMP";
          TOP PROMPT "Ultimo"; 
          ACTION  (oDefXlsCol:oBrwA:GoBottom(),oDefXlsCol:oBrwA:Setfocus())

   DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XSALIR.BMP";
          TOP PROMPT "Cerrar"; 
          ACTION  oDefXlsCol:Close()

  oDefXlsCol:oBrwA:SetColor(0,oDefXlsCol:nClrPane1)

  EVAL(oDefXlsCol:oBrwA:bChange)
 
  oBar:SetColor(CLR_BLACK,oDp:nGris)

  oBar:SetSize(NIL,100+70+30,.T.)

  AEVAL(oBar:aControls,{|o,n|o:SetColor(CLR_BLACK,oDp:nGris)})

  SetWndDefault(oBar) 

  @ 1+50+20,10 SAY " Código "  RIGHT OF oBar BORDER SIZE 060,20 PIXEL COLOR oDp:nClrLabelText,oDp:nClrLabelPane FONT oFont
  @21+50+20,10 SAY " Nombre "  RIGHT OF oBar BORDER SIZE 060,20 PIXEL COLOR oDp:nClrLabelText,oDp:nClrLabelPane FONT oFont

  @ 21+50+20,71 SAY oDefXlsCol:oNombre PROMPT " "+oDefXlsCol:IXL_DESCRI;
                 OF oBar PIXEL SIZE 280,20 BORDER COLOR oDp:nClrYellowText,oDp:nClrYellow FONT oFont

  @ 01+50+20,71 SAY " "+oDefXlsCol:IXL_CODIGO+" ";
                 OF oBar PIXEL SIZE 110+100,20 BORDER COLOR oDp:nClrYellowText,oDp:nClrYellow FONT oFont

  //
  // Campo : IXL_LININI 
  // Uso   : Línea de Inicio                  
  //
  @ 130+20, 30 GET oDefXlsCol:oIXL_LININI VAR oDefXlsCol:IXL_LININI SPINNER PICTURE "999999" ;
            FONT oFont;
            SIZE 40,20 OF oBar RIGHT PIXEL 

  //
  // Campo : IXL_LINFIN
  // Uso   : Línea Final                         
  //
  @ 130+20,30+70 GET oDefXlsCol:oIXL_LINFIN  VAR oDefXlsCol:IXL_LINFIN SPINNER PICTURE "99999";
              FONT oFont;
              SIZE 40,20 OF oBar RIGHT PIXEL 

   oDefXlsCol:oIXL_LINFIN:cMsg    :="Línea Final"
   oDefXlsCol:oIXL_LINFIN:cToolTip:="Línea Final"


  //
  // Campo : IXL_MINCOL
  // Uso   : Columna de Inicio                         
  //
  @ 130+20,30+140 GET oDefXlsCol:oIXL_MINCOL  VAR oDefXlsCol:IXL_MINCOL;
               FONT oFont;
               SIZE 20,20 OF oBar RIGHT PIXEL 


   oDefXlsCol:oIXL_MINCOL:cMsg    :="Columna de Inicio"
   oDefXlsCol:oIXL_MINCOL:cToolTip:="Columna de Inicio"

//  @ oDefXlsCol:oIXL_MINCOL:nTop-08,oDefXlsCol:oIXL_MINCOL:nLeft SAY "Columna"+CRLF+"Inicial" PIXEL;
//                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris

  //
  // Campo : IXL_MAXCOL
  // Uso   : Columna Final                         
  //
  @ 130+20,30+140+70 GET oDefXlsCol:oIXL_MAXCOL  VAR oDefXlsCol:IXL_MAXCOL;
                  FONT oFont;
                  SIZE 20,20 OF oBar RIGHT PIXEL 

  oDefXlsCol:oIXL_MAXCOL:cMsg    :="Columna Final"
  oDefXlsCol:oIXL_MAXCOL:cToolTip:="Columna Final"

  // Campo : nCantid
  // Uso   : Columna Final                         
  //
  @ 130+20,30+140+140 GET oDefXlsCol:oCantid  VAR oDefXlsCol:nCantid PICTURE "99999";
                   FONT oFont;
                   SIZE 40,20 OF oBar RIGHT PIXEL SPINNER

  oDefXlsCol:oIXL_MAXCOL:cMsg    :="Cantidad de Registros"
  oDefXlsCol:oIXL_MAXCOL:cToolTip:=oDefXlsCol:oIXL_MAXCOL:cMsg

 
  @ oDefXlsCol:oIXL_LININI:nTop-35,oDefXlsCol:oIXL_LININI:nLeft SAY oDefXlsCol:oSay1 PROMPT "Línea"+CRLF+"Inicio";
    SIZE 32,32 PIXEL FONT oFont COLOR nClrText,oDp:nGris

  @ oDefXlsCol:oIXL_LINFIN:nTop-35,oDefXlsCol:oIXL_LINFIN:nLeft SAY oDefXlsCol:oSay1 PROMPT "Línea"+CRLF+"Final";
    SIZE 32,32 PIXEL FONT oFont COLOR nClrText,oDp:nGris

  @ oDefXlsCol:oIXL_MINCOL:nTop-35,oDefXlsCol:oIXL_MINCOL:nLeft SAY oDefXlsCol:oSay3 PROMPT "Col."+CRLF+"Inicio";
    SIZE 32,35 PIXEL FONT oFont COLOR nClrText,oDp:nGris

  @ oDefXlsCol:oIXL_MAXCOL:nTop-35,oDefXlsCol:oIXL_MAXCOL:nLeft SAY oDefXlsCol:oSay3 PROMPT "Col."+CRLF+"Final";
    SIZE 32,35 PIXEL FONT oFont COLOR nClrText,oDp:nGris

  @ oDefXlsCol:oCantid:nTop-35,oDefXlsCol:oCantid:nLeft SAY oDefXlsCol:oSay3 PROMPT "Cant."+CRLF+"Reg.";
    SIZE 40,40-5 PIXEL FONT oFont COLOR nClrText,oDp:nGris

  @ 51+15,400-25 METER oDefXlsCol:oMeterR VAR oDefXlsCol:nRecord   PIXEL FONT oFont COLOR CLR_WHITE,0        SIZE 390,20
  @ 71+15,400-25 SAY oDefXlsCol:oSay PROMPT "Progreso"             PIXEL FONT oFont COLOR nClrText,oDp:nGris SIZE 320,18

  @ 82+22,400-25 GET oDefXlsCol:oCondic Var oDefXlsCol:cCondic FONT oFont;
                 HSCROLL FONT oFont MULTILINE;
                 SIZE 390,90 PIXEL OF oBar
         
  oDefXlsCol:oBar:=oBar
 
RETURN .T.

/*
// Evento para presionar CLICK
*/
FUNCTION RUNCLICK()
  LOCAL aLine:=oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt],I
  LOCAL cMax :=""
  LOCAL aData:={}

  IF Empty(aLine[6])

     // aData:=ACLONE(oDefXlsCol:oBrwA:aArrayData)
     aData:=ACLONE(oDefXlsCol:oBrwA:aData) // 17/07/2023 Utiliza la data del buscador

     ADEPURA(aData,{|a,n|Empty(a[6]) .OR. LEN(a[6])=1})

     IF Empty(aData)
       aData:=ACLONE(oDefXlsCol:oBrwA:aArrayData)
       ADEPURA(aData,{|a,n|Empty(a[6])})
     ENDIF

     aData:=ASORT(aData,,, { |x, y| x[6] > y[6] })

     IF Empty(aData)
        cMax:="A"
     ELSE
       cMax :=aData[1,6]
       cMax :=oDefXlsCol:NEXTCOL(cMax,.F.)
     ENDIF

     oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,6]:=cMax
     
     oDefXlsCol:SETBRWDATA(NIL,cMax)

// ? cMax,"cMax"

     oDefXlsCol:oBrwA:DrawLine(.T.)

  ENDIF

RETURN .T.

FUNCTION RUNCLICKD()
  LOCAL cCol:=CHR(64+oDefXlsCol:oBrwD:nColSel)

  oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,6]:=cCol
  oDefXlsCol:oBrwA:DrawLine(.T.)

RETURN .T.

FUNCTION NEXTCOL(cMax,lSay)
  LOCAL cNext,cMax2:="",cMaxIni:=cMax

  DEFAULT lSay:=.F.

  cMax :=ALLTRIM(cMax)

  IF LEN(cMax)=2
    cMax2:=ALLTRIM(RIGHT(cMaxIni,1))
    cMax :=ALLTRIM(LEFT(cMaxIni,1))
  ENDIF
  
  IF cMax="Z" .AND. Empty(cMax2)
    cMax :="AA"
    RETURN cMax
  ENDIF

  IF LEN(cMaxIni)=2

    IF cMax2="Z"
      cMax :=CHR(ASC(cMax)+1)
      cMax2:="A"
      cMax :=cMax+cMax2
    ELSE
      cMax2:=CHR(ASC(cMax2)+1)
      cMax :=cMax+cMax2
    ENDIF

  ELSE

     IF lSay
        ? cMax,"en next"
     ENDIF

     cMax:=CHR(ASC(cMax)+1)

  ENDIF

RETURN cMax

/*
// Imprimir
*/
FUNCTION IMPRIMIR()
RETURN .T.

FUNCTION LEEFECHAS()
RETURN .T.


FUNCTION HACERWHERE(dDesde,dHasta,cWhere_,lRun)
RETURN cWhere


FUNCTION LEERDATA(cWhere,oBrw,cServer,cTableA)
RETURN aData

FUNCTION SAVEPERIODO()
RETURN .T.

/*
// Permite Crear Filtros para las Búquedas
*/
FUNCTION BRWQUERY()
     EJECUTAR("BRWQUERY",oDefXlsCol)
RETURN .T.

/*
// Ejecución Cambio de Linea 
*/
FUNCTION BRWCHANGE()

/*
  oDefXlsCol:cMemo:=oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,3]

  IF oDefXlsCol:oBrwD<>NIL
    oDefXlsCol:oBrwD:VarPut(oDefXlsCol:cMemo,.T.)
  ENDIF
*/

RETURN NIL

/*
// Refrescar Browse
*/
FUNCTION BRWREFRESCAR()
    LOCAL cWhere

    IF Type("oDefXlsCol")="O" .AND. oDefXlsCol:oWnd:hWnd>0

      cWhere:=" "+IIF(!Empty("oDefXlsCol":cWhere_),"oDefXlsCol":cWhere_,"oDefXlsCol":cWhere)
      cWhere:=STRTRAN(cWhere," WHERE ","")


      oDefXlsCol:LEERDATA(oDefXlsCol:cWhere_,oDefXlsCol:oBrwA,oDefXlsCol:cServer)
      oDefXlsCol:oWnd:Show()
      oDefXlsCol:oWnd:Maximize()

    ENDIF

RETURN NIL

FUNCTION RUNFILEXLS()
  LOCAL cFile:=ALLTRIM(oDefXlsCol:IXL_FILE)

  SHELLEXECUTE(oDp:oFrameDp:hWND,"open",cFile)
  CursorWait()

RETURN .T.

FUNCTION BRWRESTOREPAR()
RETURN EJECUTAR("BRWRESTOREPAR",oDefXlsCol)


FUNCTION VERAUDITORIA()
RETURN EJECUTAR("DPREPORTRECOVER",oDefXlsCol:cCodigo,NIL,NIL,NIL,NIL,NIL,NIL,"DPAUDELIMODCNF_HIS")

FUNCTION XLSLOADDBF()
   LOCAL aDbf:={},aLine,aFields:={},nDif,oTable,cFileDbf,nLinIni
   LOCAL nHead,nColGet,lStruct:=.T.,cMaxCol:=oDefXlsCol:IXL_MAXCOL,aSelect:={},lData:=.t.,uValue,cFileXls,cTable,cMemo
   LOCAL aStruct:={}

   CLOSE ALL

   CURSORWAIT()

//  EJECUTAR("DPIMPRXLSASIENTOS",oDefXlsCol:cCodigo,.T.,.F.,oDefXlsCol:nCantid,oDefXlsCol:oMemo,oDefXlsCol:oMeterR,oDefXlsCol:oSay,.F.,.T.)
// ? oDp:cFileDbf,"oDp:cFileDbf"

   oTable  :=OpenTable("SELECT * FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",oDefXlsCol:cCodigo),.T.)
   cFileXls:=ALLTRIM(oTable:IXL_FILE  )
   cTable  :=ALLTRIM(oTable:IXL_TABLA )
   cMemo   :=ALLTRIM(oTable:IXL_MEMO  )
   nLinIni :=MAX(oDefXlsCol:IXL_LININI,1)

   oTable:End()

   cfileDbf:=STRTRAN( cFileXls,".xls",".dbf")

   // debe cargar estructura
   aStruct:=EJECUTAR("XLSTODBF",cFileXls,cFileDbf,oDefXlsCol:oMeterR,oDefXlsCol:oSay,.T.,nLinIni,oDefXlsCol:nCantid,nHead,nColGet,lStruct,cMaxCol,aSelect)

   IF Empty(aStruct)
      RETURN .F.
   ENDIF

   CLOSE ALL
   oDp:cFileDbf:=cFileDbf

   lStruct:=.F.

   EJECUTAR("XLSTODBF",cFileXls,cFileDbf,oDefXlsCol:oMeterR,oDefXlsCol:oSay,.T.,nLinIni,oDefXlsCol:nCantid,nHead,nColGet,lStruct,cMaxCol,aSelect)

   CLOSE ALL

   IF !FILE(oDp:cFileDbf)
     RETURN .F.    
   ENDIF

   SELECT A
   USE (oDp:cFileDbf)

   aFields:=DBSTRUCT()

   ASIZE(aFields,oDefXlsCol:nMaxCol) // no puede pasar de 

   nDif   :=24-LEN(aFields)+2

   WHILE !A->(EOF())
      aLine:={}     
      AEVAL(aFields,{|a,n| AADD(aLine,CTOO(FIELDGET(n),"C"))})
      // Complementa las Columnas
      IF nDif>0
        AEVAL(ARRAY(nDif),{|a,n|AADD(aLine,"")})
      ENDIF

      AADD(aDbf,ACLONE(aLine))
      A->(DBSKIP())
   ENDDO

   oDefXlsCol:oBrwD:SetArray(ACLONE(aDbf))

   AEVAL(oDefXlsCol:oBrwD:aCols,{|oCol,n|oCol:cHeader:=oDefXlsCol:aHead[n],oCol:nWidth:=40})

   oDefXlsCol:oBrwD:Gotop()
   oDefXlsCol:oBrwD:Refresh(.F.)

RETURN .T.

FUNCTION SORTXLSCOL(r,c,f,o) 
   LOCAL aData :=ACLONE(oDefXlsCol:oBrwA:aArrayData)
   LOCAL aNoSel:=ACLONE(oDefXlsCol:oBrwA:aArrayData)

   ADEPURA(aData ,{|a,n|  Empty(a[6])})
   ADEPURA(aNoSel,{|a,n| !Empty(a[6])})

   IF !(o:cOrder == "" .or. o:cOrder == "D")
     aData:=ASORT(aData,,, { |x, y| x[6] < y[6] })
   ELSE
     aData:=ASORT(aData,,, { |x, y| x[6] > y[6] })
   ENDIF

   IF !(o:cOrder == "" .or. o:cOrder == "D")
      AEVAL(aNoSel,{|a,n| AADD(aData,a)})
      o:cOrder     := "D"
      o:nHeadBmpNo := 1
   ELSE
      AEVAL(aData,{|a,n| AADD(aNoSel,a)})
      aData:=ACLONE(aNoSel)
      o:cOrder     := "A"
      o:nHeadBmpNo := 1
   ENDIF
    
   oDefXlsCol:oBrwA:aArrayData:=ACLONE(aData)
   oDefXlsCol:oBrwA:Refresh(.F.)
RETURN .T.

FUNCTION PUTCOLUMNA(oCol,uValue,nCol)
  LOCAL aLine:=oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt],I,nAt:=0,cField:=""

  uValue:=ALLTRIM(UPPER(uValue))

  IF Empty(uValue)
     oDefXlsCol:XLSDELCOL()
     RETURN .T.
  ENDIF

  IF uValue>="A" .AND. uValue<="ZZ"

    // Si filtra debemos tomar en cuenta todos datos

    FOR I=1 TO LEN(oDefXlsCol:oBrwA:aArrayData)
       IF ALLTRIM(oDefXlsCol:oBrwA:aArrayData[I,nCol])=ALLTRIM(uValue)
          oDefXlsCol:oBrwA:aArrayData[I,nCol]:=SPACE(2)
       ENDIF
    NEXT I

    uValue:=PADR(uValue,2)

    oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,nCol]:=PADR(uValue,2)
    oDefXlsCol:oBrwA:DrawLine(.t.)

    // Debe colocar la Letra en el Origen de la Data, en buscar y reemplazar se pierde

    oDefXlsCol:SETBRWDATA(cField,uValue)

  ENDIF

RETURN .T.

FUNCTION SETBRWDATA(cField,uValue)
  LOCAL nAt,nCol:=6

  IF Empty(cField)
     cField:=ALLTRIM(oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,1])
  ENDIF

  nAt   :=ASCAN(oDefXlsCol:oBrwA:aData,{|a,n| ALLTRIM(a[1])==cField})

  IF nAt>0
     oDefXlsCol:oBrwA:aData[nAt,nCol]:=PADR(uValue,2)
  ENDIF

// ? cField,uValue,"cField,uValue",nAt,nCol

RETURN .T.

FUNCTION PUTDEFAULT(oCol,uValue,nCol)
  LOCAL cField:=oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,1]

  oDefXlsCol:SETBRWDATA(cField,uValue)

  oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,nCol]:=uValue
  oDefXlsCol:oBrwA:DrawLine(.t.)

RETURN .F.

/*
// Grabar Lbx
*/
FUNCTION SAVELBX(cFileIXL)
   LOCAL cMemo:="",I,cCol,aFields,cFile:=NIL

   oDefXlsCol:oSay:SetText("Guardando...")
  
   DEFAULT cFileIXL:=oDefXlsCol:cFileIXL

   cFile:=cFileNoPath(cFileIXL)

   // aFields:=ACLONE(oDefXlsCol:oBrwA:aArrayData) 17/07/2023 debe guardar la data almacenada en Buscar
   aFields:=ACLONE(oDefXlsCol:oBrwA:aArrayData)

   ADEPURA(aFields,{|a,n|Empty(a[6])})

   // 30/05/2023
   IF Empty(oDefXlsCol:oBrwD:aArrayData[1,1])
      CursorWait()
      // 30/05/2023 oDefXlsCol:XLSLOADDBF()
   ENDIF

   // ViewArray(aFields)
   // Borrar Columnas

   FOR I=1 TO LEN(aFields)

     cCol:="COL"+STRZERO(I,2)

     cMemo:=cMemo+IF( Empty(cMemo),"",CRLF)+;
            cCol+"_HEADER :="+ALLTRIM(aFields[I,1])+CRLF+;
            cCol+"_COLUMN :="+ALLTRIM(aFields[I,6])+CRLF+;
            cCol+"_DEFAULT:="+ALLTRIM(aFields[I,5])+CRLF+;
            cCol+"_EMPTY  :="+IF(.F.,"T","F")+CRLF+;
            cCol+"_KEY    :="+IF(.F.,"T","F")
   
   NEXT I

   IF Empty(oDefXlsCol:cCodigo)
     FERASE(oDefXlsCol:cFileIXL)
     DpWrite(oDefXlsCol:cFileIXL,cMemo)
     EJECUTAR("FILESAVEBIN",oDefXlsCol:cFileIXL,cFileNoPath(oDefXlsCol:cFileIXL),NIL,oDefXlsCol:cFileIXL,"DPFILES",NIL,.F.,.T.,0)
   ENDIF

   IF ValType(oDefXlsCol:bPostSave)="B"
      EVAL(oDefXlsCol:bPostSave,cMemo)
   ENDIF

   IF !Empty(oDefXlsCol:cCodigo)

     SQLUPDATE("DPIMPRXLS",{"IXL_MEMO","IXL_LININI"         ,"IXL_LINFIN"         ,"IXL_MINCOL"          ,"IXL_MAXCOL"         },;
                           {cMemo     ,oDefXlsCol:IXL_LININI,oDefXlsCol:IXL_LINFIN,oDefXlsCol:IXL_MINCOL ,oDefXlsCol:IXL_MAXCOL},;
                           "IXL_CODIGO"+GetWhere("=",oDefXlsCol:cCodigo))

   ENDIF

   oDefXlsCol:oSay:SetText("Concluido")

   CLOSE ALL

   oDefXlsCol:CLOSE()

   EJECUTAR("DPIMPRXLSRUN",oDefXlsCol:cCodigo,oDefXlsCol:oMeterR,oDefXlsCol:oSay)
 
RETURN .T.

/*
// remover Columna
*/
FUNCTION XLSDELCOL()

  oDefXlsCol:oBrwA:aArrayData[oDefXlsCol:oBrwA:nArrayAt,6]:=SPACE(2)
  oDefXlsCol:oBrwA:DrawLine(.t.)

RETURN .T.

/*
// remover Columna
*/
FUNCTION XLSMODCOL()

  oDefXlsCol:oBrwA:aCols[6]:nEditType:=1
  oDefXlsCol:oBrwA:DrawLine(.t.)
  oDefXlsCol:oBrwA:aCols[6]:Edit()

RETURN .T.

FUNCTION BTNMENU(nOption,cOption)

   IF nOption=1
     oDefXlsCol:PRGDEFREFEQT()
   ENDIF

   IF nOption=2
   ENDIF

   IF nOption=3
   ENDIF

RETURN .T.

FUNCTION PRGDEFREFEQT()
   LOCAL bRun,cWhere,cMemo,cFile,cField:="IXL_ETQPRG",cTitle,cTable
   LOCAL cCodigo:=oDefXlsCol:cCodigo,cPrg:=""

   HrbLoad("DPXBASE.HRB") // Carga M?dulo DpXbase

   bRun  :={||MensajeErr("FINAL")}
   cWhere:="IXL_CODIGO"+GetWhere("=",cCodigo)
   cMemo :=SQLGET("DPIMPRXLS",cField,cWhere)
   cMemo :=IF(Empty(cMemo),cPrg,ALLTRIM(cMemo))

   IF Empty(cMemo) .OR. LEN(cMemo)=0
      cMemo:=MemoRead("DP\DPIMPXLSREFETQ.TXT")
   ENDIF

   cFile :="DP\DPIMPRXLS_"+cField+".TXT"

   cCodigo:=ALLTRIM(cCodigo)
   cMemo  :=STRTRAN(cMemo,"<CODIGO>",cCodigo)
   cMemo  :=STRTRAN(cMemo,"<FECHA>" ,DTOC(oDp:dFecha))

   cTitle:="Editar Programa para generar referencias de Campos para Definición: "+cCodigo

   DPXBASEEDIT(3,cCodigo,bRun,NIL,cMemo,"DPIMPRXLS",cField,cWhere,cTitle)

RETURN .T.
// EOF


// EOF
