// Programa   : DPIMPRXLSRUN
// Fecha/Hora : 11/03/2014 01:13:11
// Propósito  : Ejecutar Importar Datos desde EXCEL
// Creado Por : Juan Navas
// Llamado por: DPIMPRXLS.LBX
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo)
  LOCAL cDir    :=Lower(cFilePath(GetModuleFileName( GetInstance() )))
  LOCAL cFileXls:=SPACE(250),oFontG,nLinIni:=2
  LOCAL cTable  :="",oTable
  LOCAL cTitle  :="",aRef:={},cMemo
  LOCAL cPrg    :=""

  IF Type("oImpXls")="O" .AND. oImpXls:oWnd:hWnd>0
     RETURN EJECUTAR("BRRUNNEW",oImpXls,GetScript())
  ENDIF

  IF Empty(oDp:hDllRtf) // Carga RTF
     oDp:hDLLRtf := LoadLibrary( "Riched20.dll" )
  ENDIF

  DEFAULT cCodigo:=SQLGET("DPIMPRXLS","IXL_CODIGO")

  oTable  :=OpenTable("SELECT * FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)

  cTitle  :=ALLTRIM(oTable:IXL_DESCRI)
  cFileXls:=ALLTRIM(oTable:IXL_FILE  )
  cTable  :=ALLTRIM(oTable:IXL_TABLA )
  cMemo   :=ALLTRIM(oTable:IXL_MEMO  )
//  cPrg    :=ALLTRIM(oTable:IXL_DEFPRG   )


// ? cPrg,"cPrg"

  nLinIni :=oTable:IXL_LININI

  cFileXls:=EJECUTAR("PATHADMWIN",cFileXls)
  oTable:End()

  IF Empty(cMemo) .AND. Empty(oTable:IXL_DEFPRG)
     MsgRun("Importar Excel "+ALLTRIM(cCodigo)+" Sin Definición")
     EJECUTAR("DPIMPRXLSDEF",cCodigo)
     RETURN .F.
  ENDIF

  DEFINE FONT oFontG NAME "Tahoma"   SIZE 0, -14

  cFileXls:=STRTRAN(cFileXls,"\\","\")

//  oImpXls:=DPEDIT():New("Importar Datos desde Excel ["+cTitle+"]","IMPORTXLS.EDT","oImpXls",.T.)

  DpMdi(cTitle,"oImpXls","IMPORTXLS.EDT")

  oImpXls:Windows(0,0,oDp:aCoors[3]-180,MIN(1000,oDp:aCoors[4]-10),.T.) // Maximizado

  oImpXls:nRecord :=0
  oImpXls:oMeterR :=NIL
  oImpXls:lMsgBar :=.F.
  oImpXls:cMemo   :=SPACE(10)
  oImpXls:lChk    :=.F.
  oImpXls:lTodos  :=.T.
  oImpXls:nCantid :=1
  oImpXls:nLinIni :=nLinIni // Linea Excel, Inicio de Lectura 
  oImpXls:cCodigo :=cCodigo
  oImpXls:cTable  :=cTable
  oImpXls:lBrowse :=.F.
  oImpXls:SetTable(oTable)
  oImpXls:lIntRef  :=oTable:IXL_INTREF
  oImpXls:lBarDef  :=.T.
  oImpXls:oMemo    :=NIL
  oImpXls:cMemo    :=""
  oImpXls:nOption  :=3

  oImpXls:cFileXls:=PADR(cFileXls,250)
  oImpXls:aRef    :=ACLONE(aRef)

//  @ 10,1 GET oImpXls:oMemo VAR oImpXls:cMemo MULTI READONLY

  @ 10,0 RICHEDIT oImpXls:oMemo VAR oImpXls:cMemo OF oImpXls:oWnd HSCROLL  FONT oFontG

  oImpXls:oWnd:oClient := oImpXls:oMemo

  oImpXls:Activate({||oImpXls:IXLBOTBAR()})

  oBtn:=BMPGETBTN(oImpXls:oFileXls)

RETURN NIL

/*
// Barra de Botones
*/
FUNCTION IXLBOTBAR()
  LOCAL oBar,oBtn,oFont,oCursor,oFontG
  LOCAL nClrText:=0

  DEFINE CURSOR oCursor HAND
  DEFINE FONT oFont NAME "Tahoma"   SIZE 0, -14 BOLD

  DEFINE FONT oFontG NAME "Tahoma"   SIZE 0, -14

// DEFINE BUTTONBAR oBar SIZE 39, 39 3D OF oImpXls:oDlg

  DEFINE BUTTONBAR oBar SIZE 58,60+60+40+30 OF oImpXls:oDlg 3D CURSOR oCursor

  DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME oDp:cPathBitMaps+"run.BMP";
          TOP PROMPT "Ejecutar";
          ACTION oImpXls:IMPORTAR()

  DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME oDp:cPathBitMaps+"configura.BMP";
          TOP PROMPT "Config";
          ACTION EJECUTAR("DPIMPRXLSDEF",oImpXls:cCodigo)


  DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME oDp:cPathBitMaps+"excel.BMP";
          TOP PROMPT "Abrir";
          ACTION  EJECUTAR("RUNEXCEL",oImpXls:cFileXls)


  DEFINE BUTTON oBtn;
          OF oBar;
          NOBORDER;
          FONT oFont;
          FILENAME "BITMAPS\XSALIR.BMP";
          TOP PROMPT "Salir";
          ACTION oImpXls:Close()

  oBar:SetColor(CLR_BLACK,oDp:nGris)
  AEVAL(oBar:aControls,{|o,n|o:SetColor(CLR_BLACK,oDp:nGris)})

  oImpXls:SETBTNBAR(52,60,oBar)

   SetWndDefault( oBar) 

  @ 3,2 SAY "Nombre del Archivo"

  //
  // Campo : cFileXls  
  // Uso   : Dirección del Archivo                   
  //
  @ 10, 1 BMPGET oImpXls:oFileXls    VAR oImpXls:cFileXls   ;
             NAME "BITMAPS\FIND.BMP";
             ACTION  (cFile:=cGetFile32("Ficheros Excel X (*.XLSX) | *.XLSX| Excel (*.XLS) | *.XLS" ,;
                     "Seleccionar Ficheros (*.xlsx,*.xls)",1,cFilePath(oImpXls:cFileXls),.f.,.t.),;
                     cFile:=STRTRAN(cFile,"/","/"),;
                     oImpXls:cFileXls:=IIF(!EMPTY(cFile),cFile,oImpXls:cFileXls),;
                     oImpXls:oFileXls:KeyBoard(13));
                     FONT oFontG;
                     SIZE 200,10 OF oBar

  //
  // Campo : cFileXls  
  // Uso   : Dirección del Archivo                   
  //
  @ 94, 1 GET oImpXls:oCantid VAR oImpXls:nCantid SPINNER PICTURE "99999" ;
              FONT oFontG;
              SIZE 40,10;
              WHEN .T. SPINNER OF oBar RIGHT

//       WHEN !oImpXls:lTodos SPINNER OF oBar


  //
  // Campo : cFileXls  
  // Uso   : Línea de Inicio                  
  //
  @ 94, 1 GET oImpXls:oLinIni VAR oImpXls:nLinIni SPINNER PICTURE "99999" ;
              FONT oFontG;
              SIZE 40,10 OF oBar RIGHT

 //
  // Campo : IXL_LINFIN
  // Uso   : Línea Final                         
  //
  @ 11.5, 1.0 GET oImpXls:oIXL_LINFIN  VAR oImpXls:IXL_LINFIN  PICTURE "99";
                    WHEN .T.;
                    FONT oFontG;
                    SIZE 8,10 SPINNER;
                    RIGHT SPINNER OF oBar


    oImpXls:oIXL_LINFIN:cMsg    :="Línea Final"
    oImpXls:oIXL_LINFIN:cToolTip:="Línea Final"

  @ oImpXls:oIXL_LINFIN:nTop-08,oImpXls:oIXL_LINFIN:nLeft SAY "Línea"+CRLF+"Final" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris OF oBar



  //
  // Campo : IXL_MINCOL
  // Uso   : Columna de Inicio                         
  //
  @ 11.5, 1.0 GET oImpXls:oIXL_MINCOL  VAR oImpXls:IXL_MINCOL;
                    WHEN .T.;
                    FONT oFontG;
                    SIZE 8,10;
                    RIGHT SPINNER OF oBar 


    oImpXls:oIXL_MINCOL:cMsg    :="Columna de Inicio"
    oImpXls:oIXL_MINCOL:cToolTip:="Columna de Inicio"

  @ oImpXls:oIXL_MINCOL:nTop-08,oImpXls:oIXL_MINCOL:nLeft SAY "Columna"+CRLF+"Inicial" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris

  //
  // Campo : IXL_MAXCOL
  // Uso   : Columna Final                         
  //
  @ 11.5, 10 GET oImpXls:oIXL_MAXCOL  VAR oImpXls:IXL_MAXCOL;
                    WHEN .T.;
                    FONT oFontG;
                    SIZE 8,10;
                    RIGHT SPINNER OF oBar


    oImpXls:oIXL_MAXCOL:cMsg    :="Columna Final"
    oImpXls:oIXL_MAXCOL:cToolTip:="Columna Final"

  @ oImpXls:oIXL_MAXCOL:nTop-08,oImpXls:oIXL_MAXCOL:nLeft SAY "Columna"+CRLF+"Final" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris

  @ 4,2 SAY oImpXls:oSay PROMPT "Progreso" RIGHT

  @ 4,2 SAY oImpXls:oSay2 PROMPT "Línea"+CRLF+"Inicio" 

  @ 4,2 SAY oImpXls:oSay2 PROMPT "Cant."+CRLF+"Reg." 


  @ 01,01 METER oImpXls:oMeterR VAR oImpXls:nRecord
         

  @ 2,20 CHECKBOX oImpXls:lChk    PROMPT ANSITOOEM("Revisión")
  @ 3,20 CHECKBOX oImpXls:lTodos  PROMPT ANSITOOEM("Importar Todos los Registros") ON CHANGE (oImpXls:oCantid:ForWhen(.T.))
  @ 4,20 CHECKBOX oImpXls:lBrowse PROMPT ANSITOOEM("Mostrar Browse")
  @ 4,20 CHECKBOX oImpXls:lIntRef PROMPT ANSITOOEM("Integridad Referencial")


RETURN .T.

PROCE IMPORTAR()
   LOCAL oTable,cMemo

   CursorWait()

   oTable:=OpenTable("SELECT * FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",oImpXls:cCodigo),.T.)

   cMemo:=oTable:IXL_DEFPRG // Programa personalizado
   // oTable:lAuditar:=.F.
   oTable:cPrimary:="IXL_CODIGO"
   oTable:Replace("IXL_FILE"  ,oImpXls:cFileXls )
   oTable:Replace("IXL_LININI",oImpXls:nLinIni  )
   oTable:Replace("IXL_INTREF",oImpXls:lIntRef )
   oTable:Commit(oTable:cWhere)
   oTable:End(.t.)

   // ? LEN(cMemo),oImpXls:cCodigo,CLPCOPY(cMemo)

   IF !Empty(cMemo)
      EJECUTAR("RUNMEMO",cMemo,oImpXls:cCodigo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:oMemo)
      RETURN .T.
   ENDIF

   IF "Prod"$oImpXls:cTable .OR. "DPINV"$oTable:IXL_TABLA
     EJECUTAR("DPIMPRXLSINV",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .T.
   ENDIF

   IF "Activos"$oImpXls:cTable
     EJECUTAR("DPIMPRXLSACT",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .T.
   ENDIF

   IF "Prov"$oImpXls:cTable
     EJECUTAR("DPIMPRXLSPRO",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .T.
   ENDIF

   IF "Pagar"$oImpXls:cTable
     EJECUTAR("DPIMPRXLSCXP",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .F.
   ENDIF

   IF "Compras"$oImpXls:cTable
     EJECUTAR("DPIMPRXLSLIBC",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .T.
   ENDIF

   IF "Cobrar"$oImpXls:cTable
     EJECUTAR("DPIMPRXLSCXC",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .T.
   ENDIF

   IF "PLAN"$(UPPER(oImpXls:cTable)) .OR. "DPCTA"=ALLTRIM(oTable:IXL_TABLA)
     EJECUTAR("DPIMPRXLSCTA",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .T.
   ENDIF

//   IF ("DPCLIENTENT"$(UPPER(oImpXls:cTable)) .AND. (oDp:cDp="192.168.10.13" .OR."DATAPRO"$oDp:cEmpresa)) .OR. .T.
//     EJECUTAR("DPIMPRXLSEVALADP",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
//     RETURN .T.
//   ENDIF

   IF !Empty(SQLGET("DPTABLAS","TAB_NOMBRE","TAB_NOMBRE"+GetWhere("=",oImpXls:cTable)))
     EJECUTAR("DPIMPRXLSTABLA",oImpXls:cCodigo,oImpXls:lChk,oImpXls:lTodos,IIF(oImpXls:lTodos,0,oImpXls:nCantid),oImpXls:oMemo,oImpXls:oMeterR,oImpXls:oSay,oImpXls:lBrowse)
     RETURN .T.
   ENDIF

   MensajeErr("Tabla "+oImpXls:cTable+" no posee Proceso para Importar")


RETURN .F.
// EOF


