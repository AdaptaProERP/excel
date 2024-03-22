// Programa   : DPIMPRXLS
// Fecha/Hora : 06/03/2014 01:21:39
// Propósito  : Incluir/Modificar DPIMPRXLS
// Creado Por : DpXbase
// Llamado por: DPIMPRXLS.LBX
// Aplicación : Administración del Sistema                                  
// Tabla      : DPIMPRXLS

#INCLUDE "DPXBASE.CH"
#INCLUDE "TSBUTTON.CH"
#INCLUDE "IMAGE.CH"

FUNCTION DPIMPRXLS(nOption,cCodigo,cTable,cOption,aRef)
  LOCAL oBtn,oTable,oGet,oFont,oFontB,oFontG
  LOCAL cTitle,cSql,cFile,cExcluye:=""
  LOCAL nClrText
  LOCAL cTitle :="Definición de Importación desde Excel",;
        aItems1:={},;
        aTablas:={}

// GETOPTIONS("DPIMPRXLS","IXL_PREDEF",NIL,.T.),;
//ASQL("SELECT TAB_NOMBRE FROM DPTABLAS ")
// ? cTable,"cTable",cOption,"<-cOption"
        
  cExcluye:="IXL_CODIGO,;
             IXL_DESCRI,;
             IXL_TABLA,;
             IXL_ACTIVO,;
             IXL_FILE,;
             IXL_LININI"

  DEFAULT cCodigo:="1234"

  DEFAULT nOption:=1,;
          cTable :=""

  nOption:=IIF(nOption=2,0,nOption) 

  oDp:aTablasXls:={}
  aItems1:={}
  EJECUTAR("DPIMPXLSDEFTABLA")

  AEVAL(oDp:aTablasXls,{|a,n| AADD(aItems1,a[1])})

  IF ASCAN(aItems1,"Indefinido")=0
    AINSERTAR(aItems1,1,"Indefinido")
  ENDIF

// AEVAL(aTablas,{|a,n| AADD(aItems1,ALLTRIM(a[1]))})
// ViewArray(aItems1)
// RETURN NIL

  DEFINE FONT oFont  NAME "Tahoma" SIZE 0, -10 BOLD
  DEFINE FONT oFontB NAME "Tahoma" SIZE 0, -12 BOLD ITALIC
  DEFINE FONT oFontG NAME "Tahoma" SIZE 0, -11

  nClrText:=10485760 // Color del texto

  IF nOption=1 // Incluir
    cSql     :=[SELECT * FROM DPIMPRXLS WHERE ]+BuildConcat("IXL_CODIGO")+GetWhere("=",cCodigo)+[]
    cTitle   :=" Incluir {oDp:DPIMPRXLS}"
  ELSE // Modificar o Consultar
    cSql     :=[SELECT * FROM DPIMPRXLS WHERE ]+BuildConcat("IXL_CODIGO")+GetWhere("=",cCodigo)+[]
    cTitle   :=IIF(nOption=2,"Consultar","Modificar")+" Definición de Importación desde Excel   "
    cTitle   :=IIF(nOption=2,"Consultar","Modificar")+" {oDp:DPIMPRXLS}"
  ENDIF

  oTable   :=OpenTable(cSql,"WHERE"$cSql) // nOption!=1)

  IF nOption=1 .AND. oTable:RecCount()=0 // Genera Cursor Vacio
     oTable:End()
     cSql     :=[SELECT * FROM DPIMPRXLS]
     oTable   :=OpenTable(cSql,.F.) // nOption!=1)


  ENDIF

  oTable:cPrimary:="IXL_CODIGO" // Clave de Validación de Registro

  oIMPRXLS:=DPEDIT():New(cTitle,"DPIMPRXLS.edt","oIMPRXLS" , .F. )

  oIMPRXLS:nOption  :=nOption
  oIMPRXLS:SetTable( oTable , .F. ) // Asocia la tabla <cTabla> con el formulario oIMPRXLS
  oIMPRXLS:SetScript("DPIMPRXLS")        // Asigna Funciones DpXbase como Metodos de oIMPRXLS
  oIMPRXLS:SetDefault()       // Asume valores standar por Defecto, CANCEL,PRESAVE,POSTSAVE,ORDERBY
  oIMPRXLS:nClrPane :=oDp:nGris
  oIMPRXLS:cWhereLbx:=""
  oIMPRXLS:aRef     :={}

  IF oIMPRXLS:nOption=1 // Incluir en caso de ser Incremental
     // oIMPRXLS:RepeatGet(NIL,"IXL_CODIGO") // Repetir Valores

     oIMPRXLS:IXL_ACTIVO:=.T.
     oIMPRXLS:IXL_MEMO  :=""
     oIMPRXLS:cWhereLbx :=oIMPRXLS:oDpLbx:cWhere 

     IF "DPCTA"$oIMPRXLS:cWhereLbx .OR. "DPCTA"$cTable
        oIMPRXLS:IXL_TABLA:="DPCTA"
        oIMPRXLS:IXL_PREDEF:="Plan de Cuentas"
     ENDIF

     IF "DPCLIENTES"$oIMPRXLS:cWhereLbx .OR. "DPCLIENTES"$cTable
        oIMPRXLS:IXL_TABLA:="DPCLIENTES"
        oIMPRXLS:IXL_PREDEF:="Clientes"
     ENDIF

     IF !Empty(cTable)
        oIMPRXLS:IXL_TABLA:=cTable
     ENDIF

     oIMPRXLS:IXL_LININI:=1
     oIMPRXLS:IXL_LINFIN:=0

     // AutoIncremental 
  ENDIF

  IF !Empty(cOption)
     oIMPRXLS:IXL_PREDEF:=cOption
     aItems1:={cOption}
  ENDIF

  IF !Empty(aRef)
     oIMPRXLS:aRef:=ACLONE(aRef)
  ENDIF

  //Tablas Relacionadas con los Controles del Formulario

  oIMPRXLS:IXL_FILE:=EJECUTAR("PATHADMWIN",oIMPRXLS:IXL_FILE)
  oIMPRXLS:cPostSave:="POSTSAVE"

  oIMPRXLS:CreateWindow()       // Presenta la Ventana

  // Opciones del Formulario

  
  //
  // Campo : IXL_CODIGO
  // Uso   : Código                                  
  //
  @ 3.0, 1.0 GET oIMPRXLS:oIXL_CODIGO  VAR oIMPRXLS:IXL_CODIGO  VALID oIMPRXLS:ValUnique(oIMPRXLS:IXL_CODIGO);
                    WHEN (AccessField("DPIMPRXLS","IXL_CODIGO",oIMPRXLS:nOption);
                    .AND. oIMPRXLS:nOption!=0);
                    FONT oFontG;
                    SIZE 120,10

    oIMPRXLS:oIXL_CODIGO:cMsg    :="Código"
    oIMPRXLS:oIXL_CODIGO:cToolTip:="Código"

  @ oIMPRXLS:oIXL_CODIGO:nTop-08,oIMPRXLS:oIXL_CODIGO:nLeft SAY "Código" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris


  //
  // Campo : IXL_DESCRI
  // Uso   : Descripción                             
  //
  @ 4.8, 1.0 GET oIMPRXLS:oIXL_DESCRI  VAR oIMPRXLS:IXL_DESCRI ;
                    WHEN (AccessField("DPIMPRXLS","IXL_DESCRI",oIMPRXLS:nOption);
                    .AND. oIMPRXLS:nOption!=0);
                    FONT oFontG;
                    SIZE 1016,10

    oIMPRXLS:oIXL_DESCRI:cMsg    :="Descripción"
    oIMPRXLS:oIXL_DESCRI:cToolTip:="Descripción"

  @ oIMPRXLS:oIXL_DESCRI:nTop-08,oIMPRXLS:oIXL_DESCRI:nLeft SAY "Descripción" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris


  //
  // Campo : IXL_TABLA 
  // Uso   : Tabla                                   
  //

   IF Empty(aItems1)
     AADD(aItems1,"Indefinido")
   ENDIF 

  @ 6.1, 1.0 COMBOBOX oIMPRXLS:oIXL_PREDEF  VAR oIMPRXLS:IXL_PREDEF ITEMS aItems1;
                      WHEN (AccessField("DPIMPRXLS","IXL_PREDEF",oIMPRXLS:nOption);
                      .AND. oIMPRXLS:nOption!=0 .AND. LEN(oIMPRXLS:oIXL_PREDEF:aItems)>1);
                      FONT oFontG;
                      VALID oIMPRXLS:VALIXLPREDEF()

   ComboIni(oIMPRXLS:oIXL_PREDEF )

   oIMPRXLS:oIXL_PREDEF:cMsg    :="Tabla"
   oIMPRXLS:oIXL_PREDEF:cToolTip:="Tabla"

  @ oIMPRXLS:oIXL_PREDEF:nTop-08,oIMPRXLS:oIXL_PREDEF:nLeft SAY "Predefinido" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris


  //
  // Campo : IXL_ACTIVO
  // Uso   : Activo                                  
  //
  @ 7.9, 1.0 CHECKBOX oIMPRXLS:oIXL_ACTIVO  VAR oIMPRXLS:IXL_ACTIVO  PROMPT ANSITOOEM("Activo");
                    WHEN (AccessField("DPIMPRXLS","IXL_ACTIVO",oIMPRXLS:nOption);
                    .AND. oIMPRXLS:nOption!=0);
                     FONT oFont COLOR nClrText,NIL SIZE 76,10;
                    SIZE 4,10

    oIMPRXLS:oIXL_ACTIVO:cMsg    :="Activo"
    oIMPRXLS:oIXL_ACTIVO:cToolTip:="Activo"


  //
  // Campo : IXL_TABLA
  // Uso   : Tabla Solicitante                       
  //
  @ 1.0, 1.0 BMPGET oIMPRXLS:oIXL_TABLA  VAR oIMPRXLS:IXL_TABLA ;
             NAME "BITMAPS\find.BMP";
             ACTION (oDpLbx:=DpLbx("DPTABLAS",NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,oIMPRXLS:oIXL_TABLA),oDpLbx:GetValue("TAB_NOMBRE",oIMPRXLS:oIXL_TABLA));
             VALID oIMPRXLS:VALTABLA(.F.);
             WHEN ("Indefinido"$oIMPRXLS:IXL_PREDEF);
             FONT oFontG SIZE 200,10

  oIMPRXLS:oIXL_TABLA:cMsg    :="Tabla Solicitante"
  oIMPRXLS:oIXL_TABLA:cToolTip:="Tabla Solicitante"

  @ oIMPRXLS:oIXL_TABLA:nTop-08,oIMPRXLS:oIXL_TABLA:nLeft SAY "Tabla" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,NIL 


  //
  // Campo : IXL_FILE  
  // Uso   : Nombre del Archivo                      
  //
  @ 9.7, 1.0 BMPGET oIMPRXLS:oIXL_FILE;
             VAR oIMPRXLS:IXL_FILE   ;
             NAME "BITMAPS\FIND.BMP";
             ACTION  (cFile:=cGetFile32("Fichero(*.XLSX) |*.XLSX|Ficheros (*.XLSX) |*.XLSX",;
                      "Seleccionar Archivo (*.XLSX)",1,cFilePath(oIMPRXLS:IXL_FILE),.f.,.t.),;
                      cFile:=STRTRAN(cFile,"/","/"),;
                      oIMPRXLS:oIXL_FILE:VarPut(IIF(!EMPTY(cFile),cFile,oIMPRXLS:IXL_FILE),.T.),;
                      DPFOCUS(oIMPRXLS:oIXL_FILE),oIMPRXLS:VALIXLFILE());
                      VALID oIMPRXLS:VALIXLFILE();
                      WHEN (AccessField("DPIMPRXLS","IXL_FILE",oIMPRXLS:nOption);
                            .AND. oIMPRXLS:nOption!=0);
                      FONT oFontG;
                      SIZE 200,10

    oIMPRXLS:oIXL_FILE  :cMsg    :="Nombre del Archivo"
    oIMPRXLS:oIXL_FILE  :cToolTip:="Nombre del Archivo"

  @ oIMPRXLS:oIXL_FILE  :nTop-08,oIMPRXLS:oIXL_FILE  :nLeft SAY "Nombre del Archivo" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris


  //
  // Campo : IXL_LININI
  // Uso   : Línea de Inicio                         
  //
  @ 11.5, 1.0 GET oIMPRXLS:oIXL_LININI  VAR oIMPRXLS:IXL_LININI  PICTURE "99";
                    WHEN (AccessField("DPIMPRXLS","IXL_LININI",oIMPRXLS:nOption);
                    .AND. oIMPRXLS:nOption!=0);
                    FONT oFontG;
                    SIZE 8,10 SPINNER;
                    RIGHT


    oIMPRXLS:oIXL_LININI:cMsg    :="Línea de Inicio"
    oIMPRXLS:oIXL_LININI:cToolTip:="Línea de Inicio"

  @ oIMPRXLS:oIXL_LININI:nTop-08,oIMPRXLS:oIXL_LININI:nLeft SAY "Línea"+CRLF+"Inicial" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris

  //
  // Campo : IXL_LINFIN
  // Uso   : Línea Final                         
  //
  @ 11.5, 1.0 GET oIMPRXLS:oIXL_LINFIN  VAR oIMPRXLS:IXL_LINFIN  PICTURE "99";
                    WHEN (AccessField("DPIMPRXLS","IXL_LINFIN",oIMPRXLS:nOption);
                    .AND. oIMPRXLS:nOption!=0);
                    FONT oFontG;
                    SIZE 8,10 SPINNER;
                    RIGHT


    oIMPRXLS:oIXL_LINFIN:cMsg    :="Línea Final"
    oIMPRXLS:oIXL_LINFIN:cToolTip:="Línea Final"

  @ oIMPRXLS:oIXL_LINFIN:nTop-08,oIMPRXLS:oIXL_LINFIN:nLeft SAY "Línea"+CRLF+"Final" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris



//
  // Campo : IXL_MINCOL
  // Uso   : Columna de Inicio                         
  //
  @ 11.5, 1.0 GET oIMPRXLS:oIXL_MINCOL  VAR oIMPRXLS:IXL_MINCOL;
                    WHEN (AccessField("DPIMPRXLS","IXL_MINCOL",oIMPRXLS:nOption);
                    .AND. oIMPRXLS:nOption!=0);
                    FONT oFontG;
                    SIZE 8,10;
                    RIGHT PICT "@"



    oIMPRXLS:oIXL_MINCOL:cMsg    :="Letra Columna de Inicio"
    oIMPRXLS:oIXL_MINCOL:cToolTip:="Letra Columna de Inicio"

  @ oIMPRXLS:oIXL_MINCOL:nTop-08,oIMPRXLS:oIXL_MINCOL:nLeft SAY "Columna"+CRLF+"Inicial" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris

  //
  // Campo : IXL_MAXCOL
  // Uso   : Columna Final                         
  //
  @ 11.5, 10 GET oIMPRXLS:oIXL_MAXCOL  VAR oIMPRXLS:IXL_MAXCOL;
                    WHEN (AccessField("DPIMPRXLS","IXL_MAXCOL",oIMPRXLS:nOption);
                    .AND. oIMPRXLS:nOption!=0);
                    FONT oFontG;
                    SIZE 8,10;
                    RIGHT PICT "@"


  oIMPRXLS:oIXL_MAXCOL:cMsg    :="Columna Final"
  oIMPRXLS:oIXL_MAXCOL:cToolTip:="Columna Final"

  @ oIMPRXLS:oIXL_MAXCOL:nTop-08,oIMPRXLS:oIXL_MAXCOL:nLeft SAY "Columna"+CRLF+"Final" PIXEL;
                            SIZE NIL,7 FONT oFont COLOR nClrText,oDp:nGris

  oIMPRXLS:Activate({||oIMPRXLS:ViewDatBar()})

  STORE NIL TO oTable,oGet,oFont,oGetB,oFontG

RETURN oIMPRXLS

/*
// Barra de Botones
*/
FUNCTION ViewDatBar()
   LOCAL oCursor,oBar,oBtn
   LOCAL oDlg:=oIMPRXLS:oDlg

   DEFINE CURSOR oCursor HAND

   IF !oDp:lBtnText 
     DEFINE BUTTONBAR oBar SIZE 52-15,60-15 OF oDlg 3D CURSOR oCursor
   ELSE 
    DEFINE BUTTONBAR oBar SIZE oDp:nBtnWidth,oDp:nBarnHeight+6 OF oDlg 3D CURSOR oCursor 
   ENDIF 

   IF oIMPRXLS:nOption=2 


     DEFINE BUTTON oBtn;
            OF oBar;
            NOBORDER;
            FONT oFont;
            TOP PROMPT "Salir"; 
            FILENAME "BITMAPS\XSALIR.BMP";
            ACTION (oIMPRXLS:Close())

     oBtn:cToolTip:="Salir"

   ELSE

     DEFINE BUTTON oBtn;
            OF oBar;
            NOBORDER;
            FONT oFont;
            FILENAME "BITMAPS\DOWNLOAD.BMP";
            TOP PROMPT "Descarga"; 
            ACTION EJECUTAR("DPIMPRXIMPORTDOWN",oDpLbx)

     oBtn:cToolTip:="Descargar Definiciones desde AdaptaPro Server"

     DEFINE BUTTON oBtn;
            OF oBar;
            NOBORDER;
            FONT oFont;
            TOP PROMPT "Grabar"; 
            FILENAME "BITMAPS\XSAVE.BMP",NIL,"BITMAPS\XSAVEG.BMP";
            WHEN FILE(oIMPRXLS:IXL_FILE);
            ACTION (oIMPRXLS:Save())

     oBtn:cToolTip:="Grabar"

     oIMPRXLS:oBtnSave:=oBtn


     IF oIMPRXLS:nOption=3 

        DEFINE BUTTON oBtn;
               OF oBar;
               NOBORDER;
               FONT oFont;
               FILENAME "BITMAPS\CONFIGURA.BMP";
               ACTION EJECUTAR("DPIMPRXLSDEF",oIMPRXLS:IXL_CODIGO_)

        oBtn:cToolTip:="Configuración"

     ENDIF

     DEFINE BUTTON oBtn;
            OF oBar;
            FONT oFont;
            NOBORDER;
            TOP PROMPT "Cancelar"; 
            FILENAME "BITMAPS\XCANCEL.BMP";
            ACTION (oIMPRXLS:Cancel()) CANCEL

     oBtn:cToolTip:="Cancelar"

   ENDIF

   oBar:SetColor(CLR_BLACK,oDp:nGris)
   AEVAL(oBar:aControls,{|o,n|o:SetColor(CLR_BLACK,oDp:nGris)})

RETURN .T.


/*
// Carga de Datos, para Incluir
*/
FUNCTION LOAD()

  IF oIMPRXLS:nOption=1 // Incluir en caso de ser Incremental
     
     // AutoIncremental 
  ENDIF

RETURN .T.
/*
// Ejecuta Cancelar
*/
FUNCTION CANCEL()
RETURN .T.

/*
// Ejecución PreGrabar
*/
FUNCTION PRESAVE()
  LOCAL lResp:=.T.

  lResp:=oIMPRXLS:ValUnique(oIMPRXLS:IXL_CODIGO)

  oIMPRXLS:IXL_FECHA:=DPFECHA()
  oIMPRXLS:IXL_HORA :=DPHORA()
  oIMPRXLS:IXL_ALTER:=.T.

  IF Empty(oIMPRXLS:IXL_TABLA)
     oIMPRXLS:VALIXLPREDEF()
  ENDIF

  IF Empty(oIMPRXLS:IXL_TABLA)
     oIMPRXLS:oIXL_TABLA:MsgErr("Necesario Indicar la Tabla")
     RETURN .F.
  ENDIF

  IF !ISSQLFIND("DPTABLAS","TAB_NOMBRE"+GetWhere("=",oIMPRXLS:IXL_TABLA))
     oIMPRXLS:oIXL_TABLA:MsgErr("Necesario Indicar la Tabla")
     RETURN .F.
  ENDIF

  IF !lResp
     MsgAlert("Registro "+CTOO(oIMPRXLS:IXL_CODIGO),"Ya Existe")
     RETURN .F.
  ENDIF

// 30/05/2023
//  IF oIMPRXLS:nOption=1
//    DPWRITE("FORMS\"+ALLTRIM(oIMPRXLS:IXL_CODIGO)+".IXL","")
//  ENDIF
//  EJECUTAR("DPIMPRXLSDEF",oIMPRXLS:IXL_CODIGO)

RETURN .T.

/*
// Ejecución despues de Grabar
*/
FUNCTION POSTSAVE()

  IF oIMPRXLS:nOption=1
    DPWRITE("FORMS\"+ALLTRIM(oIMPRXLS:IXL_CODIGO)+".IXL","")
  ENDIF

  EJECUTAR("DPIMPRXLSDEF",oIMPRXLS:IXL_CODIGO)
 
RETURN .T.

FUNCTION VALIXLFILE()

   IF !FILE(oIMPRXLS:IXL_FILE) 
      oIMPRXLS:oIXL_FILE:MsgErr("Archivo "+ALLTRIM(oIMPRXLS:IXL_FILE)+" No Existe","Introduzca Nombre del Archivo")
      RETURN .F.
   ENDIF

   oIMPRXLS:oBtnSave:ForWhen(.T.)

RETURN .T.

FUNCTION VALTABLA()

  IF !ISSQLFIND("DPTABLAS","TAB_NOMBRE"+GetWhere("=",oIMPRXLS:IXL_TABLA))
     oIMPRXLS:oIXL_TABLA:MsgErr("Tabla "+ALLTRIM(oIMPRXLS:IXL_TABLA)+" No Existe","Introduzca Nombre de la Tabla")
     RETURN .F.
  ENDIF

RETURN .T.

FUNCTION VALIXLPREDEF()
  LOCAL nLen  :=LEN(oIMPRXLS:IXL_TABLA)
  LOCAL cTable:=EJECUTAR("DPIMPXLSDEFTABLA",oIMPRXLS:IXL_PREDEF)

  cTable:=PADR(cTable,nLen)
  oIMPRXLS:oIXL_TABLA:VarPut(cTable,.T.)
  oIMPRXLS:oIXL_TABLA:Refresh(.T.)
  oIMPRXLS:oIXL_TABLA:ForWhen(.T.)

RETURN .T.
/*
<LISTA:IXL_CODIGO:Y:GET:N:N:Y:Código,IXL_DESCRI:N:GET:N:N:Y:Descripción,IXL_TABLA:N:COMBO:N:N:Y:Tabla,IXL_ACTIVO:N:CHECKBOX:N:N:Y:Activo
,IXL_FILE:N:BMPGETF:N:N:Y:Nombre del Archivo,IXL_LININI:N:GET:N:N:Y:Línea de Inicio>
*/
