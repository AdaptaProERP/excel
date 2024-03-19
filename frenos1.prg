// Programa   : FRENOS1
// Fecha/Hora : 24/10/2019 06:29:28
// Propósito  : Crear registro de productos, precios, marcas y grupos utilizando archivo .CSV convertido desde excel
// Archivo    : https://github.com/AdaptaProERP/excel/blob/main/productos1.csv
// Creado Por :
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

FUNCTION MAIN()
  LOCAL cFile:="C:\frenos\productos1.csv"
  LOCAL oNew:=OpenTable("SELECT * FROM DPINV",.F.),cWhere
  LOCAL cCodInv,cCodGru,cCodMar
  LOCAL cUndMed,nCxUnd,cCodMon:="DBC",nUndMed:=1,cUndGru:="CAJA"
  LOCAL nContar:=0,nIni:=SQLGETMAX("DPINV","INV_VOLUME","INV_USO"+GetWhere("=","LISTA1"))
  LOCAL nPrecioF,nPrecioG,nCosto,nGrupo:=0,cDescri:="",cMarca:="",aLine:={}
  LOCAL oXls   :=NIL

  oXls:=EJECUTAR("CSVTORDD",cFile,.F.)

//  ViewArray(oXls:aData)
 
  IF nIni=0 .OR. .T.

    SQLDELETE("DPMOVINV")
    SQLDELETE("DPEQUIV")
    SQLDELETE("DPPRECIOS")
    SQLDELETE("DPINVMED")
    SQLDELETE("DPINV")
    SQLDELETE("DPGRU")
    SQLDELETE("DPMARCAS")

    nIni:=0

  ENDIF

  aInv:={}
  oNew:SetForeignkeyOff()

  ADEPURA(oXls:aData,{|a,n| "Refe"$a[1]})
  ADEPURA(oXls:aData,{|a,n| "CATEG"$a[10]})


  cUndMed:="UND"
  cUndGru:="CAJA"
  cCodGru:=STRZERO(0,4) 
  cCodMar:="INDEF"

  EJECUTAR("CREATERECORD","DPGRU",{"GRU_CODIGO","GRU_DESCRI","GRU_ACTIVO"},;
                                  {cCodGru     ,"Indefinida",.T.         },NIL,.T.,;
                                  "GRU_CODIGO"+GetWhere("=",cCodGru))

  SQLUPDATE("DPINV","INV_GRUPO",cCodGru,"INV_USO"+GetWhere("=","LISTA1"))

  nContar:=0
  nIni   :=IF(nIni=0,1,nIni)

  // CREAR LAS MARCAS, la primera seccion de la lista no tiene Marcas
  IF ISSQLFIND("DPMARCAS","MAR_DESCRI"+GetWhere("=","Indefinido"))
     cCodMar:=STRZERO(0,4)
     EJECUTAR("CREATERECORD","DPMARCAS",{"MAR_CODIGO","MAR_DESCRI","MAR_ACTIVO"},;
                                        {cCodMar    ,"Indefinido" ,.T.         },NIL,.T.,;
                                        "MAR_DESCRI"+GetWhere("=",cCodMar))
  ENDIF

  ViewArray(oXls:aData)

  oXls:Gotop()
  WHILE !oXls:Eof() 

     IF !Empty(oXls:COL_A) .AND. Empty(oXls:COL_B) .AND. Empty(oXls:COL_L) .AND. ;
        !ISSQLFIND("DPMARCAS","MAR_DESCRI"+GetWhere("=",oXls:COL_A))
          nContar++
          cCodMar:=STRZERO(nContar,4)
          EJECUTAR("CREATERECORD","DPMARCAS",{"MAR_CODIGO","MAR_DESCRI","MAR_ACTIVO"},;
                                              {cCodMar    ,oXls:COL_A  ,.T. },NIL,.T.,;
                                              "MAR_DESCRI"+GetWhere("=",oXls:COL_A))
     ENDIF

     oXls:COL_L:=ALLTRIM(oXls:COL_L)
     cMarca    :=ALLTRIM(oXls:COL_L)
     cWhere    :="MAR_DESCRI"+GetWhere("=",cMarca)

     IF !Empty(cMarca) .AND. !ISSQLFIND("DPMARCAS",cWhere)

          nContar++
          cCodMar:=STRZERO(nContar,4)
          EJECUTAR("CREATERECORD","DPMARCAS",{"MAR_CODIGO","MAR_DESCRI","MAR_ACTIVO"},;
                                             {cCodMar     ,cMarca       ,.T.         },NIL,.T.,;
                                              cWhere)
     ENDIF

     cDescri:=ALLTRIM(oXls:COL_J)

     IF !Empty(cDescri) .AND. !ISSQLFIND("DPGRU","GRU_DESCRI"+GetWhere("=",cDescri))
          nGrupo++
          cCodGru:=STRZERO(nGrupo,4)
          EJECUTAR("CREATERECORD","DPGRU",{"GRU_CODIGO","GRU_DESCRI","GRU_ACTIVO"},;
                                          {cCodGru     ,cDescri     ,.T.         },NIL,.T.,;
                                          "GRU_DESCRI"+GetWhere("=",cDescri))
     ENDIF

     oXls:DbSkip()

  ENDDO

  oXls:Gotop()
  IF nIni>1
    oXls:Goto(nIni)
  ENDIF

  WHILE !oXls:Eof()

    IF !Empty(oXls:COL_A) .AND. Empty(oXls:COL_B) .AND. Empty(oXls:COL_L)
       cCodMar:=SQLGET("DPMARCAS","MAR_CODIGO","MAR_DESCRI"+GetWhere("=",oXls:COL_A))
    ENDIF

    IF !Empty(oXls:COL_L) 
       cCodMar:=SQLGET("DPMARCAS","MAR_CODIGO","MAR_DESCRI"+GetWhere("=",oXls:COL_L))
    ENDIF

    IF !Empty(oXls:COL_J) 
       cCodGru:=SQLGET("DPGRU","GRU_CODIGO","GRU_DESCRI"+GetWhere("=",oXls:COL_J))
    ENDIF

    cCodInv:=oXls:COL_A
    cCodInv:=STRTRAN(cCodInv,["],"")

    IF Empty(cCodInv) .OR. Empty(oXls:COL_B)
       oXls:DbSkip()
       LOOP 
    ENDIF

    nCosto  :=STRTRAN(oXls:COL_E,",",".")
    nPrecioF:=STRTRAN(oXls:COL_F,",",".")
    nPrecioG:=STRTRAN(oXls:COL_G,",",".")

    nCosto  :=CTOO(nCosto  ,"N")
    nPrecioF:=CTOO(nPrecioF,"N")
    nPrecioG:=CTOO(nPrecioG,"N")

    nUndMed:=1
    cUndGru:="UND"
    cUndMed:="UND"

    cWhere :="INV_CODIGO"+GetWhere("=",cCodInv)

    IF oXls:Recno()%10=0
       MsgRun(cCodInv+" - ("+LSTR(oXls:Recno())+"/"+LSTR(oXls:RecCount())+")")
    ENDIF

    IF ISSQLFIND("DPINV",cWhere) 
       oXls:DbSkip()
       LOOP 
    ENDIF

    oNew:AppendBlank()
    oNew:Replace("INV_CODIGO",cCodInv)
    oNew:Replace("INV_DESCRI",oXls:COL_B)
    oNew:Replace("INV_GRUPO" ,cCodGru)
    oNew:Replace("INV_IVA"   ,"GN")
    oNew:Replace("INV_CODMAR",cCodMar)
    oNew:Replace("INV_OBS1"  ,oXls:COL_C)
    oNew:Replace("INV_OBS2"  ,oXls:COL_D)
    oNew:Replace("INV_UTILIZ","Venta")
    oNew:Replace("INV_ESTADO","A"    )
    oNew:Replace("INV_APLICA","T"    )
    oNew:Replace("INV_VOLUME",oXls:RecNo()) 
    oNew:Replace("INV_USO"   ,"LISTA1")
    oNew:Replace("INV_COSMER"  ,nCosto  )
    oNew:SetDefault()
    oNew:Commit()

    EJECUTAR("DPINVCREAUND",oNew:INV_CODIGO,cUndGru,1)
    nUndMed:=1
    EJECUTAR("DPPRECIOSCREAR",cCodInv,"A",cUndGru,"DBC",nPrecioF*nUndMed)
    EJECUTAR("DPPRECIOSCREAR",cCodInv,"B",cUndGru,"DBC",nPrecioG*nUndMed)


    SysRefresh(.T.)

    oXls:DbSkip()

  ENDDO

  oXls:End()
  oNew:End()

? "CONCLUIDA LISTA1"

RETURN  NIL
//
