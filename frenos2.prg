// Programa   : FRENOS2
// Fecha/Hora : 24/10/2019 06:29:28
// Propósito  : https://github.com/AdaptaProERP/excel/blob/main/productos2.csv
// Creado Por :
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

FUNCTION MAIN()
  LOCAL cFile:="C:\frenos\productos2.csv"
  LOCAL oNew:=OpenTable("SELECT * FROM DPINV",.F.),cWhere
  LOCAL cCodInv,cCodGru,cCodMar,cDescri:="Indefinido"
  LOCAL cUndMed,nCxUnd,cCodMon:="DBC",nUndMed:=1,cUndGru:="CAJA"
  LOCAL nContar:=0,nIni:=0
  LOCAL nPrecioH,nPrecioI,nCosto,cDescri,nGrupo:=0

  LOCAL oXls   :=NIL

  oXls:=EJECUTAR("CSVTORDD",cFile,.F.)

  ADEPURA(oXls:aData,{|a,n| "Refer"$a[1]})
  ADEPURA(oXls:aData,{|a,n| Empty(a[2])})
//  ADEPURA(oXls:aData,{|a,n| Empty(a[3])})

 // SQLDELETE("DPINV","INV_USO"+GetWhere("=","LISTA2"))
  nIni:=SQLGETMAX("DPINV","INV_VOLUME","INV_USO"+GetWhere("=","LISTA2"))

  oNew:SetForeignkeyOff()

  cUndMed:="UND"
  cUndGru:="CAJA"

  nContar:=0
  nIni   :=IF(nIni=0,1,nIni)
  cCodGru:=SQLGET("DPGRU","GRU_CODIGO","GRU_DESCRI"+GetWhere("=",cDescri))

  IF Empty(cCodGru)

     cCodGru:=SQLINCREMENTAL("DPGRU","GRU_CODIGO",NIL,NIL,NIL,.T.,4	)

     EJECUTAR("CREATERECORD","DPGRU",{"GRU_CODIGO","GRU_DESCRI","GRU_ACTIVO"},;
                                     {cCodGru     ,cDescri    ,.T.         },NIL,.T.,;
                                     "GRU_CODIGO"+GetWhere("=",cCodGru))
  ENDIF

  cCodMar:=STRZERO(0,4)

  IF !ISSQLFIND("DPMARCAS","MAR_CODIGO"+GetWhere("=",cCodMar))

     EJECUTAR("CREATERECORD","DPMARCAS",{"MAR_CODIGO","MAR_DESCRI","MAR_ACTIVO"},;
                                        {cCodMar     ,"Indefinido" ,.T.        },NIL,.T.,;
                                        "MAR_CODIGO"+GetWhere("=",cCodMar))

  ENDIF

  cDescri:="Akron"
  cCodMar:=SQLGET("DPMARCAS","MAR_DESCRI"+GetWhere("=",cDescri))

  IF Empty(cCodMar)

     cCodMar:=SQLINCREMENTAL("DPMARCAS","MAR_CODIGO",NIL,NIL,NIL,.T.,4)

     EJECUTAR("CREATERECORD","DPMARCAS",{"MAR_CODIGO","MAR_DESCRI","MAR_ACTIVO"},;
                                        {cCodMar     ,cDescri     ,.T.        },NIL,.T.,;
                                        "MAR_CODIGO"+GetWhere("=",cCodMar))
  ENDIF


  oXls:Gotop()
  WHILE !oXls:Eof() .AND. nIni=0

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


  oXls:Gotop()
  IF nIni>1
    oXls:Goto(nIni)
  ENDIF

  WHILE !oXls:Eof()

    cCodInv:=oXls:COL_A
    cCodInv:=STRTRAN(cCodInv,["],"")

    IF Empty(cCodInv) .OR. Empty(oXls:COL_B)
       oXls:DbSkip()
       LOOP 
    ENDIF

    nCosto  :=STRTRAN(oXls:COL_G,",",".")
    nPrecioH:=STRTRAN(oXls:COL_H,",",".")
    nPrecioI:=STRTRAN(oXls:COL_I,",",".")

    nCosto  :=CTOO(nCosto  ,"N")
    nPrecioH:=CTOO(nPrecioH,"N")
    nPrecioI:=CTOO(nPrecioI,"N")

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

    cCodGru:=SQLGET("DPGRU","GRU_CODIGO","GRU_DESCRI"+GetWhere("=",oXls:COL_J))

    IF Empty(cCodGru)
       cCodGru:=STRZERO(0,4)
    ENDIF

    oNew:AppendBlank()
    oNew:Replace("INV_CODIGO",cCodInv)
    oNew:Replace("INV_DESCRI",oXls:COL_B)
    oNew:Replace("INV_GRUPO" ,cCodGru)
    oNew:Replace("INV_IVA"   ,"GN")
    oNew:Replace("INV_CODMAR",cCodMar)
    oNew:Replace("INV_OBS1"  ,oXls:COL_C)
    oNew:Replace("INV_OBS2"  ,oXls:COL_D)
    oNew:Replace("INV_OBS3"  ,oXls:COL_E)

    oNew:Replace("INV_UTILIZ","Venta")
    oNew:Replace("INV_ESTADO","A"    )
    oNew:Replace("INV_APLICA","T"    )
    oNew:Replace("INV_VOLUME",oXls:RecNo()) 
    oNew:Replace("INV_USO"   ,"LISTA2")
    oNew:Replace("INV_COSMER",nCosto  )
    oNew:SetDefault()
    oNew:Commit()

    EJECUTAR("DPINVCREAUND",oNew:INV_CODIGO,cUndGru,1)
    nUndMed:=1
    EJECUTAR("DPPRECIOSCREAR",cCodInv,"A",cUndGru,"DBC",nPrecioH*nUndMed)
    EJECUTAR("DPPRECIOSCREAR",cCodInv,"B",cUndGru,"DBC",nPrecioI*nUndMed)

    IF !Empty(oXls:COL_K)

       cWhere:="EQUI_CODIG"+GetWhere("=",cCodInv      )+" AND "+;
               "EQUI_BARRA"+GetWhere("=",oXls:COL_K)

       EJECUTAR("CREATERECORD","DPEQUIV",{"EQUI_CODIG","EQUI_BARRA","EQUI_DESCR"},;
                                        {cCodInv      ,oXls:COL_K  ,"LYO"       },NIL,.T.,;
                                         cWhere)

    ENDIF
   
    SysRefresh(.T.)

    oXls:DbSkip()

  ENDDO

  oXls:End()
  oNew:End()

 ? "CONCLUIDO LISTA 2"

RETURN  NIL
/*
C001=EQUI_BARRA          ,'C',020,0,'','Equivalente',0,''
 C002=EQUI_CODIG          ,'C',022,0,'','Producto',0,''
 C003=EQUI_DESCR          ,'C',040,0,'','Descripción',0,''
 C004=EQUI_LPT            ,'C',004,0,'','Puerto de Impresión',0,''
 C005=EQUI_MED           
*/
//


