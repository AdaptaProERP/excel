// Programa   : XLSTODBF2
// Fecha/Hora : 17/02/2024 05:51:26
// Propósito  :
// Creado Por : Juan Navas
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cFileXls,cFileDbf,oMeter,oSay,lAuto,nLinIni,nCantRec,nHead,nColGet,lStruct,cMaxCol,aSelect,lOut,cCodigo,lBrowse,nMaxCol)
   LOCAL aData  :={},aLine:={}
   LOCAL aFields:={},I,nLin,nCol,cValue,uData,cField,nContar:=0,nLen,nDec,cType
   LOCAL cTable,oTable,oExcel,lEmpty:=.F. 
   LOCAL nContar:=0,aFieldDbf:={}

   DEFAULT nMaxCol :=24,;
           nCantRec:=0 ,;
           aSelect :={},;
           nLinIni :=1

//  DEFAULT cFileXls:="C:\ahiskel\propietarios.xls"

   DEFAULT cFileXls:="C:\LIBRODECOMPRAS\librodecompras.xlsx"


   IF cCodigo<>NIL

     // DEFAULT cCodigo:=SQLGET("DPIMPRXLS","IXL_CODIGO","IXL_TABLA"+GetWhere("=","DPCLIENTESREC"))

     oTable  :=OpenTable("SELECT IXL_FILE,IXL_TABLA,IXL_LININI FROM DPIMPRXLS WHERE IXL_CODIGO"+GetWhere("=",cCodigo),.T.)

// ? oTable:RecCount(),oDp:cSql

     cFileXls:=ALLTRIM(oTable:IXL_FILE  )
     cTable  :=ALLTRIM(oTable:IXL_TABLA )
     // cMemo   :=ALLTRIM(oTable:IXL_MEMO  )
     nLinIni :=MAX(oTable:IXL_LININI,1)
     oTable:End(.T.)

     aSelect:=ACLONE(EJECUTAR("IXLLOAD",cTable,NIL,cCodigo))

   ENDIF

   DEFAULT cFileDbf:=STRTRAN( cFileXls,".xls",".dbf")

   IF !FILE(ALLTRIM(cFileXls))
      MensajeErr("Archivo "+cFileXls+" no Existe")
      RETURN NIL
   ENDIF

   SET DECI TO 6

   IF Empty(aSelect)
      aSelect:=ARRAY(nMaxCol)
      AEVAL(aSelect,{|a,n|aSelect[n]:={CHR(64+n),CHR(64+n),CHR(64+n)}})
   ENDIF

   aFields:={}
   FOR I=1 TO LEN(aSelect)
     AADD(aFields,{aSelect[I,1],aSelect[I,3]+LSTR(nLinIni) ,"",0,0,"",aSelect[I,3]})
   NEXT I

   oExcel := TExcelScript():New()
   oExcel:Open( cFileXls )

   // Lectura de Campos para determina la estructura
   aLine:=ARRAY(LEN(aFields))

   FOR I=1 TO LEN(aFields)

     cValue:=SPACE(1024)
     nLin  :=VAL(SUBS(aFields[I,2],2,4))
     nCol  :=(ASC(LEFT(aFields[I,2],1))-64)
     uData :=oExcel:Get( nLin , nCol ,@cValue )

     aFields[I,4]:=nLin
     aFields[I,5]:=nCol
     aFields[I,3]:=uData
     aFields[I,6]:=VALTYPE(uData)
     aLine[I]:=uData

   NEXT I

   AADD(aData,ACLONE(aLine))

   lEmpty:=.F. 

   WHILE !lEmpty

     nLin++
     lEmpty:=.T.     

     FOR I=1 TO LEN(aFields)
         cValue  :=SPACE(1024)
         nCol    :=aFields[I,5]
         uData   :=oExcel:Get( nLin , nCol ,@cValue )
         aLine[I]:=uData
         lEmpty  :=IF(!Empty(uData),.F.,lEmpty) // Si algun Columna no está vacia Continua
     NEXT I

     IF !lEmpty
       AADD(aData,ACLONE(aLine))
     ENDIF

     nContar++

     IF nCantRec>0 .AND. nCantRec>=nContar
        EXIT
     ENDIF

   ENDDO

   oExcel:End(.F.)

   // Creamos Archivo DBF para migrar los datos

   FOR I=1 TO LEN(aFields)
       cField:=aFields[I,7]
       cType :=aFields[I,6]
       nLen  :=250
       nLen  :=IF(cType="D",08,nLen)
       nLen  :=IF(cType="N",19,nLen)
       nDec  :=IF(cType="N",06,0)
       AADD(aFieldDbf,{cField,cType,nLen,nDec})
   NEXT I  

   FERASE(cFileDbf)

   DBCREATE(cFileDbf,aFieldDbf)

   USE (cFileDbf) EXCLU

   FOR nContar=1 TO LEN(aData)
      aLine:=ACLONE(aData[nContar])
      APPEND BLANK
      AEVAL(aLine,{|a,n| IF(!Empty(a),FIELDPUT(n,a),NIL)})
   NEXT I

   USE

   DEFAULT lBrowse:=.F.

   IF lBrowse
     EJECUTAR("DBFVIEWARRAY",cFileDbf)
   ENDIF

RETURN .F.
// EOF
