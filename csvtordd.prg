// Programa   : CSVTORDD
// Fecha/Hora : 15/04/2024 03:05:48
// Propósito  : Lectura Archivo CSV hacia RDD
// Creado Por : Juan Navas
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cFileCsv,lBrowse)
  LOCAL I,aFields:={},aLine,nContar,nAt,oTable

  DEFAULT cFileCsv:="C:\LIBRODECOMPRAS\librodecompras.csv",;
          lBrowse :=.T.

  oDp:cMemoI:=MemoRead(cFileCsv)
  oDp:aDataI:=STRTRAN(oDp:cMemoI,CRLF,CHR(10))
  oDp:aDataI:=_VECTOR(oDp:cMemoI,CHR(10))

  FOR I=1 TO LEN(oDp:aDataI)
    oDp:aDataI[I]:=_VECTOR(oDp:aDataI[I],";")
  NEXT I

  oTable:=TRDDARRAY():New()

  oTable:SetArray(ACLONE(oDp:aDataI))

  AEVAL(oTable:aFields,{|a,I,cField| cField:="COL_"+CHR(64+I),oTable:aFields[I]:=cField,__objAddData( oTable,cField)})

  oTable:Gotop()

  IF lBrowse
     oTable:Browse()
  ENDIF

RETURN oTable
// EOF
