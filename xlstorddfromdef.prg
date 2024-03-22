// Programa   : XLSTORDDFROMDEF
// Fecha/Hora : 18/03/2024 23:47:52
// Propósito  :
// Creado Por :
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo)
  LOCAL cFileXls,cFileDbf,oMeter,oSay,lAuto,nLinIni,nCantRec,nHead,nColGet,lStruct,cMaxCol,aSelect,lOut,lBrowse,nMaxCol
  LOCAL oData

  oData:=EJECUTAR("XLSTORDD",cFileXls,cFileDbf,oMeter,oSay,lAuto,nLinIni,nCantRec,nHead,nColGet,lStruct,cMaxCol,aSelect,lOut,cCodigo,lBrowse,nMaxCol)

RETURN oData
