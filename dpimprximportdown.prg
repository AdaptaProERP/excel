// Programa   : DPIMPRXIMPORTDOW
// Fecha/Hora : 22/03/2024 22:48:09
// Propósito  : Descargar Definiciones desde Excel
// Creado Por : Juan Navas
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(oLbx)
  LOCAL cDir    :=Lower(cFileNoExt(cFileName(oDp:cBinExe)))
  LOCAL cUrl    :=oDp:cUrlDownLoad+cDir+"/dpimprxls.zip"
  LOCAL cSaveAs :=oDp:cBin+"temp\dpimprxls.zip"
  LOCAL cOut    :=cDir+"\datadbf"
  LOCAL cFile   :=cOut+"\dpimprxls.dbf"

  IF !MsgYesNo("Desea descargar Definiciones para Importar desde Excel")
      RETURN .T.
  ENDIF
  
  ferase(cSaveAs)

  MsgRunView("Descargando desde AdaptaPro Server","Descargando "+cUrl+" "+cSaveAs)
  SysRefresh(.t.)

  // oDp:oMsgRun:FRMTEXT("Descargando "+cUrl+" "+cSaveAs)

  URLDownLoad(cUrl, cSaveAs)

  ferase(cFile)

  HB_UNZIPFILE( cSaveAs , {|| nil }, .t., NIL, cOut , NIL )

  IF FILE(cFile)
    oDp:oMsgRun:FRMTEXT("Actualizando Definiciones ")
    EJECUTAR("DPIMPRXLS_FROMDBF")
  ENDIF

  DpMsgClose()

  IF ValType(oLbx)="O"
     oLbx:RELOAD()
  ENDIF

  IF !FILE(cFile)
     MsgMemo("Archivo "+cFile+" no pudo ser Descargadoo")
  ENDIF

RETURN .t.
// EOF
