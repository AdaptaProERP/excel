// Programa   : DPIMPRXLS_FROMDBF
// Fecha/Hora : 22/03/2024 10:41:47
// Propósito  : Actualiza el Contenido de la tabla Vs DATADBF\DPIMPRXLS.DBF               
// Creado Por : Juan Navas
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(lUpdate)
   LOCAL cFileDbf:="DATADBF\DPIMPRXLS.DBF"
   LOCAL oDb     :=OpenOdbc(oDp:cDsnConfig)
   LOCAL oNew,oTable,aFields,cWhere

   // Actualiza cuando se actualiza automaticamente y no esta personalizada

   DEFAULT lUpdate:=.T.

   IF !FILE(cFileDbf)
       RETURN .F.	
   ENDIF

   IF !EJECUTAR("ISFIELDMYSQL",oDb,"DPIMPRXLS","IXL_ALTER")
     EJECUTAR("DPCAMPOSADD","DPIMPRXLS"  ,"IXL_ALTER" ,"L",1,0,"Personalizado") // ,NIL,.T.,.T.)
     oDb:EXECUTE("UPDATE DPIMPRXLS SET IXL_ALTER=0")
   ENDIF

   oNew:=OpenTable("SELECT * FROM DPIMPRXLS",.F.)

   CLOSE ALL
   SELECT A
   USE (cFileDbf) EXCLU

   WHILE !A->(EOF())

      cWhere:="IXL_CODIGO"+GetWhere("=",A->IXL_CODIGO)

      IF !ISSQLFIND("DPIMPRXLS",cWhere)
         // AGREGAR
         oNew:AppendBlank()
         AEVAL(DBSTRUCT(),{|a,n| oNew:Replace(a[1],FIELDGET(n))})
         oNew:Replace("IXL_ALTER",.F.)
         oNew:Commit("")

      ELSE

         IF ISSQLFIND("DPIMPRXLS",cWhere+" AND IXL_ALTER=0" )

           oTable:=OpenTable("SELECT * FROM DPIMPRXLS WHERE "+cWhere,.T.) 
           AEVAL(DBSTRUCT(),{|a,n| oNew:Replace(a[1],FIELDGET(n))})
           oTable:Replace("IXL_ALTER",.F.)
           oTable:Commit(cWhere)
           oTable:End(.T.)

        ENDIF

       ENDIF

       DBSKIP()

   ENDDO

   CLOSE ALL

   oNew:End(.T.)

RETURN .T.
