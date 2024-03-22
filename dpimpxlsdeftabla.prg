// Programa   : DPIMPXLSDEFTABLA 
// Fecha/Hora : 01/11/2022 18:27:23
// Propósito  :
// Creado Por :
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cPreDef)
  LOCAL aTablas:={},nAt,cTable:=""

  DEFAULT cPreDef:="Plan de Cuentas"

  AADD(aTablas,{"Productos"         ,"DPINV"      })
  AADD(aTablas,{"Proveedores"       ,"DPPROVEEDOR"})
  AADD(aTablas,{"Clientes"          ,"DPCLIENTES" })
  AADD(aTablas,{"Cuentas por Pagar" ,"DPDOCPRO"   })
  AADD(aTablas,{"Cuentas por Cobrar","DPDOCCLI"   })
  AADD(aTablas,{"Plan de Cuentas"   ,"DPCTA"      })
  AADD(aTablas,{"Libro de Compras"  ,"DPLIBCOMPRASDET"   })
  AADD(aTablas,{"Libro de Ventas"   ,"DPLIBVENTADET"   })
  AADD(aTablas,{"Activos"           ,"DPACTIVOS"  })
  AADD(aTablas,{"Asientos"          ,"DPASIENTOS"  })
  AADD(aTablas,{"Propietarios para Condominio","DPCLIENTESREC"  })

  nAt   :=ASCAN(aTablas,{|a,n| a[1]=ALLTRIM(cPreDef)})
  cTable:=IF(nAt=0,cTable,aTablas[nAt,2])

  oDp:aTablasXls:=ACLONE(aTablas)

RETURN cTable
// EOF
