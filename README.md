# C2E (_Copy2Excel)
Autor (En parte): Walter Valle

Version: 1.0

Fecha: 13-SEP-2023


Un UDF más para copiar el contenido de un cursor o tabla a Excel para VFP.
Este es mi "Frankenstein" para exportar tablas o cursores de forma fácil, NO TODO el código es de mi autoría, en parte lo saqué de https://comunidadvfp.blogspot.com/ y adapté a mis necesidades.

# Configuración:
No requiere, pero se necesita tener una carpeta de trabajo para los archivos temporales, preferiblemente sin espacios en el nombre, actualmente es: C:\WV_TMPDir\ y La carpeta destino generalmente uso C:\WV_Excel\

Como es sabido, hay problemas con el comando COPY TO al usar espacios en blanco en el nombre, por lo cual no los uso.

OJO: Ocupas tener instalado MS Excel.

# Ejemplo de uso:
```
SET TALK OFF
*
DO C2E
*
TRY
  MD 'C:\WV_Excel'
CATCH
ENDTRY
IF !DIRECTORY('C:\WV_Excel')
  MESSAGEBOX('Error: No se creó la carpeta para guardar los archivos de Excel', 16, 'WV: _Copy2Excel')
  RETURN
ENDIF

USE IN SELECT('WV_Cursor')
CREATE CURSOR WV_Cursor (ID I AUTOINC, Nombre C(30), Precio N(12,2), STOCK I)

LOCAL x, nTotalRegs, cFile
nTotalRegs = 1000
FOR x = 1 TO nTotalRegs
  INSERT INTO WV_Cursor (Nombre, Precio, STOCK) VALUES ('Producto ' + TRANSFORM(x), RAND()*500, INT(RAND()*100))
ENDFOR
GO TOP IN SELECT('WV_Cursor')

cFile = 'C:\WV_Excel\MiExcel.XLSX'
IF _Copy2Excel('WV_Cursor', cFile)
  _OpenExcel(cFile)
ENDIF
cFile = 'C:\WV_Excel\MiExcelConTotales.XLSX'
IF _Copy2Excel('WV_Cursor', cFile, 'Precio,STOCK')
  _OpenExcel(cFile)
ENDIF

USE IN SELECT('WV_Cursor')
```
