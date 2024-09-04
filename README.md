# VBA_Reporte_Venta_Tablero_Compras
Repository to process with VBA reports processes

# Reporte: Reporte de ventas e inventarios

Lista inputs: 
	1. Reporte_Inventarios
	2. Transitos
	3. Venta
	4. Reporte_Existencias_Almacen_Matriz
	5. PP y PS
	6. Directorio_Sucursales
	7. Catalogo_Lineas_Familias

1. Abrir archivo del dia a procesar (29-08-24.xlsx)
	**Este proceso se dejará con un boton al usuario para poder seleccionar la ruta del archivo a procesar

2. Procesar input Reporte_Inventarios	
	2.1 Abrir Excel Reporte_De_Inventarios
	2.2 Iterar por cada hoja del archivo y juntar en una sola matriz todas las filas con datos en el rango de columnas A:AC
	2.3 En Excel 29-08-24.xlsx ir a hoja Inv, limpiar todos los registros que existan y pegar la extraccion de Reporte_De_Inventarios comenzando en la columna A y fila 1
	
3. Procesar input Transitos
	3.1 Abrir Excel Inventario_Transito_Traspasos
	3.2 Extraer en una matriz las columnas Cliente, Codigo, Descripcion y CantidadPendiente
	3.3 Crear columna Key, concatenando Cliente y Codigo
	3.4 Crear matriz sumarizada, pivote es la columna Key, se requiere obtener la suma de CantidadPendiente de cada Key y mostrar la columna Descripcion
	3.5 En Excel 29-08-24.xlsx ir a hoja Transitos, limpiar todos los registros que existan y pegar la matriz sumarizada de Inventario_Transito_Traspasos comenzando en la columna A y fila 1
	
4. Proceasr input Reporte_Existencias_Almacen_Matriz
	4.1 Abrir Excel Existencias_Almacen_Matriz(cc184)
	4.2 Extraer en una matriz todas las filas con datos en el rango de columnas A:AP
	4.3 En Excel 29-08-24.xlsx ir a hoja 184, limpiar todos los registros que existen y pegar la extracion de Reporte_Existencias_Almacen_Matriz comenzando en la columna A y fila 1
	
5. Procesar input PP y PS
	5.1 Abrir Excel PP_PS_Y_BO_AL_02_DE_SEPTIEMBRE
	5.2 Extraer en una matriz todas las filas con datos en el rango de las columnas A:P
	5.3 Crear al extremo izquierdo de la matriz columna Key, concatenando columnas AlmacenCliente y Codigo
	5.4 Crear al extremo izquierdo de la matriz columna Pendiente, que será el resultado de "Cant Solicitada" - "CantSurtida"
	5.5 Se agrega columna por separado para el CC 186 ¿? **necesito más detalle
	5.3 En Excel 29-08-24.xlsx ir a hoja PP, limpiar todos los registros que existan y pegar la extraccion de PP_PS_Y_BO_AL_02_DE_SEPTIEMBRE comenzando en la columna A y fila 1
	
6. Procesar input Directorio_Sucursales
	6.1 Abrir Excel DIRECTORIO_CADENA_2020.xlsx
	6.2 Ir a hoja BASE DE SUCURSALES 2
	6.3 Extraer en una matriz todas las filas con datos comenzando en la fila 3 y en el rango de columnas A:Q
	6.4 ¿En que hoja del Excel 29-08-24.xlsx se pega la informacion o se utiliza estos datos? **necesito más detalle
	
7. Procesar input Catalogo_Lineas_Familias
	7.1 ¿En que archivo se encuentra este catalogo? ¿Que proceso tendria? (copiar, crear columnas, calcular campos, pegar en alguna hoja, etc)
