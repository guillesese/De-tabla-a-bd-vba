# De tabla de Excel a tabla de Oracle
Módulo de VBA para actualizar una tabla de Oracle mediante ADODB con el contenido de una tabla de Excel. 

## INTRODUCCIÓN ##
Es necesario un procedimiento que vaya construyendo un *UPDATE* en una tabla de Oracle en función del contenido de la tabla. Para el correcto funcionamiento del módulo será necesaria la creación de una hoja de datos llamada *CORRESPONDENCIA*. 

## FUNCIONAMIENTO ##
1. Debemos establecer una hoja de *CORRESPONDENCIA*, el contenido de esta hoja sera el siguiente: 
   - Hoja: Nombre real de la hoja según Excel
   - Nombre: Nombre visual de la hoja. 
   - Tabla BD: Nombre de la tabla en BD
   - Tabla datos: Nombre que tiene la tabla de datos en Excel
   - Máscara: Codificación para hacer el *SELECT* y el *UPDATE*
   - MáscaraTipos: Codificación de tipos para evitar que Excel pueda hacer su conversión. 
   
  ![imagen](https://user-images.githubusercontent.com/16133041/140708858-3aecf935-a41e-486b-95d7-9feac5bb7c4b.png)

2. Crearemos un botón desde la hoja que queremos llamar y le asignaremos la llamada a la función *trasladarTablaaBD*

## MÁSCARA ##
Debemos indicar a través de la máscara que hacemos con cada uno de los campos de la tabla siguiendo las siguientes directrices:
![imagen](https://user-images.githubusercontent.com/16133041/140709746-675bdb69-1d2f-4101-977f-aa703db9deb9.png)

## MÁSCARA DE TIPOS ##
Tras probar varios métodos y ninguno fructífero, he optado por añadir también una máscara de tipos: 
![imagen](https://user-images.githubusercontent.com/16133041/140709977-45f4f901-88de-4110-8948-66d626f92fe5.png)

## CORRELACIÓN ##
Como es posible que finalmente la Excel la utilice personal no informático y que para mostrar los datos en la tabla de Excel se utilicen nombre distintos a los que se utilizan en BD, he creado una función que asocia el nombre de la columna con el nombre en BD. 
![imagen](https://user-images.githubusercontent.com/16133041/140710302-c931f1dc-0eaa-45d4-9726-075ae2dd4624.png)
