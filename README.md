# orioncop

Modulo principal de operación de la plataforma Panorama Legacy en VB.NET.

## Alcance
- Operacion principal de negocio.
- Integraciones con componentes comunes y reporteria.

## Estructura base
- `src/`: carpeta reservada para estabilizacion gradual de codigo.
- `docs/`: documentacion tecnica y operativa.
- `OriIntCont/`: Proyecto en el cual se definen las interfaces contables. Esto permite generar un archivo espesífico con información clave para enviar a diferentes programas contables, también contiene la interfaz para facturación electrónica con mis facturas.
- `OrionCopIU/`: Acá se define toda la Interfaz gráfica de Orión Plus y la funcionalidad de cada uno de sus elementos.
- `RecImagenes/`: Acá viven los archivos binarios de las imágenes que se referencian en la interfaz.
- `OrionCopL/`: Proyecto donde se enceuntra la mayoría de la lógica de Orión
- `RepOriCop/`: Acá se definen y construyen todos los reportes, Se usa Crystal Report para generarlo.

## Notas del desarrollador                                                                                                                                                
* El código es legacy.
* El diseño de interfaz ya se encuentra establecido.
* Clase heredable para interfaces de usuario.`ClsFormInterface.vd` es la clase padre de casi todas las ventanas.
* Todas las ventanas que heredan de la clase `ClsFormInterface.vd` deben implementar la interfáz `IWinPanoramaIU` que se enceuntra en `acOriWin.vb`.
* El diccionario `dicPanRecursos.vb` contiene todos los recursos gráficos, prefabricados que utilizan las ventanas.
* Todas las clases heredan de `ClsCBObjetoPan`, esto estandariza las acciones más comunes como leer o escribir de la base de datos, creación de logs, entre otros.
* Las propiedades que de la clase, que se representa en base de datos, se encuentran en la #Region "Clases de Propiedad" de cada clase
* Cada propiedad es una clase que hereda de `ClsCBPropiedad.vb` 
* Tanto en código cómo en base de datos, el cliente hereda de forma indirecta de terceros.
* Las convenciones no se siguen al pié de la letra pero es lo ideal.
* Los prefijos de las convenciones en algunas aprtes están en minusculas y en otras en mayusculas.
* Para las variables se debe usar camellCase, con incial en minuscula.
* Para las funciones, métodos y clases se usa CamellCase, con inicial mayuscula.
* Para las cosntantes se usa MAYUSCULASOSTENIDA.


## Interfaces

La interfaz de usuario se encuentra establecida y se debe usar siempre el mismo estilo de diseño, adaptando forma y tamaño según el caso.

Las Interfaces se encuentran en la siguiente estructura de código. 
```
|- orioncop ->   *carpeta*
| |- OrionCopIU ->   *proyecto*
|   |- MWOrionCop.xaml
|   |- acOrionCopIU.vb
|   |- mOrionCopIU.vb
|   |- FrmAyuda.vb -> WindowsForms
|   |- winNombreDescriptivo
|- comunes ->   *carpeta*
| |- OriWin
|   |- ClsFormInterface.vb
|   |- acOriWin.vb
|   |- mdefOriWin.vb
|   |- dicPanRecursos.xaml
|   |- otros archivos
| |- WinCom
|   |- winCopiaSeg.xaml
|   |- winTerceros.xaml
|   |- otras ventanas compartidas entre proyectos.

```
## Proyectos de Interfaz de Usuario
* **OrionCopIU** -> Se enceuntran todas las ventanas de Orión
* **OriWin** -> Proyecto compartido entre ***Admin Orion** (adminorion) y **Orion Plus** (orioncop)* 

## Proyectos en carpetas comunes
* **OriWin** -> [Interfaz de usuario](#proyectos_de_interfaz_de_usuario)
* **PanDat** -> Base de datos compartida, acá se encuentra la conexción, estructura, actualización de la estructura, de la base de datos `pan...`.
* **PanL** -> Lógica de programación compartida.
* **WinCom** -> Ventanas comunes a los proyectos.

## Convenciones
### Prefijos

#### Variables
* Las variables se componen de al menos dos prefijos. Scop, tipo de dato y nombre descriptivo
**Ejemplo**
```
lstrNombreCompleto
mbolEsProgramable
MCSTRNOMBRECAMPOBD
GCSTRCLAVEDESEGURIDAD
```

M = Módulo
l = Local
G = Global
H = Es heredada y es heredable

***Tipo de la variable***
str = string
int = entero
dec = decimal
txt = texto
bln = boolean

rec = recurso o asset
img = imágen
btt = botón
sep = separador
txt = Texto o texbox
mnu = Menú
*Por nombrar algunos*

#### Función o método
* Los métodos y unciones se componen de uno o dos prefijos

F = Función
S = Método
* Define si es función o método, luego si es de un tipo de dato espesifico lo agrega, y finalmente un nombre descriptivo.

#### Constantes
- Igual que las variables pero en mayuscula sostenida.
MCSTR... = Módulo,Constante, Tipo string

#### Archivos
* Los archivos se conforman de uno o dos prefijos y de un nombre descriptivo.

I = Interfaz
ac = archivo de código
dic = Diccionario
cls o cls = Clase
m =  Archivo de módulo
win = Ventana de WPF
frm = Ventana de Forms
clsCB / CB = Clase Base, Clases que son heredables.
rep = Reportes en CrystalReport. Cada reporte usa `FrmReportes.vb`


#### Proyectos
Los proyectos se componene de un prefijo y nombre descriptivo.

Ori / Orion = Son específicos de Orión
Pan = Son comunes y se compartían con varios proyectos. Este solo se usa para compartirlo OriónCop y AdminOrión

### Sufijos

#### Variables
L = Lógica de programación, Capa de lógica del programa.
IU = Interfaz de usuario
Com = Comunes a las soluciones
Con = Contabilidad

#### Archivos - Clases
#### Tipo de dato que está contenido en la calse.
Shr = Short
Dbl = Double
Str = String
Dec = Decimal
*Por nombrar algunos*