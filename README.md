\# Claude Assistant para Excel



Este proyecto es un \*\*complemento personalizado de Excel\*\* que integra el modelo Claude en un panel lateral (Task Pane).  

Permite enviar consultas directamente desde Excel y recibir respuestas que se muestran en el panel o se insertan en la celda activa.



---



\## Archivos principales

\- `manifest.xml` → Manifiesto del complemento, describe cómo se integra en Excel.

\- `panel.html` → Interfaz web con cuadro de texto y botón para enviar prompts a Claude.

\- `icon.png` → Ícono opcional para el complemento (puedes agregarlo más adelante).

\- `readme.md` → Documentación del proyecto.



---



\## Instalación en Excel



\### Opción 1: Usar versión pública desde GitHub

1\. Copia la URL \*\*Raw\*\* del archivo `manifest.xml` en este repositorio:  



https://raw.githubusercontent.com/gfernandez-engineer/excel-claude-addin/main/manifest.xml





2\. Abre Excel y ve a:  

\- \*\*Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza → Catálogo de complementos confiables\*\*  

3\. Agrega la URL anterior como catálogo confiable.  

4\. Reinicia Excel.  

5\. En la pestaña \*\*Insertar → Complementos → Mis complementos\*\*, aparecerá \*\*Claude Assistant\*\*.  



\### Opción 2: Usar versión local con tu API key privada

1\. Mantén el repositorio público con el `panel.html` que contiene el placeholder:

```js

const apiKey = "TU\_API\_KEY";



2\. En tu PC, guarda una copia privada de panel.html con tu clave real:

const apiKey = "sk-ant-xxxxxxxxxxxxxxxx";





3\. Ajusta tu manifest.xml local para que apunte a tu copia privada:

<SourceLocation DefaultValue="file:///C:/Users/Gian/excel-claude-addin/panel.html"/>



4\. Carga ese manifest.xml en Excel desde el catálogo confiable.



5\. De esta forma, tu Excel consume solo tu versión local con la clave, y tu repo público sigue siendo seguro.





