# Markdownificator

[English](README.md) · **Español**

Markdownificator v1.2 exporta a Markdown el documento abierto en Word para Windows, incluidos los cambios pendientes de guardar. La conversión se realiza en tu equipo y genera un archivo `.md` compatible con GitHub Flavored Markdown (GFM).

## Antes de empezar

Necesitas **Word de escritorio para Windows** y **Pandoc**, el programa que convierte el documento a Markdown. Markdownificator v1.2 se probó con Word para Microsoft 365, Windows 11 y Pandoc 3.14. La macro muestra sus mensajes en español cuando los menús de Word están en español y en inglés en los demás casos.

## 1. Instala Pandoc

Si ya tienes Pandoc, abre **Terminal** o **Símbolo del sistema** desde el menú Inicio y ejecuta `pandoc --version`. Si aparece `pandoc` seguido de un número de versión, pasa al paso 2.

Para instalarlo, cierra Word, abre **Terminal** o **Símbolo del sistema**, pega este comando y pulsa **Entrar**. Sigue las indicaciones que aparezcan.

```powershell
winget install --source winget --exact --id JohnMacFarlane.Pandoc
```

Si Windows no reconoce `winget` o prefieres usar el asistente gráfico, descarga el **instalador de Windows (`.msi`)** desde la página oficial de Pandoc y ejecútalo. Con cualquiera de las dos opciones, abre una terminal nueva al terminar y ejecuta `pandoc --version` para comprobar que Windows encuentra el programa. [Instalación oficial de Pandoc](https://pandoc.org/installing.html#windows)

Si PowerShell sigue sin reconocer `pandoc` en la ventana que usaste para instalarlo con WinGet, cierra esa ventana, abre otra y vuelve a ejecutar `pandoc --version` antes de reinstalar nada. La ventana anterior puede seguir usando la configuración que tenía antes de instalar Pandoc.

## 2. Descarga e importa la macro

1. Abre [Markdownificator_v1.2.bas](Markdownificator_v1.2.bas) en GitHub. Pulsa el icono de descarga de la esquina superior derecha del archivo (*Download raw file*) y guarda el `.bas` en **Descargas**. [Descargar un archivo desde GitHub](https://docs.github.com/en/repositories/working-with-files/using-files/viewing-and-understanding-files#viewing-or-copying-the-raw-file-content)
2. Abre Word y crea un documento en blanco. Pulsa **Alt + F11** para abrir el editor de Visual Basic. También puedes entrar desde **Programador → Visual Basic**. Si la pestaña está oculta, ve a **Archivo → Opciones → Personalizar cinta de opciones**, marca **Programador** en **Pestañas principales** y pulsa **Aceptar**. [Mostrar la pestaña Programador](https://support.microsoft.com/es-es/word/show-the-developer-tab-in-word)
3. En el panel izquierdo del editor, haz clic en el proyecto **Normal**. Si el panel está oculto, pulsa **Ctrl + R**. **Normal** es la plantilla personal que Word ya usa. Al guardar ahí la macro podrás ejecutarla desde otros documentos. [Explorador de proyectos](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/view-menu) · [Macros disponibles en todos los documentos](https://support.microsoft.com/es-es/word/create-or-run-a-macro)
4. En el editor, elige **Archivo → Importar archivo** (*File → Import File*), selecciona `Markdownificator_v1.2.bas` y pulsa **Abrir**. Comprueba que en el panel aparece **Normal → Módulos → Markdownificator_v1_2**. [Importar un módulo en Visual Basic](https://learn.microsoft.com/es-es/office/vba/language/reference/user-interface-help/file-menu)
5. Con **Normal** seleccionado, pulsa el icono **Guardar** del editor. Vuelve a Word con **Alt + F11** y cierra Word. Si Word pregunta si deseas guardar los cambios de la plantilla **Normal**, acéptalo.

Esta guía es para una primera instalación. Si ya instalaste una versión anterior o de prueba de Markdownificator, al añadir el botón fíjate en el nombre completo de la macro. Las versiones anteriores también pueden aparecer como `ExportarAMarkdown`. Si ya ves `Markdownificator_v1_2` en **Normal → Módulos**, comprueba qué versión tienes antes de repetir la importación. Word puede añadir un número al nombre del módulo duplicado. [Comportamiento al importar un módulo duplicado](https://learn.microsoft.com/es-es/office/vba/language/reference/user-interface-help/file-menu)

## 3. Añade el botón «Exportar a Markdown»

1. Vuelve a abrir Word y entra en **Archivo → Opciones → Barra de herramientas de acceso rápido** (*Quick Access Toolbar*).
2. En **Comandos disponibles en**, elige **Macros**. Busca `Normal.Markdownificator_v1_2.ExportarAMarkdown`. El nombre puede mostrarse abreviado.
3. En la lista de la derecha, elige **Para todos los documentos**. Selecciona la macro a la izquierda y pulsa **Agregar**.
4. Selecciona la macro añadida a la derecha, pulsa **Modificar**, elige un icono y escribe **Exportar a Markdown** como nombre para mostrar. Pulsa **Aceptar** para cerrar cada cuadro. [Asignar un botón a una macro](https://support.microsoft.com/es-es/word/create-or-run-a-macro)

Si no ves la barra, haz clic con el botón derecho sobre la cinta y elige **Mostrar barra de herramientas de acceso rápido**. También puedes ejecutar la macro desde **Vista → Macros → Ver macros**. [Mostrar la barra de acceso rápido](https://support.microsoft.com/es-es/office/customize-the-quick-access-toolbar) · [Ejecutar una macro](https://support.microsoft.com/es-es/word/create-or-run-a-macro)

## 4. Haz tu primera exportación

1. Abre un documento en Word. Para una prueba sencilla, crea uno en blanco y escribe una línea de texto.
2. Pulsa **Exportar a Markdown** en la barra de acceso rápido. Debe abrirse **Guardar como** con **Markdownificator v1.2** en el título.
3. Elige dónde guardarlo, pulsa **Guardar** y abre el archivo `.md` con el Bloc de notas para comprobar que contiene tu texto. Si el documento incluye imágenes, la macro puede avisarte de que el Markdown conserva las referencias y los archivos de imagen no se exportan.

Está previsto añadir la exportación de los archivos de imagen en una próxima versión.

Si la macro indica que falta Pandoc, comprueba `pandoc --version` en una terminal nueva. Si el comando funciona, cierra todas las ventanas de Word y vuelve a abrirlo antes de reintentar la exportación. Si Word indica que tu organización bloquea las macros, consulta con su equipo de soporte. [Configuración de macros en Office](https://support.microsoft.com/es-es/office/vba/enable-or-disable-macros-in-microsoft-365-files)

## Dónde se guarda el Markdown

Cada vez que exportas, la macro pregunta a Windows dónde está la carpeta **Descargas** de tu cuenta. Si la localiza, propone esa carpeta en **Guardar como**, incluso si has trasladado Descargas a otra unidad o has guardado la exportación anterior en otra carpeta. Puedes cambiar tanto la carpeta como el nombre del archivo antes de pulsar **Guardar**.

Si Windows no devuelve una ruta válida o la macro no puede confirmar que la carpeta existe, la macro propone solo el nombre del archivo. En ese caso, Windows elige qué carpeta mostrar. Comprueba la ubicación en **Guardar como** y elige la que prefieras. Este caso se ha simulado con el diálogo de Windows. Queda por comprobar cómo se presenta en Word en una cuenta con Descargas inaccesible.

## Ayuda y errores

Si aparece un problema, abre una incidencia en la pestaña **Issues** del repositorio. Indica tus versiones de Windows, Word y Pandoc, el paso en el que ocurrió y el mensaje de error si lo hubo. Si **Guardar como** abre otra carpeta, indica si puedes abrir **Descargas** desde el Explorador de archivos y qué carpeta apareció (por ejemplo, «la del documento» o «la última utilizada»). Omite los datos personales que aparezcan en el mensaje o en las rutas. Basta con describir lo ocurrido sin adjuntar el documento. [Cómo abrir una incidencia en GitHub](https://docs.github.com/en/issues/tracking-your-work-with-issues/using-issues/creating-an-issue)

## Créditos

- **J.F. Franco (@AMTWorks)** · Dirección del proyecto y pruebas en Word.
- **Claude (Anthropic)** · Desarrollo de la v1.1.
- **ChatGPT (OpenAI)** · Desarrollo de la v1.1.1 y la v1.2.

## Licencia

Markdownificator se distribuye bajo la [licencia MIT](LICENSE).
