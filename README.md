# 📋 Guía de Configuración y Ejecución del Proyecto

Este proyecto contiene un ecosistema de herramientas para el análisis de clientes, generación de reportes y dashboards interactivos. Dado que algunos archivos de datos pesados (`.xlsx`) están configurados en el `.gitignore` para evitar saturar el repositorio de GitHub, es necesario seguir esta guía para compartir y ejecutar todo correctamente en otra laptop.

---

## 📂 Estructura del Proyecto

El proyecto está compuesto por los siguientes módulos principales:

| Componente / Archivo | Tipo | Descripción |
| :--- | :--- | :--- |
| **`TRX WU_BP.xlsx`** | Base de Datos | Archivo Excel con las transacciones históricas (insumo principal). |
| **`ejecutar_analisis.bat`** | Automatización | Ejecuta los scripts de análisis de datos y sube los resultados CSV a GitHub. |
| **`iniciar_dashboard.bat`** | Dashboard | Inicia el **Dashboard de Alertas Inteligentes de Clientes** (`app.py`). |
| **`iniciar_webapp.bat`** | Dashboard | Inicia el **Dashboard de Inteligencia de Negocio y Salud** (`dashboard.py`). |
| **`iniciar_servidor.bat`** | Web App | Levanta un servidor local en el puerto 8080 para la **Web App de Proyección** (`webapp/`). |
| **`actualizar_nube.bat`** | Despliegue | Sube la versión actualizada del Dashboard y la base de datos a Streamlit Cloud. |
| **`requirements.txt`** | Dependencias | Archivo con las librerías de Python necesarias para correr el proyecto. |

---

## 📤 Paso 1: Cómo Compartir el Proyecto

Dado que el archivo de datos `TRX WU_BP.xlsx` y los reportes generados no están en el repositorio de Git por su peso, la forma más rápida y recomendada de compartir el proyecto completo es:

1. Ve a la carpeta raíz del proyecto en tu laptop.
2. Comprime toda la carpeta en un archivo **`.zip`** (puedes excluir las carpetas ocultas `.git` y `__pycache__` si quieres reducir el peso, pero asegúrate de incluir `TRX WU_BP.xlsx`).
3. Envía el archivo `.zip` a tu compañero a través de **Google Drive**, **OneDrive**, **Teams**, o una memoria USB.

---

## 📥 Paso 2: Configuración en la Nueva Laptop

Para que tu compañero pueda ejecutar todo sin problemas, debe preparar su entorno siguiendo estos pasos:

### 1. Instalar Python
* Descargar e instalar la última versión de **Python** (versión 3.9 o superior) desde la página oficial: [python.org](https://www.python.org/downloads/).
* ⚠️ **MUY IMPORTANTE**: Durante la instalación en Windows, asegúrate de marcar la casilla **"Add Python to PATH"** (Agregar Python al PATH) en el primer instalador. De lo contrario, los comandos de automatización no funcionarán.

### 2. Instalar Git (para las actualizaciones)
* Descargar e instalar **Git** desde: [git-scm.com](https://git-scm.com/).
* Configurar las credenciales de Git en la terminal (Powershell o CMD):
  ```bash
  git config --global user.name "Nombre del Compañero"
  git config --global user.email "correo@ejemplo.com"
  ```
* Asegurarse de que su cuenta de GitHub tenga permisos de acceso a los siguientes repositorios:
  * `https://github.com/muresbrian/proyecci-n-pw.git` (Repositorio del Proyecto / Webapp)
  * `https://github.com/muresbrian/dashboard-alertas-diarias.git` (Repositorio de Streamlit Cloud)

### 3. Extraer y Preparar el Proyecto
1. Descomprimir el archivo `.zip` recibido en una carpeta local (por ejemplo, en `Documentos` o `Escritorio`).
2. Abrir la consola de comandos (CMD o PowerShell) en esa carpeta.
3. Instalar las dependencias de Python ejecutando:
   ```bash
   pip install -r requirements.txt
   ```

---

## ⚙️ Paso 3: Cómo Ejecutar y Actualizar las Plataformas

Una vez configurado el entorno, todo se controla de forma sencilla mediante los archivos `.bat` (doble clic en cada uno):

### 1. Actualizar Datos e Historial de Análisis
Si se actualiza la base de datos `TRX WU_BP.xlsx` con transacciones nuevas:
* Haz doble clic en **`ejecutar_analisis.bat`**.
* Este script validará Python, instalará las dependencias necesarias y ejecutará consecutivamente los scripts de tendencias, salud de cartera y clientes perdidos.
* Finalmente, exportará los archivos de visualización e intentará subir los reportes actualizados a GitHub para alimentar la web app.

### 2. Ejecutar Dashboard de Alertas Inteligentes (Streamlit Local)
* Haz doble clic en **`iniciar_dashboard.bat`**.
* Se abrirá una ventana de comandos e iniciará Streamlit. Automaticamente se abrirá tu navegador con el panel interactivo de alertas.
* *Nota: Mantén abierta la ventana negra de consola mientras uses el Dashboard.*

### 3. Ejecutar Dashboard de Inteligencia y Salud (Streamlit Local)
* Haz doble clic en **`iniciar_webapp.bat`**.
* Se abrirá el Dashboard general que contiene las pestañas de **Inteligencia de Negocio**, **Salud de la Cartera 2026** y **Clientes Perdidos**.

### 4. Lanzar la Web App de Proyección (HTML/CSS/JS)
* Haz doble clic en **`iniciar_servidor.bat`**.
* Esto iniciará un servidor web local en el puerto `8080` y abrirá la interfaz premium de la Web App de Proyección en tu navegador (`http://localhost:8080`).

### 5. Actualizar el Dashboard en la Nube (Streamlit Cloud)
* Si quieres que tus cambios en el código o en los datos de transacciones se reflejen en la versión web pública del dashboard:
* Haz doble clic en **`actualizar_nube.bat`**.
* Este script clonará (si no existe) la carpeta `Nube_Dashboard`, copiará el nuevo código y base de datos, y los subirá mediante Git al repositorio vinculado a Streamlit Cloud.
