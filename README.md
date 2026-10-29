# Consulta Miembro de Mesa ONPE
**Extracción automatizada de datos y contenedorización con Docker**

Sistema automatizado que consulta el portal oficial de la ONPE para determinar si una lista de ciudadanos (identificados por DNI) son miembros de mesa en las Elecciones Generales 2026. Los resultados se muestran en una interfaz web y se exportan automáticamente a un archivo Excel.

---

## Descripción

El sistema:
1. Lee una lista de DNIs desde el archivo `dnis.xlsx`
2. Consulta automáticamente el portal [consultaelectoral.onpe.gob.pe](https://consultaelectoral.onpe.gob.pe/inicio) usando Selenium
3. Detecta si cada ciudadano es miembro de mesa (Presidente, Secretario, Tercer Miembro o Suplente)
4. Extrae: nombre completo, región, provincia, distrito, dirección del local y número de mesa
5. Muestra los resultados en una interfaz web y actualiza el Excel automáticamente

---

## Estructura del Proyecto

```
onpe-scraping-docker/
├── app.py                   # Aplicación principal Flask + Selenium
├── dnis.xlsx                # Lista de DNIs a consultar
├── requirements.txt         # Dependencias Python
├── Dockerfile               # Imagen base
├── Dockerfile.optimizado    # Imagen optimizada con capas reducidas
├── Dockerfile.multistage    # Imagen multistage (build separado)
└── README.md                # Este archivo
```

---

## Requisitos Previos

- [Docker Desktop](https://www.docker.com/products/docker-desktop/) instalado y corriendo
- Conexión a internet (para acceder al portal de ONPE)
- El archivo `dnis.xlsx` con DNIs reales en la columna B (desde la fila 5)

---

## Pasos de Instalación y Ejecución

### Paso 1 — Clonar el repositorio

```bash
git clone https://github.com/sebasalvadev/onpe-scraping-docker.git
cd sem04
```

### Paso 2 — Agregar los DNIs al Excel

Abre `dnis.xlsx` con Excel y coloca los DNIs de 8 dígitos en la columna B desde la fila 5. Guarda y cierra el archivo.

### Paso 3 — Construir la imagen Docker

```bash
docker build -t onpe-app:v1.0 .
```

### Paso 4 — Ejecutar el contenedor

```bash
docker run -d -p 5000:5000 -v "%cd%\dnis.xlsx:/app/dnis.xlsx" --name onpe-v1 onpe-app:v1.0
```

> En PowerShell usar `${PWD}` en lugar de `%cd%`

### Paso 5 — Abrir la interfaz web

Abre tu navegador y ve a:

```
http://localhost:5000
```

### Paso 6 — Ejecutar la consulta

Presiona el botón **"Consultar todos los DNIs"** y espera. Cada DNI tarda entre 15 y 30 segundos. Para 5 DNIs el proceso completo toma entre 2 y 3 minutos.

### Paso 7 — Ver resultados

Los resultados aparecen en la tabla de la interfaz web y quedan guardados automáticamente en el archivo `dnis.xlsx`.

---

## Variantes de Dockerfile

### Dockerfile (base)
Imagen estándar basada en `python:3.11-slim`. Instala Chromium, ChromeDriver y todas las dependencias necesarias para ejecutar Selenium en modo headless dentro del contenedor.

```bash
docker build -f Dockerfile -t onpe-app:v1.0 .
```

### Dockerfile.optimizado
Reduce el tamaño de la imagen combinando comandos RUN en una sola capa y limpiando cachés de apt y pip en el mismo paso. Resultado: imagen más ligera y menor tiempo de descarga.

```bash
docker build -f Dockerfile.optimizado -t onpe-app:optimizado .
```

### Dockerfile.multistage
Usa construcción en dos etapas: la primera instala dependencias y compila lo necesario; la segunda copia solo los artefactos finales a una imagen limpia. Ideal para producción.

```bash
docker build -f Dockerfile.multistage -t onpe-app:multistage .
```

---

## Detener y eliminar el contenedor

```bash
docker stop onpe-v1 && docker rm onpe-v1
```

---

## Campos extraídos por el sistema

| Campo | Descripción |
|---|---|
| DNI | Número de documento consultado |
| ¿Miembro? | SÍ / NO según el portal de ONPE |
| Nombre | Nombres y apellidos completos |
| Región | Región de votación |
| Provincia | Provincia de votación |
| Distrito | Distrito de votación |
| Dirección del Local | Nombre y dirección del local de votación |
| N° Mesa | Número de mesa asignada |
| Rol | Presidente / Secretario / Tercer Miembro / Suplente |

---

## ¿Cómo detecta si es miembro de mesa?

El portal de ONPE muestra un banner con el texto **"ERES PRESIDENTE"**, **"ERES SECRETARIO"**, **"ERES TERCER MIEMBRO"** o **"ERES SUPLENTE"** únicamente para quienes han sido designados. Si ese banner no aparece, el ciudadano no es miembro de mesa. El sistema detecta exactamente ese texto para determinar el resultado.

---

## URL del Repositorio

```
https://github.com/sebasalvadev/onpe-scraping-docker.git
```

---

