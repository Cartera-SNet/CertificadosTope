# Descarga Automática Certificados Tope SOAT
## Activa IT / RGC · Grupo Campbell

App web local que automatiza la descarga de Certificados de Tope desde Activa IT
usando Playwright, con interfaz web para monitorear el progreso en tiempo real.

---

## Instalación (una sola vez)

```
pip install flask playwright pandas openpyxl
playwright install chromium
```

---

## Arrancar

Doble clic en `iniciar.bat`  
O desde terminal: `python app.py`

Luego abre: **http://127.0.0.1:5001**

---

## Uso

1. Sube el Excel con columnas: **TIPO · ID · TIPO DE AMPARO**
2. Ingresa usuario y contraseña de Activa IT
3. Pulsa **Iniciar descarga automática**
4. Monitorea el progreso en tiempo real en el log y la tabla
5. Al terminar descarga **Excel** (reporte) o **ZIP** (PDFs + reporte)

---

## Formato del Excel de entrada

| TIPO | ID         | TIPO DE AMPARO |
|------|------------|----------------|
| CC   | 22465561   | MED            |
| CC   | 1007187279 | MED            |

Códigos de amparo: MED · TRA · FUN · PER · MUE

---

## Puerto

Esta app corre en el puerto **5001** para no chocar con la app de lectura de PDFs (5000).
# CertificadosTope
