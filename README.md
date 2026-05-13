# Auditor IVA Streamlit v2.0

App Streamlit para auditar IVA Compras comparando **AFIP/ARCA Mis Comprobantes** contra el **Libro IVA del sistema**.

## Auditorías incluidas

### 1. Auditoría de Comprobantes
Control documental por comprobante:

**CUIT + Tipo + Punto de Venta + Número**

Estados:

- OK
- Sólo AFIP/ARCA
- Sólo Libro IVA
- IVA distinto
- Duplicados en AFIP/ARCA
- Duplicados en Libro IVA

### 2. Auditoría IVA del Mes
Valida el IVA computado en el período fiscal según el **Libro IVA del sistema cargado** y lo cruza contra AFIP/ARCA como respaldo documental.

Importante: el Libro IVA cargado se toma como universo del período fiscal. La fecha visible de cada comprobante se usa para alertar comprobantes con fecha anterior o posterior, no para excluirlos automáticamente.

## Historial obligatorio

Cada corrida puede guardarse en historial con:

- Fecha/hora
- Usuario
- Sociedad/CUIT
- Período fiscal auditado
- Tipo de auditoría
- Archivos cargados
- Totales de IVA
- Diferencia neta y diferencia bruta
- Cantidad de OK, sólo AFIP, sólo Libro, IVA distinto y duplicados
- Descarga posterior del Excel/PDF generado

## Formatos soportados

- AFIP/ARCA: `.xlsx`, `.csv`, `.xls`
- Libro IVA sistema: `.xlsx`, `.csv`, `.xls`

Para leer `.xls` antiguo se incluye `xlrd>=2.0.1` en `requirements.txt`.

## Login inicial

```text
Usuario: dancona2016@gmail.com
Contraseña: Dancona2026*
```

En Streamlit Cloud también se pueden usar secrets:

```toml
APP_USER = "dancona2016@gmail.com"
APP_PASSWORD = "Dancona2026*"
```

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```
