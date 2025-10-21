# Contabilidad de Gestión — Práctica P1

Este repositorio incluye una plantilla para la entrega de los ejercicios de la Práctica P1 (temas 1 y 2), de acuerdo con las instrucciones oficiales:

- Contenido:
  - Ejercicios tema 1: ejercicios 2 y 4
  - Ejercicios tema 2: ejercicios A, B y C
- Calificación:
  - Test: 5 puntos
  - Ejercicios: 2 puntos cada ejercicio (total 10 puntos)
- Entrega:
  - Fecha tope: 22 de octubre (improrrogable).
  - Formato: Excel (preferente). Un único archivo, con cada ejercicio en una hoja diferente, y cada hoja con el nombre del ejercicio.
  - Alternativa: PDF (las tablas no deben aparecer partidas; si se parten, esos ejercicios no se evaluarán).
  - No se permiten archivos comprimidos.
  - La entrega en la convocatoria ordinaria sin realizar el test implica perder la puntuación del test. No se admiten entregas en la extraordinaria.

## Plantilla Excel

Este repositorio incluye un script que genera la plantilla de Excel con las hojas:
- Portada (para tus datos: nombre, DNI, grupo, fecha)
- Ejercicio 2
- Ejercicio 4
- Ejercicio A
- Ejercicio B
- Ejercicio C

Cada hoja de ejercicio contiene:
- Un encabezado con el nombre del ejercicio.
- Una tabla de trabajo con columnas "Paso", "Concepto", "Cálculo/Explicación", "Resultado".
- Ajustes de impresión para facilitar la exportación a PDF sin cortes horizontales (modo apaisado y ajuste a una página de ancho).

### Requisitos

- Python 3.8+
- Paquete `openpyxl` (ver `requirements.txt`)

Instalación rápida:
```bash
pip install -r requirements.txt
```

### Generar la plantilla

```bash
python scripts/generar_plantilla_p1.py -o P1_Ejercicios.xlsx
```

Si no especificas `-o`, se guardará como `P1_Ejercicios.xlsx` por defecto.

### Consejos para exportar a PDF sin tablas partidas

En Excel:
- Establece Área de impresión para abarcar la tabla (por ejemplo A1:D60 si no la ajustas dinámicamente).
- Configura:
  - Diseño: Horizontal (Apaisado)
  - Ajustar a: 1 página de ancho por X de alto (o "Ajustar al ancho")
  - Márgenes estrechos si es necesario.
- Inserta saltos de página manuales si la tabla es muy larga, para evitar particiones incómodas.

### Notas importantes

- No cambies los nombres de las hojas.
- No entregues archivos comprimidos.
- Si entregas en PDF, comprueba que ninguna tabla esté partida entre páginas.