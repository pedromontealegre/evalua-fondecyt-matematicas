# Evaluador FONDECYT Matemáticas

**App web:** https://evaluabot.streamlit.app/

Prototipo en Python + Streamlit para evaluar publicaciones en revistas del grupo de estudios de Matemáticas FONDECYT.

## Qué hace

- Carga automáticamente el listado FONDECYT desde `listado.xlsx`, incluido en la carpeta de la app.
- Pide solamente la bibliografía: se puede subir un archivo `.bib` o pegar el contenido BibTeX en un campo de texto.
- Filtra publicaciones desde un año dado, por defecto 2021.
- Clasifica revistas reconocidas como `MB`, `B` o `R`.
- Ordena por defecto en el orden `MB`, luego `B`, luego `R`.
- Muestra una tabla compacta con trazabilidad del match hecho contra el listado.
- Permite mostrar opcionalmente `Título` y `Autores`.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso web local con Streamlit

```bash
streamlit run app_streamlit.py
```

La interfaz permite:

1. Subir un archivo `.bib`, o pegar la bibliografía BibTeX en el campo de texto ubicado justo debajo.
2. Escoger el año inicial, por defecto 2021.
3. Ajustar el umbral de coincidencia aproximada, si hiciera falta.
4. Elegir campos opcionales para la tabla:
   - Título
   - Autores
5. Presionar **Actualizar resultados**.

El archivo `listado.xlsx` debe estar en la misma carpeta que `app_streamlit.py`. En este paquete ya viene incluido.

## Tabla principal

Por defecto, la tabla muestra solo:

- Clasificación
- Año
- Revista en BibTeX
- Match en listado

Los campos `Título` y `Autores` están ocultos por defecto.

## Uso por consola

```bash
python fondecyt_eval.py --bib cv.bib --since 2021 --show-review
```

Por defecto, el modo consola busca `listado.xlsx` en la carpeta actual. También puedes indicar otra ruta explícitamente:

```bash
python fondecyt_eval.py --listado otro_listado.xlsx --bib cv.bib --since 2021
```

Esto genera:

- `resultados.csv`: publicaciones reconocidas.
- `para_revisar.csv`: entradas que requieren revisión manual.
- `diagnostico.csv`: tabla completa de depuración.

## Notas

El programa intenta reconocer abreviaturas frecuentes de revistas, por ejemplo:

- `Theor. Comput. Sci.` ↔ `Theoretical Computer Science`
- `J. Comput. Syst. Sci.` ↔ `Journal of Computer and System Sciences`
- `Inf. Comput.` ↔ `Information and Computation`
- `SIAM J. Comput.` ↔ `SIAM Journal on Computing`

Las entradas no reconocidas quedan en `para_revisar.csv` y en el expansor **Para revisar** dentro de la app.
