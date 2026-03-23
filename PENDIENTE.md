# Pendiente — Generador de Evaluaciones AMR

## Bugs a corregir

1. **Header repetido en multipágina**
   - Problema: cuando la evaluación es larga y ocupa más de 1 página, el header (logo+título+Nombre/Curso/Fecha) aparece en cada página
   - Solución: usar 2 tablas separadas — tabla_header (fija, solo pág 1) + tabla_content (fluye entre páginas sin repetir el header)

2. **Margen inferior muy grande**
   - Reducir `sec.bottom_margin` de `Cm(1.5)` a `Cm(1.2)`

3. **Logo AMR no visible en la app**
   - El logo queda invisible porque el fondo del header es muy oscuro
   - Solución: poner el label del logo sobre `bg="white"`

---

## Features a agregar

4. **Soporte de imágenes en las evaluaciones**
   - En el JSON agregar campo opcional `"imagen": "nombre_archivo.png"`
   - La app busca el archivo en la carpeta de imágenes seleccionada
   - La imagen se inserta entre el pasaje y las preguntas
   - JSON ejemplo:
     ```json
     {
       "numero": 1,
       "instruccion": "Observa la imagen y responde:",
       "pasaje": "Texto opcional.",
       "imagen": "grafico1.png",
       "imagen_ancho_cm": 10,
       "preguntas": [...]
     }
     ```

5. **Título personalizable**
   - Agregar campo opcional `"titulo"` en el JSON
   - Si no se pone, usa `"FICHA DE COMPRENSIÓN LECTORA N° {numero}"`
   - Ejemplo: `"titulo": "EVALUACIÓN DE CIENCIAS N° 1"`

---

## UX a mejorar

6. **Selector de carpeta de imágenes más visible** — hacerlo más prominente en la UI
7. **Actualizar los formatos JSON** (botones de copiar) para incluir los campos nuevos (imagen, titulo)

---

## Para empaquetar (.exe)

- Agregar pantalla "Acerca de" o splash con nombre del autor
- Comprimir con PyInstaller incluyendo los logos
- Definir nombre del autor antes de empaquetar

---

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `generador_app.py` | App principal |
| `logo_gabriela_mistral.png` | Logo del colegio |
| `amr logo.png` | Logo AMR (marca del profesor) |
| `requirements.txt` | Dependencias Python |
