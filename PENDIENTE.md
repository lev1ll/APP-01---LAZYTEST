# Pendiente — Generador de Evaluaciones AMR

## Bugs corregidos
- [x] Header repetido en multipágina — separado en 2 tablas (hdr_outer + content_outer)
- [x] Margen inferior muy grande — reducido a Cm(1.2)
- [x] Logo AMR no visible en la app — bg="white" en el label
- [x] Alternativas sueltas entre páginas — una fila por pregunta con cantSplit=true
- [x] Header repetido en modo normal — una sola cabecera, contenido fluye continuo

---

## Features a agregar

1. **Soporte de imágenes en las evaluaciones**
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

2. **Título personalizable**
   - Agregar campo opcional `"titulo"` en el JSON
   - Si no se pone, usa `"FICHA DE COMPRENSIÓN LECTORA N° {numero}"`
   - Ejemplo: `"titulo": "EVALUACIÓN DE CIENCIAS N° 1"`

---

## UX a mejorar

3. **Selector de carpeta de imágenes más visible** — hacerlo más prominente en la UI
4. **Actualizar los formatos JSON** (botones de copiar) para incluir los campos nuevos (imagen, titulo)

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
