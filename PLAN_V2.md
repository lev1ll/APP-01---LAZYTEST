# Plan de desarrollo — Generador de Evaluaciones v2

## Estado actual (rama `v2-gemini-integrado`)

La app funciona con:
- UI de 3 paneles: historial | chat con Gemini | opciones
- Modo oscuro
- Historial de conversaciones persistido en `%APPDATA%/GeneradorAMR/conversations/`
- Clic derecho en conversación → renombrar / eliminar
- Input de chat auto-expandible (máx. 3 líneas)
- Selector de asignatura
- Indicador "Pensando…" mientras Gemini trabaja
- Versión B corregida (baraja alternativas y actualiza respuesta_correcta)
- API Key verificación asíncrona (no bloquea la UI)

---

## Etapa 1 — Burbujas de chat [ PENDIENTE ]

Rediseño visual estilo Gemini:

- [ ] Ancho máximo 65% usuario / 78% IA (no full-width)
- [ ] Esquinas muy redondeadas (20px)
- [ ] Padding interno generoso
- [ ] Ícono pequeño (punto de color o "G") junto a burbuja IA
- [ ] Timestamp (hora) debajo de cada burbuja
- [ ] Sombra sutil en burbujas IA

---

## Etapa 2 — Sistema de imágenes [ PENDIENTE ]

### Flujo completo
```
Profe sube PDF → app extrae texto + imágenes
Profe elige imagen (o sube una manualmente)
Gemini recibe texto + imagen (Vision API) → genera preguntas sobre el contenido visual
Word generado incluye la imagen entre la instrucción y las preguntas
```

### Sub-tareas

**2a — Fuentes de imagen (3 opciones)**
- [ ] Auto desde PDF: al adjuntar PDF, extraer imágenes con `pymupdf` y mostrarlas como miniaturas para elegir
- [ ] Selector manual: botón para abrir JPG/PNG directamente desde el computador
- [ ] Sin imagen: comportamiento por defecto (como ahora)

**2b — Panel de vista previa**
- [ ] Mostrar miniatura de la imagen seleccionada con botón para quitarla

**2c — Gemini Vision**
- [ ] Actualizar `gemini_client.py` para enviar imagen + texto a la API (multimodal)
- [ ] La imagen se pasa como `bytes` junto con el prompt

**2d — Imagen en el Word**
- [ ] Insertar la imagen seleccionada en el documento entre la instrucción y las preguntas
- [ ] Respetar ancho de página (máx. 16cm)
- [ ] Funciona en ambos modos: con encabezado y sin encabezado

**2e — Dependencia nueva**
- [ ] Agregar `pymupdf` a `requirements_v2.txt`

---

## Etapa 3 — Soporte de Matemáticas [ PENDIENTE ]

### Estrategia por nivel

| Nivel | Enfoque |
|-------|---------|
| Básica (1°–8°) | Texto Unicode: `x²`, `√x`, `½`, `÷`, `×`, `π` |
| Media (1°–4°) | Unicode extendido + fracciones como `(2x+1)/(x-3)` |
| Avanzado (futuro) | Render de ecuaciones como imagen con `matplotlib` |

### Sub-tareas

- [ ] Actualizar `_SYSTEM` en `gemini_client.py` con instrucciones de formato matemático según asignatura
- [ ] Cuando asignatura = "Matemáticas", Gemini usa Unicode para potencias, raíces y símbolos
- [ ] Para geometría: el profe puede subir una figura como imagen (usa el sistema de Etapa 2)
- [ ] Evaluar si agregar render LaTeX→imagen con `matplotlib` para nivel medio avanzado

---

## Backlog (ideas futuras)

- [ ] Splash screen / pantalla "Acerca de" para el .exe empaquetado
- [ ] Empaquetar con PyInstaller (.exe) incluyendo logos e icono
- [ ] Modo claro / modo oscuro conmutable desde la UI
- [ ] Vista previa del Word antes de descargar
- [ ] Soporte para más tipos de pregunta (desarrollo, verdadero/falso)

---

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `generador_v2.py` | App principal (entry point) |
| `generador_app.py` | Motor de generación Word + clase App legacy |
| `gemini_client.py` | Cliente Gemini con historial y structured output |
| `config.py` | API Key persistida en `%APPDATA%/GeneradorAMR/` |
| `file_parser.py` | Extracción de texto desde PDF, DOCX, TXT |
| `requirements_v2.txt` | Dependencias Python |
| `logo_gabriela_mistral.png` | Logo del colegio |
| `amr logo.png` | Logo AMR |
| `icono.ico` | Ícono de la app |
