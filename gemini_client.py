"""
gemini_client.py
Integracion con Gemini API para generacion de evaluaciones.
- Structured output: Gemini devuelve siempre el formato exacto que necesita Generador
- Historial de chat: el profe puede refinar la prueba sin perder el contexto
Usa el SDK oficial: google-genai
"""

import json
from google import genai
from google.genai import types

# ── Schema de respuesta ───────────────────────────────────────────────────────
_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    required=["fichas"],
    properties={
        "fichas": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                required=["numero", "instruccion", "pasaje", "preguntas", "banco_palabras"],
                properties={
                    "numero":         types.Schema(type=types.Type.INTEGER),
                    "instruccion":    types.Schema(type=types.Type.STRING),
                    "pasaje":         types.Schema(type=types.Type.STRING),
                    "banco_palabras": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.STRING),
                    ),
                    "preguntas": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(
                            type=types.Type.OBJECT,
                            required=["tipo", "numero", "enunciado",
                                      "alternativas", "respuesta_correcta", "lineas"],
                            properties={
                                "tipo":               types.Schema(type=types.Type.STRING),
                                "numero":             types.Schema(type=types.Type.INTEGER),
                                "enunciado":          types.Schema(type=types.Type.STRING),
                                "alternativas":       types.Schema(
                                    type=types.Type.ARRAY,
                                    items=types.Schema(type=types.Type.STRING),
                                ),
                                "respuesta_correcta": types.Schema(type=types.Type.STRING),
                                "lineas":             types.Schema(type=types.Type.INTEGER),
                            },
                        ),
                    ),
                },
            ),
        )
    },
)

# ── Prompt de sistema ─────────────────────────────────────────────────────────
_SYSTEM = """\
Eres un asistente experto en crear evaluaciones escolares para profesores chilenos de educacion basica y media.
Responde siempre en español neutro (sin modismos argentinos ni de otro pais).
Si la asignatura es Ingles, redacta el pasaje y las preguntas en ingles.

REGLAS GENERALES:
1. El pasaje debe ser un texto completo, coherente y apropiado para el nivel indicado.
2. Si el texto es un poema, usa saltos de linea reales (\\n) entre los versos dentro del campo "pasaje".
3. Cuando el usuario pida modificar algo, devuelve la evaluacion COMPLETA con el cambio aplicado.
4. Numera las fichas de forma consecutiva comenzando en 1.
5. Adapta el vocabulario y complejidad al curso indicado.
6. El campo "banco_palabras" va siempre en el objeto de la ficha (no en cada pregunta).

TIPOS DE PREGUNTA — usa el campo "tipo" en cada pregunta:

"seleccion_multiple":
  - 4 alternativas A, B, C, D en el campo "alternativas".
  - "respuesta_correcta" = "A", "B", "C" o "D".
  - "lineas" = 0. "banco_palabras" de la ficha = [].

"verdadero_falso":
  - El enunciado es una afirmacion (no una pregunta).
  - "alternativas" = []. "respuesta_correcta" = "V" o "F".
  - "lineas" = 0. "banco_palabras" de la ficha = [].

"completar":
  - El enunciado contiene _____ donde va la palabra que falta.
  - "alternativas" = []. "respuesta_correcta" = la palabra correcta que completa el espacio.
  - "lineas" = 0.
  - "banco_palabras" de la FICHA: incluye todas las respuestas correctas de las preguntas completar MAS 2 o 3 palabras distractor, en orden aleatorio.

"desarrollo":
  - El enunciado es una pregunta abierta que requiere respuesta escrita.
  - "alternativas" = []. "respuesta_correcta" = "" (respuesta libre).
  - "lineas" = numero de lineas en blanco sugeridas (entre 3 y 6 segun la extension esperada).
  - "banco_palabras" de la ficha = [].

"mixta":
  - Si el tipo solicitado es "mixta", elige el tipo mas apropiado para cada pregunta segun el contexto pedagogico.
  - Puedes mezclar seleccion_multiple, verdadero_falso, completar y desarrollo en la misma ficha.
  - El "banco_palabras" de la ficha incluye las palabras de todas las preguntas tipo completar.
"""

_MODEL = "gemini-2.5-flash-lite"


class GeminiClient:
    """Cliente de Gemini con historial de conversacion y structured output."""

    def __init__(self, api_key: str):
        self._client = genai.Client(api_key=api_key)
        self._historial: list[types.Content] = []
        self._ultima_fichas: list | None = None

    # ── API publica ───────────────────────────────────────────────────────────

    def verificar_api_key(self) -> tuple[bool, str]:
        """Verifica que la API key sea valida con una llamada minima."""
        try:
            resp = self._client.models.generate_content(
                model=_MODEL,
                contents="Responde solo: ok",
                config=types.GenerateContentConfig(max_output_tokens=5),
            )
            _ = resp.text
            return True, ""
        except Exception as e:
            return False, str(e)

    def generar(
        self,
        prompt: str,
        n_fichas: int = 1,
        n_preguntas: int = 4,
        curso: str = "3° Básico",
        asignatura: str = "Lenguaje",
        tipo_pregunta: str = "Selección múltiple",
        texto_base: str = "",
        imagen_bytes: bytes | None = None,
    ) -> list:
        """
        Llama a Gemini y devuelve la lista de fichas generadas.
        Mantiene historial para refinamientos posteriores.
        """
        prompt_final = self._construir_prompt(
            prompt, n_fichas, n_preguntas, curso, asignatura, tipo_pregunta, texto_base)

        if imagen_bytes:
            mime = "image/png" if imagen_bytes[:8] == b'\x89PNG\r\n\x1a\n' else "image/jpeg"
            partes = [
                types.Part(inline_data=types.Blob(data=imagen_bytes, mime_type=mime)),
                types.Part(text=prompt_final + "\nAnaliza la imagen y crea preguntas que hagan referencia directa a ella."),
            ]
        else:
            partes = [types.Part(text=prompt_final)]

        self._historial.append(
            types.Content(role="user", parts=partes)
        )

        response = self._client.models.generate_content(
            model=_MODEL,
            contents=self._historial,
            config=types.GenerateContentConfig(
                system_instruction=_SYSTEM,
                response_mime_type="application/json",
                response_schema=_SCHEMA,
                temperature=0.7,
            ),
        )

        texto = response.text
        self._historial.append(
            types.Content(role="model", parts=[types.Part(text=texto)])
        )

        data = json.loads(texto)
        fichas = data.get("fichas", data) if isinstance(data, dict) else data
        self._ultima_fichas = fichas
        return fichas

    def limpiar_historial(self) -> None:
        self._historial = []
        self._ultima_fichas = None

    def get_historial_raw(self) -> list:
        """Serializa el historial para guardarlo en disco."""
        result = []
        for content in self._historial:
            parts = []
            for p in content.parts:
                if hasattr(p, "text") and p.text:
                    parts.append({"text": p.text})
            if parts:
                result.append({"role": content.role, "parts": parts})
        return result

    def set_historial_raw(self, data: list) -> None:
        """Restaura el historial desde datos serializados."""
        self._historial = []
        for item in data:
            try:
                self._historial.append(
                    types.Content(
                        role=item["role"],
                        parts=[types.Part(text=p["text"])
                               for p in item.get("parts", []) if p.get("text")]
                    )
                )
            except Exception:
                pass

    @property
    def ultima_fichas(self) -> list | None:
        return self._ultima_fichas

    @property
    def tiene_fichas(self) -> bool:
        return self._ultima_fichas is not None

    # ── Helpers privados ──────────────────────────────────────────────────────

    def _construir_prompt(
        self, prompt: str, n_fichas: int, n_preguntas: int,
        curso: str, asignatura: str, tipo_pregunta: str, texto_base: str
    ) -> str:
        # Mapear nombre UI → tipo interno
        _MAPA = {
            "Selección múltiple": "seleccion_multiple",
            "Verdadero / Falso":  "verdadero_falso",
            "Completar texto":    "completar",
            "Desarrollo":         "desarrollo",
            "Mixta":              "mixta",
        }
        tipo_interno = _MAPA.get(tipo_pregunta, "seleccion_multiple")
        partes = [
            f"Asignatura: {asignatura}",
            f"Curso: {curso}",
            f"Tipo de preguntas: {tipo_interno}",
            f"Cantidad de fichas: {n_fichas}",
            f"Preguntas por ficha: {n_preguntas}",
        ]
        if texto_base.strip():
            partes.append(f"\nTexto base del profesor:\n{texto_base.strip()}")
            partes.append("Crea las preguntas basandote en ese texto.")
        else:
            partes.append("Crea el texto (pasaje) y las preguntas desde cero.")

        partes.append(f"\nPedido: {prompt.strip()}")
        return "\n".join(partes)
