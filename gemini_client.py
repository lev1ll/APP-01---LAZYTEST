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
                required=["numero", "instruccion", "pasaje", "preguntas"],
                properties={
                    "numero":      types.Schema(type=types.Type.INTEGER),
                    "instruccion": types.Schema(type=types.Type.STRING),
                    "pasaje":      types.Schema(type=types.Type.STRING),
                    "preguntas": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(
                            type=types.Type.OBJECT,
                            required=["numero", "enunciado", "alternativas", "respuesta_correcta"],
                            properties={
                                "numero":             types.Schema(type=types.Type.INTEGER),
                                "enunciado":          types.Schema(type=types.Type.STRING),
                                "alternativas":       types.Schema(
                                    type=types.Type.ARRAY,
                                    items=types.Schema(type=types.Type.STRING),
                                ),
                                "respuesta_correcta": types.Schema(type=types.Type.STRING),
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

REGLAS:
1. Crea evaluaciones de seleccion multiple con 4 alternativas (A, B, C, D) por pregunta.
2. El campo "respuesta_correcta" debe ser exactamente "A", "B", "C" o "D".
3. El pasaje debe ser un texto completo, coherente y apropiado para el nivel indicado.
4. Si el texto es un poema, usa saltos de linea reales (\\n) entre los versos dentro del campo "pasaje".
5. La instruccion siempre es: "Lee atentamente el siguiente texto:"
6. Cuando el usuario pida modificar algo, devuelve la evaluacion COMPLETA con el cambio aplicado.
7. Numera las fichas de forma consecutiva comenzando en 1.
8. Adapta el vocabulario y complejidad al curso indicado.
9. Las alternativas incorrectas deben ser plausibles pero claramente erroneas para ese nivel.
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
        texto_base: str = "",
    ) -> list:
        """
        Llama a Gemini y devuelve la lista de fichas generadas.
        Mantiene historial para refinamientos posteriores.
        """
        prompt_final = self._construir_prompt(prompt, n_fichas, n_preguntas, curso, texto_base)

        self._historial.append(
            types.Content(role="user", parts=[types.Part(text=prompt_final)])
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

    @property
    def ultima_fichas(self) -> list | None:
        return self._ultima_fichas

    @property
    def tiene_fichas(self) -> bool:
        return self._ultima_fichas is not None

    # ── Helpers privados ──────────────────────────────────────────────────────

    def _construir_prompt(
        self, prompt: str, n_fichas: int, n_preguntas: int, curso: str, texto_base: str
    ) -> str:
        partes = [
            f"Curso: {curso}",
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
