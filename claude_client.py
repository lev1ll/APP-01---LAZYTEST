"""
claude_client.py
Integración con Claude API (Anthropic) para generación de evaluaciones.
Misma interfaz que GeminiClient para intercambio transparente.
"""

import json
import base64
import anthropic

_MODEL = "claude-haiku-4-5-20251001"

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

FORMATO DE RESPUESTA:
Responde ÚNICAMENTE con JSON válido, sin texto adicional, sin markdown, sin bloques de código.
El JSON debe tener esta estructura exacta:
{
  "fichas": [
    {
      "numero": 1,
      "instruccion": "...",
      "pasaje": "...",
      "banco_palabras": [],
      "preguntas": [
        {
          "tipo": "seleccion_multiple",
          "numero": 1,
          "enunciado": "...",
          "alternativas": ["A) ...", "B) ...", "C) ...", "D) ..."],
          "respuesta_correcta": "A",
          "lineas": 0
        }
      ]
    }
  ]
}
"""


class ClaudeClient:
    """Cliente de Claude con historial de conversacion."""

    def __init__(self, api_key: str):
        self._client = anthropic.Anthropic(api_key=api_key)
        self._historial: list = []
        self._ultima_fichas: list | None = None

    # ── API publica ───────────────────────────────────────────────────────────

    def verificar_api_key(self) -> tuple[bool, str]:
        try:
            self._client.messages.create(
                model=_MODEL,
                max_tokens=5,
                messages=[{"role": "user", "content": "ok"}],
            )
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
        imagenes_bytes: list | None = None,
    ) -> list:
        # Si ya hay historial → es un refinamiento, mandar solo el prompt del profe
        if self._historial:
            prompt_final = prompt.strip()
        else:
            prompt_final = self._construir_prompt(
                prompt, n_fichas, n_preguntas, curso, asignatura, tipo_pregunta, texto_base)

        # Construir contenido del mensaje
        if imagenes_bytes:
            content = []
            for img_bytes in imagenes_bytes:
                mime = "image/png" if img_bytes[:8] == b'\x89PNG\r\n\x1a\n' else "image/jpeg"
                content.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": mime,
                        "data": base64.standard_b64encode(img_bytes).decode("utf-8"),
                    }
                })
            content.append({"type": "text",
                             "text": prompt_final + "\nAnaliza las imágenes y crea preguntas que hagan referencia directa a ellas."})
        else:
            content = prompt_final

        self._historial.append({"role": "user", "content": content})

        response = self._client.messages.create(
            model=_MODEL,
            max_tokens=8192,
            system=_SYSTEM,
            messages=self._historial,
        )

        texto = response.content[0].text
        self._historial.append({"role": "assistant", "content": texto})

        # Limpiar posibles bloques markdown
        texto_limpio = texto.strip()
        if texto_limpio.startswith("```"):
            texto_limpio = texto_limpio.split("\n", 1)[-1]
            texto_limpio = texto_limpio.rsplit("```", 1)[0]

        try:
            data = json.loads(texto_limpio)
        except json.JSONDecodeError:
            raise ValueError(
                "La respuesta fue demasiado larga y se cortó.\n"
                "Intentá con menos fichas o menos preguntas por ficha."
            )
        fichas = data.get("fichas", data) if isinstance(data, dict) else data
        self._ultima_fichas = fichas
        return fichas

    def limpiar_historial(self) -> None:
        self._historial = []
        self._ultima_fichas = None

    def get_historial_raw(self) -> list:
        result = []
        for msg in self._historial:
            content = msg["content"]
            if isinstance(content, str):
                result.append({"role": msg["role"], "parts": [{"text": content}]})
            elif isinstance(content, list):
                textos = [p["text"] for p in content if p.get("type") == "text"]
                if textos:
                    result.append({"role": msg["role"], "parts": [{"text": t} for t in textos]})
        return result

    def set_historial_raw(self, data: list) -> None:
        self._historial = []
        for item in data:
            role = item.get("role", "user")
            # Mapear "model" (Gemini) → "assistant" (Claude)
            if role == "model":
                role = "assistant"
            textos = [p["text"] for p in item.get("parts", []) if p.get("text")]
            if textos:
                self._historial.append({"role": role, "content": "\n".join(textos)})

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
