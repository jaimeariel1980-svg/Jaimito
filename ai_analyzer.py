import json
import re

import anthropic

from config import Config

SYSTEM_PROMPT = """Eres un analista de Recursos Humanos experto en evaluación de candidatos.
Tu tarea es analizar el currículum de un candidato y compararlo con el perfil del puesto requerido.

Debes responder ÚNICAMENTE con un objeto JSON válido con esta estructura exacta:
{
  "candidate_name": "Nombre completo del candidato",
  "candidate_email": "email@ejemplo.com",
  "candidate_phone": "+54 9 11 1234-5678",
  "years_experience": 5,
  "key_skills": ["habilidad1", "habilidad2", "habilidad3"],
  "education": "Título universitario, institución",
  "score": 7,
  "applies": true,
  "strengths": ["fortaleza 1", "fortaleza 2"],
  "weaknesses": ["debilidad 1", "debilidad 2"],
  "summary": "Resumen conciso en 2-3 oraciones del perfil del candidato y por qué aplica o no."
}

Criterios de evaluación:
- score: del 1 al 10 (1=no cumple nada, 10=perfil ideal)
- applies: true si el score es >= 6 Y cumple los requisitos excluyentes
- Si no encontrás un dato en el CV, usá null para ese campo
- Las habilidades clave deben ser las más relevantes para el puesto (máximo 8)
- El resumen debe ser en español y objetivo

Responde SOLO con el JSON, sin texto adicional ni bloques de código."""


class AIAnalyzer:
    def __init__(self, config: Config):
        self.client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
        self.model = config.CLAUDE_MODEL

    def analyze(self, cv_text: str, job_profile: dict) -> dict:
        """Analiza el CV contra el perfil del puesto usando Claude con prompt caching."""
        job_profile_text = self._format_job_profile(job_profile)

        response = self.client.messages.create(
            model=self.model,
            max_tokens=1024,
            system=[
                {
                    "type": "text",
                    "text": SYSTEM_PROMPT,
                    # El system prompt se cachea: es igual para todos los candidatos del mismo proceso
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": f"PERFIL DEL PUESTO:\n\n{job_profile_text}",
                            # El perfil del puesto se cachea porque es el mismo para todos los candidatos
                            "cache_control": {"type": "ephemeral"},
                        },
                        {
                            "type": "text",
                            "text": f"CV DEL CANDIDATO:\n\n{cv_text}",
                        },
                    ],
                }
            ],
        )

        return self._parse_response(response.content[0].text)

    def _format_job_profile(self, profile: dict) -> str:
        lines = [
            f"Puesto: {profile.get('title', '')}",
            f"Área: {profile.get('department', '')}",
            f"Descripción: {profile.get('description', '')}",
            f"Experiencia requerida: {profile.get('required_experience_years', 'No especificada')} años",
            f"Educación requerida: {profile.get('education_required', 'No especificada')}",
        ]

        if profile.get("required_skills"):
            lines.append("Habilidades requeridas (excluyentes):")
            for skill in profile["required_skills"]:
                lines.append(f"  - {skill}")

        if profile.get("nice_to_have_skills"):
            lines.append("Habilidades deseables:")
            for skill in profile["nice_to_have_skills"]:
                lines.append(f"  - {skill}")

        if profile.get("responsibilities"):
            lines.append("Responsabilidades principales:")
            for resp in profile["responsibilities"]:
                lines.append(f"  - {resp}")

        if profile.get("disqualifiers"):
            lines.append("Criterios excluyentes (descalifican al candidato):")
            for d in profile["disqualifiers"]:
                lines.append(f"  - {d}")

        return "\n".join(lines)

    def _parse_response(self, text: str) -> dict:
        text = text.strip()

        # Extrae JSON si viene envuelto en bloques de código
        match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
        if match:
            text = match.group(1)

        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            # Fallback si el modelo no devuelve JSON válido
            data = {
                "candidate_name": None,
                "candidate_email": None,
                "candidate_phone": None,
                "years_experience": None,
                "key_skills": [],
                "education": None,
                "score": 0,
                "applies": False,
                "strengths": [],
                "weaknesses": [],
                "summary": f"Error al parsear respuesta de IA: {text[:200]}",
            }

        # Normalización de tipos
        data["score"] = int(data.get("score") or 0)
        data["applies"] = bool(data.get("applies", False))
        data["key_skills"] = data.get("key_skills") or []
        data["strengths"] = data.get("strengths") or []
        data["weaknesses"] = data.get("weaknesses") or []

        return data
