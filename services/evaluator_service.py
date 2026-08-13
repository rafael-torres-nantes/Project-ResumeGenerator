"""Serviço headless para auditar e avaliar o currículo via IA."""

import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)


class CurriculumEvaluator:
    """Carrega o currículo e usa templates de prompt para auditá-lo via Gemini."""

    def __init__(self, data_path: Path, prompts_dir: Path) -> None:
        """Inicializa o avaliador.

        Args:
            data_path: Caminho para o JSON de dados do currículo.
            prompts_dir: Diretório contendo os templates de prompt.
        """
        self._data_path = data_path
        self._prompts_dir = prompts_dir

    def _read_curriculum_text(self) -> str:
        """Lê o currículo em formato de string plana do JSON."""
        if not self._data_path.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {self._data_path}")
        return self._data_path.read_text(encoding="utf-8")

    def _read_prompt(self, prompt_name: str) -> str:
        """Lê o conteúdo de um template de prompt."""
        prompt_path = self._prompts_dir / prompt_name
        if not prompt_path.exists():
            raise FileNotFoundError(f"Template não encontrado: {prompt_path}")
        return prompt_path.read_text(encoding="utf-8")

    def evaluate(self, prompt_filename: str) -> str:
        """Avalia o currículo usando um template específico.

        Requer a variável de ambiente GEMINI_API_KEY.

        Args:
            prompt_filename: Nome do arquivo de template (ex: 01_auditoria_curriculo.md).

        Returns:
            A resposta da IA com as sugestões.
        """
        try:
            from google import genai
            from google.genai import types
        except ImportError:
            logger.error("Pacote 'google-genai' não encontrado. Instale com: pip install google-genai")
            return ""

        if "GEMINI_API_KEY" not in os.environ:
            logger.error("Variável GEMINI_API_KEY não configurada no ambiente.")
            return ""

        curriculum_text = self._read_curriculum_text()
        prompt_text = self._read_prompt(prompt_filename)

        full_prompt = (
            f"{prompt_text}\n\n"
            f"--- INÍCIO DO CURRÍCULO ---\n{curriculum_text}\n--- FIM DO CURRÍCULO ---"
        )

        logger.info("Enviando requisição ao Gemini (gemini-2.5-pro)...")
        client = genai.Client()
        response = client.models.generate_content(
            model="gemini-2.5-pro",
            contents=full_prompt,
        )
        return response.text


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    base_dir = Path(__file__).resolve().parent.parent
    data = base_dir / "personal_info" / "curriculum_data.json"
    prompts = base_dir / "prompts_template"
    
    logger.info("Executando teste local de CurriculumEvaluator...")
    evaluator = CurriculumEvaluator(data_path=data, prompts_dir=prompts)
    
    # Simula o fluxo (pode falhar se não houver GEMINI_API_KEY)
    resultado = evaluator.evaluate("01_auditoria_curriculo.md")
    if resultado:
        logger.info("Avaliação concluída:\n%s", resultado)
    else:
        logger.info("Teste concluído sem requisição real.")
