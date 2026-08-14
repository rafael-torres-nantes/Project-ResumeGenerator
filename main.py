"""Entrypoint para o Project-ResumeGenerator."""

import argparse
import logging
from pathlib import Path

from services.curriculum_service import CurriculumGenerator
from services.github_service import GithubRepoFetcher
from services.evaluator_service import CurriculumEvaluator

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


def main() -> None:
    """Função principal que orquestra a geração de currículo e exportação de repos."""
    parser = argparse.ArgumentParser(description="Resume Generator & GitHub Repo Fetcher")
    
    subparsers = parser.add_subparsers(dest="command", help="Comandos disponíveis", required=True)
    
    # Comando de currículo
    parser_resume = subparsers.add_parser("resume", help="Gera o currículo em DOCX e PDF")
    
    # Comando de github
    parser_github = subparsers.add_parser("repos", help="Busca repositórios do GitHub e salva em JSON")
    parser_github.add_argument("--user", "-u", help="GitHub username (default: authenticated user)")
    parser_github.add_argument("--limit", "-l", type=int, default=100, help="Max repos (default: 100)")
    parser_github.add_argument("--output", "-o", default="output/repos.json", help="Output file (default: output/repos.json)")
    parser_github.add_argument("--no-forks", action="store_true", help="Exclude forked repositories")
    parser_github.add_argument("--public-only", action="store_true", help="Exclude private repositories")
    
    # Comando de avaliador headless
    parser_eval = subparsers.add_parser("evaluate", help="Audita o currículo usando os prompts template")
    parser_eval.add_argument("--prompt", "-p", default="01_auditoria_curriculo.md", help="Nome do arquivo de prompt na pasta prompts_template")
    parser_eval.add_argument("--provider", choices=["gemini", "claude"], default="gemini", help="Provedor de IA (gemini ou claude)")
    
    args = parser.parse_args()
    
    base_dir = Path(__file__).resolve().parent
    
    if args.command == "resume":
        try:
            output_dir = base_dir / "output"
            data_path = base_dir / "personal_info" / "curriculum_data.json"
            generator = CurriculumGenerator(output_dir=output_dir, data_path=data_path)
            generator.build(filename_base="curriculo___rafael_torres_nantes")
        except Exception as e:
            logger.error("Erro na geração de currículo: %s", e)
            
    elif args.command == "evaluate":
        try:
            data_path = base_dir / "personal_info" / "curriculum_data.json"
            prompts_dir = base_dir / "prompts_template"
            evaluator = CurriculumEvaluator(data_path=data_path, prompts_dir=prompts_dir)
            
            logger.info("Iniciando avaliação com o prompt: %s via %s", args.prompt, args.provider)
            result = evaluator.evaluate(args.prompt, provider=args.provider)
            if result:
                out_file = base_dir / "output" / f"avaliacao_{args.prompt.replace('.md', '')}_{args.provider}.md"
                out_file.parent.mkdir(parents=True, exist_ok=True)
                out_file.write_text(result, encoding="utf-8")
                logger.info("Avaliação salva em: %s", out_file)
        except Exception as e:
            logger.error("Erro na avaliação: %s", e)
            
    elif args.command == "repos":
        try:
            fetcher = GithubRepoFetcher(limit=args.limit)
            output_path = Path(args.output)
            fetcher.export_repos(
                output_path=output_path,
                user=args.user,
                no_forks=args.no_forks,
                public_only=args.public_only
            )
        except Exception as e:
            logger.error("Erro ao buscar repositórios: %s", e)


if __name__ == "__main__":
    main()
