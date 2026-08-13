"""Serviço para busca e exportação de repositórios GitHub."""

import json
import logging
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Union

logger = logging.getLogger(__name__)


class GithubRepoFetcher:
    """Busca repositórios via GitHub CLI e exporta para JSON."""

    def __init__(self, limit: int = 100) -> None:
        """Inicializa o fetcher.

        Args:
            limit: Número máximo de repositórios a retornar por padrão.
        """
        self._limit = limit

    def run_gh(self, args: List[str]) -> Union[Dict, List]:
        """Executa um subcomando do gh CLI e retorna o JSON parseado.

        Args:
            args: Argumentos passados ao gh após o binário.

        Returns:
            Objeto Python resultante do parse do JSON retornado pelo gh.

        Raises:
            subprocess.CalledProcessError: Se o gh retornar código de saída não-zero.
        """
        logger.info("Executando gh %s", " ".join(args))
        result = subprocess.run(
            ["gh", *args],
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=True,
        )
        return json.loads(result.stdout)

    def fetch_repos(self, user: Optional[str] = None, limit: Optional[int] = None) -> List[Dict]:
        """Busca repositórios de um usuário via gh CLI.

        Args:
            user: Nome do usuário GitHub; se None, usa o usuário autenticado.
            limit: Número máximo de repositórios a retornar.

        Returns:
            Lista de dicionários com metadados de cada repositório.
        """
        effective_limit = limit or self._limit
        base_args = [
            "repo", "list", "--limit", str(effective_limit), "--json",
            "name,description,url,primaryLanguage,isPrivate,isFork,stargazerCount,updatedAt",
        ]
        
        if user:
            base_args.insert(2, user)
            
        return self.run_gh(base_args)

    def fetch_current_user(self) -> str:
        """Retorna o login do usuário autenticado no gh CLI.

        Returns:
            Nome de usuário GitHub autenticado.
        """
        result = subprocess.run(
            ["gh", "api", "user", "--jq", ".login"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=True,
        )
        return result.stdout.strip()

    def build_output(self, repos: List[Dict], user: str) -> Dict:
        """Constrói o dicionário de saída normalizado.

        Args:
            repos: Lista de repositórios no formato retornado pelo gh CLI.
            user: Nome do usuário dono dos repositórios.

        Returns:
            Dicionário estruturado para exportação.
        """
        return {
            "github_user": user,
            "total_repos": len(repos),
            "repositories": [
                {
                    "name": r.get("name", ""),
                    "description": r.get("description") or "",
                    "url": r.get("url", ""),
                    "language": (r.get("primaryLanguage") or {}).get("name", ""),
                    "private": r.get("isPrivate", False),
                    "fork": r.get("isFork", False),
                    "stars": r.get("stargazerCount", 0),
                    "updated_at": r.get("updatedAt", ""),
                }
                for r in repos
            ],
        }

    def export_repos(self, output_path: Path, user: Optional[str] = None, no_forks: bool = False, public_only: bool = False) -> None:
        """Orquestra a busca, filtragem e exportação dos repositórios.

        Args:
            output_path: Caminho onde salvar o arquivo JSON.
            user: Nome do usuário (None para usuário atual).
            no_forks: Se verdadeiro, ignora repositórios do tipo fork.
            public_only: Se verdadeiro, ignora repositórios privados.
        """
        target_user = user or self.fetch_current_user()
        logger.info("Buscando repositórios do usuário: %s", target_user)
        
        repos = self.fetch_repos(user=target_user)
        
        if no_forks:
            repos = [r for r in repos if not r.get("isFork")]
        if public_only:
            repos = [r for r in repos if not r.get("isPrivate")]
            
        output = self.build_output(repos, target_user)
        
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("Exportados %d repositórios -> %s", output["total_repos"], output_path.resolve())


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    test_out = Path(__file__).resolve().parent.parent / "output" / "repos_teste_local.json"
    logger.info("Executando teste local de GithubRepoFetcher...")
    fetcher = GithubRepoFetcher(limit=5)
    fetcher.export_repos(output_path=test_out)
    logger.info("Teste concluído. Verifique o arquivo gerado em: %s", test_out)
