# Project-ResumeGenerator

Este repositório contém os scripts para geração do currículo do Rafael Torres Nantes e extração de repositórios GitHub.

## Comandos Úteis

- Gerar o currículo: `python main.py resume`
- Buscar repositórios: `python main.py repos --limit 150`

## Estrutura de Documentação

- `docs/current-state/`: Descreve o estado atual dos geradores de currículo e exportação de repositórios.
- `docs/planning/`: Próximas features e integrações (ex: exportação JSON de currículo).
- `docs/implementation/`: Detalhes sobre as decisões técnicas, como o uso de win32com para converter DOCX para PDF de maneira local.
