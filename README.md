# Project-ResumeGenerator

## 👨‍💻 Projeto desenvolvido por:
[Rafael Torres Nantes](https://github.com/rafael-torres-nantes)

## Índice

* [📚 Contextualização do projeto](#-contextualização-do-projeto)
* [🛠️ Tecnologias/Ferramentas utilizadas](#%EF%B8%8F-tecnologiasferramentas-utilizadas)
* [🖥️ Funcionamento do sistema](#%EF%B8%8F-funcionamento-do-sistema)
   * [🧩 Parte 1 - Backend](#parte-1---backend)
* [🔀 Arquitetura da aplicação](#arquitetura-da-aplicação)
* [📁 Estrutura do projeto](#estrutura-do-projeto)
* [📌 Como executar o projeto](#como-executar-o-projeto)
* [🕵️ Dificuldades Encontradas](#%EF%B8%8F-dificuldades-encontradas)

## 📚 Contextualização do projeto

O **Project-ResumeGenerator** é uma ferramenta de automação para a geração do currículo de Rafael e busca de seus repositórios no GitHub. O projeto resolve a necessidade de centralizar as descrições de experiência e formação e **exportar rapidamente em DOCX e PDF** o currículo usando templates nativos com **Python**.

## 🛠️ Tecnologias/Ferramentas utilizadas

[<img src="https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=white">](https://www.python.org/)
[<img src="https://img.shields.io/badge/GitHub-181717?logo=github&logoColor=white">](https://github.com/)

## 🖥️ Funcionamento do sistema

### 🧩 Parte 1 - Backend

O sistema é construído inteiramente em Python e utiliza a estrutura orientada a objetos orientada pelo estilo rafael-coding-style.

* **Controller**: `main.py` — Ponto de entrada CLI (entrypoint) responsável por parsear os comandos (`resume` e `repos`).
* **Serviços**: `services/curriculum_service.py` — Constrói o layout do documento DOCX usando `python-docx` e utiliza `win32com.client` para exportar para PDF nativamente.
* **Serviços**: `services/github_service.py` — Wrapper em torno do GitHub CLI (`gh`) para buscar os repositórios públicos de forma autenticada e exportar o resultado final em JSON normalizado.

## 🔀 Arquitetura da aplicação

A aplicação consiste de um simples pipeline acionado via linha de comando (`main.py`). Se o comando `resume` for passado, invoca a classe `CurriculumGenerator` para renderizar o currículo. Se for `repos`, utiliza `GithubRepoFetcher` instanciando o processo `gh` e gravando em `output/`. Todo fluxo de I/O é despejado no diretório de log/saída padrão configurado.

## 📁 Estrutura do projeto

```
.
├── docs/
│   ├── current-state/
│   ├── implementation/
│   └── planning/
├── services/
│   ├── curriculum_service.py
│   └── github_service.py
├── output/
│   ├── curriculo___rafael_torres_nantes.docx
│   ├── curriculo___rafael_torres_nantes.pdf
│   └── repos.json
├── CLAUDE.md
├── main.py
└── README.md
```

## 📌 Como executar o projeto

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/rafael-torres-nantes/Project-ResumeGenerator.git
   ```

2. **Instale as dependências:**
   ```bash
   pip install python-docx pywin32
   ```

3. **Inicie a aplicação:**
   ```bash
   # Para gerar o currículo (DOCX e PDF) em /output
   python main.py resume

   # Para gerar os repositórios em /output/repos.json
   python main.py repos --limit 150
   ```

## 🕵️ Dificuldades Encontradas

- **Conversão nativa DOCX para PDF:** Em vez de usar APIs de terceiros que adicionam marcas d'água ou quebram o layout restrito de uma página, foi implementado localmente o pacote nativo Windows COM (`win32com.client`) que permite ao Office instalado rodar a renderização com fidelidade absoluta.
