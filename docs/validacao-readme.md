# Validação da documentação — Slides NSPAZ

Verificações da base revisada em 02/11/2026, Windows, Python 3.14.3 em ambiente virtual isolado.

| Verificação | Resultado |
| :-- | :-- |
| Instalação de `requirements.txt` | Concluída |
| `python -m unittest discover -s tests -v` | 15 testes passaram: 13 de geração e 2 de isolamento de arquivos |
| Geração e download de PPTX | Cobertos pela suíte, com e sem modelo sintético |
| Página inicial e imagens referenciadas | HTTP 200 nos testes |
| Arquivos da raiz / `.git` / `.env` / uploads / saídas pelo acesso estático | HTTP 404 nos testes |
| Links locais e âncoras do README / SVGs | Conferidos por script e prévia local |

Não foram executados backend PHP, Gunicorn/Render, sequência Linux/macOS, revisão completa dos 404 slides históricos ou teste de disponibilidade da API de liturgia externa. As ilustrações são diagramas do fluxo, não capturas de execução da conta de alguém.

## Exportação pública

A versão pública foi criada em histórico novo. Nomes, links e guias foram adaptados; fixtures do Classroom usam identificadores fictícios. Verificações desta exportação e seus limites constam em `docs/versao-publica.md`. Os resultados anteriores não certificam deploy, OAuth, notificações ou E2E nesta cópia.

Nesta cópia pública, a suíte foi executada novamente: 15 testes passaram. Os links locais e SVGs foram conferidos após a adaptação dos nomes.
