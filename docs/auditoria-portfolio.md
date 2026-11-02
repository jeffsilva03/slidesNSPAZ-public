# Auditoria da versão pública — Slides NSPAZ

A cópia pública foi exportada somente com arquivos versionados revisados, sem copiar `.git`, histórico antigo, ambientes locais, tokens ou dados pessoais. Foi inicializado um repositório independente com um único commit raiz; não é um fork.

O repositório e a hospedagem de produção continuam separados. Esta revisão não acessou bancos pessoais nem alterou deploys, domínios ou credenciais de produção.

## Conteúdo público

- Somente as duas imagens usadas pela interface, sem GPS identificado na revisão, e os SVGs de documentação.
- Sem modelos PPTX privados, galeria ou fotos removidas, caches Python, uploads e apresentações geradas.
- A pasta estática do Flask serve somente `assets/images/`; modelos e arquivos de trabalho não são servidos por esse roteamento.
- Código, testes e arquivos de hospedagem preservados. O nome de exemplo do serviço Render foi diferenciado para a versão pública.

## Limites de segurança

O app continua sem autenticação e isolamento por usuário, com CORS amplo e rotas de diagnóstico. A revisão não certifica uma instância para exposição pública, direitos autorais de todas as mídias ou disponibilidade dos serviços externos. Use ambiente local ou configure controles apropriados antes de hospedar.
