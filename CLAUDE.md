# CLAUDE.md

Este repositório é um **segundo cérebro** (vault de notas em Markdown, compatível com Obsidian),
organizado pelo método PARA. Não é um projeto de código.

## Idioma

Escreva notas e respostas em **português do Brasil**, salvo pedido em contrário.

## Estrutura

- `00-Inbox/` — captura rápida; notas sem destino definido vão aqui.
- `01-Diario/` — notas diárias, nome `AAAA-MM-DD.md`, criadas a partir de `Templates/Diario.md`.
- `02-Projetos/` — uma pasta ou nota por projeto ativo (tem objetivo e prazo).
- `03-Areas/` — responsabilidades contínuas.
- `04-Recursos/` — referências, conceitos, leituras.
- `05-Arquivo/` — itens inativos. Nunca apague notas; mova-as para cá.
- `Templates/` — modelos. Use-os ao criar notas novas.

## Regras ao editar notas

- Toda nota nova começa com o frontmatter do template correspondente (`tipo`, `criado`, `tags`, e `status` quando aplicável). Datas em `AAAA-MM-DD`.
- Use links wiki `[[Nome da nota]]` para conectar notas; ao criar uma nota, procure notas relacionadas existentes e adicione links.
- Tags em minúsculas, sem acento, com hífen: `#saude-mental`, `#financas`.
- Nomes de arquivo legíveis, sem `/ \ : * ? " < > |`.
- Não reorganize pastas nem renomeie notas em massa sem pedir confirmação (quebra links).
- Ao triar o Inbox, proponha o destino de cada nota antes de mover.

## Tarefas comuns

- **"Nota do dia"**: criar/abrir `01-Diario/<data de hoje>.md` a partir de `Templates/Diario.md`.
- **"Triar inbox"**: ler cada nota em `00-Inbox/`, sugerir destino (Projetos/Áreas/Recursos/Arquivo) e tags.
- **"Novo projeto X"**: criar `02-Projetos/X.md` a partir de `Templates/Projeto.md`.
- **Buscar conhecimento**: use Grep em todo o vault antes de responder perguntas sobre "o que eu já anotei".
