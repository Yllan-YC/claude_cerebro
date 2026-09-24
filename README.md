# claude_cerebro

Meu segundo cérebro: um vault de notas em Markdown, organizado no método **PARA**
e pensado para funcionar tanto no [Obsidian](https://obsidian.md) quanto com o Claude.

## Estrutura

| Pasta | Para quê |
|---|---|
| `00-Inbox/` | Captura rápida. Tudo entra aqui primeiro e é triado depois. |
| `01-Diario/` | Notas diárias (`AAAA-MM-DD.md`). |
| `02-Projetos/` | Esforços com objetivo e prazo definidos. |
| `03-Areas/` | Responsabilidades contínuas, sem data de término (saúde, finanças, carreira…). |
| `04-Recursos/` | Referências e temas de interesse (livros, cursos, artigos, conceitos). |
| `05-Arquivo/` | Itens inativos de qualquer outra pasta. |
| `Templates/` | Modelos de nota. |

## Fluxo

1. **Capturar** no `00-Inbox/` sem se preocupar com organização.
2. **Triar** periodicamente: mover cada nota para Projetos, Áreas ou Recursos (ou apagar).
3. **Conectar** notas com links `[[Nome da nota]]` e tags.
4. **Arquivar** o que deixou de estar ativo em `05-Arquivo/`.

## Convenções

- Nomes de arquivo em português, legíveis, sem caracteres especiais problemáticos (`/ \ : * ? " < > |`).
- Toda nota começa com frontmatter YAML (veja `Templates/`).
- Links internos no formato wiki: `[[Nota]]`.
