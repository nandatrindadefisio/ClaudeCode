# CLAUDE.md — Instruções do Projeto

## Repositório GitHub

Este projeto é sincronizado automaticamente com o GitHub:
- **URL**: https://github.com/nandatrindadefisio/ClaudeCode
- **Usuário**: nandatrindadefisio

## Auto-push automático

Um hook `post-commit` está configurado em `.git/hooks/post-commit`.
**A cada commit feito localmente, o código é enviado automaticamente para o GitHub.**

Não é necessário rodar `git push` manualmente.

## Como commitar alterações

Após fazer alterações no projeto, use:

```bash
git add .
git commit -m "descrição das alterações"
# o push para o GitHub acontece automaticamente
```

## Configuração do remote

O remote `origin` aponta para:
```
https://github.com/nandatrindadefisio/ClaudeCode.git
```

As credenciais estão embutidas na URL do remote com o Personal Access Token do GitHub.

## Instruções para o Claude

- Sempre que fizer alterações no projeto, criar um commit e fazer push para o GitHub.
- O hook post-commit cuida do push automaticamente após cada commit.
- Manter o repositório GitHub sempre atualizado com as últimas alterações.
