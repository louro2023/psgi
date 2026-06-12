# PGSI - Sistema de Gestão de Segurança da Informação

## Visão Geral

Flask web app para gestão de controles de segurança, riscos, incidentes, ações, documentos e logs de auditoria. Framework PGSI (16 eixos temáticos de segurança da informação).

**Stack:** Python 3 + Flask + SQLite3 + Jinja2 + Vanilla JS  
**Repo:** https://github.com/louro2023/psgi.git  
**Port padrão:** 5000

---

## Estrutura do Projeto

```
app.py                    # Aplicação Flask monolítica (~3.4k linhas)
requirements.txt          # Flask, pypdf, openpyxl, reportlab
static/
  app.js                  # Charts e interação frontend
  styles.css              # Design system CSS
templates/                # Jinja2 templates (14 arquivos)
  base.html               # Layout base com navegação
  login.html
  dashboard.html
  control_center.html
  area_detail.html
  entity_form.html        # CRUD genérico
  entity_list.html
  documents.html
  reports.html
  logs.html
  backup.html
  error.html
instance/
  pgsi.sqlite3            # DB SQLite (auto-criado no primeiro run)
uploads/                  # Arquivos enviados pelos usuários
backups/                  # Backups do banco de dados
tests/
  test_core.py            # Testes unitários (cálculo de risco)
```

---

## Executar o Projeto

```bash
pip install -r requirements.txt
python -m flask --app app run --host 127.0.0.1 --port 5000
```

Acesso: `http://127.0.0.1:5000`  
Admin padrão: `admin@pgsi.local` / `Admin@123`

---

## Banco de Dados

SQLite3. Auto-inicializado via `init_db()` no primeiro run.

**Tabelas principais:**

| Tabela | Propósito |
|--------|-----------|
| `usuarios` | Credenciais de login |
| `perfis` | Perfis RBAC com permissões JSON |
| `pessoas` | Contatos organizacionais |
| `controles` | Controles de segurança |
| `riscos` | Registro de riscos (P×I) |
| `incidentes` | Incidentes de segurança |
| `acoes` | Planos de ação/remediação |
| `arquivos` | Documentos/evidências enviados |
| `ativos_inventario` | Inventário de ativos |
| `indicadores` | KPIs e métricas |
| `logs_auditoria` | Trilha de auditoria completa |

**Soft-delete:** campo `ativo` em todas as entidades (nunca delete físico).

---

## Autenticação e Autorização

**Auth:** Session-based Flask + Werkzeug PBKDF2 password hash.

**RBAC - 6 perfis padrão:**

| Perfil | Nível |
|--------|-------|
| Administrador | Acesso total |
| Gestor | Gerenciar controles, riscos, ações, relatórios |
| Técnico | Criar/editar controles, riscos, incidentes, evidências |
| Auditor | Somente leitura + evidências e relatórios |
| Consulta | Somente leitura limitada |
| Usuário comum | Upload de documentos apenas |

**Módulos de permissão:** dashboard, usuarios, perfis, pessoas, controles, riscos, incidentes, acoes, arquivos, relatorios, logs, backup

**Decoradores:** `@login_required`, `@permission_required(module, action)`

---

## Rotas Principais

```
GET/POST /login                    # Autenticação
GET      /dashboard                # KPIs e alertas executivos
GET      /controles/painel         # Hub de gestão de controles
GET      /areas                    # 16 áreas de segurança
GET      /areas/<key>              # Detalhe da área (tabs: visao/controles/riscos/acoes/evidencias)
GET/POST /<entity>                 # CRUD genérico (controles, riscos, incidentes, acoes, etc.)
GET/POST /arquivos                 # Upload e listagem de documentos
GET      /arquivos/<id>/download   # Download de arquivo
POST     /arquivos/<id>/analisar   # Extrair texto e importar controles
GET      /relatorios               # Construtor de relatórios
GET      /relatorios/exportar/<fmt># Exportar CSV/XLSX/PDF
GET      /logs                     # Visualizador de auditoria
GET/POST /backup                   # Gerenciamento de backups
```

---

## 16 Eixos PGSI

1. Gestão de Vulnerabilidades
2. Segurança Física e do Ambiente
3. Gestão de Incidentes de SI
4. Gestão de Ativos
5. Infraestrutura de Rede
6. Controle de Acesso
7. Gestão de Software
8. Gestão de Mudanças
9. Fornecedores e Terceiros
10. Continuidade de Negócios
11. Recuperação de Dados
12. Gestão de Riscos
13. Logs de Auditoria
14. Auditoria e Conformidade
15. Conscientização e Treinamento
16. Gestão de Dados

---

## Cálculo de Risco

```
Pontuação = Probabilidade (1-3) × Impacto (1-3)  → max 9
Baixo:  ≤ 2
Médio:  3–5
Alto:   6–9
```

Riscos Altos exigem plano de tratamento preenchido.

---

## Upload de Documentos

- **Tamanho máximo:** 25MB
- **Extensões permitidas:** csv, doc, docx, gif, jpeg, jpg, pdf, png, txt, xls, xlsm, xlsx
- **Nomenclatura:** UUID + extensão original
- **Extração de texto:** PDF (pypdf), Excel (openpyxl), CSV/TXT
- **Cache de texto:** primeiros 50k chars no DB
- **Classificações:** Interno, Restrito, Confidencial, Sensível
- **Vinculação:** controle, risco, incidente, ação, pessoa

---

## Configuração

**Variáveis de ambiente:**

| Variável | Padrão | Descrição |
|----------|--------|-----------|
| `PGSI_SECRET_KEY` | `pgsi-dev-change-me` | Chave secreta Flask |

**ATENÇÃO:** `SECRET_KEY` padrão é insegura para produção. Definir `PGSI_SECRET_KEY` no ambiente.

---

## Exportação de Relatórios

7 tipos: executivo, geral, controles, riscos, incidentes, melhorias, evidencias  
3 formatos: CSV (UTF-8 BOM + delimitador `;`), XLSX, PDF (landscape)

---

## Padrões de Código

- Arquivo monolítico: `app.py` (~3.4k linhas) — toda lógica de rotas, DB e helpers
- Raw SQL com sqlite3 (sem ORM) — usar parâmetros `?` para evitar SQL injection
- `g.user` e `g.permissions` carregados via `@app.before_request`
- Soft-delete: `ativo = 0` (nunca `DELETE` físico em entidades)
- Auditoria: chamar `registrar_log(acao, tabela, id, detalhes)` em todo CRUD
- Templates: herdam de `base.html` via `{% extends %}`

---

## Testes

```bash
python -m pytest tests/
```

Cobertura atual: apenas cálculo de risco (`tests/test_core.py`).

---

## Segurança - Pontos de Atenção

- `PGSI_SECRET_KEY` deve ser definida como variável de ambiente em produção
- Admin padrão `admin@pgsi.local` / `Admin@123` deve ser alterado após deploy
- Sem timeout de sessão configurado — considerar adicionar
- Backups ficam no filesystem local — considerar armazenamento externo
- DB SQLite: adequado para uso institucional, não para alta concorrência
