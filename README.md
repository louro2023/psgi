# PGSI - Sistema Web de Gestão de Segurança da Informação

Sistema Flask/SQLite para gestão de pessoas, perfis, controles, riscos, incidentes, ações, documentos, indicadores, auditoria e relatórios.

## Executar

```powershell
python -m flask --app app run --host 127.0.0.1 --port 5000
```

Ao iniciar, o banco `instance/pgsi.sqlite3` é criado automaticamente com perfis, usuário administrador e controles iniciais baseados no PGSI.

## Acesso inicial

- E-mail: `admin@pgsi.local`
- Senha: `Admin@123`

## Principais recursos

- Login com senha criptografada e controle de sessão.
- Perfis e permissões por módulo.
- CRUD de pessoas, usuários, controles, riscos, incidentes e plano de ação.
- Área central de Controles com subabas para controles, riscos, ações, evidências e melhorias.
- Mapa PGSI integrado à visão de Controles, com 16 ambientes personalizados por eixo, abas internas, indicadores, histórico e upload contextualizado.
- Inventário próprio em Gestão de Ativos para registrar itens, responsáveis, criticidade, status, localização e revisões.
- Cálculo automático do nível de risco por probabilidade x impacto.
- Upload protegido de evidências e documentos com classificação.
- Extração de texto para PDF, Excel, CSV e TXT.
- Importação de controles a partir de planilhas com coluna de título/controle.
- Dashboard com indicadores, gráficos, pontos de atenção e melhorias sugeridas.
- Relatórios executivo, geral, controles, riscos, incidentes, melhorias e evidências.
- Exportação em PDF, Excel e CSV.
- Logs de auditoria e backup local do banco.
