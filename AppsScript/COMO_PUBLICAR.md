# Como Publicar a Web App no Google Apps Script

## Passo 1 — Criar o projeto

1. Vai a **script.google.com**
2. Clica em **"Novo projeto"**
3. Muda o nome do projeto (ex: `Portal Envio Emails`)

---

## Passo 2 — Colar o código

### Ficheiro Code.gs
- O ficheiro `Code.gs` já existe por defeito
- Apaga o conteúdo que lá está
- Cola o conteúdo do ficheiro **`Code.gs`** deste projeto

### Ficheiro Index.html
- Clica em **"+"** ao lado de "Ficheiros" → escolhe **"HTML"**
- Chama-lhe exactamente **`Index`** (sem .html, o Apps Script adiciona sozinho)
- Cola o conteúdo do ficheiro **`Index.html`** deste projeto

---

## Passo 3 — Publicar como Web App

1. Clica em **"Implementar"** (canto superior direito) → **"Nova implementação"**
2. Clica no ícone de engrenagem ⚙️ → escolhe **"Aplicação Web"**
3. Configura assim:
   - **Descrição:** Portal de Envio de Emails
   - **Executar como:** Eu (o teu email da empresa) ← IMPORTANTE
   - **Quem tem acesso:** Qualquer pessoa da [nome da empresa] ← só colegas
4. Clica em **"Implementar"**
5. **Copia o URL** que aparece — é esse o link para partilhar!

---

## Notas Importantes

- **Sem App Password:** Os emails são enviados com a tua conta Google directamente. Não precisas de App Password.
- **Confidencialidade:** Tudo corre dentro do Google Workspace da empresa. Nenhum dado sai para servidores externos.
- **Limites Gmail:** O Google Workspace permite até ~1500 emails/dia. A app já tem pausa de 1 segundo entre envios para respeitar os limites.
- **Actualizar código:** Se modificares o código, clica em "Implementar" → "Gerir implementações" → editar → "Nova versão".
- **Anexos via Drive:** Para usar a opção Google Drive, a pasta tem de ser partilhada com o email que executa a app (ou ser do mesmo utilizador).
