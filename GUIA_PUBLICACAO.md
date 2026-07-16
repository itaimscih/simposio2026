# Guia de Publicacao — II Simposio IRAS & Stewardship

## 1. Criar Google Forms (formulario de inscricao)

> Google Forms nativo = mais seguro. Sem codigo exposto, sem script intermediario.
> Os dados ficam armazenados diretamente no Google Sheets, gerenciados pelo Google.

### 1.1 Criar formulario

1. Acesse forms.google.com > Formulario em branco
2. Titulo: `Inscricao — II Simposio de Prevencao de IRAS e Stewardship`
3. Descricao: `Evento gratuito. Vagas limitadas (100). 14 e 15 de agosto de 2026.`

Campos (todos obrigatorios):
```
1. Nome completo       — Resposta curta
2. E-mail              — Resposta curta (ativar validacao de e-mail)
3. Profissao           — Multipla escolha:
                          Medico(a) / Enfermeiro(a) / Farmaceutico(a) / Biomedico(a) / Outro
4. Instituicao         — Resposta curta
5. CPF                 — Resposta curta
6. Conselho de Classe  — Resposta curta
```

### 1.2 Configuracoes

- ⚙️ Configuracoes > **Coletar e-mails** (ativar)
- ⚙️ Configuracoes > **Limitar a 1 resposta** (ativar — requer login Google)
- ⚙️ Apresentacao > **Mostrar barra de progresso** (ativar)

### 1.3 Vincular ao Google Sheets

1. Clique na aba **Respostas**
2. Clique em **Criar planilha** (icone verde do Sheets)
3. Nome: `Inscricoes_II_Simposio_IRAS`

### 1.4 Obter link de incorporacao

1. Clique em **Enviar** (canto superior direito)
2. Clique no icone **< >** (incorporar)
3. Copie o atributo `src` do iframe (apenas a URL entre aspas)
4. Exemplo: `https://docs.google.com/forms/d/e/1FAIpQLS.../viewform?embedded=true`

### 1.5 Ativar no index.html

1. Abra `index.html`
2. Encontre a linha:
   ```html
   src="https://docs.google.com/forms/d/e/SEU-FORM-ID-AQUI/viewform?embedded=true"
   ```
3. Substitua `SEU-FORM-ID-AQUI` pelo ID real do seu formulario

---

## 2. GitHub Pages (hospedagem gratuita)

### 2.1 Criar repositorio no GitHub

1. Acesse github.com/new
2. Nome: `ii-simposio-iras-stewardship-2026`
3. Visibilidade: **Publico**
4. NAO marcar "Add a README file"

### 2.2 Enviar os arquivos

```bash
cd "PaginaInscricao"
git init
git add index.html
git commit -m "Landing page de inscricoes"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/ii-simposio-iras-stewardship-2026.git
git push -u origin main
```

### 2.3 Ativar GitHub Pages

1. No repositorio: Settings > Pages
2. Source: **Deploy from a branch**
3. Branch: **main**, pasta: **/ (root)**
4. Clicar **Save**
5. Aguardar 1-2 minutos. A URL sera:
   ```
   https://SEU-USUARIO.github.io/ii-simposio-iras-stewardship-2026
   ```

---

## 3. Atualizar QR Code

Edite o script `criar_qrcode_e_cards.py`:
```python
LANDING_URL = "https://SEU-USUARIO.github.io/ii-simposio-iras-stewardship-2026"
```

Execute:
```bash
python criar_qrcode_e_cards.py
```

Isso regera o QR Code e os cards de inscricao com a URL real.

---

## URLs finais

| Recurso | URL |
|---|---|
| Landing page | `https://itaimscih.github.io/simposio2026` |
| Google Forms | `https://docs.google.com/forms/d/SEU-FORM-ID/viewform` |
| Google Sheets | `https://docs.google.com/spreadsheets/d/SEU-SHEET-ID/edit` |
