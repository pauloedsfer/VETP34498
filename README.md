# Controlados — Fórmula Animal

Site para geração automática da planilha de controlados veterinários
a partir dos relatórios exportados do Farma Fácil.

## Como fazer o deploy no Vercel (uma vez só)

### Opção A — Via GitHub (recomendado)

1. Crie uma conta grátis em https://github.com e em https://vercel.com
2. No GitHub, crie um repositório novo (ex: `controlados-formula-animal`)
3. Faça o upload de todos os arquivos desta pasta para o repositório
4. No Vercel, clique em "Add New Project" e selecione o repositório
5. Clique em "Deploy" — o site fica no ar em ~30 segundos
6. Você receberá um link como `https://controlados-formula-animal.vercel.app`

### Opção B — Via Vercel CLI

1. Instale o Node.js (https://nodejs.org)
2. No terminal, dentro desta pasta:
   ```
   npm install -g vercel
   vercel login
   vercel --prod
   ```

## Estrutura do projeto

```
controlados-vercel/
├── vercel.json          ← configuração do Vercel
├── README.md            ← este arquivo
└── public/
    ├── index.html       ← página principal
    └── app.js           ← toda a lógica de processamento
```

## Como usar o site

1. Acesse o link do site no Vercel
2. Exporte do Farma Fácil:
   - **MOVIMENTO.XLS** → Movimento Controlados Período
   - **CLIENTE_END.XLS** → Registro de Receituário Geral
3. Faça o upload dos dois arquivos nas zonas indicadas
4. Informe o estoque inicial do período (em gramas)
5. Clique em **Gerar Planilha**
6. Baixe o arquivo `Controlados_FormulaAnimal.xlsx`

## Abas geradas

| Aba | Conteúdo |
|-----|----------|
| `RESUMO` | Totais por substância para o Anexo IX (MAPA 837/2025) |
| `CONTROLE` | Banco de dados completo com todos os campos cruzados |
| `CORPO_*` | Livro de Registro por substância (Anexo IV) |
| `FICHAS_IMPRIMIR` | Fichas individuais para colar no livro físico |

## Privacidade

Todo o processamento é feito **localmente no navegador**.
Nenhum dado é enviado a servidores externos.

## Responsável Técnico

Paulo Edson Fernandes — CRF 9303  
R S O Manipulação Animal — GO 0198-8
