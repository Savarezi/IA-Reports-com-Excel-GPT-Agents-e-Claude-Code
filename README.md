
# IA Reports: Automação de Dados Porsche
<img width="500" height="300" alt="image" src="https://github.com/user-attachments/assets/c33c9593-72c6-41a1-9ab0-026c46386e12" />



Projeto de aceleração voltado para a criação de fluxos de trabalho inteligentes, utilizando a integração de dados em Excel, agentes baseados em GPT e Claude Code para sanitização, análise e visualização de dados.

## 🚀 Objetivo do Projeto
O **IA Reports** visa otimizar o processamento de bases de dados (ex: *Porsche Sales Data*), aplicando regras de negócio automatizadas para transformar dados brutos em informações tratadas, prontas para dashboards e insights estratégicos.
<img width="500" height="200" alt="image" src="https://github.com/user-attachments/assets/124d165b-8909-4402-b304-ccbb4a8a2d18" />


## 🛠️ Tecnologias e Ferramentas
* **Linguagem:** Python
* **Bibliotecas:** Pandas, OpenPyXL
* **Integração:** GPT Agents, Claude Code
* **Ambiente de Desenvolvimento:** GitHub Codespaces

## 📋 Estrutura do Repositório
* `/data`: Contém a `planilha_base_porsche.xlsx` (origem) e a `planilha_base_porsche_sanitized.xlsx` (saída tratada).
* `schema.md`: O "coração" do projeto. Define o mapeamento de colunas, regras de sanitização e validações de qualidade.
* `sanitize_porsche.py`: O agente responsável pela leitura, processamento e formatação visual dos dados.

## ⚙️ Regras de Qualidade e Sanitização
Para garantir a integridade dos dados, seguimos o esquema definido no `schema.md`:
1. **Preservação:** A coluna original nunca é renomeada.
2. **Nova Coluna:** Sempre criamos uma coluna `Sanitized` para comparação.
3. **Padrões:**
   - **Estados (UF):** Apenas siglas (ex: "NY") ou "INVALID".
   - **Preços:** Formato decimal com 2 casas.
   - **Milhagem:** Valores sempre inteiros.
   - **Normalização:** Caso o dado seja ambíguo, o agente retorna "INVALID".

## 🛡️ Boas Práticas (Compliance e Segurança)
* **Amostragem:** Sempre inicie a análise com uma amostra do volume total de dados antes de processar toda a base.
* **Privacidade:** **Nunca** utilize dados sensíveis (PII) em interações com IAs públicas. Sempre crie uma cópia anonimizada (removendo dados pessoais) antes de processar.

## 🔗 Acesso ao Dashboard
Visualize os dados processados e os insights gerados em tempo real através do link abaixo:
[Acessar Dashboard Porsche Sales](https://porschesalesdashboard.netlify.app/)

## 💻 Como rodar o projeto
1. Certifique-se de estar na raiz do projeto.
2. Instale as dependências:
   ```bash
   pip install pandas openpyxl
