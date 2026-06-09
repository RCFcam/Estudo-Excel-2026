# Estudo-Excel-2026
Análise e estudo no uso de Excel em diversos ambientes.
# 🛠️ Documentação Técnica do Projeto: Simulador de Investimentos

Este projeto aplica conceitos avançados de **Spreadsheet Engineering** para transformar dados brutos em uma ferramenta de simulação financeira dinâmica, auditável e visualmente profissional, utilizando o Microsoft Excel como motor de cálculo corporativo.

---

## 📌 Índice
1. [Arquitetura e Engenharia de Dados](#-arquitetura-e-engenharia-de-dados)
2. [Módulos de Simulação e Negócio](#-módulos-de-simulação-e-negócio)
3. [UI/UX e Uniformidade Visual](#-uiux-e-uniformidade-visual)
4. [Stack Técnico Utilizado](#-stack-técnico-utilizado)

---

## 📐 Arquitetura e Engenharia de Dados

### 1. Separação de Camadas (Data Isolation)
O projeto segue o padrão rigoroso de arquitetura em três camadas para garantir a integridade dos cálculos e a escalabilidade do modelo:
* **Camada de Dados (Inputs):** Concentrada na aba `tab-auxiliar`, isolando matrizes de alíquotas de impostos, parâmetros de fundos e taxas de mercado.
* **Camada de Processamento/Cenários:** Localizada na aba `Proj-Investimento`, onde as variáveis globais interagem com as fórmulas dinâmicas ao longo da linha do tempo.
* **Camada de Apresentação (UI):** Formatação limpa orientada a *dashboards*, ocultando a complexidade das fórmulas por trás de uma interface amigável.

### 2. Governança e Variáveis Globais
* **Eliminação de Hardcoding:** Valores fixos (como alíquotas de IR por tempo de aplicação ou taxas de administração) foram completamente extraídos de dentro das fórmulas e centralizados na aba de parâmetros.
* **Nomeando Intervalos:** Uso de intervalos nomeados no Excel para transformar fórmulas crípticas em lógica de negócio legível e auditável. 
  > **Exemplo Prático:** Em vez de utilizar `=B12 * 'tab-auxiliar'!$C$4`, o modelo adota a sintaxe `=Aporte_Inicial * Taxa_Rendimento`.

---

## 📊 Módulos de Simulação e Negócio

### 3. Simulador de Patrimônio e Cenários
A lógica financeira foi projetada especificamente para responder dinamicamente a **Perguntas de Negócio** essenciais ("E se...?"):
* **Projeção Base:** Estruturação cronológica indexando aportes recorrentes, juros compostos e custos de gestão de fundos.
* **Análise de Sensibilidade:** Modelagem preparada para alternar variáveis macroeconômicas (taxas de juros, inflação) e tipos de fundos para avaliar instantaneamente o impacto no montante líquido final.

---

## 🎨 UI/UX e Uniformidade Visual

### 4. Design de Interface Corporativa
A estética do modelo foi planejada para mitigar a fadiga cognitiva do usuário e transmitir confiança técnica na apresentação dos dados:

| Técnica Aplicada | Objetivo Prático |
| :--- | :--- |
| **Tipos de Fundo Coesos** | Uso de uma paleta de cores escura/grafite para destacar elementos de navegação e blocos de dados essenciais. |
| **Hierarquia Visual** | Títulos marcantes, fontes padronizadas de leitura limpa e tamanhos estritos para diferenciar inputs (células de digitação) de outputs (fórmulas). |
| **Linhas de Grade Ocultas** | Remoção das linhas de grade padrão do Excel para dar ao projeto o aspecto visual de um software ou sistema dedicado, e não de uma planilha comum. |

---

## 🚀 Stack Técnico Utilizado
* **Estruturas de Busca:** `XLOOKUP` / `VLOOKUP` para cruzamento de dados automatizado entre as tabelas auxiliares.
* **Funções Lógicas e Financeiras:** `IF`, `AND` e `VF` (Valor Futuro) para simulações precisas de juros compostos acumulados.
* **Recursos Nativos:** Gerenciador de Nomes, Validação de Dados (listas suspensas dinâmicas) e Tabelas Estruturadas (`Ctrl + ALT + T`).
