# 📊 Simulador de Investimentos em Fundos Imobiliários

Projeto desenvolvido como parte de um desafio prático da DIO, com o objetivo de aplicar conceitos de Excel na criação de uma ferramenta simples para simulação de investimentos em Fundos de Investimento Imobiliário (FIIs).

A planilha foi estruturada para permitir a inserção de parâmetros de investimento e apresentar projeções relacionadas ao valor investido, patrimônio acumulado e dividendos mensais.

O projeto também contempla a documentação da solução e sua organização para compartilhamento por meio do GitHub.

---

## 🎯 Sobre o projeto

O projeto consiste na criação de uma ferramenta simples em Excel para simular investimentos em fundos imobiliários.

A planilha permite:

* calcular o valor total investido;
* projetar o patrimônio acumulado;
* estimar dividendos mensais;
* simular diferentes períodos de investimento;
* selecionar um perfil de investimento;
* distribuir o investimento entre diferentes tipos de FIIs.

A solução foi desenvolvida com foco em **simplicidade, organização e aplicação prática dos conceitos apresentados durante o curso**.

O modelo também pode servir como base para futuras expansões e personalizações.

---

# 🎓 Contexto do desafio

Este projeto foi desenvolvido como parte de um laboratório da **DIO**, cujo objetivo é aplicar conceitos de Excel na construção de uma ferramenta prática de simulação de investimentos em fundos imobiliários.

O desafio propõe utilizar conhecimentos relacionados ao funcionamento dos FIIs e às principais variáveis consideradas em uma simulação de investimentos, como:

* valor a investir;
* período de investimento;
* taxa de rendimento;
* patrimônio acumulado;
* dividendos.

A proposta é utilizar o Excel para estruturar esses dados e automatizar os cálculos necessários para apresentar os resultados de forma organizada.

---

# 📚 Objetivos de aprendizagem

O projeto foi desenvolvido com o objetivo de praticar os seguintes conhecimentos:

* criação de ferramentas de simulação em Excel;
* aplicação de cálculos financeiros;
* utilização de fórmulas e funções do Excel;
* utilização de tabelas auxiliares;
* aplicação de regras de negócio;
* utilização de validação de dados;
* organização de informações;
* documentação de processos técnicos;
* utilização do GitHub para compartilhamento do projeto.

---

# 📑 Estrutura da planilha

A solução possui duas abas principais:

### 1. Calculadora de investimentos

É a interface principal da ferramenta.

Nessa aba estão concentrados:

* parâmetros de entrada;
* cálculos financeiros;
* projeções de investimento;
* cenários de diferentes períodos;
* seleção do perfil;
* distribuição dos investimentos.

### 2. Tabela de Apoio perfil

É utilizada como base auxiliar para armazenar as regras de distribuição dos investimentos de acordo com o perfil selecionado.

A tabela relaciona:

* perfil;
* tipo de FII;
* percentual de distribuição;
* chave de pesquisa.

---

# ⚙️ Funcionamento da solução

O funcionamento da planilha pode ser representado pelo seguinte fluxo:

```text
             ENTRADA DE DADOS
                    │
                    ↓
          Parâmetros do investimento
                    │
                    ↓
        Sugestão de investimento
                    │
                    ↓
           Cálculos financeiros
                    │
          ┌─────────┼─────────┐
          ↓         ↓         ↓
       Total     Patrimônio  Dividendos
      investido  acumulado    mensais
          │         │         │
          └─────────┼─────────┘
                    ↓
             Cenários futuros
                    │
                    ↓
            Seleção de perfil
                    │
                    ↓
            Tabela de apoio
                    │
                    ↓
          Distribuição dos FIIs
```

---

# 🧮 Principais cálculos

A planilha utiliza fórmulas do Excel para automatizar os cálculos.

Entre as principais funções utilizadas estão:

### `FV`

Utilizada para calcular o valor futuro dos aportes considerando uma taxa de rendimento e determinado período.

```excel
=FV(taxa_mensal,qtd_anos*12,aporte*-1)
```

### `VLOOKUP`

Utilizada para consultar na tabela auxiliar o percentual correspondente à combinação entre perfil e tipo de FII.

```excel
=VLOOKUP($E$34&"-"&B38,'Tabela de Apoio perfil'!B:E,4,FALSE)
```

### `SUM`

Utilizada para realizar a soma dos valores distribuídos entre as categorias.

```excel
=SUM(E38:E43)
```

---

# 📊 Cenários de investimento

A ferramenta permite visualizar diferentes horizontes de investimento.

A estrutura dos cenários considera:

| Informação           | Descrição                                  |
| -------------------- | ------------------------------------------ |
| Período              | Tempo considerado para a simulação         |
| Total investido      | Soma dos aportes realizados no período     |
| Patrimônio acumulado | Projeção do valor futuro dos investimentos |
| Dividendos mensais   | Estimativa baseada no patrimônio projetado |

Os períodos são parametrizados na própria tabela de cenários, permitindo que a ferramenta seja posteriormente expandida ou personalizada.

---

# 👤 Perfis de investimento

A planilha possui uma seleção de perfil utilizando **validação de dados**.

Os perfis cadastrados são:

* Conservador;
* Moderado;
* Agressivo.

A escolha do perfil influencia a distribuição percentual entre os diferentes tipos de FIIs.

---

# 🏢 Categorias de FIIs

A ferramenta considera as seguintes categorias:

* Papel;
* Tijolo;
* Híbridos;
* FOFs;
* Desenvolvimento;
* Hotelaria.

A distribuição é realizada a partir das informações armazenadas na aba **Tabela de Apoio perfil**.

---

# 🔎 Tabela de apoio

A tabela auxiliar possui a seguinte estrutura:

| Campo          | Descrição                             |
| -------------- | ------------------------------------- |
| Chave composta | Combinação entre perfil e tipo de FII |
| Perfil         | Perfil selecionado                    |
| Tipo de FII    | Categoria do fundo                    |
| %              | Percentual definido para distribuição |

A chave composta é criada por meio da concatenação:

```excel
=C3&"-"&D3
```

Essa estrutura permite relacionar o perfil escolhido com a categoria correspondente e realizar a busca do percentual por meio da função `VLOOKUP`.

---

# 🧠 Nomes definidos

A planilha utiliza **nomes definidos** para algumas células importantes.

| Nome                  | Referência |
| --------------------- | ---------- |
| `aporte`              | `E17`      |
| `qtd_anos`            | `E18`      |
| `Rendimento_carteira` | `E11`      |
| `salario`             | `E10`      |
| `sugest_invest`       | `E13`      |
| `sugest_taxa`         | `E12`      |
| `taxa_mensal`         | `E19`      |

O uso de nomes definidos contribui para melhorar a legibilidade das fórmulas e facilitar a manutenção da planilha.

---

# 🛠️ Recursos utilizados

* Microsoft Excel;
* Fórmulas financeiras;
* Função `FV`;
* Função `VLOOKUP`;
* Função `SUM`;
* Validação de dados;
* Nomes definidos;
* Tabelas auxiliares;
* Referências absolutas e relativas;
* Modelagem de regras de negócio;
* Git;
* GitHub;
* Markdown.

---

# 📁 Estrutura do repositório

A proposta de organização do projeto no GitHub é:

```text
calculadora-investimentos/
│
├── README.md
│
├── projeto/
│   └── Projeto planilha financeira.xlsx
│
├── docs/
│   └── documentacao-tecnica.md
│
└── images/
    └── screenshots/
```

### `README.md`

Apresenta o projeto, seu objetivo, funcionamento e tecnologias utilizadas.

### `projeto/`

Contém a planilha desenvolvida durante o desafio.

### `docs/`

Pode conter a documentação técnica detalhada da solução.

### `images/`

Pode armazenar capturas de tela da planilha para facilitar a visualização do projeto.

---

# 📌 Resultado do desafio

O projeto demonstra a aplicação prática dos conhecimentos adquiridos durante o curso, transformando conceitos de Excel em uma ferramenta funcional de simulação.

Além da construção da planilha, o projeto contempla a **documentação técnica e organização do código/arquivo para compartilhamento no GitHub**, conforme proposto no desafio.

---

# ⚠️ Observação

Esta ferramenta possui finalidade **educacional e de simulação**.

Os resultados apresentados pela planilha dependem dos parâmetros inseridos e das premissas utilizadas no modelo. As projeções não representam garantia de rentabilidade ou de recebimento futuro de dividendos.

---

## 👩‍💻 Status do projeto

**Concluído — Projeto desenvolvido para desafio da DIO**

O projeto poderá ser atualizado posteriormente com novas funcionalidades, melhorias na visualização dos dados e recursos de análise.
