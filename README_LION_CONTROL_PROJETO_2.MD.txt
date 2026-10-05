# 🦁 LION CONTROL — Organizador de Dados para Declaração de Imposto de Renda

## 📌 Sobre o projeto

O **LION CONTROL** é uma ferramenta desenvolvida em **Microsoft Excel** com o objetivo de organizar e centralizar informações importantes para auxiliar na preparação da declaração de Imposto de Renda.

A proposta é transformar informações que normalmente ficam espalhadas em documentos, comprovantes e anotações em uma estrutura única, organizada e de fácil preenchimento.

Este projeto foi desenvolvido como parte de um desafio prático de aprendizagem, com foco na aplicação de conceitos de **Excel, organização de dados, validação de informações e documentação técnica**.

---

## 🎯 Objetivo

Criar uma ferramenta simples e funcional para reunir informações essenciais utilizadas como apoio na preparação da declaração de Imposto de Renda, permitindo:

- Centralizar os dados do titular;
- Organizar informes de rendimento bancário;
- Registrar entradas provenientes de notas, holerites e outros recebimentos;
- Utilizar listas de seleção para reduzir erros de preenchimento;
- Realizar cálculos automáticos;
- Manter tabelas auxiliares para padronização das informações;
- Facilitar a consulta e organização dos documentos utilizados na declaração.

---

## 🧩 Estrutura da planilha

A solução está organizada em quatro abas principais:

### 1. 👤 TITULAR

Área destinada ao cadastro das principais informações do titular.

**Informações contempladas:**

- Nome;
- CPF;
- Data de nascimento;
- Título de eleitor;
- Cônjuge;
- Endereço;
- CEP;
- Telefone;
- Celular;
- E-mail;
- Outras informações cadastrais.

A aba possui **validação de dados** para campos que utilizam opções predefinidas.

---

### 2. 🏦 INFORMES

Área destinada ao registro dos **informes de rendimento bancário**.

É possível organizar informações de diferentes instituições financeiras, incluindo:

- Banco;
- Valor atual;
- Identificação do documento/anexo;
- Total dos valores registrados.

A planilha utiliza fórmula para realizar o cálculo automático do total dos valores informados.

**Exemplo de cálculo utilizado:**

```excel
=SUM(D10,D15,D20)
```

A estrutura pode ser expandida futuramente para incluir outras instituições ou categorias de investimentos.

---

### 3. 🧾 NOTAS

Área destinada ao registro de entradas e informações provenientes de documentos e comprovantes.

A tabela contempla:

| Campo | Descrição |
|---|---|
| DATA | Data do recebimento ou documento |
| CATEGORIA | Tipo de entrada/documento |
| VALOR | Valor registrado |

Para a categoria, foi utilizada uma **lista de validação**, com opções como:

- HOLERITE
- CNPJ
- FREELANCE

Essa estrutura ajuda a padronizar os registros e reduzir erros de digitação.

---

### 4. 📚 TABELAS

A aba **TABELAS** funciona como uma área auxiliar para armazenar informações utilizadas pela planilha.

Entre os dados disponíveis está uma relação de instituições bancárias, utilizada como base para padronização dos registros.

A separação das tabelas auxiliares em uma aba própria facilita futuras alterações e expansões do projeto.

---

## ⚙️ Recursos utilizados

O projeto utiliza recursos nativos do Excel, entre eles:

- Formatação de células;
- Organização por abas;
- Fórmulas;
- Validação de dados;
- Listas suspensas;
- Tabelas auxiliares;
- Cálculos automáticos;
- Estruturação e padronização de informações.

---

## 🧠 Conceitos praticados

Durante o desenvolvimento do projeto, foram aplicados conceitos relacionados a:

- Organização de dados;
- Estruturação de planilhas;
- Validação de informações;
- Padronização de entradas;
- Fórmulas no Excel;
- Referenciamento de tabelas auxiliares;
- Automação de cálculos;
- Documentação técnica;
- Organização de projetos para publicação no GitHub.

---

## 🚀 Como utilizar

1. Faça o download do arquivo Excel.
2. Abra a planilha no Microsoft Excel.
3. Acesse a aba **TITULAR** e preencha os dados cadastrais.
4. Acesse **INFORMES** para registrar os dados dos bancos e respectivos valores.
5. Utilize a aba **NOTAS** para registrar entradas e documentos.
6. Utilize as listas disponíveis para selecionar categorias e evitar inconsistências.
7. Consulte os totais calculados automaticamente.
8. Mantenha os documentos comprobatórios organizados para facilitar a conferência das informações.

> **Importante:** esta ferramenta tem finalidade de **organização e controle de informações**. Ela não substitui a declaração oficial do Imposto de Renda nem a orientação de um profissional contábil.

---

## 📚 Aprendizados

Este projeto permitiu colocar em prática conhecimentos de Excel em uma situação próxima de uma necessidade real, trabalhando não apenas a construção da planilha, mas também a **organização, padronização e documentação de uma solução**.

Além do desenvolvimento técnico, o projeto faz parte do processo de aprendizagem sobre o uso do **GitHub como ferramenta para versionamento e compartilhamento de documentação e projetos**.

## 👩‍💻 Projeto

**LION CONTROL — Organizador de Dados para Declaração de Imposto de Renda**

Projeto desenvolvido para fins de estudo e prática profissional.

---

## 📄 Licença

Este projeto foi desenvolvido para fins educacionais e de portfólio.

