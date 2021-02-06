# Biblioteca Pessoal - 1.1
Aplicação de conceitos de Excel voltada a organização de um acervo pessoal de livros.

Este arquivo é a primeira tentativa do uso conhecimentos básicos e intermediários do Excel, visando uma aplicação prática dentro da BCI.
A catalogação e classificação aqui utilizadas são simplificadas para fins de prática de ferramentas do excel.
Há limitações de acesso ao software na construção inicial desse arquivo (utilização apenas do Excel Online disponibilizado para estudantes).

Resumo de funções utilizadas: ÍNDICE, SE, SEERRO, PROCV, MENOR, MAIOR, DESLOC, CONT.VALORES.
________________________________________________________________________________________________________________________________________________________
### VERSÃO ATUAL
### **1.3 - Implementação de Dashboard**

- A planilha está em processo de implementação de uma Dashboard.
  - Agora, o quadro "Informações gerais" foi transferido da planilha Consulta para planilha Dashboard.
  - O menu de navegação foi remodelado para a Dashboard. Todas as planilhas contém um hyperlink de fácil acesso para redirecionamento para a Dashboard.

- A dashboard contém, além das informações gerais: gráficos dinâmicos, botão com Macro para atualização das planilhas dinâmicas, menu de navegação.
  - A formatação e visibilidade da informação também é muito mais agradável ao usuário.
  
________________________________________________________________________________________________________________________________________________________
### HISTÓRICO
### **1.2 - Grandes ajustes de funcionalidade**

- A planilha foi atualizada para a versão desktop do Excel 2016.
  - Agora, os intervalos dinâmicos para a validação de dados são nomeados (seguindo a lógica *bd_*) assim como as tabelas dinâmicas (*td_*). Isto permitiu que fórmulas que referenciavam esse intervalo fossem drasticamente reduzidas, com muito maior visibilidade, funcionalidade e estabilidade.
  
- Bugfix na fórmula de ano com menos obras ainda pendente.

### **1.11 - Ajustes, bugfixes e compatibilidade com o Calc**

- Reformulação da forma de cálculo das informações gerais: agora utiliza as funções MAIOR/MENOR e ÍNDICE para retornar os valores mais presentes, não fazendo referência apenas à célula da tabela dinâmica.
  - Agora, os dados são exibidos corretamente, independente da forma de classificação da tabela dinâmica.
  - Isto também permitiu a inserção a exibição dos valores menos presentes (necessita de bugfix apenas no de "ano com menos obras");
  - Obs: A função DESLOC é utilizada pois o Excel Online não permite a nomeação de intervalos.
    
- Bugfix nas funções de busca (maior intervalo na função ÍNDICE e ajuste de funções LIN indefinidas);

- Teste de compatibilidade com Libreoffice Calc -> as tabelas dinâmicas **NÃO** são mais compatíveis, pois o intervalo não acompanha o crescimento da base de dados, inviabilizando a planilha como um todo.
  - A compatibilidade é possível, mas necessita de ajustes (nomeação de intervalo da BD, etc), que são bem diferentes em ambos os softwares.
  
#### **1.1 - Implementação de informações gerais**

- Mais mudanças visuais (Clareza na funcionalidade);
- Inserção de uma resumo de informações gerais do acervo, como número de obras, de autores e editoras; assunto mais comum, autores mais prominentes e etc.
  - Uso da função SE no caso de coleções para eliminar o "n/a" caso este seja o dado com maior prevalência, mostrando então o segundo valor.
- Ajuste nas indicações de autoria (exemplo: Coordernadação de NOME ajustado para NOME (coord.)).

#### **1.0 - Acervo funcional**

- Todas as obras foram inseridas.
- Todas as funções essenciais foram programadas e testadas.
- Inserção de classificação condicional para os dados incertos na base de dados (contém a expressão (?) ), ressaltando-as através de uma borda vermelha.

#### **0.4 - Validação de Dados**

- Mudanças visuais na planilha de consulta;
- Inserção de consulta por Editora;
  - Também foram implementados mais campos para exibição do resultados da consulta.
    
- Validação de dados com intervalos dinâmicos implementadas (uso das Funções DESLOC e CONT.Valores).
  - Agora as listas suspensas para seleção de título, assunto, autor, país de origem e editora da aba de consultas atualizam automaticamente conforme inserção de novos dados na base de dados. Limite atual de até 5000 linhas, mas facilmente modificável.
    
  (Obs: Ainda é necessário que as tabelas dinâmicas sejam atualizadas.)
  
#### **0.3 - Teste de compatibilidade**

- Inserção da consulta por País da Obra Original;
- Teste de compatibilidade com LibreOffice Calc -> Todas as funções, tabelas dinâmicas e validações estão funcionando 
corretamente.
- Inserção de mais obras na base de dados;

#### **0.2 - Aplicação prática de funções**
- Inserção de mais obras na base de dados;
- Mudanças visuais (maior clareza e usabilidade);
- Inserção da possibilidade de consulta de obras por Assunto e Autor (uso das funções ÍNDICE, SE e MENOR);

#### **0.1 - Inicial**
- Aplicação básicas de formatação de tabelas, intervalos e validação de dados;
- Construção da base dados inicial;
- Consulta de informações das obras (Função PROCV);
- Base inicial das tabelas dinâmicas;


