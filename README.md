# Biblioteca Pessoal - 1.62
Aplicação de conceitos de Excel voltada a organização de um acervo pessoal de livros.

Este arquivo é a primeira tentativa do uso conhecimentos básicos e intermediários do Excel, visando uma aplicação prática dentro da BCI.
A catalogação e classificação aqui utilizadas são simplificadas para fins de prática de ferramentas do excel.

Resumo de funções utilizadas: ÍNDICE, SE, SEERRO, PROCV, MENOR, MAIOR, DESLOC, CONT.VALORES, Tabelas dinâmicas, Gráficos Dinâmicos, Macro, VBA Básico e Hyperlinks.



## Estruturação de nomeação de Intervalos:

*bd_Main* -> Refere-se a Base de Dados em sua totalidade.

*td_XXXX* -> Intervalo das tabelas dinâmicas (ambas as colunas).

*bd_XXXX* -> Intervalo dinâmico dos valores preenchidos das tabelas dinâmicas (1a coluna apenas, isto é, os rótulos).

*seg_XXXX* -> Refere-se às segmentações de dados.

*macro_XXXX* -> Refere-se às macros.
________________________________________________________________________________________________________________________________________________________
# VERSÃO ATUAL
### **1.62 - Bugfixes**

- Botão de "Resetar Filtros" corrigido. O nome das segmentações de dados na fórmula da Macro estavam incorretas.
    

________________________________________________________________________________________________________________________________________________________
### HISTÓRICO
### **1.61 - Melhorias na Dashboard**

*Referentes à Dashboard*
 - Inserção das bandeiras dos países como ponto de dados no gráfico de Obras por País da Obra Original;
 - Inserção de mais dois campos de informação na Dashboard: Total de Obras, Autores e Editoras, refletindo a do Painel do controle, mais com visual melhorado.
 - Ranking de Top 3 Assuntos, Autores e Editoras da seleção escolhida.
 
*Referentes à pasta de trabalho*
 - Foram inseridas funções SEERRO para todas as fórmulas do painel de controle, melhorando a visibilidade em caso de aplicação de filtros que não possuem resultados.
 - Foi criada uma planilha de Tabelas de Apoio, onde seram inseridas planilhas não-dinâmicas auxiliares.
 - 
### **1.6 - Dashboard 2.0**

- A Planilha Dashboard agora tem duas partes:
  - Painel de controle superior, onde constam as informações gerais, o menu de navegação e cadastro, assim como a de consulta rápida;
  - A Dashboard propriamente dita, com os gráficos dinâmicos, filtros e disposição de informações de maneira mais visual.
 
- Toda a estruturação gráfica da Dashboard foi atualizada.
  - Os seguintes gráficos foram removidos: Número de Obras por Ano, Número de Obras por idioma (this is a buff);
  - A visibilidade, riqueza visual e a paleta de cores foram drasticamente melhoradas.
  - O Painel de Filtro foi reformulado e inserido dentro da área da dashboard.
  - A Dashboard se encontra  levemente separada do Painel de Superior, sendo possível enquadrá-la com sua totalidade no Zoom Padrão do Excel (100%).
  
### **1.51 - Filtros**

- Agora há a possibilidade de filtrar as informações na dashboard.
- Textos foram alterados para indicar que as informações dispostas se referem a atual seleção.

### **1.5 - Melhorias na Dashboard**

- A Base de Dados agora possui uma nova Coluna: Extensão, que classifica a obra conforme o número de páginas.

- A Dashboard agora possui mais um gráfico dinâmico, referente ao percentual de obras conforme sua Extensão.
  - A posição e tamanhos dos gráficos foram ajustados para melhor harmonia visual.
  - Também foi inserido um campo de consulta de informações por título na dashboard, igual ao da Planilha Consulta.
  - O visual dos botões do menu de cadastro e de navegação foram melhorados.

- Bugfix na aba de consultas: alguns campos que estavam com a fórmula ÍNDICE sem a aplicação de fórmula matricial foram corrigidos.

### **1.4 - Formulário de Cadastro**

- A Dashboard agora possui mais gráficos dinâmicos.

- A Dashboard agora possui dois botões: 
  - Atualizar (atualiza todas os dados/tabelas dinâmicas da planilha.)
  - **Cadastro**: Chama um Forms para inserção de dados na Base de Dados. Isso permite a inserção de dados de maneira mais eficiente. O Forms utiliza uma programação básica em VBA, de fácil compreensão e leitura.
  
- O Formulário de Cadastro ainda necessita de melhorias visuais.

### **1.3 - Implementação de Dashboard**

- A planilha está em processo de implementação de uma Dashboard.
  - Agora, o quadro "Informações gerais" foi transferido da planilha Consulta para planilha Dashboard.
  - O menu de navegação foi remodelado para a Dashboard. Todas as planilhas contém um hyperlink de fácil acesso para redirecionamento para a Dashboard.

- A dashboard contém, além das informações gerais: gráficos dinâmicos, botão com Macro para atualização das planilhas dinâmicas, menu de navegação.
  - A formatação e visibilidade da informação também é muito mais agradável ao usuário.
  
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
- Inserção de um resumo de informações gerais do acervo, como número de obras, de autores e editoras; assunto mais comum, autores mais prominentes e etc.
  - Uso da função SE no caso de coleções para eliminar o "n/a" caso este seja o dado com maior prevalência, mostrando então o segundo valor.
- Ajuste nas indicações de autoria (exemplo: Coordernadação de NOME ajustado para NOME (coord.)).

#### **1.0 - Acervo funcional**

- Todas as obras foram inseridas.
- Todas as funções essenciais foram programadas e testadas.

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


