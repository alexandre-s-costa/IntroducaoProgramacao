# IntroducaoProgramacao
Repositório para arquivos usados na disciplina Introdução de Programação, do Programa de Residência em Tecnologia da Informação da UFG.

#### PROJETO CONSULTA PÚBLICA DE PROCESSOS DO PROJUDI COM WEB SCRAPING PARA ANÁLISE E VISUALIZAÇÃO DE DADOS
O projeto visa automatizar a extração de dados do Projudi, usando como filtro os campos *Número do Processo, Nome da Parte ou CPF/CNPJ da Parte,* permitindo armazenamento em um arquivo *Excel*. O usuário pode ainda escolher se deseja fornecer um arquivo Excel com números de processos para que o script faça a consulta no Projudi, salvando os resultados no *Excel.*
 Público alvo
Servidores que necessitam consultar vários processos no Projudi diariamente, mas não têm acesso ao banco de dados.

 #### Descrição do Arquivo
Como se trata de um código para Web Scraping, os dados são disponibilizados publicamente no Projudi, através de consulta pública. O script possui as seguintes características:
1. Utiliza a biblioteca __Selenium__ __WebDriver__ para automatizar a navegação no site e controlar o navegador __Google__ __Chrome__;
2. Permite ao usuário a consulta por Nome da Parte (Número do Processo ou CPF/CNPJ da Parte) ou ainda informando o caminho de um arquivo *Excel,* contendo números de processos;
3. Captura dados detalhados de cada processo encontrado como número, área, classe, assunto(s), valor da causa, fase processual, data de distribuição, custas e status;
4. Captura de múltiplas páginas de resultados (somente quando usada a opção Nome da Parte);
5. Exporta todos os dados coletados para um arquivo *Excel* com a data e hora atual;
6. Apresenta mensagem com o número de registros coletados corresponde ao total informado pelo sistema no rodapé (total de processos).

O código também inclui tratamentos de erro e mecanismos para lidar com possíveis problemas durante a navegação, como elementos que não carregam ou páginas que demoram para responder.
