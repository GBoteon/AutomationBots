# AutomationBots
Author: Gustavo Rosas Boteon

*Bem vindo(a) a documentação do Bot de cadastros de representantes*

Pra começar os Arquivos CreateRoutine001 e CreateRoutine006 são Rotinas que utilizam as bibliotecas Selenium, Pandas, os, Datetime e também utiliza o Chromedriver como software para navegação do chrome automatizada.
Criadas para estarem operando por tempo inderteminado mas no entanto há pontos importantes a serem ressaltados como:
- Não executar com uma planinha do excel de entrada aberta!
 - No arquivo usuario.txt a informações utilizadas pelo programa que precisam ser alteradas para execução em     outras máquinas a primeira é a coluna usuario, a primeira linha é o usuario do windows, que é utilizado para o profile do Chrome, em baixo do user é o contador da localização do elemento HTML botão collapse usado para indentificar erros; a coluna seguinte é o registro, em baixo é o numero de linhas criadas no excel do onedrive Entrada 001, e em baixo o número de linhas criadas no excel Entrada 006 também no Onedrive.
 - O Executavel UpdateFromOneDrive é responsavel por ler e criar arquivos de entrada da pasta do OneDrive e também utiliza o usuario do windows
 - Caso haja necessidade de excluir um registro no navegador será necessario rebootar o progama para que seja executado corretamente
 - Cada execução de um arquivo de entrada gera um log informando quais cadastros foram executados e quais falharam, em casos de erro ele também ira gerar uma planinha com os não cadastrados para uma outra tentativa ao serem inseridos na pasta de entrada uma segunda vez
 - O programa pode executar tanto no UAT quanto no AX bastanto alterar o link que esta na 4° linha da coluna usuario do usuario.txt entre enesa-uat.sandbox.operations.dynamics.com e ensp-365-axprod.operations.dynamics.com.
 - Para a Execução do progama é necessario que na mesma pasta do executavel esteje também uma pasta chamada Cadastro Representantes, o chromedriver.exe, o usuario.txt, e dentro da pasta Cadastro Representantes ter as pastas: Entrada, Erros e logData.
 - para a utilização do Bot basta deixa-lo rodando e acresentar arquivos Excel com uma coluna NOME e outra REPRESENTANTE na pasta Entrada dento da pasta Cadastro Representantes