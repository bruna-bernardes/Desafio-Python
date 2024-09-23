# Desafio-Python
Criação de código em Python para checagem de estatísticas de empresas no Reclame Aqui de acordo com critérios indicados.

Tecnologias Utilizadas:
Python 3.12+
Selenium: Automação do navegador.
BeautifulSoup: Extração de dados HTML.
Pandas: Manipulação de dados e exportação para Excel.
openpyxl: Manipulação e formatação de arquivos Excel.
webdriver-manager: Gerenciamento automático do WebDriver (para Chrome).

Requisitos de Instalação:
Python 3.12+ deve estar instalado. Baixe e instale o Python.
Clone este repositório, navegue até o diretório do projeto, crie e ative um ambiente virtual.
Instale as dependências do projeto.
Para executar o script e iniciar a coleta de dados, basta rodar o seguinte comando: python Desafio.py
O script navegará pelo site do Reclame Aqui, coletará as informações das empresas e gerará um arquivo Excel chamado empresas_dados_formatado.xlsx na pasta do projeto.

Saída:
Após a execução bem-sucedida, um arquivo Excel será gerado com as informações das 3 melhores e 3 piores empresas de planos de internet: empresas_dados_formatado.xlsx
A planilha terá formatação condicional:
. As 3 melhores empresas estarão destacadas em verde claro.
. As 3 piores empresas estarão destacadas em vermelho.

Dependências:
As principais dependências utilizadas no projeto estão listadas abaixo:
. selenium
. beautifulsoup4
. pandas
. openpyxl
. webdriver-manager

Como Funciona o Código:
. Navegação e Extração de Links: O Selenium navega pelo site do Reclame Aqui e localiza as seções de "Melhores Empresas" e "Piores     
Empresas" no ramo de Planos de Internet.
. Extração de Informações: Usamos BeautifulSoup para parsear o HTML das páginas das empresas e coletar as estatísticas desejadas.
. Armazenamento: As informações são armazenadas em um DataFrame do Pandas e depois exportadas para um arquivo Excel.
. Formatação: O arquivo Excel é formatado com o openpyxl para destacar as melhores e piores empresas.

