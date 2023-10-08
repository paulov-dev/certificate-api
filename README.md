# Projeto de Automação de Downloads de Certificados e Interação com a Web
Este projeto consiste em um script de automação desenvolvido em JavaScript e Node.js, que tem como objetivo simplificar o processo de download de certificados de funcionários de uma plataforma online chamada "conquerplus.com.br". O script utiliza a biblioteca Puppeteer para interagir com o site de forma automatizada, preenchendo formulários de login, autenticando-se e baixando os certificados em formato PDF.

Através da leitura de um arquivo Excel contendo informações dos funcionários, o projeto automatiza as etapas de acesso à plataforma, verificação da existência de certificados e organização dos arquivos baixados em pastas correspondentes. Isso economiza tempo e minimiza erros humanos, tornando o processo de documentação mais eficiente e confiável.

O script também registra eventos e informações relevantes em arquivos de log, proporcionando uma visão detalhada do processo de automação. Este projeto é uma demonstração prática de como tecnologias como Node.js e Puppeteer podem ser aplicadas para otimizar tarefas repetitivas e manuais em um ambiente de trabalho.

## Pré-requisitos
Certifique-se de ter as seguintes ferramentas instaladas em seu ambiente de desenvolvimento:

Node.js: Este projeto foi desenvolvido usando Node.js, portanto, você precisará tê-lo instalado em sua máquina

## Instalação
1. Clone este repositório:

   ```bash
   git clone https://github.com/paulov-dev/certificate-api
   cd pulse-app
   ```

2. Instale as dependências do Node.js:

   ```bash
   npm install
   ```

3. Inicie o servidor:

   ```bash
     node script.js
   ```


### Uso
Para usar este projeto, siga estas etapas:

- Certifique-se de ter os arquivos de dados corretos. O script espera um arquivo Excel contendo informações de funcionários.

- Edite o arquivo ``login.xlsx`` com as informações de login dos funcionários.

- Personalize a lista de certificados esperados no código, se necessário:

   ```bash
     const certificadosEsperados = ['Certificado 1', 'Certificado 2', ...];
   ```

## Configuração
Você pode ajustar algumas configurações no arquivo index.js conforme necessário:

- ``headless``: Defina como true para executar o navegador no modo headless (sem interface gráfica) ou como false para ver a interface gráfica do navegador.
- ``certificadosEsperados``: Defina a lista de certificados esperados que você deseja arquivar.
- ``delay``: Ajuste o tempo de espera entre as ações, se necessário.

### Personalização
Você pode personalizar este projeto de acordo com suas necessidades, adicionando recursos extras, melhorando a lógica de organização de certificados ou implementando tratamentos de erros adicionais.

## Arquivos de Log
O projeto gera arquivos de log para registrar o status da emissão de certificados:

``log.txt``: Registra erros e informações sobre certificados não emitidos.
``logNew.txt``: Registra informações sobre certificados emitidos com sucesso.

## Aviso
Este projeto foi desenvolvido para fins educacionais e pode ser adaptado para atender às suas necessidades específicas. Certifique-se de cumprir os termos de uso do site automatizado e de respeitar a privacidade e a segurança dos dados dos funcionários, conforme exigido pela Lei Geral de Proteção de Dados (LGPD).
