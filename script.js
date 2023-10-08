const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const fs = require('fs');
const moment = require('moment');
const { log } = require('console');

const delay = ms => new Promise(res => setTimeout(res, ms));

// Carregar o arquivo Excel
const workbook = xlsx.readFile('login.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Extrair os dados de e-mail, senha e certificados do Excel
const funcionarios = xlsx.utils.sheet_to_json(worksheet);

// Certificados esperados
const certificadosEsperados = ['Excel Básico', 'Excel Intermediário', 'Excel Avançado', 'Finanças Pessoais', 'Learning Agility - Aprendendo a Aprender', 'Produtividade e Gestão do Tempo', 'Negociação e Influência', 'Comunicação e pensamento assertivo', 'Escrita Assertiva', 'Escutatória', 'Comunicação assertiva à distância', 'Ferramentas para aumentar a criatividade', 'Criatividade para resolução de problemas', 'Mídias pagas: otimizando resultados', 'Facebook Ads e Instagram Ads', 'Anúncios pelo Linkedin', 'Feedback e Feedforward', 'Os superpoderes da inteligência emocional', 'Trabalhando a inteligência emocional com os filhos', 'Planejamento Financeiro Familiar', 'Inteligência Emocional 2.0', 'Gestão Financeira I', 'Gestão Financeira II', 'Gente: Relações Humanas'];

const createFolder = (name) => {
    if (!fs.existsSync(name))
        fs.mkdirSync(name);
}

(async () => {
    // Iniciar o navegador Puppeteer
    const browser = await puppeteer.launch({ headless: false });

    // Iterar sobre a lista de funcionários
    for await (const funcionario of funcionarios) {
        // Abrir uma nova página
        const context = await browser.createIncognitoBrowserContext();
        const page = await context.newPage();
        // Acessar página de login
        await page.goto('https://conquerplus.com.br/login');

        // Preencher e-mail
        await page.type('input[name="email"]', funcionario.email);

        // Preencher senha
        await page.type('input[name="password"]', funcionario.senha);

        // Submeter o formulário de login
        await page.keyboard.press('Enter');

        // Aguardar o login ser concluído
        await page.waitForNavigation();

        let cookies = await page.cookies();
        const lmsToken = cookies.find(cookie => cookie.name === "lms.token");

        //Necessária validação semanal da URL abaixo, pois toda semana é feito a alteração.
        let certificateResponse = await page.goto("https://conquerplus.com.br/_next/data/MkBzaAYi-NNvyUC2n47dw/certificates.json", { referer: "https://conquerplus.com.br/home" });
        //await page.waitForNavigation()
        let certificateResponseJson = await certificateResponse.json();

        // TODO: Verificar se o JSON foi retornado corretamente e contém as propriedades a serem chamadas.
        let certificateList = certificateResponseJson.pageProps.certificateList;

        const outputPath = "output"
        createFolder(outputPath);

        const basePath = `${outputPath}/${funcionario.nome}/`; // Use a propriedade correta para o nome do colaborador
        createFolder(basePath);

        let dataHoraAtual = moment().format('YYYY-MM-DD HH:mm:ss');

        // FUNCIONÁRIO SEM INICIAR CURSO
        if (certificateList.length == 0) {

            const linha = `${dataHoraAtual} - ERROR - ${funcionario.nome} - NÃO EXISTE CERTIFICADOS`;

            try {
                fs.appendFileSync('log.txt', linha + '\n');
            } catch (error) {
                console.log('Ocorreu um erro ao adicionar a linha ao arquivo:', error);
            }


            console.log("Nao existe certificados para o colaborador."); continue;
        }

        let contextAPI = await browser.createIncognitoBrowserContext();
        const pageAPI = await contextAPI.newPage();

        const certificatesInfoArray = [];
        for await (let certificate of certificateList) {
            let title = certificate.title;
            let enrollmentId = certificate.enrollmentId;
            await pageAPI.setExtraHTTPHeaders({
                'authorization': decodeURI(lmsToken.value),
            });

            let apiResponse = await pageAPI.goto(`https://api.conquerplus.com.br/api/certificate/getByEnrollmentId/${enrollmentId}`, { referer: "https://conquerplus.com.br/" });

            let apiResponseJson = await apiResponse.json();

            title = title.replace(/[^a-z0-9\u00C0-\u017F]/gi, '_').toLowerCase()

            let dataHoraAtual = moment().format('YYYY-MM-DD HH:mm:ss');

            //Verifica se possui certificado emitido
            if (apiResponseJson.status == "EMITTING") {

                await pageAPI.goto(`https://api.conquerplus.com.br/api/certificate/enrollment/${enrollmentId}/emmit?testing=true`, { referer: "https://conquerplus.com.br/" });

                const linha = `${dataHoraAtual} - ERROR - ${funcionario.nome} - Certificado em emissão`;

                try {
                    fs.appendFileSync('log.txt', linha + '\n');
                } catch (error) {
                    console.log('Ocorreu um erro ao adicionar a linha ao arquivo:', error);
                }
                console.log(`ERRO | ${funcionario.nome} - Certificado em emissão`); continue;
            }

            certificatesInfoArray.push({ "title": title, "url": apiResponseJson.result.src, "name": certificate.title })
            await delay(300);
        }

        pageAPI.close();
        contextAPI.close();

        //TODO: VERIFICAR SE JÁ EXISTE O ARQUIVO
        let certificadosEsperadosExistentes = certificatesInfoArray.filter(x => certificadosEsperados.find(y => y == x.name));
        for await (let certificateInfo of certificatesInfoArray) {
            let pastaDestino = basePath;
            if (certificadosEsperadosExistentes.find(x => x.name == certificateInfo.name))
                pastaDestino += "/realizados/";
            else
                pastaDestino += "/outros/";

            createFolder(pastaDestino);

            //Verifica se já foi criado o arquivo
            if (fs.existsSync(`${pastaDestino}/${funcionario.nome} - ${certificateInfo.title}.pdf`)) {

                console.log(`RETORNO | ${funcionario.nome} - ${certificateInfo.title} já emitido`); continue;

            }

            const linhaNew = `${dataHoraAtual} - SUCESSO - ${funcionario.nome} - ${certificateInfo.title}`;

            try {
                fs.appendFileSync('logNew.txt', linhaNew + '\n');
            } catch (error) {
                console.log('Ocorreu um erro ao adicionar a linha ao arquivo:', error);
            }
            console.log(`SUCESSO | ${funcionario.nome} - ${certificateInfo.title}`)


            await page.goto(certificateInfo.url);
            await page.pdf({ path: `${pastaDestino}/${funcionario.nome} - ${certificateInfo.title}.pdf`, landscape: true, scale: 2.0 });
        }

        page.close();
        context.close();
        await delay(3000);
    }
    // Fechar o navegador
    await browser.close();
})();