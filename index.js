const XLSX = require('xlsx');
const fs = require('fs');
const puppeteer = require('puppeteer');


const workbook = XLSX.readFile('IMPLANTAÇÃO_PONTOFOPAG.xlsx', {
  type: 'binary',
  cellDates: true,
  cellNF: false,
  cellText: false
});
const sheetName = 'Dados_Funcionários'; // nome do sheet
const worksheet = workbook.Sheets[sheetName];

// Converter a planilha para um array de objetos JavaScript
const jsonData = XLSX.utils.sheet_to_json(worksheet);
const jsonLeght = jsonData.length;
let count = 0;

// Converter o array em uma string JSON
const jsonString = JSON.stringify(jsonData);


// Gravar a string JSON em um arquivo
fs.writeFileSync('dados_funcionarios.json', jsonString, (err) => {
  if (err) {
    console.error('Erro ao gravar arquivo:', err);
  } else {
    console.log('Arquivo gravado com sucesso!');
  }
});


// acessos para o pontofopag 
const login = 'epaysmyrh';
const senha = 'Z6GdVzLzEgFtz9KH@2023';


// Puppeteer
(async () => {
    const browser = await puppeteer.launch({
      headless: false,
    });
    
    const page = await browser.newPage();
    await page.setViewport({ width: 1366, height: 720});

    await page.goto('https://prd.pontofopag.com.br');
    await page.type('[name="login"]', login ); // informação de login e senha 
    await page.type('[name="Password"]', senha );
    await page.click('[type="submit"]');
    await page.waitForSelector('[title="Sair"]');
    
    while (count < jsonLeght){

      // DE - PARA
        let numberMatricula = jsonData[count]['* Matrícula:'];
        let matricula = numberMatricula.toString();
        let nPis = jsonData[count].PIS;
        let cpf = jsonData[count]['* CPF'];
        let nome = jsonData[count]['* Nome'];
        let dataAdmissao = jsonData[count]["* Data Admissao"];
        const dia = String(dataAdmissao.getDate()).padStart(2, '0');
        const mes = String(dataAdmissao.getMonth() + 1).padStart(2, '0');
        const ano = String(dataAdmissao.getFullYear());
        const dataFormatada = dia + mes + ano;
        let empresa = jsonData[count]["* Empresa (Razão Social)"];
        let cnpj = jsonData[count]["* CNPJ "];
        let departamento = jsonData[count]["* Departamento"];
        let alocacao = jsonData[count]["Alocação"];
        let contrato = jsonData[count]["Contrato"];
        let supervisor = jsonData[count]["Supervisor"];
        let funcao = jsonData[count]["* Função"];
        let horario = jsonData[count]["* Nome do Horário"];
        let bh = jsonData[count]["* Vincula ao Banco de Horas?"]
        let registradorWeb = jsonData[count]["* Utiliza Registrador Web?"]
        let senhaApp = jsonData[count]["Senha Registrador WebApp/App Ponto"]
        let appUse = jsonData[count]["Utiliza aplicativo Pontofopag?"]
        let senhaEnable = false;
        const errorPopup = await page.$('.modal-header');

      await page.goto('https://prd.pontofopag.com.br/Funcionario/Grid');
      await page.waitForSelector('#btIncluir');
      await page.click('#btIncluir');
      await page.waitForSelector('[name="Matricula"]');
      await page.type('[name="Matricula"]', matricula);
      await page.type('[name="Pis"]', nPis);
      await page.type('[name="CPF"]', cpf);
      await page.type('[name="Nome"]', nome);
      // await page.type('[name="Dataadmissao"]', dataFormatada);
      // await page.type('[name="Empresa"]', empresa);
      // await page.keyboard.press("Tab");
      // await page.type('[name="Departamento"]', departamento);
      // await page.keyboard.press("Tab");
      // await page.type('[name="Funcao"]', funcao);
      // await page.keyboard.press("Tab");
      // await page.type('[name="Horario"]', horario);
      // await page.keyboard.press("Tab");
      
      if ( alocacao === null){
        console.log(`Alocação ${count} é nulo`);
      } else if (alocacao === undefined){
        console.log(`Alocação ${count} não foi definida`);
      } else {
        await page.type('[name="Alocacao"]', alocacao);
        await page.keyboard.press("Tab");
        if (errorPopup) {
          console.log(`Alocação ${alocacao} do funcionário ${nome} nao cadastrada`)
        }
      }

      // if ( contrato === null){
      //   console.log(`contrato ${count} é nulo`);
      // } else if (contrato === undefined){
      //   console.log(`contrato ${count} não foi definido`);
      // } else {
      //   await page.type('[name="Contrato"]', contrato);
      //   await page.keyboard.press("Tab");
      // }

      // if ( supervisor === null){
      //   console.log(`supervisor ${count} é nulo`);
      // } else if (supervisor === undefined){
      //   console.log(`supervisor ${count} não foi definido`);
      // } else {
      //   await page.type('[name="Supervisor"]', supervisor);
      //   await page.keyboard.press("Tab");
      //}

      // if (bh === 'nao' || bh === 'não' || bh === 'NÃO' || bh === 'NAO') {
      //   await page.click('[name="bNaoentrarbanco"]');
      // }

      // if (registradorWeb === 'sim' || registradorWeb === 'SIM') {
      //   await page.click('[name="UtilizaWebAppPontofopag"]');
      //   senhaEnable = true;
      // }

      // if (appUse === 'sim' || appUse === 'SIM') {
      //   await page.click('[name="UtilizaAppPontofopag"]');
      //   senhaEnable = true;
      // }

      // if (senhaEnable === true) {
      //   await page.type('[name="Mob_Senha"]', senhaApp);
      // }
      await page.waitForTimeout(3000);

      count++;
      console.log(`Feito ${count} de ${jsonLeght}`)
    }


    // await browser.close();
})();
