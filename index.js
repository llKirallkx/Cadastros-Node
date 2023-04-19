const XLSX = require('xlsx');
const fs = require('fs');
const puppeteer = require('puppeteer');


const workbook = XLSX.readFile('AUTOMPONTO.xlsx', {
  type: 'binary',
  cellDates: true,
  cellNF: false,
  cellText: false
});
const sheetName = 'automação'; // nome do sheet
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
const login = 'epaysclimatich';
const senha = 'Z6GdVzLzEgFtz9KH@2023';

// Puppeteer
(async () => {
    const browser = await puppeteer.launch({
      headless: false,
    });
    
    const page = await browser.newPage();
    await page.setViewport({ width: 1366, height: 612});

    await page.goto('https://prd.pontofopag.com.br');
    await page.type('[name="login"]', login ); // informação de login e senha 
    await page.type('[name="Password"]', senha );
    await page.click('[type="submit"]');
    await page.waitForSelector('[title="Sair"]');
    
    while (count < jsonLeght){

      // DE - PARA
        let numberMatricula = jsonData[count]['Matrícula:'];
        let matricula = numberMatricula.toString();
        
        let numberPIS = jsonData[count]['PIS'];
        let nPis = null;
        if ( numberPIS === null){
          nPis = numberPIS;
        } else if (numberPIS === undefined){
          nPis = numberPIS;
        } else {
          nPis = numberPIS.toString();
        }

        let cpf = jsonData[count]['CPF'];
        let nome = jsonData[count]['Nome'];
        let dataAdmissao = jsonData[count]["Data Admissao"];
        let dia = String(dataAdmissao.getDate()).padStart(2, '0');
        let mes = String(dataAdmissao.getMonth() + 1).padStart(2, '0');
        let ano = String(dataAdmissao.getFullYear());
        let dataFormatada = dia + mes + ano;
        let empresa = jsonData[count]["Empresa"];
        let cnpj = jsonData[count]["* CNPJ "];
        let departamento = jsonData[count]["Departamento"];
        let numberAlocacao = jsonData[count]["Alocação"];
        let alocacao = null;
        if ( numberAlocacao === null){
          alocacao = numberAlocacao;
        } else if (alocacao === undefined){
          alocacao = numberAlocacao;
        } else {
          alocacao = numberAlocacao.toString();
        }
        let contrato = jsonData[count]["Contrato"];
        let numberSupervisor = jsonData[count]["Supervisor"];
        let supervisor = numberSupervisor.toString();
        let numberFuncao = jsonData[count]["Função"];
        let funcao = numberFuncao.toString();

        let numberHorario = jsonData[count]["Horário"];
        let horario = numberHorario.toString();
        let bh = jsonData[count]["* Vincula ao Banco de Horas?"]
        let registradorWeb = jsonData[count]["* Utiliza Registrador Web?"]
        let numberSenhaApp = jsonData[count]["Senha"]
        let senhaApp = numberSenhaApp.toString();
        let appUse = jsonData[count]["Utiliza aplicativo Pontofopag?"]
        let senhaEnable = false;
        const errorMessage = "Registro não encontrado.";

      await page.goto('https://prd.pontofopag.com.br/Funcionario/Grid');
      await page.waitForSelector('#btIncluir');
      await page.click('#btIncluir');
      await page.waitForSelector('[name="Matricula"]');
      await page.waitForTimeout(500);
      await page.type('[name="Matricula"]', matricula);
      if ( nPis === null){
        console.log(`PIS do funcioario ${nome} é nulo`);
      } else if (nPis === undefined){
        console.log(`PIS do funcioario ${nome} não foi preenchido`);
      } else {
        await page.type('[name="Pis"]', nPis);;
      }
      await page.type('[name="CPF"]', cpf);
      await page.type('[name="Nome"]', nome);
      await page.type('[name="Dataadmissao"]', dataFormatada);

      await page.type('[name="Empresa"]', empresa);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(500);

      await page.type('[name="Departamento"]', departamento);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(300);


      await page.type('[name="Funcao"]', funcao);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(300);


      await page.type('[name="Horario"]', horario);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(300);

      
      
      if ( alocacao === null){
        console.log(`Alocação ${count} é nulo`);
      } else if (alocacao === undefined){
        console.log(`Alocação ${count} não foi definida`);
      } else {
        await page.type('[name="Alocacao"]', alocacao);
        await page.keyboard.press("Tab");
        await page.waitForTimeout(300);
      }

      if ( contrato === null){
        console.log(`contrato ${count} é nulo`);
      } else if (contrato === undefined){
        console.log(`contrato ${count} não foi definido`);
      } else {
        await page.type('[name="Contrato"]', contrato);
        await page.keyboard.press("Tab");
        await page.waitForTimeout(300);
      }

      if ( supervisor === null){
        console.log(`supervisor ${count} é nulo`);
      } else if (supervisor === undefined){
        console.log(`supervisor ${count} não foi definido`);
      } else {
        await page.type('[name="Supervisor"]', supervisor);
        await page.keyboard.press("Tab");
        await page.waitForTimeout(333);
      }

      if (bh === 'nao' || bh === 'não' || bh === 'NÃO' || bh === 'NAO') {
        await page.click('[name="bNaoentrarbanco"]');
      }

      if (registradorWeb === 'sim' || registradorWeb === 'SIM') {
        await page.waitForTimeout(500);
        await page.click('[name="UtilizaWebAppPontofopag"]');
        senhaEnable = true;
      }

      if (appUse === 'sim' || appUse === 'SIM') {
        await page.waitForTimeout(500);
        await page.click('[name="UtilizaAppPontofopag"]');
        senhaEnable = true;
      }

      if (senhaEnable === true) {
        await page.waitForTimeout(500);
        await page.type('[name="Mob_Senha"]', senhaApp);
      }


      await page.waitForTimeout(1000);

      await page.click('[type="submit"]');
      await page.waitForSelector('[title="Sair"]');

      count++;
      console.log(`Feito ${count} de ${jsonLeght}`);

    }

    console.log('Cadastros finalizados');
    await browser.close();
})();
