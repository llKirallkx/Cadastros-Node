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

// DE - PARA
let numberMatricula = jsonData[count]['* Matrícula:'];
let matricula = numberMatricula.toString();
let pis = jsonData[count].pis;
let cpf = jsonData[count]['* CPF'];
let nome = jsonData[count]['* Nome'];
let dataAdmissao = jsonData[count]["* Data Admissao"];
let empresa = jsonData[count]["* Empresa (Razão Social)"];
let cnpj = jsonData[count]["* CNPJ "];
let departamento = jsonData[count]["* Departamento"];
let alocação = jsonData[count]["Alocação"];
let contrato = jsonData[count]["Contrato"];
let supervisor = jsonData[count]["Supervisor"];
let funcao = jsonData[count]["* Função"];
let horario = jsonData[count]["* Nome do Horário"];

// "* Vincula ao Banco de Horas?": "SIM",
// "* Utiliza Registrador Web?": "SIM",
// "Senha Registrador WebApp/App Ponto": "Informe uma senha exemplo: CPF, 12345, 6 primeiros digitos",
// "Utiliza aplicativo Pontofopag?": "SIM",
// "* Utiliza Registrador c/ Reconhecimento Facial?": "SIM",
// "Utiliza Cerca Virtual no APP?": "SIM",



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
    await page.setViewport({ width: 1080, height: 720});

    await page.goto('https://prd.pontofopag.com.br');
    await page.type('[name="login"]', login ); // informação de login e senha 
    await page.type('[name="Password"]', senha );
    await page.click('[type="submit"]');
    await page.waitForSelector('[title="Sair"]');
    
    while (count < jsonLeght){

        await page.goto('https://prd.pontofopag.com.br/Funcionario/Grid');
        await page.waitForSelector('#btIncluir');
        await page.click('#btIncluir');
        await page.type('[name="Matricula"]', matricula);

        count++;
        console.log(`Feito ${count} de ${jsonLeght}`)
    }


    // await browser.close();
})();
