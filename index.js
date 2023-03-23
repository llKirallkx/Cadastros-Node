const XLSX = require('xlsx');
const fs = require('fs');
const puppeteer = require('puppeteer');


const workbook = XLSX.readFile('IMPLANTAÇÃO_PONTOFOPAG.xlsx');
const sheetName = 'Dados_Funcionários'; // nome do sheet
const worksheet = workbook.Sheets[sheetName];

// Converter a planilha para um array de objetos JavaScript
const jsonData = XLSX.utils.sheet_to_json(worksheet);
const jsonLeght = jsonData.length;
let count = 0;

// DE - PARA
const matricula = '* Matrícula:';

// // Converter o array em uma string JSON
// const jsonString = JSON.stringify(jsonData);


// // Gravar a string JSON em um arquivo
// fs.writeFileSync('dados_funcionarios.json', jsonString, (err) => {
//   if (err) {
//     console.error('Erro ao gravar arquivo:', err);
//   } else {
//     console.log('Arquivo gravado com sucesso!');
//   }
// });


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
    await page.waitForTimeout(2000);
    await page.waitForSelector('[title="Sair"]');
    
    while (count < jsonLeght){

        await page.goto('https://prd.pontofopag.com.br/Funcionario/Grid');
        await page.waitForSelector('#btIncluir');
        await page.click('#btIncluir');
        await page.type('[name="Matricula"]', matricula );


    }


    // await browser.close();
})();
