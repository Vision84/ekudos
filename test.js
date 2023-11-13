const { chromium } = require('playwright'), fs = require('fs'), HTMLParser = require('node-html-parser'), axios = require('axios');

// senderNames = ["Alistair Keiller"];
// senderEmails = ["alistair@keiller.net"];
senderNames = [];
senderEmails = [];
recipientNames = [];
recipientEmails = [];

counter = 0;

// for(let i = 1; i < 36; i++) {
//     const d = HTMLParser.parse(fs.readFileSync(`recipient/p${i}.html`, 'utf8'));
//     const names = d.querySelectorAll('.box-member__title');
//     for(let j = 0; j < names.length; j++)
//         recipientNames.push(names[j].firstChild._rawText);
//     const emails = d.querySelectorAll('.icon-text-sidebar__content.link-has-underline.list-in-article.typography-space-small');
//     for(let j = 0; j < emails.length; j++)
//         if(emails[j].firstChild && emails[j].firstChild._rawText.includes("@ohs.stanford.edu"))
//             recipientEmails.push(emails[j].firstChild._rawText);
// }
// for(let i = 1; i < 8; i++) {
//     const d = HTMLParser.parse(fs.readFileSync(`sender/p${i}.html`, 'utf8'));
//     const names = d.querySelectorAll('.box-member__title');
//     for(let j = 0; j < names.length; j++)
//         senderNames.push(names[j].firstChild._rawText);
//     const emails = d.querySelectorAll('.icon-text-sidebar__content.link-has-underline.list-in-article.typography-space-small');
//     for(let j = 0; j < emails.length; j++)
//         if(emails[j].firstChild && emails[j].firstChild._rawText.includes("@ohs.stanford.edu"))
//             senderEmails.push(emails[j].firstChild._rawText);
// }
const Excel = require('exceljs');

var workbook = new Excel.Workbook();
workbook.xlsx.readFile("ninth.xlsx")
    .then(function() {
        const worksheet = workbook.getWorksheet('Sheet1');
        for (let i = 1; i <= 181; i++) {
            if (i < 90) {
                const row = worksheet.getRow(i);
                senderEmails.push(row.getCell(3).value);
                senderNames.push(row.getCell(2).value);
            } else {
                const row = worksheet.getRow(i);
                recipientEmails.push(row.getCell(3).value);
                recipientNames.push(row.getCell(2).value);
            }
        }
    });

    //proxy: {server: proxyServer}

async function fillForm(proxyServer) {
    let browser;
    try {
    browser = await chromium.launch({headless: false, proxy: {server: proxyServer}});
    const page = await browser.newPage();
    await page.goto('https://forms.zohopublic.com/stanfordonlinehighschoolsg/form/SpiritWeekEkudos/formperma/WceKhelCTqzeq8et4MP-5Ghp-mq2CfY-PzxhJoIRPvA', {referer: "https://spiritweek.ohsstudentgov.com/"});

    await page.getByText('Spirit Week Ekudos').waitFor();

    senderIDX = Math.floor(Math.random()*senderNames.length);
    //await page.waitForTimeout(Math.random()*3000+3000);
    await page.isVisible("#Email-arialabel");
    console.log(senderEmails[senderIDX]);
    await page.locator("#Email-arialabel").fill(senderEmails[senderIDX])
    //await page.fill('#Email-arialabel', senderEmails[senderIDX]);
    // await page.waitForTimeout(Math.random()*3000+3000);
    const nameArray = senderNames[senderIDX].split(" ");
    console.log(nameArray[0])
    await page.isVisible("[aria-labelledby='Name-arialabel aria-showelemslabel-Name0 ariarequired-Name0']");
    await page.locator("[aria-labelledby='Name-arialabel aria-showelemslabel-Name0 ariarequired-Name0']").fill(nameArray[0]);
    //await page.waitForTimeout(Math.random()*1000+3000);
    await page.locator("[aria-labelledby='Name-arialabel aria-showelemslabel-Name1 ariarequired-Name1']").fill(nameArray[nameArray.length - 1]);

    recipientIDX = Math.floor(Math.random()*recipientNames.length)
    const recipientArray = recipientNames[recipientIDX].split(" ");
    //await page.waitForTimeout(Math.random()*3000+3000);
    await page.locator("[aria-labelledby='Name1-arialabel aria-showelemslabel-Name10 ariarequired-Name10']").fill(recipientArray[0]);
    //await page.waitForTimeout(Math.random()*3000+3000);
    await page.locator("[aria-labelledby='Name1-arialabel aria-showelemslabel-Name11 ariarequired-Name11']").fill(recipientArray[recipientArray.length - 1]);
    // await page.waitForTimeout(Math.random()*3000+3000);
    await page.fill('#Email1-arialabel', recipientEmails[recipientIDX]);

    //await page.waitForTimeout(Math.random()*3000+3000);
    await page.click('[for="Radio_2"]');

    const possibleMessages = ["Hi", "Hello"];

    let message = possibleMessages[Math.floor(Math.random()*possibleMessages.length)];

    //await page.waitForTimeout(message.length*1000/500);
    await page.fill('#MultiLine-arialabel', possibleMessages[Math.floor(Math.random()*possibleMessages.length)]);

    //await page.waitForTimeout(Math.random()*3000+3000);
    await page.click('[for="Radio1_2"]');

    //await page.waitForTimeout(Math.random()*3000+3000);
    await page.click('[value="submit"]');
    //await page.waitForTimeout(Math.random()*3000+3000);
    console.log(++counter);
    } catch (e) {}
    try{
    await browser.close();
    } catch (e) {}
}

(async () => {
    const proxyServers = (await axios.get('https://api.proxyscrape.com/v2/?request=displayproxies&protocol=http&timeout=1000&country=all&ssl=all&anonymity=all')).data.split("\r\n");
    while(true) {
        fillForm(proxyServers[Math.floor(Math.random()*proxyServers.length)]);
        await new Promise(r => setTimeout(r, 2000));
    }
})()