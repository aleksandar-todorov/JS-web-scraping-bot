const request = require('request');
const rp = require('request-promise');
const cheerio = require('cheerio');
const xl = require('excel4node');

const wb = new xl.Workbook();
const ws = wb.addWorksheet('Sheet 1');
const style = wb.createStyle({
    font: {
        bold: true
    },
});
ws.column(1).setWidth(40);
ws.column(2).setWidth(20);
ws.column(3).setWidth(15);
ws.column(4).setWidth(15);
ws.column(5).setWidth(15);
ws.column(6).setWidth(15);
/**
 * first row of data in the excel table with bold style
 */
ws.cell(1, 1).string("Company").style(style);
ws.cell(1, 2).string("Address").style(style);
ws.cell(1, 3).string("Phone").style(style);
ws.cell(1, 4).string("Fax").style(style);
ws.cell(1, 5).string("Email").style(style);
ws.cell(1, 6).string("Contact Person").style(style);
ws.cell(1, 7).string("Website").style(style);

let names = [];
let addresses = [];
let phones = [];
let faxes = [];
let emails = [];
let contactPeople = [];
let websites = [];
/**
 * try to get the data from the website
 */
rp('https://www.ok-power.de/fuer-strom-kunden/anbieter-uebersicht.html', function (err, resp, html) {
    if (!err) {
        const $ = cheerio.load(html);
        $('table').each(function (index) {
            const name = $(this).find('.row_0 .col_0').text().trim();
            const address = $(this).find('.row_2 .col_0').text().trim()
                + " " + $(this).find('.row_3 .col_0').text().trim();
            const phone = $(this).find('.row_2 .col_1').text().trim().split("Tel. ")[1] || "";
            const dataRow3Col1 = $(this).find('.row_3 .col_1').text().trim();
            let fax = "";
            /**
             * if there is Fax Number, it is always in .row_3 .col_1
             */
            if (dataRow3Col1.includes("Fax")) fax = dataRow3Col1.substring("4").trim();

            let email;
            let contactPerson = "";
            let website;
            let dataCol1;
            /**
             * take the data from .col_1 but only from row2 to row6
             */
            for (let i = 2; i < 7; i++) {
                $(this).find(`.row_${i} .col_1`).each(function () {
                    dataCol1 += $(this).text() + " ";
                });
            }
            email = dataCol1.match("[\\w-.]+@[\\w-]+[.de|.com|.net]+") || "";
            website = dataCol1.match("www.[\\w-]+.\\w+") || "";
            const regex = /((Ansprechpartner:|Ansprechpartnerin:|Kontaktperson:|Abspraechpartnerin:)\s+)*([A-Z][a-zäöüß]+\s([A-Z][a-zA-Zäöüß-]+\s)+)/g;
            const match = regex.exec(dataCol1);
            if (match) {
                contactPerson = match[3].trim();
            }
            /**
             * add the data into arrays
             */
            names.push(name);
            addresses.push(address);
            phones.push(phone);
            faxes.push(fax);
            emails.push(email[0] || "");
            contactPeople.push(contactPerson);
            websites.push(website[0] || "");

        });
    } else {
        console.log("Couldn't take the data from the website.");
    }
}).then(function () {
    /**
     * transfer the data from the arrays into the excel table
     */
    for (let i = 0; i < names.length ; i++) {
            ws.cell(i + 2 , 1).string(names[i]);
            ws.cell(i + 2 , 2).string(addresses[i]);
            ws.cell(i + 2 , 3).string(phones[i]);
            ws.cell(i + 2 , 4).string(faxes[i]);
            ws.cell(i + 2 , 5).string(emails[i]);
            ws.cell(i + 2 , 6).string(contactPeople[i]);
            ws.cell(i + 2 , 7).string(websites[i]);
    }
    wb.write('Excel.xlsx');
});
