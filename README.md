# JS-web-scraping-bot
Scraping bot to read and download information from public website and save it to excel file.

Да се напише скрипт Javascript (NodeJS) за да свалите данни от публичен сайт - 
https://www.ok-power.de/fuer-strom-kunden/anbieter-uebersicht.html

Трябва да се свалят контактите от таблица Ökostromanbieter mit zertifzierten Produkten и да се генерира excel-ски документ със следните колини: Company, Address, Phone, Fax, Email, Contact Person, Website. Да се има в предвид, че не за всички контакти е посочена тази информация.

За парсването на HTML-a на сайта да се използва библиотеката cheerio.
За генерирането на excel файл да се използва библиотеката excel4node.
Резултатът трябва да е excel-ски документ в табличен вид.
