//Load libraries 
const puppeteer = require('puppeteer');
const Excel = require('exceljs');

(async ()=>{

    const browser = await puppeteer.launch();

    //Open new tab
    const page = await browser.newPage();
    // Deku Deals URL public wishlist
    await page.goto("Insert URL here");

    let gamesHandle = await page.$$(".items-table3 .items-list-row");
    let promises = [];

    gamesHandle.forEach(game => {
        let pricePromise = game.$$eval('.price', (nodes) => {
            return nodes.map(n => n.textContent);
        });
        let namePromise = game.$$eval('.main .w-100 .name', (nodes) => {
            return nodes.map(n => n.textContent);
        });
        
        promises.push(Promise.all([pricePromise, namePromise]));
    });

    let gamesArray = await Promise.all(promises);
    //Print gamesArray raw
    console.log(gamesArray);

    //Let's organize array
    let games = gamesArray.map(g=>{
        let price = g[0][0].trim();
        if (price.indexOf("-") >= 0) {
            price = price.replace("-", " -").replace("\n", " -> ").replace("%", "% ");
        }
        return {
            name: g[1][0].trim(),
            price: price,
        };
    });
    
    //Print games list
    console.log(games);
    await browser.close();

    let workbook = new Excel.Workbook(); 
    await workbook.xlsx.readFile("GamePrices.xlsx").catch((err) => {
    // If file does not exist, let's create it
        // Create a sheet
        var sheet = workbook.addWorksheet("Wishlist");
    });
    
    let sheet = workbook.getWorksheet("Wishlist");
    sheet.columns = [
        { header: 'Name', key: 'name', width: 50 },
        { header: 'Price', key: 'price', width: 40 },
        { header: 'Date', key: 'date', width: 12 },
    ];

    for (game of games) {
        // Add rows in the above header
        let row = sheet.addRow({
            name: game.name,
            price: game.price,
            date: new Date(),
        });
        if (game.price.indexOf("Lowest price ever") >= 0) {
            row.font = {
                color: { argb: 'FF4E62'},
                bold: true
            }
        };
    }

    workbook.xlsx.writeFile("GamePrices.xlsx").then(function () {
        // Success Message
        console.log("Games & Prices Saved");
    });

})().catch((err) => {
    console.error(err);
});