const axios = require("axios");
const cheerio = require("cheerio");
const excelJs = require("exceljs");

async function scrapeData() {
    try {
        const response = await axios.get("https://www.amazon.in/s?k=keyboard", {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
        });
        const $ = cheerio.load(response.data);
        const products = [];

        $(".s-main-slot .s-result-item").each((index, element) => {
            const name = $(element).find("h2 .a-text-normal").text().trim();
            const price = $(element).find(".a-price-whole").text().trim();
            const rating = $(element).find(".a-icon-alt").text().trim();
            const availability = $(element).find(".a-size-medium.a-color-price").text().trim() || 'In Stock';

            if (name || price || rating || availability) {
                products.push({ name, price, rating, availability });
            }
        });

        console.log(products); // Debug log to check extracted data
        return products;
    } catch (error) {
        console.error("Error scraping data:", error);
        return [];
    }
}

async function saveToExcel(products) {
    const workbook = new excelJs.Workbook();
    const worksheet = workbook.addWorksheet("Products");
    worksheet.columns = [
        { header: 'Product Name', key: 'name', width: 30 },
        { header: 'Price', key: 'price', width: 15 },
        { header: 'Availability', key: 'availability', width: 15 },
        { header: 'Product Rating', key: 'rating', width: 15 },
    ];

    // Log the products data to verify it's being processed correctly
    console.log('Saving to Excel:', products);

    products.forEach(product => {
        worksheet.addRow(product);
    });

    try {
        await workbook.xlsx.writeFile("products.xlsx");
        console.log('Data successfully saved to products.xlsx');
    } catch (error) {
        console.error('Error saving to Excel:', error);
    }
}

async function main() {
    try {
        const products = await scrapeData();
        await saveToExcel(products);
    } catch (err) {
        console.log(err.message);
    }
}

main();
