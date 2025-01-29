import { chromium } from "playwright";
import XLSX from "xlsx";
import { eraseLines } from "ansi-escapes";

interface book {
  title: string;
  category: string;
  price: string;
  stock: "In stock" | "Out stock";
  available: number;
  rating: number;
  upc: string;
  description: string;
}

declare global {
  interface Window {
    parseWordToNum: (word: string) => number;
  }
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.exposeFunction("parseWordToNum", parseWordToNum);

  await page.goto("https://books.toscrape.com/");

  const BookList: Array<book> = [];

  // get categories
  const categoriesEl = await page.locator("ul.nav ul li a").all();
  for (const categoryEl of categoriesEl) {
    let category = await categoryEl.innerText();

    await categoryEl.click();

    // get books in a each category page
    const booksEl = await page.locator(".product_pod").all();

    for (const bookEl of booksEl) {
      let newBook: book = {} as book;

      await bookEl.locator("h3 a").click();
      newBook.title = await page.locator(".product_main h1").innerText();
      newBook.category = category;
      newBook.price = await page
        .locator(".product_main .price_color")
        .innerText();
      newBook.stock = (await page.locator(".product_main .instock").isVisible())
        ? "In stock"
        : "Out stock";
      newBook.available = (await page
        .locator(".product_main .instock")
        .isVisible())
        ? await page
            .locator(".product_main .instock")
            .evaluate((v) =>
              parseInt(v.textContent?.match(/\d+/)?.[0] || "0", 10)
            )
        : 0;
      newBook.rating = await page
        .locator(".product_main .star-rating")
        .evaluate((v) =>
          window.parseWordToNum(
            v.getAttribute("class")?.replace("star-rating", "")?.trim() ||
              "zero"
          )
        );
      newBook.upc = await page.locator('th:text("UPC")')
                              .locator('xpath=following-sibling::td')
                              .innerText();
      newBook.description = await page.locator('#product_description + p').innerText()

      BookList.push(newBook);
      writeLog("Added Book #"+BookList.length, ":", newBook.title, "...");

      await page.goBack();
    }
  }

  writeExcel(BookList);

  await browser.close();
})();

function parseWordToNum(word: string): number {
  const wordsToNumbers: Record<string, number> = {
    zero: 0,
    one: 1,
    two: 2,
    three: 3,
    four: 4,
    five: 5,
  };
  return wordsToNumbers[word.toLowerCase()] ?? NaN;
}

function writeExcel(jsonData: Array<any>) {
  // 2. Convert JSON to worksheet
  const worksheet = XLSX.utils.json_to_sheet(jsonData);

  // 3. Create workbook & append sheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Book");

  // 4. Write to file
  XLSX.writeFile(workbook, "book.xlsx");

  writeLog("Excel file created from JSON!🎉");
  console.log("There", jsonData.length, "in book.xlsx")
}

function writeLog(...str: Array<string>) {
  process.stdout.write(eraseLines(10));
  console.log(...str);
}

