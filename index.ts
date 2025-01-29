import { chromium, type Page } from "playwright";
import XLSX from "xlsx";
import { eraseLines } from "ansi-escapes";

interface IBook {
  title: string;
  category: string;
  price: string;
  stock: "In stock" | "Out stock" | "N/A";
  available: number;
  rating: number;
  upc: string;
  description: string;
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.exposeFunction("parseWordToNum", parseWordToNum);

  await page.goto("https://books.toscrape.com/");

  const BookList: Array<IBook> = [];

  // get categories
  const categoriesEl = await page.locator("ul.nav ul li a").all();

  for (const categoryEl of categoriesEl) {
    let category = await categoryEl.innerText();
    await categoryEl.click();

    // get books in a each category page
    const booksEl = await page.locator(".product_pod").all();
    for (const bookEl of booksEl) {
      await bookEl.locator("h3 a").click();

      const newBook = new Book(page, { category });
      await newBook.scrape();

      BookList.push(newBook.toJSON());

      writeLog(
        "Added Book #" + BookList.length,
        ":",
        newBook.title.length > 50 ? newBook.title.slice(0, 50) : newBook.title,
        "..."
      );

      await page.goBack();
    }
  }

  writeExcel(BookList);

  await browser.close();
})();

function parseWordToNum(word: string): number {
  const wordsToNumbers = new Map([
    ["zero", 0],
    ["one", 1],
    ["two", 2],
    ["three", 3],
    ["four", 4],
    ["five", 5],
  ]);
  return wordsToNumbers.get(word.toLowerCase()) ?? NaN;
}

function writeExcel(jsonData: Array<any>) {
  const worksheet = XLSX.utils.json_to_sheet(jsonData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Books");

  XLSX.writeFile(workbook, "books.xlsx");
  writeLog(`📂 Excel file created! (${jsonData.length} books saved)`);
}

function writeLog(...str: string[]) {
  process.stdout.write(eraseLines(2));
  console.log(...str);
}

class Book implements IBook {
  readonly page: Page;
  title = "";
  category = "";
  price = "";
  stock: "In stock" | "Out stock" | "N/A" = "N/A";
  available = 0;
  rating = 0;
  upc = "";
  description = "";

  constructor(page: Page, bookData: Partial<IBook>) {
    this.page = page;
    Object.assign(this, bookData);
  }
  async scrape() {
    try {
      [
        this.title,
        this.price,
        this.stock,
        this.available,
        this.rating,
        this.upc,
        this.description,
      ] = await Promise.all([
        this.getText(".product_main h1"),
        this.getText(".product_main .price_color"),
        this.getStock(),
        this.getAvailable(),
        this.getRating(),
        this.getUPC(),
        this.getText("#product_description + p"),
      ]);
    } catch (err) {
      console.error("❌ Error scraping book:", err);
    }
  }
  async getText(selector: string) {
    return await this.page.locator(selector).innerText();
  }
  async getStock() {
    const available = await this.page
      .locator(".product_main .availability")
      .innerText();
    return available.split("(")[0].trim() == "In stock"
      ? "In stock"
      : "Out stock";
  }
  async getAvailable() {
    const stockText = await this.getText(".product_main .availability");
    return parseInt(stockText.match(/\d+/)?.[0] || "0", 10);
  }
  async getRating() {
    const ratingClass = await this.page
      .locator(".product_main .star-rating")
      .getAttribute("class");

    return parseWordToNum(
      ratingClass?.replace("star-rating", "").trim() || "zero"
    );
  }
  async getUPC() {
    return await this.page.locator('th:text("UPC") + td').innerText();
  }
  toJSON(): IBook {
    return {
      title: this.title,
      category: this.category,
      price: this.price,
      stock: this.stock,
      available: this.available,
      rating: this.rating,
      upc: this.upc,
      description: this.description,
    };
  }
}
