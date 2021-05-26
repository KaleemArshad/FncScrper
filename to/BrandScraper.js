// const ProxyGenerator = require("./proxyRotator");
const puppeteer = require("puppeteer-extra");
const pluginStealth = require("puppeteer-extra-plugin-stealth");
const userAgent = require("user-agents");
const proxyChain = require("proxy-chain");
const Excel = require("exceljs");
// const translate = require("free-google-translate-with-puppeteer");
const translate = require("@iamtraction/google-translate");
// const { proxyRequest } = require("puppeteer-proxy");
// const useProxy = require("puppeteer-page-proxy");
puppeteer.use(pluginStealth());

const ScrapeBrands = async () => {
  try {
    var BrandLinks = [];
    var BrandName = [];
    // const oldProxyUrl = `http://lum-customer-c_ef78a635-zone-data_center-ip-${ProxyGenerator()}@zproxy.lum-superproxy.io:22225`;
    const oldProxyUrl =
      "http://lum-customer-c_ef78a635-zone-data_center:qid5pp9zjd0e@zproxy.lum-superproxy.io:22225";
    const newProxyUrl = await proxyChain.anonymizeProxy(oldProxyUrl);
    const browser = await puppeteer.launch({
      headless: false,
      executablePath:
        "C://Program Files//Google//Chrome//Application//chrome.exe",
      // userDataDir:
      //   "C://Users//sajaw//AppData//Local//Google//Chrome//User Data//Default",
      ignoreDefaultArgs: ["--disable-extensions", "--enable-automation"],
      args: [
        "--start-maximized",
        `--proxy-server=${newProxyUrl}`,
        // '--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36"',
      ],
      ignoreHTTPSErrors: true,
      slowMo: 250,
    });
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(0);
    await page.setUserAgent(userAgent.toString());
    await page.viewport({
      width: 1024 + Math.floor(Math.random() * 100),
      height: 768 + Math.floor(Math.random() * 100),
    });
    await page.setRequestInterception(true);
    page.on("request", async (req) => {
      //   useProxy(req, `${newProxyUrl}`);
      const url = req.url();
      if (
        // req.resourceType() == "stylesheet" ||
        // req.resourceType() == "font" ||
        url.endsWith("oct8ne-api-2.3.js") ||
        url.endsWith("launch-385ae949e4fb.min.js") ||
        url.endsWith("U9GF8-WZXVN-ZKGGU-XY7JQ-2MKU8") ||
        url.endsWith("otBannerSdk.js") ||
        url.endsWith("script.js") ||
        url.endsWith("web.use-ext-href.js") ||
        url.endsWith("web.use-ext-href.js")
        // url.endsWith("c.js")
      ) {
        req.abort();
      } else {
        req.continue();
      }
    });
    const url = "https://www.fnac.es/marcas";
    await page.goto(url, { timeout: 0 });
    if (await page.$(".Header__nav")) {
      const xpath_expression =
        '//*[contains(concat( " ", @class, " " ), concat( " ", "Marques_StrateIndex-link", " " ))]';
      await page.waitForXPath(xpath_expression, { timeout: 0 });
      const links = await page.$x(xpath_expression);
      const link_urls = await page.evaluate((...links) => {
        return links.map((e) => e.href);
      }, ...links);

      link_urls.forEach((link, index) => {
        BrandLinks[index] = {
          url: link,
        };
      });
      console.log(link_urls);

      const xpath_expression2 =
        '//*[contains(concat( " ", @class, " " ), concat( " ", "Marques_StrateIndex-link", " " ))]';
      await page.waitForXPath(xpath_expression2, { timeout: 0 });
      const names = await page.$x(xpath_expression2);
      const brand_names = await page.evaluate((...names) => {
        return names.map((f) => f.textContent);
      }, ...names);
      brand_names.forEach((brand, i) => {
        BrandLinks[i] = {
          name: brand,
        };
      });
      console.log(brand_names);
      return {
        brand: BrandName,
        link: BrandLinks,
      };
    } else {
      //   await page.waitForTimeout(20000);
      await page.waitForSelector(".Header__nav", { timeout: 0 });
      const xpath_expression =
        '//*[contains(concat( " ", @class, " " ), concat( " ", "Marques_StrateIndex-link", " " ))]';
      await page.waitForXPath(xpath_expression, { timeout: 0 });
      const links = await page.$x(xpath_expression);
      const link_urls = await page.evaluate((...links) => {
        return links.map((e) => e.href);
      }, ...links);

      link_urls.forEach((link, index) => {
        BrandLinks[index] = {
          url: link,
        };
      });
      console.log(BrandLinks);

      const xpath_expression2 =
        '//*[contains(concat( " ", @class, " " ), concat( " ", "Marques_StrateIndex-link", " " ))]';
      await page.waitForXPath(xpath_expression2, { timeout: 0 });
      const names = await page.$x(xpath_expression2);
      const brand_names = await page.evaluate((...names) => {
        return names.map((f) => f.textContent);
      }, ...names);
      brand_names.forEach((brand, i) => {
        BrandName[i] = {
          name: brand,
        };
      });
      console.log(BrandName);
      return {
        brand: BrandName,
        link: BrandLinks,
      };
    }
  } catch (error) {
    console.log(error);
  }
};

// ScrapeLinks();

const insertData = async () => {
  try {
    const Data = await ScrapeBrands();
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet("Sheet1");
    worksheet.columns = [{ header: "Brand Link", key: "url" }];
    worksheet.columns.forEach((column) => {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    });
    worksheet.getRow(1).font = { bold: true };
    Data.link.forEach((e) => {
      worksheet.addRow({
        ...e,
      });
    });
    workbook.xlsx.writeFile("Links.xlsx");

    let workbook2 = new Excel.Workbook();
    let worksheet2 = workbook2.addWorksheet("Sheet1");
    worksheet2.columns = [{ header: "Brand Name", key: "name" }];
    worksheet2.columns.forEach((column) => {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    });
    worksheet2.getRow(1).font = { bold: true };
    Data.brand.forEach((e) => {
      worksheet2.addRow({
        ...e,
      });
    });
    workbook2.xlsx.writeFile("Brands.xlsx");
  } catch (err) {
    console.log(err);
  }
};

insertData();

// captcha__human__title
