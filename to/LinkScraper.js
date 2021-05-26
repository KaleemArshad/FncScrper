const reader = require("xlsx");
const puppeteer = require("puppeteer-extra");
const pluginStealth = require("puppeteer-extra-plugin-stealth");
const userAgent = require("user-agents");
const proxyChain = require("proxy-chain");
const Excel = require("exceljs");
puppeteer.use(pluginStealth());

const ScrapeLinks = async () => {
  try {
    // const Brands = [];
    // const brands = reader.readFile(__dirname + "/Brands.xlsx");
    // const keyField = reader.utils.sheet_to_csv(
    //   brands.Sheets[brands.SheetNames[0]]
    // );
    // keyField.split("\n").forEach((row) => {
    //   Brands.push(row);
    // });

    const Links = [];
    const links = reader.readFile(__dirname + "/Links.xlsx");
    const keyField2 = reader.utils.sheet_to_csv(
      links.Sheets[links.SheetNames[0]]
    );
    keyField2.split("\n").forEach((row) => {
      Links.push(row);
    });
    // const Links = [
    //   "https://www.fnac.es/e8066/Saldana",
    //   "https://www.fnac.es/e6940/Magix",
    // ];

    const oldProxyUrl =
      "http://lum-customer-c_ef78a635-zone-data_center:qid5pp9zjd0e@zproxy.lum-superproxy.io:22225";
    const newProxyUrl = await proxyChain.anonymizeProxy(oldProxyUrl);
    const browser = await puppeteer.launch({
      headless: false,
      executablePath:
        "C://Program Files//Google//Chrome//Application//chrome.exe",
      ignoreDefaultArgs: ["--disable-extensions", "--enable-automation"],
      args: [
        "--start-maximized",
        `--proxy-server=${newProxyUrl}`,
        "--force-device-scale-factor=0.75",
      ],
      ignoreHTTPSErrors: true,
      slowMo: 250,
    });
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(0);
    await page.setDefaultTimeout(0);
    await page.setUserAgent(userAgent.toString());
    await page.viewport({
      width: 1024 + Math.floor(Math.random() * 100),
      height: 768 + Math.floor(Math.random() * 100),
      //   deviceScaleFactor: 0.75,
    });
    await page.setRequestInterception(true);
    page.on("request", async (req) => {
      const url = req.url();
      if (
        url.endsWith("oct8ne-api-2.3.js") ||
        url.endsWith("launch-385ae949e4fb.min.js") ||
        url.endsWith("U9GF8-WZXVN-ZKGGU-XY7JQ-2MKU8") ||
        url.endsWith("otBannerSdk.js") ||
        url.endsWith("script.js") ||
        url.endsWith("web.use-ext-href.js") ||
        url.endsWith("web.use-ext-href.js")
      ) {
        req.abort();
      } else {
        req.continue();
      }
    });
    var Data = [];
    // var Names = [];
    // var names = [];
    for (let i = 0; i <= Links.length - 1; i++) {
      // console.log(Brands[i]);
      const link = Links[i];
      await page.goto(link, { timeout: 0 });
      if (await page.$(".Header__nav")) {
        if (await page.$(".js-ToolbarPager > li.nextLevel > a")) {
          while (await page.$(".js-ToolbarPager > li.nextLevel > a")) {
            const xpath_expression =
              '//*[(@id = "dontTouchThisDiv")]//*[contains(concat( " ", @class, " " ), concat( " ", "js-Search-hashLink", " " ))]';
            await page.waitForXPath(xpath_expression, { timeout: 0 });
            const links = await page.$x(xpath_expression);
            const link_urls = await page.evaluate((...links) => {
              return links.map((e) => e.href);
            }, ...links);
            link_urls.forEach((e) => {
              var data = {
                link: e,
              };
              Data.push(data);
              // Names.push(Brands[i]);
            });
            // Names.forEach((e) => {
            //   var data = {
            //     brand: e,
            //   };
            //   names.push(data);
            // });
            let workbook = new Excel.Workbook();
            let worksheet = workbook.addWorksheet("Sheet1");
            worksheet.columns = [{ header: "Link", key: "link" }];
            worksheet.columns.forEach((column) => {
              column.width =
                column.header.length < 12 ? 12 : column.header.length;
            });
            worksheet.getRow(1).font = { bold: true };
            Data.forEach((e) => {
              worksheet.addRow({
                ...e,
              });
            });
            workbook.xlsx.writeFile("Product__Links.xlsx");

            // let workbook2 = new Excel.Workbook();
            // let worksheet2 = workbook2.addWorksheet("Sheet1");
            // worksheet2.columns = [{ header: "Brand", key: "brand" }];
            // worksheet2.columns.forEach((column) => {
            //   column.width =
            //     column.header.length < 12 ? 12 : column.header.length;
            // });
            // worksheet2.getRow(1).font = { bold: true };
            // names.forEach((e) => {
            //   worksheet2.addRow({
            //     ...e,
            //   });
            // });
            // workbook2.xlsx.writeFile("Brands___Names.xlsx");
            console.log(link_urls);
            console.log("\n");
            await page.waitForTimeout(3000);
            const btn = await page.evaluate(() => {
              let Btn = document.querySelector(
                ".js-ToolbarPager > li.nextLevel > a"
              );
              return Btn;
            });
            if (btn) {
              await page.evaluate(() => {
                document
                  .querySelector(".js-ToolbarPager > li.nextLevel > a")
                  .click();
              });
            }
            // await page.waitForTimeout(7000);
            await page.waitForNavigation({
              waitUntil: "networkidle0",
            });
          }
        } else {
          const xpath_expression =
            '//*[(@id = "dontTouchThisDiv")]//*[contains(concat( " ", @class, " " ), concat( " ", "js-Search-hashLink", " " ))]';
          await page.waitForXPath(xpath_expression, { timeout: 0 });
          const links = await page.$x(xpath_expression);
          const link_urls = await page.evaluate((...links) => {
            return links.map((e) => e.href);
          }, ...links);
          link_urls.forEach((e) => {
            var data = {
              link: e,
            };
            Data.push(data);
            // Names.push(Brands[i]);
          });
          // Names.forEach((e) => {
          //   var data = {
          //     brand: e,
          //   };
          //   names.push(data);
          // });
          let workbook = new Excel.Workbook();
          let worksheet = workbook.addWorksheet("Sheet1");
          worksheet.columns = [{ header: "Link", key: "link" }];
          worksheet.columns.forEach((column) => {
            column.width =
              column.header.length < 12 ? 12 : column.header.length;
          });
          worksheet.getRow(1).font = { bold: true };
          Data.forEach((e) => {
            worksheet.addRow({
              ...e,
            });
          });
          workbook.xlsx.writeFile("Product__Links.xlsx");

          // let workbook2 = new Excel.Workbook();
          // let worksheet2 = workbook2.addWorksheet("Sheet1");
          // worksheet2.columns = [{ header: "Brand", key: "brand" }];
          // worksheet2.columns.forEach((column) => {
          //   column.width =
          //     column.header.length < 12 ? 12 : column.header.length;
          // });
          // worksheet2.getRow(1).font = { bold: true };
          // names.forEach((e) => {
          //   worksheet2.addRow({
          //     ...e,
          //   });
          // });
          // workbook2.xlsx.writeFile("Brands___Names.xlsx");
          console.log(link_urls);
          console.log("\n");
        }
      } else {
        // await page.waitForTimeout(20000);
        await page.waitForSelector(".Header__nav", { timeout: 0 });
        if (await page.$(".js-ToolbarPager > li.nextLevel > a")) {
          while (await page.$(".js-ToolbarPager > li.nextLevel > a")) {
            const xpath_expression =
              '//*[(@id = "dontTouchThisDiv")]//*[contains(concat( " ", @class, " " ), concat( " ", "js-Search-hashLink", " " ))]';
            await page.waitForXPath(xpath_expression, { timeout: 0 });
            const links = await page.$x(xpath_expression);
            const link_urls = await page.evaluate((...links) => {
              return links.map((e) => e.href);
            }, ...links);
            link_urls.forEach((e) => {
              var data = {
                link: e,
              };
              Data.push(data);
              // Names.push(Brands[i]);
            });
            // Names.forEach((e) => {
            //   var data = {
            //     brand: e,
            //   };
            //   names.push(data);
            // });
            let workbook = new Excel.Workbook();
            let worksheet = workbook.addWorksheet("Sheet1");
            worksheet.columns = [{ header: "Link", key: "link" }];
            worksheet.columns.forEach((column) => {
              column.width =
                column.header.length < 12 ? 12 : column.header.length;
            });
            worksheet.getRow(1).font = { bold: true };
            Data.forEach((e) => {
              worksheet.addRow({
                ...e,
              });
            });
            workbook.xlsx.writeFile("Product__Links.xlsx");

            // let workbook2 = new Excel.Workbook();
            // let worksheet2 = workbook2.addWorksheet("Sheet1");
            // worksheet2.columns = [{ header: "Brand", key: "brand" }];
            // worksheet2.columns.forEach((column) => {
            //   column.width =
            //     column.header.length < 12 ? 12 : column.header.length;
            // });
            // worksheet2.getRow(1).font = { bold: true };
            // names.forEach((e) => {
            //   worksheet2.addRow({
            //     ...e,
            //   });
            // });
            // workbook2.xlsx.writeFile("Brands___Names.xlsx");
            console.log(link_urls);
            console.log("\n");
            await page.waitForTimeout(3000);
            const btn = await page.evaluate(() => {
              let Btn = document.querySelector(
                ".js-ToolbarPager > li.nextLevel > a"
              );
              return Btn;
            });
            if (btn) {
              await page.evaluate(() => {
                document
                  .querySelector(".js-ToolbarPager > li.nextLevel > a")
                  .click();
              });
            }
            // await page.waitForTimeout(7000);
            await page.waitForNavigation({
              waitUntil: "networkidle0",
            });
          }
        } else {
          const xpath_expression =
            '//*[(@id = "dontTouchThisDiv")]//*[contains(concat( " ", @class, " " ), concat( " ", "js-Search-hashLink", " " ))]';
          await page.waitForXPath(xpath_expression, { timeout: 0 });
          const links = await page.$x(xpath_expression);
          const link_urls = await page.evaluate((...links) => {
            return links.map((e) => e.href);
          }, ...links);
          link_urls.forEach((e) => {
            var data = {
              link: e,
            };
            Data.push(data);
            // Names.push(Brands[i]);
          });
          // Names.forEach((e) => {
          //   var data = {
          //     brand: e,
          //   };
          //   names.push(data);
          // });
          let workbook = new Excel.Workbook();
          let worksheet = workbook.addWorksheet("Sheet1");
          worksheet.columns = [{ header: "Link", key: "link" }];
          worksheet.columns.forEach((column) => {
            column.width =
              column.header.length < 12 ? 12 : column.header.length;
          });
          worksheet.getRow(1).font = { bold: true };
          Data.forEach((e) => {
            worksheet.addRow({
              ...e,
            });
          });
          workbook.xlsx.writeFile("Product__Links.xlsx");

          // let workbook2 = new Excel.Workbook();
          // let worksheet2 = workbook2.addWorksheet("Sheet1");
          // worksheet2.columns = [{ header: "Brand", key: "brand" }];
          // worksheet2.columns.forEach((column) => {
          //   column.width =
          //     column.header.length < 12 ? 12 : column.header.length;
          // });
          // worksheet2.getRow(1).font = { bold: true };
          // names.forEach((e) => {
          //   worksheet2.addRow({
          //     ...e,
          //   });
          // });
          // workbook2.xlsx.writeFile("Brands___Names.xlsx");
          console.log(link_urls);
          console.log("\n");
        }
      }
    }
  } catch (error) {
    console.log(error);
  }
};

ScrapeLinks();
