const reader = require("xlsx");
const puppeteer = require("puppeteer-extra");
const pluginStealth = require("puppeteer-extra-plugin-stealth");
const userAgent = require("user-agents");
const proxyChain = require("proxy-chain");
const Excel = require("exceljs");
const translate = require("@iamtraction/google-translate");
puppeteer.use(pluginStealth());

const Translate = async (text) => {
  const Text = await translate(text.trim(), { to: "en" });
  const res = Text.text;
  return res;
};

const ScrapeLinks = async () => {
  try {
    const Brands = [];
    const brands = reader.readFile(__dirname + "/Brands.xlsx");
    const keyField = reader.utils.sheet_to_csv(
      brands.Sheets[brands.SheetNames[0]]
    );
    keyField.split("\n").forEach((row) => {
      Brands.push(row);
    });

    const Links = [];
    const links = reader.readFile(
      "C://Users//sajaw//OneDrive//Desktop//Products_Links.xlsx"
    );
    const keyField2 = reader.utils.sheet_to_csv(
      links.Sheets[links.SheetNames[0]]
    );
    keyField2.split("\n").forEach((row) => {
      Links.push(row);
    });

    // const Links = [
    //   "https://www.fnac.es/a8031082/Maria-Duenas-Sira#int=S:Destacados|Literatura%20universal:%20narrativa,%20poes%C3%ADa%20y%20teatro|39531|8031082|BL1|L1",
    //   "https://www.fnac.es/Assassin-s-Creed-Valhalla-PS4-Juego-PS4/a7430430#int=S:Ofertas%20gaming|Gaming:%20Videojuegos%20y%20Consolas|126671|7430430|BL1|L1",
    //   "https://www.fnac.es/Barra-de-sonido-Sony-HT-G700-Periferico-Electronica-Alta-fidelidad/a7511638",
    //   "https://www.fnac.es/Samsung-Galaxy-Tab-A7-10-4-64GB-Wi-Fi-Gris-Tablet-Tablet/a7654129",
    //   "https://www.fnac.es/mp8370161/Robot-Aspirador-Xiaomi-Mi-Robot-1C-Vacuum-EU/w-4?oref=653852ce-2271-976f-73c3-56ccc50c1bfa#omnsearchpos=1",
    //   "https://www.fnac.es/Tarjeta-regalo-Fnac-Felicidades-15/a7798210#omnsearchpos=1",
    //   "https://www.fnac.es/a8058614/Joaquin-Sabina-Mentiras-piadosas-Vinilo-Picture-Disc-Disco#int=:NonApplicable|NonApplicable|NonApplicable|8058614|NonApplicable|NonApplicable",
    //   "https://www.fnac.es/a697144/Juego-de-Tronos-Juego-de-tronos-Temporada-1-Blu-Ray-Emilia-Clarke?oref=00000000-0000-0000-0000-000000000000#int=:NonApplicable|NonApplicable|NonApplicable|697144|NonApplicable|NonApplicable",
    //   "https://www.fnac.es/mp8903924/Frigorifico-Combi-Bosch-KGN39VIEA-INOX-E-VitaFresh/w-4?CtoPid=127623#int=:NonApplicable|NonApplicable|NonApplicable|8903924|NonApplicable|NonApplicable",
    //   "https://www.fnac.es/mp8262330/Microondas-AEG-MSB2547DM-grill-900W-25L-acero-inox/w-4?CtoPid=127623#int=:NonApplicable|NonApplicable|NonApplicable|8262330|NonApplicable|NonApplicable",
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
    await page.viewport({
      width: 1024 + Math.floor(Math.random() * 100),
      height: 768 + Math.floor(Math.random() * 100),
      //   deviceScaleFactor: 0.75,
    });
    // await page.evaluateOnNewDocument(() => {
    //   Object.defineProperty(navigator, "platform", { get: () => "Win32" });
    //   Object.defineProperty(navigator, "productSub", { get: () => "20100101" });
    //   Object.defineProperty(navigator, "vendor", { get: () => "" });
    //   Object.defineProperty(navigator, "oscpu", {
    //     get: () => "Windows NT 10.0; Win64; x64",
    //   });
    // });
    await page.setUserAgent(userAgent.toString());
    // await page.setRequestInterception(true);
    // page.on("request", async (req) => {
    //   const url = req.url();
    //   if (
    //     // url.endsWith("stage2.min.js")
    //     // url.endsWith("oct8ne-api-2.3.js") ||
    //     // url.endsWith("launch-385ae949e4fb.min.js") ||
    //     // url.endsWith("U9GF8-WZXVN-ZKGGU-XY7JQ-2MKU8") ||
    //     // url.endsWith("otBannerSdk.js") ||
    //     // url.endsWith("script.js") ||
    //     // url.endsWith("web.use-ext-href.js") ||
    //     // url.endsWith("web.use-ext-href.js")
    //   ) {
    //     req.abort();
    //   } else {
    //     req.continue();
    //   }
    // });
    var Data = [];
    const Title = [];
    var Brand = [];
    const Price = [];
    const EAN = [];
    const Cat_Path = [];
    const Description = [];
    var Images = [];

    for (let i = 0; i <= Links.length - 1; i++) {
      const link = Links[i];
      await page.goto(link, { timeout: 0 });
      if (await page.$(".Header__nav")) {
        const p = await page.$(
          "#Characteristics > div > div > div.characteristicsStrate__lists"
        );
        console.log(i);
        const title = await page.$eval(
          ".f-productHeader-Title",
          (el) => el.textContent
        );
        const Title_Text = await Translate(title);
        Title.push(Title_Text);
        console.log(Title_Text);
        const x = await page.$(".checked");
        if (x) {
          const price = await page.$eval(".checked", (e) => e.textContent);
          Price.push(price.replace(",", "."));
          console.log(price.replace(",", "."));
        } else {
          Price.push("No Price Found");
          console.log("No Price Found");
        }
        const catPath = await page.$eval(".f-breadcrumb", (e) => e.textContent);
        const tCatPath = await Translate(catPath);
        const oCatPath = tCatPath.split("\n");
        const fCatPath = oCatPath.filter(function (str) {
          return /\S/.test(str);
        });
        let a = [];
        fCatPath.forEach((e) => {
          const f = e.trim();
          a.push(f);
        });
        Cat_Path.push(a.join(" > "));
        console.log(a.join(" > "));
        if (p) {
          await page.waitForSelector("section#Characteristics dt");
          const dt__key = await page.evaluate(() => {
            const dts = Array.from(
              document.querySelectorAll("section#Characteristics dt")
            ).map((e) => e.textContent.trim());
            return dts;
          });

          const dd__vle = await page.evaluate(() => {
            const dds = Array.from(
              document.querySelectorAll("section#Characteristics dd")
            ).map((e) => e.textContent.trim());
            return dds;
          });

          const ean__len = EAN.length;
          for (let n = 0; n <= dt__key.length; n++) {
            if (dt__key[n] !== undefined) {
              const e = dt__key[n].trim();
              if (e.indexOf("EAN") != -1) {
                EAN.push(dd__vle[n].trim());
                console.log(dd__vle[n].trim());
              }
            }
          }
          if (ean__len === EAN.length) {
            EAN.push("No EAN Found");
            console.log("No EAN Found");
          }
        } else {
          EAN.push("No EAN Found");
          console.log("No EAN Found");
        }
        const h = await page.$(".summaryStrate");
        if (h) {
          const description = await page.$eval(
            ".summaryStrate__raw",
            (e) => e.textContent
          );
          const tDes = await Translate(description);
          Description.push(tDes);
        } else {
          Description.push("No Description Found");
        }
        const images = await page.evaluate(() => {
          const srcs = Array.from(
            document.querySelectorAll(
              "body > div.Main.Main--fullWidth > div > div.productPageTop > div.f-productPage-colLeft.clearfix > section.f-productVisuals.js-articleVisuals > div.f-productVisuals-thumbnailsWrapper > div > div > div > div > img"
            )
          ).map((image) => image.getAttribute("src"));
          return srcs;
        });
        Images.push(images);

        const b__len = Brand.length;

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (p) {
          await page.waitForSelector("section#Characteristics dt");
          const dt__key = await page.evaluate(() => {
            const dts = Array.from(
              document.querySelectorAll("section#Characteristics dt")
            ).map((e) => e.textContent.trim());
            return dts;
          });

          const dd__vle = await page.evaluate(() => {
            const dds = Array.from(
              document.querySelectorAll("section#Characteristics dd")
            ).map((e) => e.textContent.trim());
            return dds;
          });
          let b = [];
          for (let n = 0; n <= dt__key.length; n++) {
            if (dt__key[n] !== undefined) {
              const e = dt__key[n].trim();
              if (e.indexOf("Fabricante/Marca") != -1) {
                b.push(dd__vle[n].trim());
                const Brand1 = b[0].toUpperCase();
                Brand.push(Brand1);
              }
            }
          }
        }

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (b__len === Brand.length) {
          Brands.forEach((e) => {
            Brand = Brand.filter(function (str) {
              return /\S/.test(str);
            });
            const st = a.join(" > ").toUpperCase().split(" ");
            st.forEach((h) => {
              Brand = Brand.filter(function (str) {
                return /\S/.test(str);
              });
              if (h == e && b__len === Brand.length) {
                Brand.push(e);
              }
            });
          });
        }

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (b__len === Brand.length) {
          const subT = await page.$eval(
            ".f-productHeader-subTitle",
            (e) => e.textContent
          );
          const f_sub = subT.split("\n").filter(function (str) {
            return /\S/.test(str);
          });
          for (let n = 0; n <= f_sub.length; n++) {
            Brand = Brand.filter(function (str) {
              return /\S/.test(str);
            });
            if (f_sub[n] !== undefined) {
              const e = f_sub[n].trim();
              Brands.forEach((g) => {
                Brand = Brand.filter(function (str) {
                  return /\S/.test(str);
                });
                if (e.trim().toUpperCase() == g && b__len === Brand.length) {
                  Brand.push(g);
                }
              });
            }
          }
        }

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (b__len === Brand.length) {
          Brand.push("No Brand Found");
        }
        console.log(Brand[i]);

        if (images.length == 1) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: "No Image",
            image3: "No Image",
            image4: "No Image",
            image5: "No Image",
          };
        } else if (images.length == 2) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: "No Image",
            image4: "No Image",
            image5: "No Image",
          };
        } else if (images.length == 3) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: Images[i][2],
            image4: "No Image",
            image5: "No Image",
          };
        } else if (images.length == 4) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: Images[i][2],
            image4: Images[i][3],
            image5: "No Image",
          };
        } else if (images.length == 5 || images.length > 5) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: Images[i][2],
            image4: Images[i][3],
            image5: Images[i][4],
          };
        }

        console.log("\n");

        let workbook = new Excel.Workbook();
        let worksheet = workbook.addWorksheet("Sheet1");
        worksheet.columns = [
          { header: "Product Title", key: "title" },
          { header: "Price", key: "price" },
          { header: "EAN", key: "ean" },
          { header: "Brand", key: "brand" },
          { header: "Category Path", key: "catPath" },
          { header: "Description", key: "desc" },
          { header: "Image_1", key: "image1" },
          { header: "Image_2", key: "image2" },
          { header: "Image_3", key: "image3" },
          { header: "Image_4", key: "image4" },
          { header: "Image_5", key: "image5" },
        ];
        worksheet.columns.forEach((column) => {
          column.width = column.header.length < 12 ? 12 : column.header.length;
        });
        worksheet.getRow(1).font = { bold: true };
        Data.forEach((e) => {
          worksheet.addRow({
            ...e,
          });
        });
        workbook.xlsx.writeFile("FnacProductsData.xlsx");
      } else {
        await page.waitForSelector(".Header__nav", { timeout: 0 });
        const p = await page.$(
          "#Characteristics > div > div > div.characteristicsStrate__lists"
        );
        console.log(i);
        const title = await page.$eval(
          ".f-productHeader-Title",
          (el) => el.textContent
        );
        const Title_Text = await Translate(title);
        Title.push(Title_Text);
        console.log(Title_Text);
        const price = await page.$eval(".checked", (e) => e.textContent);
        Price.push(price.replace(",", "."));
        console.log(price.replace(",", "."));
        const catPath = await page.$eval(".f-breadcrumb", (e) => e.textContent);
        const tCatPath = await Translate(catPath);
        const oCatPath = tCatPath.split("\n");
        const fCatPath = oCatPath.filter(function (str) {
          return /\S/.test(str);
        });
        let a = [];
        fCatPath.forEach((e) => {
          const f = e.trim();
          a.push(f);
        });
        Cat_Path.push(a.join(" > "));
        console.log(a.join(" > "));
        if (p) {
          await page.waitForSelector("section#Characteristics dt");
          const dt__key = await page.evaluate(() => {
            const dts = Array.from(
              document.querySelectorAll("section#Characteristics dt")
            ).map((e) => e.textContent.trim());
            return dts;
          });

          const dd__vle = await page.evaluate(() => {
            const dds = Array.from(
              document.querySelectorAll("section#Characteristics dd")
            ).map((e) => e.textContent.trim());
            return dds;
          });

          const ean__len = EAN.length;
          for (let n = 0; n <= dt__key.length; n++) {
            if (dt__key[n] !== undefined) {
              const e = dt__key[n].trim();
              if (e.indexOf("EAN") != -1) {
                EAN.push(dd__vle[n].trim());
                console.log(dd__vle[n].trim());
              }
            }
          }
          if (ean__len === EAN.length) {
            EAN.push("No EAN Found");
            console.log("No EAN Found");
          }
        } else {
          EAN.push("No EAN Found");
          console.log("No EAN Found");
        }
        const h = await page.$(".summaryStrate");
        if (h) {
          const description = await page.$eval(
            ".summaryStrate__raw",
            (e) => e.textContent
          );
          const tDes = await Translate(description);
          Description.push(tDes);
        } else {
          Description.push("No Description Found");
        }
        const images = await page.evaluate(() => {
          const srcs = Array.from(
            document.querySelectorAll(
              "body > div.Main.Main--fullWidth > div > div.productPageTop > div.f-productPage-colLeft.clearfix > section.f-productVisuals.js-articleVisuals > div.f-productVisuals-thumbnailsWrapper > div > div > div > div > img"
            )
          ).map((image) => image.getAttribute("src"));
          return srcs;
        });
        Images.push(images);

        const b__len = Brand.length;

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (p) {
          await page.waitForSelector("section#Characteristics dt");
          const dt__key = await page.evaluate(() => {
            const dts = Array.from(
              document.querySelectorAll("section#Characteristics dt")
            ).map((e) => e.textContent.trim());
            return dts;
          });

          const dd__vle = await page.evaluate(() => {
            const dds = Array.from(
              document.querySelectorAll("section#Characteristics dd")
            ).map((e) => e.textContent.trim());
            return dds;
          });
          let b = [];
          for (let n = 0; n <= dt__key.length; n++) {
            if (dt__key[n] !== undefined) {
              const e = dt__key[n].trim();
              if (e.indexOf("Fabricante/Marca") != -1) {
                b.push(dd__vle[n].trim());
                const Brand1 = b[0].toUpperCase();
                Brand.push(Brand1);
              }
            }
          }
        }

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (b__len === Brand.length) {
          Brands.forEach((e) => {
            Brand = Brand.filter(function (str) {
              return /\S/.test(str);
            });
            const st = a.join(" > ").toUpperCase().split(" ");
            st.forEach((h) => {
              Brand = Brand.filter(function (str) {
                return /\S/.test(str);
              });
              if (h == e && b__len === Brand.length) {
                Brand.push(e);
              }
            });
          });
        }

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (b__len === Brand.length) {
          const subT = await page.$eval(
            ".f-productHeader-subTitle",
            (e) => e.textContent
          );
          const f_sub = subT.split("\n").filter(function (str) {
            return /\S/.test(str);
          });
          for (let n = 0; n <= f_sub.length; n++) {
            Brand = Brand.filter(function (str) {
              return /\S/.test(str);
            });
            if (f_sub[n] !== undefined) {
              const e = f_sub[n].trim();
              Brands.forEach((g) => {
                Brand = Brand.filter(function (str) {
                  return /\S/.test(str);
                });
                if (e.trim().toUpperCase() == g && b__len === Brand.length) {
                  Brand.push(g);
                }
              });
            }
          }
        }

        Brand = Brand.filter(function (str) {
          return /\S/.test(str);
        });

        if (b__len === Brand.length) {
          Brand.push("No Brand Found");
        }
        console.log(Brand[i]);
        if (images.length == 1) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: "No Image",
            image3: "No Image",
            image4: "No Image",
            image5: "No Image",
          };
        } else if (images.length == 2) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: "No Image",
            image4: "No Image",
            image5: "No Image",
          };
        } else if (images.length == 3) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: Images[i][2],
            image4: "No Image",
            image5: "No Image",
          };
        } else if (images.length == 4) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: Images[i][2],
            image4: Images[i][3],
            image5: "No Image",
          };
        } else if (images.length == 5 || images.length > 5) {
          Data[i] = {
            title: Title[i],
            price: Price[i],
            ean: EAN[i],
            brand: Brand[i],
            catPath: Cat_Path[i],
            desc: Description[i],
            image1: Images[i][0],
            image2: Images[i][1],
            image3: Images[i][2],
            image4: Images[i][3],
            image5: Images[i][4],
          };
        }

        console.log("\n");

        let workbook = new Excel.Workbook();
        let worksheet = workbook.addWorksheet("Sheet1");
        worksheet.columns = [
          { header: "Product Title", key: "title" },
          { header: "Price", key: "price" },
          { header: "EAN", key: "ean" },
          { header: "Brand", key: "brand" },
          { header: "Category Path", key: "catPath" },
          { header: "Description", key: "desc" },
          { header: "Image_1", key: "image1" },
          { header: "Image_2", key: "image2" },
          { header: "Image_3", key: "image3" },
          { header: "Image_4", key: "image4" },
          { header: "Image_5", key: "image5" },
        ];
        worksheet.columns.forEach((column) => {
          column.width = column.header.length < 12 ? 12 : column.header.length;
        });
        worksheet.getRow(1).font = { bold: true };
        Data.forEach((e) => {
          worksheet.addRow({
            ...e,
          });
        });
        workbook.xlsx.writeFile("FnacProductsData.xlsx");
      }
    }
  } catch (error) {
    console.log(error);
  }
};

ScrapeLinks();
