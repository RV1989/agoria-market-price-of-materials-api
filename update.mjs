import { got } from "got";
import { JSDOM } from "jsdom";
import { createWriteStream } from "fs";
import _ from "lodash";
import pkg from "xlsx";
const { readFile, utils } = pkg;
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import { unlink, writeFile } from "fs/promises";
const __dirname = dirname(fileURLToPath(import.meta.url));

const getToolUrl = async (url) => {
  const response = await got.get(url);
  const dom = new JSDOM(response.body);
  const result = [...dom.window.document.querySelectorAll("a.btn")]
    .map((link) => {
      return { text: link.textContent, url: link.href };
    })
    .find((x) => x.text.includes("Consult"));
  return Promise.resolve(result.url);
};

const getExcelLinks = async (url) => {
  const response = await got.get(url);
  const dom = new JSDOM(response.body);
  const tables = [...dom.window.document.querySelectorAll("table")];
  const data = [];
  for (let table of tables) {
    const category = _.toLower(
      _.camelCase(table.querySelector("h3").textContent)
    );
    const dataRows = table.querySelectorAll("tr.trgrey");
    const productData = [];
    for (let row of dataRows) {
      const cells = row.querySelectorAll("td");
      const product = _.toLower(_.camelCase(cells[0].textContent));
      const xlsx = `https://tools.agoria.be${cells[1].querySelector("a").href}`;
      productData.push({ product, xlsx });
    }
    data.push({ category, productData });
  }
  return Promise.resolve(data);
};

const getJson = async ({ product, xlsx }) => {
  await got.stream(xlsx).pipe(createWriteStream(`./temp/${product}.xlsx`));
  await new Promise((r) => setTimeout(r, 1000));
  const workbookPath = join(__dirname, `./temp/${product}.xlsx`);
  const workbook = readFile(workbookPath);
  const sheets = workbook.SheetNames;
  const result = [];
  for (let type of sheets) {
    var range = utils.decode_range(workbook.Sheets[type]["!ref"]);
    let firstRow = 8;
    for (let i = 2; i < 20; i++) {
      if (_.toLower(workbook.Sheets[type][`A${i}`]?.v) == "january") {
        firstRow = i - 2;
        break;
      }
    }
    const rangeToRead = utils.encode_range({
      s: { r: firstRow, c: 1 },
      e: { r: 22, c: range.e.c },
    });
    const opt = { range: rangeToRead };
    const json = utils.sheet_to_json(workbook.Sheets[type], opt);
    const data = {
      type: _.toLower(_.camelCase(type)),
      avg: json[12],
      history: [
        { jan: json[0] },
        { feb: json[1] },
        { mar: json[2] },
        { apr: json[3] },
        { may: json[4] },
        { jun: json[5] },
        { jul: json[6] },
        { aug: json[7] },
        { sep: json[8] },
        { oct: json[9] },
        { nov: json[10] },
        { dec: json[11] },
      ],
    };
    result.push(data);
  }
  await unlink(workbookPath);
  return Promise.resolve(result);
};

(async () => {
  const url = await getToolUrl(
    "https://www.agoria.be/en/services/data-research/market-prices-of-materials"
  );
  const excelLinks = await getExcelLinks(url);
  const apiData = [];
  for (let excelLink of excelLinks) {
    for (let data of excelLink.productData) {
      apiData.push({
        category: excelLink.category,
        product: data.product,
        types: await getJson(data),
      });
    }
  }

  await writeFile("./data.json", JSON.stringify(apiData, null, 2));
})();
