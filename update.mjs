import { got } from "got";
import { JSDOM } from "jsdom";
import { createWriteStream } from "fs";

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
    const category = table.querySelector("h3").textContent;
    const dataRows = table.querySelectorAll("tr.trgrey");
    const productData = [];
    for (let row of dataRows) {
      const cells = row.querySelectorAll("td");
      const product = cells[0].textContent;
      const xlsx = `https://tools.agoria.be${cells[1].querySelector("a").href}`;
      productData.push({ product, xlsx });
    }
    data.push({ category, productData });
  }
  return Promise.resolve(data);
};

(async () => {
  const url = await getToolUrl(
    "https://www.agoria.be/en/services/data-research/market-prices-of-materials"
  );
  const excelLinks = await getExcelLinks(url);
  const excelUrl = excelLinks[0].productData[0].xlsx;

  await got
    .stream(excelUrl)
    .pipe(
      createWriteStream(`./temp/${excelLinks[0].productData[0].product}.xlsx`)
    );

  console.log("fin");
})();
