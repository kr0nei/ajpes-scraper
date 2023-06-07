import puppeteer from 'puppeteer-extra';
import StealthPlugin from 'puppeteer-extra-plugin-stealth';
import fs from 'fs';
import XLSX from 'xlsx';

puppeteer.use(StealthPlugin());

const dataToExport = [];
const date = new Date();

const saveCookie = async (page) => {
  try {
    const cookies = await page.cookies();
    const cookieJson = JSON.stringify(cookies, null, 2);
    fs.writeFile('cookies.json', cookieJson);
  } catch (error) { }
};
const loadCookie = async (page) => {
  try {
    const cookieJson = await fs.readFile('cookies.json');
    const cookies = JSON.parse(cookieJson);
    await page.setCookie(...cookies);
  } catch (error) { }

};
const delay = async (time) => {
  return new Promise(function (resolve) {
    setTimeout(resolve, time)
  });
};
const getCompanyData = async (page, url) => {
  const data = await page.evaluate((url) => {
    const selectors = {
      podjetje: ".col-md-9 > small",
      zastopniki: ".col-sm-7.div-border-bottom.no-last",
      trr: ".table.table-bordered.table-white > tbody > tr",
    }
    function getPodjetjeDat(document) {
      const podjetje = document.querySelector(selectors.podjetje)
      const podjetjeDat = podjetje.innerHTML.split("<br>")
      podjetjeDat.forEach((element, index) => {
        podjetjeDat[index] = element.replace(/\n/g, "").replace("&nbsp;", " ").trim();
        podjetjeDat[index] = podjetjeDat[index].replace(/<b>|<\/b>/g, "");
      });
      const datObject = {
        "Polno ime": podjetjeDat[0],
        "Naslov": podjetjeDat[1],
        "Pošta": podjetjeDat[2],
        "Matična številka": podjetjeDat[3].split(":")[1].trim(),
        "Davčna številka": podjetjeDat[4].split(":")[1].trim()
      };
      return datObject;
    }
    function getZastopnikiDat(document) {
      const zastopniki = document.querySelectorAll(selectors.zastopniki)[2].getElementsByTagName("div");
      const zastopnikiDat = [];
      for (let i = 0; i < zastopniki.length; i++) {
        zastopnikiDat.push(zastopniki[i].getElementsByTagName('a')[0].innerHTML.replace("&nbsp;", " ").trim());
      }
      // console.log(zastopnikiDat);
      return { "Zastopniki": zastopnikiDat };
    }
    function getTrrDat(document) {
      const trr = document.querySelectorAll(selectors.trr);
      let trrDat = [];
      for (let i = 0; i < trr.length; i++) {
        trrDat.push(trr[i].getElementsByTagName('td'));
      }
      const trrDatEdited = []
      for (let i = 0; i < trrDat.length; i++) {
        for (let j = 0; j < trrDat[i].length; j++) {
          trrDatEdited.push(trrDat[i][j].innerText);
        }
        trrDatEdited.push('-');
      }
      const znak = trrDatEdited.indexOf("-") + 1;
      // return [trrDatEdited.length,znak];
      const trrDatObject = [];
      for (let i = 0; i < trrDatEdited.length; i = i + znak) {
        trrDatObject.push({
          "Račun": trrDatEdited[i],
          "Banka": trrDatEdited[i + 1],
          "Vrsta računa": trrDatEdited[i + 2],
          "Datum odprtja": trrDatEdited[i + 3],
          "Datum zaprtja": trrDatEdited[i + 4],
          "Nep.obv.": trrDatEdited[i + 5],
        });
      }
      return trrDatObject;
    }
    const dt = [];
    dt.push(getPodjetjeDat(document));
    dt.push(getZastopnikiDat(document));
    dt.push(getTrrDat(document));
    dt.push({ "url": url })
    return dt;
  }, url);
  return data;
};
function exportData(data) {
  const ime_datoteke = "podatki.xlsx";// Ime datoteke kjer se bodo shranili podatki
  const workbook = XLSX.utils.book_new();
  for (let i = 0; i < data.length; i++) {
    console.log(data[i][0]['Polno ime']);
    const editedData = [];
    editedData.push(["Osnovni podatki"]);
    for (const [key, value] of Object.entries(data[i][0])) {
      editedData.push(["", key, value]);
    }
    editedData.push([""]);
    editedData.push(["Zastopniki"]);
    data[i][1].Zastopniki.forEach(el => editedData.push(["", el]));

    editedData.push([""]);
    editedData.push(["TRR"]);
    editedData.push(["", "Račun", "Banka", "Vrsta Računa", "Datum odprtja", "Datum zaprtja", "Nep.obv"]);
    data[i][2].forEach(el => {
      editedData.push(["", el["Račun"], el["Banka"], el["Vrsta računa"], el["Datum odprtja"], el["Datum zaprtja"], el["Nep.obv."]]);
    });
    editedData.push([""]);
    editedData.push(["Url do AJPESOVE strani", data[i][3].url]);
    const worksheet = XLSX.utils.aoa_to_sheet(editedData);
    XLSX.utils.book_append_sheet(workbook, worksheet, data[i][0]['Polno ime'].split(" ")[0].trim());
    console.log(date.toLocaleTimeString(), " - Podatki za ", data[i][0]['Polno ime'].split(" ")[0].trim(), " izvoženi.");
  }
  XLSX.writeFile(workbook, ime_datoteke);
  console.log(date.toLocaleTimeString(), " - Končano, vsi podatki izvoženi v podatki.xlsx");
}
const main = async () => {
  const browser = await puppeteer.launch({
    headless: 'new',
    executablePath: puppeteer.executablePath()
  });
  const companies_urls = ["https://www.ajpes.si/podjetje/CALMO_d.o.o.?enota=231289&EnotaStatus=1#", "https://www.ajpes.si/podjetje/MOBI_-_COMP_SISTEMI_d.o.o.?enota=539737&EnotaStatus=1"];
  try {
    const page = await browser.newPage();
    await loadCookie(page);
    for (let i = 0; i < companies_urls.length; i++) {
      await page.goto(companies_urls[i], { waitUntil: 'load' });
      if (await page.cookies() == []) {
        await page.click("[class='header-item login']", { waitUntil: 'load' });
        await page.click("[class='btn btn-default btn-lg btn-block']", { waitUntil: 'load' });
        await delay(Math.floor(Math.random() * 2500 + 1500));//Počakaj med 1,5 do 2,5 sekund
        await page.click("[class='btn btn-success']", { waitUntil: 'load' });
        await delay(Math.floor(Math.random() * 2500 + 1500));
        await saveCookie(page);
      }
      const companyData = await getCompanyData(page, companies_urls[i]);
      dataToExport.push(companyData);
    }
    exportData(dataToExport);
  } catch (error) {
    console.log(error);
  } finally {
    await browser.close();
  }
}
main();