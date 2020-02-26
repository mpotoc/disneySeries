const puppeteer = require('puppeteer');
const config = require('./config');
const Excel = require('exceljs');
const fs = require('fs');
const program = require('commander');

program
  .option('--start <number>', 'Start range number for series array (optional).')
  .option('--stop <number>', 'Stop range number for series array (optional).')
  .option('--startSeason <number>', 'Start season range number for current series (optional).')
  .option('--stopSeason <number>', 'Stop season range number for current series (optional).')
  .option('--recreate', 'Delete current excel file and create empty one (optional).')
  .option('--worksheet <name>', 'Add a new worksheet to excel file (optional).')
  .parse(process.argv);

const start = program.start;
const stop = program.stop;
const startSeason = program.startSeason;
const stopSeason = program.stopSeason;
const recreate = program.recreate;
const worksheet_name = program.worksheet;

(async () => {

  await excelHandle(recreate);

  const browser = await puppeteer.launch({
    /*args: [
      '--proxy-server=socks5://198.211.99.227:46437'
    ],*/
    //headless: false,
    //slowMo: 200,
    //devtools: true,
    defaultViewport: {
      width: 1920,
      height: 1080
    }
  });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.0 Safari/537.36');
  page.setDefaultTimeout(0);

  await page.goto(config.url, {
    waitUntil: 'networkidle0'
  });
  console.log('Login page initiated.');
  await page.waitForSelector(config.USERNAME_SELECTOR);
  console.log('Login page username selector exists!');
  //await page.waitFor(config.USERNAME_SELECTOR);

  // Additional step if we start at https://www.disneyplus.com
  //await page.screenshot({path: 'screen.png'});
  //await page.click('button[kind="outline"]');

  //await page.screenshot({ path: 'screenEmail.png' });

  await page.click(config.USERNAME_SELECTOR);
  await page.keyboard.type(config.username);
  await page.click('button[name="dssLoginSubmit"]');
  console.log('Login page username form submitted.');
  
  await page.waitForNavigation({
    waitUntil: 'networkidle0'
  });

  //await page.screenshot({ path: 'screenLogIn.png' });

  await page.click(config.PASSWORD_SELECTOR);
  await page.keyboard.type(config.password);
  await page.click('button[name="dssLoginSubmit"]');
  console.log('Login page password form submitted.');

  await page.waitForNavigation({
    waitUntil: 'networkidle0'
  });
  console.log('Main loggedin page initiated.');

  //await page.screenshot({ path: 'screenLoggedIn.png' });

  await page.goto(config.seriesUrl, {
    waitUntil: 'networkidle0'
  });
  console.log('Series page initiated.');
  await page.waitFor('.sc-jTzLTM');
  //await page.screenshot({ path: 'screenSeries.png' });

  console.log('Series page scroll to end, to get all series visible in DOM.');
  await autoScroll(page);
  await page.waitFor(1000);

  //await page.screenshot({
  //  path: 'scrollCheckSeries.png',
  //  fullPage: true
  //});

  let linksSeries = await page.$$('a.sc-ckVGcZ');
  let dataSeries = [];

  let fromNumber = !start ? 0 : (parseInt(start) > (linksSeries.length - 1) ? 0 : parseInt(start));
  let toNumber = !stop ? linksSeries.length : (parseInt(stop) > linksSeries.length ? linksSeries.length : parseInt(stop));
  let startNo = fromNumber > toNumber ? 0 : fromNumber;

  console.log('Series scraping started.');
  for (let i = startNo; i < toNumber; i++) {
    await page.evaluate((i) => {
      return ([...document.querySelectorAll('a.sc-ckVGcZ')][i]).click();
    }, i);

    await page.waitFor('a.sc-ckVGcZ');

    let urlSerie = await page.url();

    let resultSerie = await page.evaluate(async (urlSerie) => {
      let res = document.querySelector('div > p[class~="sc-gzVnrw"]').innerText;
      let resArr = res.split('\u2022');

      let genre = resArr[2].trim();

      let release_year = resArr[0].trim();

      let rating_step1 = document.querySelector('p > img');
      let rating_step2 = rating_step1.alt.split('_');
      let rating = rating_step2[2].toUpperCase();

      let title_step = document.querySelector('div[class~="sc-iujRgT"] > img');
      let title = title_step.alt;

      let source_id_step = urlSerie.split('/');
      let source_id = source_id_step[source_id_step.length - 1];

      // setting date for capture date field
      const dateCapture = new Date();
      var dateOptions = {
        year: "numeric",
        month: "2-digit",
        day: "numeric"
      };

      return {
        'bot_system': 'disneyplus',
        'bot_version': '1.0.0',
        'bot_country': 'us',
        'capture_date': dateCapture.toLocaleString('en', dateOptions),
        'offer_type': 'SVOD',
        'purchase_type': '',
        'picture_quality': '',
        'program_price': '',
        'bundle_price': '',
        'currency': '',
        'addon_name': '',
        'is_movie': 0,
        'season_number': '',
        'episode_number': '',
        'title': '',
        'genre': '',
        'source_id': '',
        'program_url': '',
        'maturity_rating': rating,
        'release_date': '',
        'release_year': '',
        'viewable_runtime': '',
        'series_title': title,
        'series_release_year': release_year,
        'series_source_id': source_id,
        'series_url': urlSerie,
        'series_genre': genre,
        'season_source_id': ''
      };
    }, urlSerie);

    let seasonsArr = await page.$$('div.sc-gtfDJT');

    let fromSeasonNumber = !startSeason ? 0 : (parseInt(startSeason) > (seasonsArr.length - 1) ? 0 : parseInt(startSeason));
    let toSeasonNumber = !stopSeason ? seasonsArr.length : (parseInt(stopSeason) > seasonsArr.length ? seasonsArr.length : parseInt(stopSeason));
    let startSeasonNo = fromSeasonNumber > toSeasonNumber ? 0 : fromSeasonNumber;
    let endSeasonNo = toSeasonNumber === fromSeasonNumber ? toSeasonNumber + 1 : toSeasonNumber;

    let seasonsData = [];

    for (const obj of seasonsArr) {
      let seasonObj = await obj.$$('a.sc-ckVGcZ');
      seasonsData.push(seasonObj.length);
    }
  
    for (let j = startSeasonNo; j < endSeasonNo; j++) {
      let seasonNo = j + 1;

      let seasonArr = await seasonsArr[j].$$('a.sc-ckVGcZ');
      let seasonLength = seasonArr.length;
      let startIndex = 0;

      // if we start not from first season
      for (let x = 0; x < j; x++) {
        startIndex += seasonsData[x];
      }

      console.log('Seasons page scroll to end, to get all episodes visible in DOM.');
      await autoScroll(page);
      await page.waitFor(1000);    

      if (seasonLength > 12) {
        let repCount = Math.round((seasonLength - 5) / 5) + 1;

        for (let r = 0; r < repCount; r++) {
          await page.evaluate((j) => {
            return ([...document.querySelectorAll('button.slick-next')][j]).click();
          }, j);
          
          await page.waitFor(1000);
        }
      }

      let seasonsArr2 = await page.$$('div.sc-gtfDJT');
      let season = await seasonsArr2[j].$$('a.sc-ckVGcZ');

      for (let k = 0; k < season.length; k++) {
        let episodeData = await season[k].$eval('p.sc-gzVnrw', el => el.innerText);
        let epArr = episodeData.split('.');
        let epArr_step_1 = epArr[epArr.length - 1].split('(');
        let epArr_step_2 = epArr_step_1[epArr_step_1.length - 1].replace(/[\D]/g, '');

        let title = epArr_step_1[0].trim();
        let viewable_runtime = parseInt(epArr_step_2) * 60;

        await page.evaluate((startIndex) => {
          return ([...document.querySelectorAll('a.sc-ckVGcZ')][startIndex]).click();
        }, startIndex);

        let urlEpisode = await page.url();
        let urlArr = urlEpisode.split('/');
        let episode_source_id = urlArr[urlArr.length - 1];

        let episodeObj = {
          'bot_system': 'disneyplus',
          'bot_version': '1.0.0',
          'bot_country': 'us',
          'capture_date': resultSerie.capture_date,
          'offer_type': 'SVOD',
          'purchase_type': '',
          'picture_quality': '',
          'program_price': '',
          'bundle_price': '',
          'currency': '',
          'addon_name': '',
          'is_movie': 0,
          'season_number': seasonNo,
          'episode_number': k + 1,
          'title': title,
          'genre': '',
          'source_id': episode_source_id,
          'program_url': urlEpisode,
          'maturity_rating': resultSerie.maturity_rating,
          'release_date': '',
          'release_year': '',
          'viewable_runtime': viewable_runtime,
          'series_title': resultSerie.series_title,
          'series_release_year': resultSerie.series_release_year,
          'series_source_id': resultSerie.series_source_id,
          'series_url': resultSerie.series_url,
          'series_genre': resultSerie.series_genre,
          'season_source_id': ''
        };

        startIndex++;

        dataSeries.push(episodeObj);

        await addDataToExcel(episodeObj, worksheet_name);

        console.log(episodeObj);

        await page.goBack();
      }
    }

    await page.goBack();
  }

  console.log('Series scraping finished.');

  await browser.close();
})();

async function autoScroll(page) {
  await page.evaluate(async () => {
    await new Promise((resolve, reject) => {
      var totalHeight = 0;
      var distance = 100;
      var timer = setInterval(() => {
        var scrollHeight = document.body.scrollHeight;
        window.scrollBy(0, distance);
        totalHeight += distance;

        if (totalHeight >= scrollHeight) {
          clearInterval(timer);
          resolve();
        }
      }, 100);
    });
  });
};

async function excelHandle(recreate) {
  const filePath = './disneySeries.xlsx';

  try {
    if (fs.existsSync(filePath)) {
      if (recreate) {
        fs.unlinkSync(filePath);
        await createExcel();
      } else {
        console.log('File already exists, will use existing file!');
      }
    } else {
      await createExcel();
    }
  } catch (e) {
    console.error(e);
  }
};

async function createExcel() {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('FEB 2020');

  worksheet.columns = [
    { header: 'bot_system', key: 'bot_system', width: 10 },
    { header: 'bot_version', key: 'bot_version', width: 10 },
    { header: 'bot_country', key: 'bot_country', width: 10 },
    { header: 'capture_date', key: 'capture_date', width: 10 },
    { header: 'offer_type', key: 'offer_type', width: 10 },
    { header: 'purchase_type', key: 'purchase_type', width: 10 },
    { header: 'picture_quality', key: 'picture_quality', width: 10 },
    { header: 'program_price', key: 'program_price', width: 10 },
    { header: 'bundle_price', key: 'bundle_price', width: 10 },
    { header: 'currency', key: 'currency', width: 10 },
    { header: 'addon_name', key: 'addon_name', width: 10 },
    { header: 'is_movie', key: 'is_movie', width: 10 },
    { header: 'season_number', key: 'season_number', width: 10 },
    { header: 'episode_number', key: 'episode_number', width: 10 },
    { header: 'title', key: 'title', width: 10 },
    { header: 'genre', key: 'genre', width: 10 },
    { header: 'source_id', key: 'source_id', width: 10 },
    { header: 'program_url', key: 'program_url', width: 10 },
    { header: 'maturity_rating', key: 'maturity_rating', width: 10 },
    { header: 'release_date', key: 'release_date', width: 10 },
    { header: 'release_year', key: 'release_year', width: 10 },
    { header: 'viewable_runtime', key: 'viewable_runtime', width: 10 },
    { header: 'series_title', key: 'series_title', width: 10 },
    { header: 'series_release_year', key: 'series_release_year', width: 10 },
    { header: 'series_source_id', key: 'series_source_id', width: 10 },
    { header: 'series_url', key: 'series_url', width: 10 },
    { header: 'series_genre', key: 'series_genre', width: 10 },
    { header: 'season_source_id', key: 'season_source_id', width: 10 }
  ];

  // save under disneySeries.xlsx
  await workbook.xlsx.writeFile('disneySeries.xlsx');

  console.log('File is created.');
};

async function addDataToExcel(data, worksheet_name) {
  //load a copy of disneySeries.xlsx
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile('disneySeries.xlsx');
  let worksheet = null;

  if (!worksheet_name) {
    worksheet = workbook.getWorksheet('FEB 2020');
  } else {
    if (!workbook.getWorksheet(worksheet_name)) {
      worksheet = workbook.addWorksheet(worksheet_name);
    } else {
      worksheet = workbook.getWorksheet(worksheet_name);
    }
  }

  worksheet.columns = [
    { header: 'bot_system', key: 'bot_system', width: 10 },
    { header: 'bot_version', key: 'bot_version', width: 10 },
    { header: 'bot_country', key: 'bot_country', width: 10 },
    { header: 'capture_date', key: 'capture_date', width: 10 },
    { header: 'offer_type', key: 'offer_type', width: 10 },
    { header: 'purchase_type', key: 'purchase_type', width: 10 },
    { header: 'picture_quality', key: 'picture_quality', width: 10 },
    { header: 'program_price', key: 'program_price', width: 10 },
    { header: 'bundle_price', key: 'bundle_price', width: 10 },
    { header: 'currency', key: 'currency', width: 10 },
    { header: 'addon_name', key: 'addon_name', width: 10 },
    { header: 'is_movie', key: 'is_movie', width: 10 },
    { header: 'season_number', key: 'season_number', width: 10 },
    { header: 'episode_number', key: 'episode_number', width: 10 },
    { header: 'title', key: 'title', width: 10 },
    { header: 'genre', key: 'genre', width: 10 },
    { header: 'source_id', key: 'source_id', width: 10 },
    { header: 'program_url', key: 'program_url', width: 10 },
    { header: 'maturity_rating', key: 'maturity_rating', width: 10 },
    { header: 'release_date', key: 'release_date', width: 10 },
    { header: 'release_year', key: 'release_year', width: 10 },
    { header: 'viewable_runtime', key: 'viewable_runtime', width: 10 },
    { header: 'series_title', key: 'series_title', width: 10 },
    { header: 'series_release_year', key: 'series_release_year', width: 10 },
    { header: 'series_source_id', key: 'series_source_id', width: 10 },
    { header: 'series_url', key: 'series_url', width: 10 },
    { header: 'series_genre', key: 'series_genre', width: 10 },
    { header: 'season_source_id', key: 'season_source_id', width: 10 }
  ];

  await worksheet.addRow({
    bot_system: data.bot_system,
    bot_version: data.bot_version,
    bot_country: data.bot_country,
    capture_date: data.capture_date,
    offer_type: data.offer_type,
    purchase_type: data.purchase_type,
    picture_quality: data.picture_quality,
    program_price: data.program_price,
    bundle_price: data.bundle_price,
    currency: data.currency,
    addon_name: data.addon_name,
    is_movie: data.is_movie,
    season_number: data.season_number,
    episode_number: data.episode_number,
    title: data.title,
    genre: data.genre,
    source_id: data.source_id,
    program_url: data.program_url,
    maturity_rating: data.maturity_rating,
    release_date: data.release_date,
    release_year: data.release_year,
    viewable_runtime: data.viewable_runtime,
    series_title: data.series_title,
    series_release_year: data.series_release_year,
    series_source_id: data.series_source_id,
    series_url: data.series_url,
    series_genre: data.series_genre,
    season_source_id: data.season_source_id
  });

  await workbook.xlsx.writeFile('disneySeries.xlsx');

  console.log("Data is written to file.");
};