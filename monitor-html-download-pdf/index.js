const chokidar = require('chokidar');
const puppeteer = require('puppeteer-core');
const chromePaths = require('chrome-paths');
const PATH = require('path');
const fs = require('fs')

//监控这个文件夹内新出现的html文件（也就是从邮件中保存的附件），需要设置的和outlook中保存附件的文件夹一致
const monitoredFolder = "C:/Data/Archive/";

//监控下载到的pdf的文件夹。配置了这一项才能修改pdf文件名，增加时间戳
const pdfFolder = "C:/Data/hsbc-pdf-files/";

//chrome 的 user profile文件夹  具体文件夹路径可以通过打开 chrome://version来查看
//需要先把chrome 设置为pdf默认下载而不是默认在页面内打开，否则无法下载，
const chromeProfilePath = "C:\\Users\\cpf\\AppData\\Local\\Google\\Chrome\\User Data";

// 用来下载pdf的密码，需替换
const pwd = 'Kelvin9731';

// 页面跳转时 最多 等待的时长，单位毫秒。先设定为两分钟，如果还会超时，可适当增加。
const waitTimeInMs = 120000;

// 如果下载失败(例如，由于超时)，等待多久再次重试，单位毫秒
const retryWaitTime = 10000;

// Initialize watcher.
const watcher = chokidar.watch(monitoredFolder, {
    persistent: true,
    ignoreInitial: true
});

const log = console.log.bind(console);

let browserWSEndpoint;

// Add event listeners.
watcher
    .on('add', async path => {
        //只监听新增的html文件（也就是从邮件中保存的附件）
        if (path.includes(".html")) {
            log(`File ${path} has been added`);

            await downloadPdfForHtmlFile(path);
        }
    })

async function downloadPdfForHtmlFile(path) {
    let page;

    try {
        let browser = await initBrowser();

        page = await browser.newPage();

        //load this page first, then the second time it maybe faster，since js files has been loaded once
        await operatePageSteps(page, path);

        var { downloaded, pdfWatcher } = watchPdfFiles();

        await page.waitForTimeout(20000);
        await page.screenshot({ path: '7.png' });

        if (!downloaded) {
            throw "pdf NOT downloaded, closing this page for now";
        } else {
            log("waited 20 secs, will close this page");
        }

        await pdfWatcher.close();
        await page.close();
    }
    catch (error) {
        retry(error, path, page);
    }
}

function watchPdfFiles() {
    const pdfWatcher = chokidar.watch(pdfFolder, {
        persistent: true,
        ignoreInitial: true
    });

    let downloaded = false;
    pdfWatcher
        .on('add', async (pdfPath) => {
            //只监听新增的pdf
            if (pdfPath.includes(".pdf") && !pdfPath.includes("DOWNLOADED-AT-")) {
                log(`new pdf file detected: ${pdfPath}`);
                downloaded = true;

                fs.stat(pdfPath, function (err, stat) {
                    if (!err) {
                        setTimeout(checkFileCopyComplete, 1000, pdfPath, stat);
                    }
                });
            }
        });
    return { downloaded, pdfWatcher };
}

async function initBrowser() {
    let browser;
    if (browserWSEndpoint) {
        browser = await puppeteer.connect({
            browserWSEndpoint,
        });
    } else {
        browser = await puppeteer.launch({
            executablePath: chromePaths.chrome,
            userDataDir: chromeProfilePath,
            headless: false
        });
        browserWSEndpoint = browser.wsEndpoint();
    }
    return browser;
}

async function operatePageSteps(page, path) {
    await page.goto("https://www.asiapacific.ecorrespondence.hsbc.com/hsbc/main.jsp", { timeout: waitTimeInMs });
    await page.waitForSelector(".buttonInner", { timeout: waitTimeInMs });
    log("pre loaded page resources");
    await page.screenshot({ path: '1.png' });

    await page.goto("file:///" + path, { timeout: waitTimeInMs });
    log("opened local html file");
    await page.screenshot({ path: '2.png' });

    await page.click('#openForm');
    log("clicked form");

    await page.waitForNavigation({
        waitUntil: 'domcontentloaded', timeout: waitTimeInMs
    });
    log("navigated to new page");
    await page.screenshot({ path: '3.png' });

    await page.waitForSelector("#dijit_form_ValidationTextBox_0", { timeout: waitTimeInMs });
    log("password input is displayed now");
    await page.screenshot({ path: '4.png' });

    await page.type('#dijit_form_ValidationTextBox_0', pwd);
    log("typed password");
    await page.screenshot({ path: '5.png' });

    await page.click('.submit_input');
    log("clicked continue");
    await page.screenshot({ path: '6.png' });    
}

function retry(error, path, page) {
    log('error: ', error);

    log(path + " pdf NOT downloaded for this html file, will retry");
    log("\n");

    if (page) {
        page.close();
    }

    // 10秒后重新尝试为此html文件下载pdf
    setTimeout(
        function () {
            log(path + " retry download for this file");
            downloadPdfForHtmlFile(path);
        }, retryWaitTime);
}

function checkFileCopyComplete(path, prev) {
    fs.stat(path, function (err, stat) {

        if (err) {
            throw err;
        }
        if (stat.mtime.getTime() === prev.mtime.getTime()) {
            log('File download complete, beginning rename');

            const baseName = PATH.basename(path);
            const newName = path.replace(baseName, "DOWNLOADED-AT-" + Date.now() + "-" + baseName);

            fs.renameSync(path, newName, function (err) {
                if (err) throw err;
            });

            log("pdf file name changed to " + newName + "\n");

        }
        else {
            setTimeout(checkFileCopyComplete, 1000, path, stat);
        }
    });
}