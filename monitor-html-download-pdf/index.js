const chokidar = require('chokidar');
const puppeteer = require('puppeteer-core');
const chromePaths = require('chrome-paths');

//监控这个文件夹内新出现的html文件（也就是从邮件中保存的附件），需要设置的和outlook中保存附件的文件夹一致
const monitoredFolder = "/Users/cpf/Downloads/";

//chrome 的 user profile文件夹  具体文件夹路径可以通过打开 chrome://version来查看
//需要先把chrome 设置为pdf默认下载而不是默认在页面内打开，否则无法下载，
const chromeProfilePath = "/Users/cpf/Library/Application Support/Google/Chrome/Default";

// 用来下载pdf的密码，需替换
const pwd = 'xxx';

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

        page = await browser.newPage();

        //load this page first, then the second time it maybe faster，since js files has been loaded once
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

        await page.waitForTimeout(20000);
        log("waited 20 secs");
        await page.screenshot({ path: '7.png' });

        log("pdf downloaded for: " + path);
        log("==================================================");

        await page.close();
    }
    catch (error) {

        log('error: ', error);

        log(path + " 未成功为此附件下载pdf 稍后将会重试");
        page.close();

        // 10秒后重新尝试为此html文件下载pdf
        setTimeout(
            function () {
                log(path + " 重试为此文件下载pdf");
                downloadPdfForHtmlFile(path);
            }, retryWaitTime);

    }
}
