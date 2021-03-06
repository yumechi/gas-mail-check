type GeneralDate = Date | GoogleAppsScript.Base.Date;


function main() {
    /**
     * start point method
     */
    run();
}

function createQuery(): string {
    /**
     * Create filter query for gmail.
     * BASE_QUERY can be replaced your secret search condition.
     * And, this method search range default 1 days ago.
     */
    const query_base = PropertiesService.getScriptProperties().getProperty("BASE_QUERY");
    const yesterday = (() => {
        const today = new Date();
        const _yesterday = new Date(today.setDate(today.getDate() - 1));
        return japaneseDateFormat(_yesterday);
    })();

    return [
        query_base,
        `after:${yesterday}`,
    ].join(" ");
}

function japaneseDateFormat(date: GeneralDate, isFull = false): string {
    /**
     * Change Date to Japanese normaly string format.
     * For example, "2020-06-30"
     */
    const formatString = isFull ? "yyyy-MM-dd HH:mm:ss" : "yyyy-MM-dd";
    return Utilities.formatDate(date, "JST", formatString);
}

function run(): void {
    /**
     * main domain function
     */

    const query: string = createQuery();
    const result: GoogleAppsScript.Gmail.GmailThread[] = GmailApp.search(query);

    // message in thread ...
    for (const elem of result) {
        const messages: GoogleAppsScript.Gmail.GmailMessage[] = elem.getMessages();
        for (const message of messages) {
            const messageId: string = message.getId();
            const date: string = japaneseDateFormat(message.getDate());
            const name: string = cleansingUserName(message.getFrom());
            const subject: string | null = cleansingSubject(message.getSubject());

            if (!subject) {
                continue;
            }

            if (existsMessageId(messageId, date)) {
                continue;
            }

            const postDate: string = japaneseDateFormat(message.getDate(), true);
            postWebHook([
                postDate,
                subject,
                name,
            ]);

            writeMessageId(messageId, date);
        }
    }
}

function getSheet(spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet, name: string): GoogleAppsScript.Spreadsheet.Sheet {
    /**
     * If the target sheet does not exist, create a new sheet and return it
     * refer: https://qiita.com/crawd4274/items/13120429cb3328e8ace2
     */
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadSheet.getSheetByName(name);
    if (!sheet) {
        const _sheet = spreadSheet.insertSheet();
        _sheet.setName(name);
        return _sheet;
    }
    return sheet;
}

function existsMessageId(messageId: string, date: string): boolean {
    /**
     * Search the message ID in the A1 cell of the sheet named with the date
     */
    const spreadSheetId = PropertiesService.getScriptProperties().getProperty("SPREAD_SHEET_ID");
    const spreadSheet = SpreadsheetApp.openById(spreadSheetId);

    const sheet: GoogleAppsScript.Spreadsheet.Sheet = getSheet(spreadSheet, date);
    const cells = sheet.getRange("A1:A1");
    const cell = cells.getCell(1, 1);
    const data = cell.getDisplayValue();

    if (data.includes(messageId)) {
        return true;
    }
    return false;
}

function writeMessageId(messageId: string, date: string): void {
    /**
     * Write the message ID in the A1 cell of the sheet named with the date
     */
    const spreadSheetId = PropertiesService.getScriptProperties().getProperty("SPREAD_SHEET_ID");
    const spreadSheet = SpreadsheetApp.openById(spreadSheetId);

    const sheet: GoogleAppsScript.Spreadsheet.Sheet = getSheet(spreadSheet, date);
    const cells = sheet.getRange("A1:A1");
    const cell = cells.getCell(1, 1);
    const data = cell.getDisplayValue();

    if (data.length < 1) {
        cell.setValue(`${messageId}`);
    } else {
        cell.setValue(`${data},${messageId}`);
    }
}


function postWebHook(data) {
    /**
     * send to webhook(Assuming Discord)
     */
    const urls = PropertiesService.getScriptProperties().getProperty("WEB_HOOK_URL");
    const thanks = getRandomThanksWord();
    data.unshift(thanks);
    const payloadData = {
        "content": data.join("\n"),
    };
    for (const url of urls.split(",")) {
        const options = {
            'contentType': 'application/json',
            'payload': JSON.stringify(payloadData),
        };
        UrlFetchApp.fetch(url, options);
    }
}

function cleansingSubject(subject: string): string | null {
    /**
     * Preprocessing subject string
     */


    // ng word filter
    const blackList: string[] = [
        "mentioned you",
        "Confluence changes in the last 24 hours",
    ];
    for (const word of blackList) {
        if (subject.includes(word)) {
            return null;
        }
    }

    // `[Confluence] こんぺこ`  -> `こんぺこ`
    const _subject: string = subject.trim();
    const m = _subject.match(/\[Confluence\](.*)/);
    if (m) {
        return m[1].trim();
    }
    return _subject;
}

function cleansingUserName(name: string): string {
    /**
     * Preprocessing UserName String
     */

    // `ぺこーら (Confluence)" <hoge>`  -> `ぺこーら`
    const _name: string = name.trim();
    const m = _name.match(/.*"(.+) \(Confluence\)".*/);
    if (m) {
        return m[1].trim();
    }
    return _name;
}

function getRandomThanksWord(): string {
    /**
     * Get peko-ra
     * refer: https://seesaawiki.jp/hololivetv/d/%C5%C6%C5%C4%A4%DA%A4%B3%A4%E9%A1%DA%B8%EC%CF%BF%A1%DB
     * TODO: get date from spread sheet
     */
    const wordList: string[] = [
        "お疲れ様ぺこ～",
        "進捗ばっちりぺこ～",
        "こーゆーこと！",
        "更新きｔらあああああ",
        "ニヤリ",
        "任せなさい、と",
        "ありぺこぉ↑",
        "きゅるるんぺこぺこぉ",
        "こころもと強くなった気がする",
        "ぺーこぺこぺこ",
    ];
    return wordList[Math.floor(Math.random() * wordList.length)];
}