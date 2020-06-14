type GeneralDate = Date | GoogleAppsScript.Base.Date;


function main() {
    /**
     * start point method
     */
    run()
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

function japaneseDateFormat(date: GeneralDate): string {
    /**
     * Change Date to Japanese normaly string format.
     * For example, "2020-06-30"
     */
    return Utilities.formatDate(date, "GMT", "yyyy-MM-dd");
}

function run(): null {
    /**
     * main domain function
     */

    const query: string = createQuery();
    const result: GoogleAppsScript.Gmail.GmailThread[] = GmailApp.search(query);

    // message in thread ...
    for (const elem of result) {
        const messages: GoogleAppsScript.Gmail.GmailMessage[] = elem.getMessages();
        for (const message of messages) {
            const messageId = message.getId();
            // TODO: read message ID from Google Spread Sheet

            const subject = cleansingSubject(message.getSubject());
            if (!subject) {
                continue;
            }
            const date = japaneseDateFormat(message.getDate());
            const name = cleansingUserName(message.getFrom());

            postWebHook({
                "content": `${date}\n${subject}\n${name}\n${messageId}`,
            });

            // TODO: write message ID to Google Spread Sheet

        }
    }

    return null;
}

function postWebHook(data) {
    /**
     * send to webhook(Assuming Discord)
     */
    const urls = PropertiesService.getScriptProperties().getProperty("WEB_HOOK_URL");
    for (const url of urls.split(",")) {
        const options = {
            'contentType': 'application/json',
            'payload': JSON.stringify(data),
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