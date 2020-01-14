/**
 *@constructor
 *@param {object} e - event from Google.form trigger by append response
 *@returns {object} a new object with properties and methods to fill the template
 */

var Builder = function(e) {
    this.form = FormApp.getActiveForm();

    /** @property spreadsheet with patterns to replace in template
     *  @returns spreadsheet
     */
    this.dictTable = SpreadsheetApp.openById('1zP5MmCa9TbHgS2PEJelKdrLCGESppE15W4WkU9iM6Wo');

    /** @property
     * sheets with ru and en patterns to replace in template
     */
    this.dict = {
        ru: this.dictTable.getSheets()[0],
        en: this.dictTable.getSheets()[1]
    };

    /** header of ready document */
    this.title = {
        ru: 'СОГЛАШЕНИЕ',
        en: 'AGREEMENT'
    };

    /** ID of spreadsheets to append responses */
    this.tableForRespsID = {
        ru: '1I4eGCZC1nb9Nm27EnjS01Wc993NePAwI08UYXUof0Do',
        en: '1mOj5zPetOQH4tAuIvSKKa-KV8eyFPZElll-Ui7mSri4'
    };

    /** spreadsheets to append responses */
    this.tableForResps = {
        ru: SpreadsheetApp.openById(this.tableForRespsID.ru),
        en: SpreadsheetApp.openById(this.tableForRespsID.en)
    };

    /** event from trigger (the object with responses and other data from form) */
    this.e = e;

    /** IDs of  doc templates */
    this.templateId = {
        en: '1OckKEn1PGboP4argutuxqJISlF6SohWSj4FQEBSX6ZU',
        ru: '1UjPCX27H_rOBt1w7CAl8OUAXlYOHrAIzhIdERl25rzs'
    };

    /** ID of folder to ready docs */
    this.targetFolderId = '1mFf-kId-ZvmAgagO3M0UCWM2G-WaZgdt';

    /** folder to ready docs */
    this.targetFolder = DriveApp.getFolderById(this.targetFolderId);

    /** @returns mail to send ready docs */
    this.mailTo = function() {
        return this.e.response.getRespondentEmail();
    };
};

/**
 *do a copy of template in need language
 *@param {String} lang "ru" or "en"
 *@returns newDoc in need language
 */
Builder.prototype.getNewDoc = function(lang) {
    return DocumentApp.openById(this.getNewFileId(lang));
};

/**
 *put patterns to replace in template from sheet in spreadsheet in need language
 *if pattern does not exist (cell is empty) add "zzzzzzzzzzzzzzzzzzzzz"
 *@param {String} lang "ru" or "en"
 *@returns [array with patterns]
 */
Builder.prototype.getPatterns = function(lang) {
    var emptyArr = [];
    var arr = this.dict[lang].getRange(2, 3, this.dict[lang].getLastRow(), 1).getValues().reduce(function(acc, v) {
        if (v == '') { v = '"zzzzzzzzzzzzzzzzzzzzzzzzz"'; }
        return acc.concat(v);
    }, emptyArr);
    return arr;
};

/**
 *make a copy of template in need language and put ID of new file in cache
 *@param {String} lang "ru" or "en"
 *@returns new file in target folder
 */
Builder.prototype.copyFile = function(lang) { // "ru" or "en"
    var title = this.title[lang] + ' ' + this.responses()[18].translate(lang) + ' ' + this.responses()[17].translate(lang);
    var file = DriveApp.getFileById(this.templateId[lang]).makeCopy(title, this.targetFolder); // .setTrashed(true);
    CacheService.getScriptCache().put(lang, file.getId());
    return file;
};

/**
 *get ID of copy template in need language
 *@param {String} lang "ru" or "en"
 *@returns {string} ID of copy template
 */
Builder.prototype.getNewFileId = function(lang) { // 'ru' or 'en'
    var id = CacheService.getScriptCache().get(lang);
    var sleepTime = 100;
    if (id == null) {
        this.copyFile(lang);
    }
    id = CacheService.getScriptCache().get(lang);
    while (id == null) {
        Utilities.sleep(sleepTime);
        sleepTime *= 1, 5;
        id = CacheService.getScriptCache().get(lang);
    }
    return id;
};

/**
 *replace patterns in copy of template in need language
 *format date in normal view
 *translate to need language (if this need)
 *@param {String} lang "ru" or "en"
 *@param {string} [targetLanguage] if you need to translate responses
 *@returns new document with replaced patterns
 */
Builder.prototype.replacePatterns = function(lang, targetLanguage) {
    var responses = this.responses();
    var re = /\d{4}\-\d{2}\-\d{2}/;
    var body = this.getNewDoc(lang).getBody();
    var hardPatterns = this.dict[lang].getRange(2, 4, this.dict[lang].getLastRow(), 1).getValues();
    this.getPatterns(lang).reduce(function(acc, v, i, arr) {
        if (hardPatterns[i] != '') {
            body.replaceText('{{' + v + '}}', hardPatterns[i]);
        } else if (responses[i] != undefined) {
            if (responses[i].search(re) != -1) {
                body.replaceText('{{' + v + '}}', Utilities.formatDate(new Date(responses[i]), 'GMT+03:00', 'dd.MM.YYYY'));
            } else {
                body.replaceText('{{' + v + '}}', responses[i].translate(targetLanguage));
                Utilities.sleep(1000)
            }
        }
    }, 0)
    this.getNewDoc(lang).saveAndClose();
    return this.getNewDoc(lang);
}

/**
     *open a sheet in need lang of spreadsheet in need lang for append responses
     *@param {String} langTable "ru" or "en"
     *@param {String} langSheet "ru" or "en"
     @returns sheet
     */
Builder.prototype.getRespSheet = function(langTable, langSheet) {
    return this.tableForResps[langTable].getSheetByName(langSheet);
}

/**
 *pul responses
 *@param {[optionArr]} [['lang of table','lang of sheet', 'lang to translate']]
 */
Builder.prototype.appendResponses = function(optionArr) { // [['lang of table','lang of sheet', 'lang to translate']]
    optionArr.map(function(option) {
        var sheet = this.getRespSheet(option[0], option[1])
        var lastRow = sheet.getLastRow()
        var row = [lastRow, new Date(), this.mailTo()].concat(this.responses(option[2]))
        sheet.appendRow(row)
        sheet.insertRowAfter(sheet.getMaxRows())
    }, this)
}

/**
 * get responses in need lang
 * @param {String} lang "ru" or "en"
 * @returns [array]
 */
Builder.prototype.responses = function(lang) {
    return this.e.response.getItemResponses().map(function(v) { return v.getResponse().translate(lang) })
}

Builder.prototype.responsesRow = function() {
    return this.responses().reduce(function(a, b) { return a[0].concat(b[0]) })
}

Builder.prototype.getTemplate = function(lang) {
    DocumentApp.openById(this.getFileId())
}

Builder.prototype.ruPDF = function() {
    return this.replacePatterns('ru').getAs(MimeType.PDF)
}

Builder.prototype.engPDF = function() {
    return this.replacePatterns('en', 'en').getAs(MimeType.PDF)
}

/**
 * get new docs with replaced patterns
 * [["язык исходного документа", "язык, на который нужно перевести"]]
 * @param {array} [['ru','en']] or [['en', 'ru']] or [['ru', '']]
 * @returns [array of PDFs]
 */
Builder.prototype.getPDF = function(optionArr) {
    var attachmentsArr = []
    optionArr.map(
        function(option) {
            option = this.replacePatterns(option[0], option[1]).getAs(MimeType.PDF)
            attachmentsArr.push(option)
        }, this
    )

    return {
        attachments: attachmentsArr
            // name: "Vasya!", // можно указать имя отправителя
    }
}

/**
 * send email from our acc with ready docs in PDF as attachments
 * [["язык исходного документа", "язык, на который нужно перевести"]]
 * @param {array} [['ru','en']] or [['en', 'ru']] or [['ru', '']]
 */
Builder.prototype.sendMail = function(optionArr) {
    MailApp.sendEmail(this.mailTo(), 'Documents', 'Направляем Вам необходимые документы', this.getPDF(optionArr))
}

/* Метод для перевода текста с помощью гугл-переводчика аргумент: обозначение языка, на который нужно перевести  */
/**
 * translate string with Google.translate
 * @param {string} [targetLanguage] a lang to translate "ru" or "en"
 * @returns {string} string
 */
String.prototype.translate = function(targetLanguage) {
    if (targetLanguage == 'en') {
        return LanguageApp.translate(this, 'ru', 'en');
    } else if (targetLanguage == 'ru') {
        return LanguageApp.translate(this, 'en', 'ru');
    }
    return this;
}

/**
 * @constructor
 * @param {string} formId ID of Google form
 */
var FORM = function(formId) {
    this.formId = formId;
    this.form = function() { return FormApp.openById(this.formId) }

    /**
     * get titles of questions in form in column
     * @returns [[title],[title]]
     */
    this.getTitles = function() {
        return this.form().getItems().reduce(
            function(acc, v) {
                if (v.getType() != 'SECTION_HEADER' && v.getType() != 'PAGE_BREAK') {
                    acc.push([v.getTitle()])
                }
                return acc
            }, [])
    }

    /**
     * get titles of questions in form in row
     * @returns [title,title]
     */
    this.getTitlesRow = function() {
        var row = ['№', 'date', 'mail'].concat(this.getTitles().map(function(v) { return v[0] }))
        return row
    }

    this.getTitlesLength = function() { return this.getTitles().length; }
}