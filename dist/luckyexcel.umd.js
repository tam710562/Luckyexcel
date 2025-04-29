(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory(require('nanoid'), require('dayjs'), require('@univerjs/core'), require('@progress/jszip-esm'), require('@zwight/exceljs'), require('papaparse')) :
    typeof define === 'function' && define.amd ? define(['nanoid', 'dayjs', '@univerjs/core', '@progress/jszip-esm', '@zwight/exceljs', 'papaparse'], factory) :
    (global = typeof globalThis !== 'undefined' ? globalThis : global || self, global.LuckyExcel = factory(global.nanoid, global.dayjs, global.core, global.JSZip, global.exceljs, global.Papa));
})(this, (function (nanoid, dayjs, core, JSZip, exceljs, Papa) { 'use strict';

    function _interopDefaultLegacy (e) { return e && typeof e === 'object' && 'default' in e ? e : { 'default': e }; }

    var dayjs__default = /*#__PURE__*/_interopDefaultLegacy(dayjs);
    var JSZip__default = /*#__PURE__*/_interopDefaultLegacy(JSZip);
    var exceljs__default = /*#__PURE__*/_interopDefaultLegacy(exceljs);
    var Papa__default = /*#__PURE__*/_interopDefaultLegacy(Papa);

    const columeHeader_word = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    const columeHeader_word_index = { 'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19, 'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25 };
    const coreFile = "docProps/core.xml";
    const appFile = "docProps/app.xml";
    const workBookFile = "xl/workbook.xml";
    const calcChainFile = "xl/calcChain.xml";
    const stylesFile = "xl/styles.xml";
    const sharedStringsFile = "xl/sharedStrings.xml";
    const worksheetFilePath = "xl/worksheets/";
    const theme1File = "xl/theme/theme1.xml";
    const workbookRels = "xl/_rels/workbook.xml.rels";
    const cellImages = 'xl/cellimages.xml';
    const cellImagesRels = 'xl/_rels/cellimages.xml.rels';
    //Excel Built-In cell type
    const ST_CellType = {
        "Boolean": "b",
        "Date": "d",
        "Error": "e",
        "InlineString": "inlineStr",
        "Number": "n",
        "SharedString": "s",
        "String": "str",
    };
    let numFmtDefault = {
        "0": 'General',
        "1": '0',
        "2": '0.00',
        "3": '#,##0',
        "4": '#,##0.00',
        "9": '0%',
        "10": '0.00%',
        "11": '0.00E+00',
        "12": '# ?/?',
        "13": '# ??/??',
        "14": 'm/d/yy',
        "15": 'd-mmm-yy',
        "16": 'd-mmm',
        "17": 'mmm-yy',
        "18": 'h:mm AM/PM',
        "19": 'h:mm:ss AM/PM',
        "20": 'h:mm',
        "21": 'h:mm:ss',
        "22": 'm/d/yy h:mm',
        "37": '#,##0 ;(#,##0)',
        "38": '#,##0 ;[Red](#,##0)',
        "39": '#,##0.00;(#,##0.00)',
        "40": '#,##0.00;[Red](#,##0.00)',
        "45": 'mm:ss',
        "46": '[h]:mm:ss',
        "47": 'mmss.0',
        "48": '##0.0E+0',
        "49": '@'
    };
    const indexedColors = {
        "0": '00000000',
        "1": '00FFFFFF',
        "2": '00FF0000',
        "3": '0000FF00',
        "4": '000000FF',
        "5": '00FFFF00',
        "6": '00FF00FF',
        "7": '0000FFFF',
        "8": '00000000',
        "9": '00FFFFFF',
        "10": '00FF0000',
        "11": '0000FF00',
        "12": '000000FF',
        "13": '00FFFF00',
        "14": '00FF00FF',
        "15": '0000FFFF',
        "16": '00800000',
        "17": '00008000',
        "18": '00000080',
        "19": '00808000',
        "20": '00800080',
        "21": '00008080',
        "22": '00C0C0C0',
        "23": '00808080',
        "24": '009999FF',
        "25": '00993366',
        "26": '00FFFFCC',
        "27": '00CCFFFF',
        "28": '00660066',
        "29": '00FF8080',
        "30": '000066CC',
        "31": '00CCCCFF',
        "32": '00000080',
        "33": '00FF00FF',
        "34": '00FFFF00',
        "35": '0000FFFF',
        "36": '00800080',
        "37": '00800000',
        "38": '00008080',
        "39": '000000FF',
        "40": '0000CCFF',
        "41": '00CCFFFF',
        "42": '00CCFFCC',
        "43": '00FFFF99',
        "44": '0099CCFF',
        "45": '00FF99CC',
        "46": '00CC99FF',
        "47": '00FFCC99',
        "48": '003366FF',
        "49": '0033CCCC',
        "50": '0099CC00',
        "51": '00FFCC00',
        "52": '00FF9900',
        "53": '00FF6600',
        "54": '00666699',
        "55": '00969696',
        "56": '00003366',
        "57": '00339966',
        "58": '00003300',
        "59": '00333300',
        "60": '00993300',
        "61": '00993366',
        "62": '00333399',
        "63": '00333333',
        "64": null, //system Foreground n/a
        "65": null, //system Background n/a
    };
    const borderTypes = {
        "none": 0,
        "thin": 1,
        "hair": 2,
        "dotted": 3,
        "dashed": 4,
        "dashDot": 5,
        "dashDotDot": 6,
        "double": 7,
        "medium": 8,
        "mediumDashed": 9,
        "mediumDashDot": 10,
        "mediumDashDotDot": 11,
        "slantDashDot": 12,
        "thick": 13
    };
    let numFmtDefaultMap = {
        "yyyy/m/d;@": "yyyy/MM/dd",
        "yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;;@": "yyyy&quot;年&quot;MM&quot;月&quot;dd&quot;日&quot;",
        "[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@": "yyyy/MM/dd hh:mm AM/PM",
    };
    const fontFamilys = {
        "0": "defualt",
        "1": "Roman",
        "2": "Swiss",
        "3": "Modern",
        "4": "Script",
        "5": "Decorative"
    };
    const DATA_VERIFICATION_MAP = {
        list: "dropdown",
        whole: "number_integer",
        decimal: "number_decimal",
        custom: "text_content",
        textLength: "text_length",
        date: "date",
        "unknown1": "number", // no match yet
        "unknown2": "checkbox", // no match yet
        "unknown3": "validity", // no match yet
    };
    const COMMON_TYPE2 = [
        "number",
        "number_integer",
        "number_decimal",
        "text_length",
    ];
    const DATA_VERIFICATION_TYPE2_MAP = {
        common: {
            between: "bw",
            notBetween: "nb",
            equal: "eq",
            notEqualTo: "ne",
            moreThanThe: "gt",
            lessThan: "lt",
            greaterOrEqualTo: "gte",
            lessThanOrEqualTo: "lte",
        },
        text_content: {
            include: "include",
            exclude: "exclude",
            equal: "equal",
        },
        date: {
            between: "bw",
            notBetween: "nb",
            equal: "eq",
            notEqualTo: "ne",
            earlierThan: "bf",
            noEarlierThan: "nbf",
            laterThan: "af",
            noLaterThan: "naf",
        },
        validity: {
            card: "card",
            phone: "phone",
        },
    };
    var ChartAttributeBits;
    (function (ChartAttributeBits) {
        ChartAttributeBits[ChartAttributeBits["Stack"] = 1073741824] = "Stack";
        ChartAttributeBits[ChartAttributeBits["PercentStack"] = 1610612736] = "PercentStack";
        ChartAttributeBits[ChartAttributeBits["Horizontal"] = 268435456] = "Horizontal";
    })(ChartAttributeBits || (ChartAttributeBits = {}));
    var ChartTypeBits;
    (function (ChartTypeBits) {
        ChartTypeBits[ChartTypeBits["None"] = 0] = "None";
        ChartTypeBits[ChartTypeBits["Line"] = 2] = "Line";
        ChartTypeBits[ChartTypeBits["Column"] = 4] = "Column";
        ChartTypeBits[ChartTypeBits["ColumnStacked"] = 1073741828] = "ColumnStacked";
        ChartTypeBits[ChartTypeBits["ColumnPercentStacked"] = 1610612740] = "ColumnPercentStacked";
        ChartTypeBits[ChartTypeBits["Bar"] = 268435460] = "Bar";
        ChartTypeBits[ChartTypeBits["BarStacked"] = 1342177284] = "BarStacked";
        ChartTypeBits[ChartTypeBits["BarPercentStacked"] = 1879048196] = "BarPercentStacked";
        ChartTypeBits[ChartTypeBits["Pie"] = 8] = "Pie";
        ChartTypeBits[ChartTypeBits["Doughnut"] = 264] = "Doughnut";
        ChartTypeBits[ChartTypeBits["Area"] = 16] = "Area";
        ChartTypeBits[ChartTypeBits["AreaStacked"] = 1073741840] = "AreaStacked";
        ChartTypeBits[ChartTypeBits["AreaPercentStacked"] = 1610612752] = "AreaPercentStacked";
        ChartTypeBits[ChartTypeBits["Radar"] = 32] = "Radar";
        ChartTypeBits[ChartTypeBits["Scatter"] = 64] = "Scatter";
        ChartTypeBits[ChartTypeBits["Combination"] = 128] = "Combination";
    })(ChartTypeBits || (ChartTypeBits = {}));
    var LabelContentType;
    (function (LabelContentType) {
        LabelContentType[LabelContentType["Empty"] = 0] = "Empty";
        LabelContentType[LabelContentType["CategoryName"] = 2] = "CategoryName";
        LabelContentType[LabelContentType["SeriesName"] = 4] = "SeriesName";
        LabelContentType[LabelContentType["Value"] = 8] = "Value";
        LabelContentType[LabelContentType["Percentage"] = 16] = "Percentage";
    })(LabelContentType || (LabelContentType = {}));

    function getRangetxt(range, sheettxt) {
        let row0 = range["row"][0], row1 = range["row"][1];
        let column0 = range["column"][0], column1 = range["column"][1];
        if (row0 == null && row1 == null) {
            return sheettxt + numberToABC(column0) + ":" + numberToABC(column1);
        }
        else if (column0 == null && column1 == null) {
            return sheettxt + (row0 + 1) + ":" + (row1 + 1);
        }
        else {
            if (column0 == column1 && row0 == row1) {
                return sheettxt + numberToABC(column0) + (row0 + 1);
            }
            else {
                return sheettxt + numberToABC(column0) + (row0 + 1) + ":" + numberToABC(column1) + (row1 + 1);
            }
        }
    }
    function getcellrange(txt) {
        let val = txt.split("!");
        let rangetxt = "";
        if (val.length > 1) {
            val[0];
            rangetxt = val[1];
        }
        else {
            rangetxt = val[0];
        }
        if (rangetxt.indexOf(":") == -1) {
            let row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
            let col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
            if (!isNaN(row) && !isNaN(col)) {
                return {
                    "row": [row, row],
                    "column": [col, col],
                };
            }
            else {
                return null;
            }
        }
        else {
            let rangetxtArray = rangetxt.split(":");
            let row = [], col = [];
            row[0] = parseInt(rangetxtArray[0].replace(/[^0-9]/g, "")) - 1;
            row[1] = parseInt(rangetxtArray[1].replace(/[^0-9]/g, "")) - 1;
            // if (isNaN(row[0])) {
            //     row[0] = 0;
            // }
            // if (isNaN(row[1])) {
            //     row[1] = sheetdata.length - 1;
            // }
            if (row[0] > row[1]) {
                return null;
            }
            col[0] = ABCatNum(rangetxtArray[0].replace(/[^A-Za-z]/g, ""));
            col[1] = ABCatNum(rangetxtArray[1].replace(/[^A-Za-z]/g, ""));
            // if (isNaN(col[0])) {
            //     col[0] = 0;
            // }
            // if (isNaN(col[1])) {
            //     col[1] = sheetdata[0].length - 1;
            // }
            if (col[0] > col[1]) {
                return null;
            }
            return {
                "row": row,
                "column": col,
            };
        }
    }
    //列下标  字母转数字
    function ABCatNum(abc) {
        abc = abc.toUpperCase();
        let abc_len = abc.length;
        if (abc_len == 0) {
            return NaN;
        }
        let abc_array = abc.split("");
        let wordlen = columeHeader_word.length;
        let ret = 0;
        for (let i = abc_len - 1; i >= 0; i--) {
            if (i == abc_len - 1) {
                ret += columeHeader_word_index[abc_array[i]];
            }
            else {
                ret += Math.pow(wordlen, abc_len - i - 1) * (columeHeader_word_index[abc_array[i]] + 1);
            }
        }
        return ret;
    }
    //列下标  数字转字母
    function chatatABC(index) {
        let wordlen = columeHeader_word.length;
        if (index < wordlen) {
            return columeHeader_word[index];
        }
        else {
            let last = 0, ret = "";
            let i = 1, n = 0;
            while (index >= (wordlen / (wordlen - 1)) * (Math.pow(wordlen, i++) - 1)) {
                n = i;
            }
            let index_ab = index - (wordlen / (wordlen - 1)) * (Math.pow(wordlen, n - 1) - 1); //970
            last = index_ab + 1;
            for (let x = n; x > 0; x--) {
                let last1 = last; //-702=268, 3
                if (x == 1) {
                    last1 = last1 % wordlen;
                    if (last1 == 0) {
                        last1 = 26;
                    }
                    return ret + columeHeader_word[last1 - 1];
                }
                last1 = Math.ceil(last1 / Math.pow(wordlen, x - 1));
                //last1 = last1 % wordlen;
                ret += columeHeader_word[last1 - 1];
                if (x > 1) {
                    last = last - (last1 - 1) * wordlen;
                }
            }
        }
    }
    /**
     * @return ratio, default 0.75 1in = 2.54cm = 25.4mm = 72pt = 6pc,  pt = 1/72 In, px = 1/dpi In
    */
    function getptToPxRatioByDPI() {
        return 72 / 96;
    }
    /**
     * @emus EMUs, Excel drawing unit
     * @return pixel
    */
    function getPxByEMUs(emus) {
        if (emus == null) {
            return 0;
        }
        let inch = emus / 914400;
        let pt = inch * 72;
        let px = pt / getptToPxRatioByDPI();
        return px;
    }
    /**
     * @emus EMUs, Excel drawing unit
     * @return pixel
    */
    function getEmusByPx(px) {
        const pt = px * getptToPxRatioByDPI();
        const inch = pt / 72;
        let emus = inch * 914400;
        return emus;
    }
    /**
     * @dom xml attribute object
     * @attr attribute name
     * @d if attribute is null, return default value
     * @return attribute value
    */
    function getXmlAttibute(dom, attr, d) {
        let value = dom[attr];
        value = value == null ? d : value;
        return value;
    }
    /**
     * @columnWidth Excel column width
     * @return pixel column width
    */
    function getColumnWidthPixel(columnWidth) {
        let pix = Math.round((columnWidth - 0.83) * 8 + 5);
        return pix;
    }
    /**
     * @rowHeight Excel row height
     * @return pixel row height
    */
    function getRowHeightPixel(rowHeight) {
        let pix = Math.round(rowHeight / getptToPxRatioByDPI());
        return pix;
    }
    function LightenDarkenColor(sixColor, tint) {
        let hex = sixColor.substring(sixColor.length - 6, sixColor.length);
        let rgbArray = hexToRgbArray("#" + hex);
        let hslArray = rgbToHsl(rgbArray[0], rgbArray[1], rgbArray[2]);
        if (tint > 0) {
            hslArray[2] = hslArray[2] * (1.0 - tint) + tint;
        }
        else if (tint < 0) {
            hslArray[2] = hslArray[2] * (1.0 + tint);
        }
        else {
            return "#" + hex;
        }
        let newRgbArray = hslToRgb(hslArray[0], hslArray[1], hslArray[2]);
        return rgbToHex("RGB(" + newRgbArray.join(",") + ")");
    }
    function rgbToHex(rgb) {
        //十六进制颜色值的正则表达式
        var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
        // 如果是rgb颜色表示
        if (/^(rgb|RGB)/.test(rgb)) {
            var aColor = rgb.replace(/(?:\(|\)|rgb|RGB)*/g, "").split(",");
            var strHex = "#";
            for (var i = 0; i < aColor.length; i++) {
                var hex = Number(aColor[i]).toString(16);
                if (hex.length < 2) {
                    hex = '0' + hex;
                }
                strHex += hex;
            }
            if (strHex.length !== 7) {
                strHex = rgb;
            }
            return strHex;
        }
        else if (reg.test(rgb)) {
            var aNum = rgb.replace(/#/, "").split("");
            if (aNum.length === 6) {
                return rgb;
            }
            else if (aNum.length === 3) {
                var numHex = "#";
                for (var i = 0; i < aNum.length; i += 1) {
                    numHex += (aNum[i] + aNum[i]);
                }
                return numHex;
            }
        }
        return rgb;
    }
    function hexToRgbArray(hex) {
        var sColor = hex.toLowerCase();
        //十六进制颜色值的正则表达式
        var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
        // 如果是16进制颜色
        if (sColor && reg.test(sColor)) {
            if (sColor.length === 4) {
                var sColorNew = "#";
                for (var i = 1; i < 4; i += 1) {
                    sColorNew += sColor.slice(i, i + 1).concat(sColor.slice(i, i + 1));
                }
                sColor = sColorNew;
            }
            //处理六位的颜色值
            var sColorChange = [];
            for (var i = 1; i < 7; i += 2) {
                sColorChange.push(parseInt("0x" + sColor.slice(i, i + 2)));
            }
            return sColorChange;
        }
        return null;
    }
    /**
     * HSL颜色值转换为RGB.
     * 换算公式改编自 http://en.wikipedia.org/wiki/HSL_color_space.
     * h, s, 和 l 设定在 [0, 1] 之间
     * 返回的 r, g, 和 b 在 [0, 255]之间
     *
     * @param   Number  h       色相
     * @param   Number  s       饱和度
     * @param   Number  l       亮度
     * @return  Array           RGB色值数值
     */
    function hslToRgb(h, s, l) {
        var r, g, b;
        if (s == 0) {
            r = g = b = l; // achromatic
        }
        else {
            var hue2rgb = function hue2rgb(p, q, t) {
                if (t < 0)
                    t += 1;
                if (t > 1)
                    t -= 1;
                if (t < 1 / 6)
                    return p + (q - p) * 6 * t;
                if (t < 1 / 2)
                    return q;
                if (t < 2 / 3)
                    return p + (q - p) * (2 / 3 - t) * 6;
                return p;
            };
            var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            var p = 2 * l - q;
            r = hue2rgb(p, q, h + 1 / 3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1 / 3);
        }
        return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
    }
    /**
     * RGB 颜色值转换为 HSL.
     * 转换公式参考自 http://en.wikipedia.org/wiki/HSL_color_space.
     * r, g, 和 b 需要在 [0, 255] 范围内
     * 返回的 h, s, 和 l 在 [0, 1] 之间
     *
     * @param   Number  r       红色色值
     * @param   Number  g       绿色色值
     * @param   Number  b       蓝色色值
     * @return  Array           HSL各值数组
     */
    function rgbToHsl(r, g, b) {
        r /= 255, g /= 255, b /= 255;
        var max = Math.max(r, g, b), min = Math.min(r, g, b);
        var h, s, l = (max + min) / 2;
        if (max == min) {
            h = s = 0; // achromatic
        }
        else {
            var d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case r:
                    h = (g - b) / d + (g < b ? 6 : 0);
                    break;
                case g:
                    h = (b - r) / d + 2;
                    break;
                case b:
                    h = (r - g) / d + 4;
                    break;
            }
            h /= 6;
        }
        return [h, s, l];
    }
    function generateRandomIndex(prefix) {
        if (prefix == null) {
            prefix = "Sheet";
        }
        let userAgent = navigator.userAgent.replace(/[^a-zA-Z0-9]/g, "").split("");
        let mid = "";
        for (let i = 0; i < 5; i++) {
            mid += userAgent[Math.round(Math.random() * (userAgent.length - 1))];
        }
        let time = new Date().getTime();
        return prefix + "_" + mid + "_" + time;
    }
    function escapeCharacter(str) {
        if (str == null || str.length == 0) {
            return str;
        }
        return str.replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&#34;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&apos;/g, "'").replace(/&iexcl;/g, "¡").replace(/&cent;/g, "¢").replace(/&pound;/g, "£").replace(/&curren;/g, "¤").replace(/&yen;/g, "¥").replace(/&brvbar;/g, "¦").replace(/&sect;/g, "§").replace(/&uml;/g, "¨").replace(/&copy;/g, "©").replace(/&ordf;/g, "ª").replace(/&laquo;/g, "«").replace(/&not;/g, "¬").replace(/&shy;/g, "­").replace(/&reg;/g, "®").replace(/&macr;/g, "¯").replace(/&deg;/g, "°").replace(/&plusmn;/g, "±").replace(/&sup2;/g, "²").replace(/&sup3;/g, "³").replace(/&acute;/g, "´").replace(/&micro;/g, "µ").replace(/&para;/g, "¶").replace(/&middot;/g, "·").replace(/&cedil;/g, "¸").replace(/&sup1;/g, "¹").replace(/&ordm;/g, "º").replace(/&raquo;/g, "»").replace(/&frac14;/g, "¼").replace(/&frac12;/g, "½").replace(/&frac34;/g, "¾").replace(/&iquest;/g, "¿").replace(/&times;/g, "×").replace(/&divide;/g, "÷").replace(/&Agrave;/g, "À").replace(/&Aacute;/g, "Á").replace(/&Acirc;/g, "Â").replace(/&Atilde;/g, "Ã").replace(/&Auml;/g, "Ä").replace(/&Aring;/g, "Å").replace(/&AElig;/g, "Æ").replace(/&Ccedil;/g, "Ç").replace(/&Egrave;/g, "È").replace(/&Eacute;/g, "É").replace(/&Ecirc;/g, "Ê").replace(/&Euml;/g, "Ë").replace(/&Igrave;/g, "Ì").replace(/&Iacute;/g, "Í").replace(/&Icirc;/g, "Î").replace(/&Iuml;/g, "Ï").replace(/&ETH;/g, "Ð").replace(/&Ntilde;/g, "Ñ").replace(/&Ograve;/g, "Ò").replace(/&Oacute;/g, "Ó").replace(/&Ocirc;/g, "Ô").replace(/&Otilde;/g, "Õ").replace(/&Ouml;/g, "Ö").replace(/&Oslash;/g, "Ø").replace(/&Ugrave;/g, "Ù").replace(/&Uacute;/g, "Ú").replace(/&Ucirc;/g, "Û").replace(/&Uuml;/g, "Ü").replace(/&Yacute;/g, "Ý").replace(/&THORN;/g, "Þ").replace(/&szlig;/g, "ß").replace(/&agrave;/g, "à").replace(/&aacute;/g, "á").replace(/&acirc;/g, "â").replace(/&atilde;/g, "ã").replace(/&auml;/g, "ä").replace(/&aring;/g, "å").replace(/&aelig;/g, "æ").replace(/&ccedil;/g, "ç").replace(/&egrave;/g, "è").replace(/&eacute;/g, "é").replace(/&ecirc;/g, "ê").replace(/&euml;/g, "ë").replace(/&igrave;/g, "ì").replace(/&iacute;/g, "í").replace(/&icirc;/g, "î").replace(/&iuml;/g, "ï").replace(/&eth;/g, "ð").replace(/&ntilde;/g, "ñ").replace(/&ograve;/g, "ò").replace(/&oacute;/g, "ó").replace(/&ocirc;/g, "ô").replace(/&otilde;/g, "õ").replace(/&ouml;/g, "ö").replace(/&oslash;/g, "ø").replace(/&ugrave;/g, "ù").replace(/&uacute;/g, "ú").replace(/&ucirc;/g, "û").replace(/&uuml;/g, "ü").replace(/&yacute;/g, "ý").replace(/&thorn;/g, "þ").replace(/&yuml;/g, "ÿ");
    }
    class fromulaRef {
        static trim(str) {
            if (str == null) {
                str = "";
            }
            return str.replace(/(^\s*)|(\s*$)/g, "");
        }
        static functionCopy(txt, mode, step) {
            let _this = this;
            if (_this.operatorjson == null) {
                let arr = _this.operator.split("|"), op = {};
                for (let i = 0; i < arr.length; i++) {
                    op[arr[i].toString()] = 1;
                }
                _this.operatorjson = op;
            }
            if (mode == null) {
                mode = "down";
            }
            if (step == null) {
                step = 1;
            }
            if (txt.substr(0, 1) == "=") {
                txt = txt.substr(1);
            }
            let funcstack = txt.split("");
            let i = 0, str = "", function_str = "";
            let matchConfig = {
                "bracket": 0,
                "comma": 0,
                "squote": 0,
                "dquote": 0
            };
            while (i < funcstack.length) {
                let s = funcstack[i];
                if (s == "(" && matchConfig.dquote == 0) {
                    matchConfig.bracket += 1;
                    if (str.length > 0) {
                        function_str += str + "(";
                    }
                    else {
                        function_str += "(";
                    }
                    str = "";
                }
                else if (s == ")" && matchConfig.dquote == 0) {
                    matchConfig.bracket -= 1;
                    function_str += _this.functionCopy(str, mode, step) + ")";
                    str = "";
                }
                else if (s == '"' && matchConfig.squote == 0) {
                    if (matchConfig.dquote > 0) {
                        function_str += str + '"';
                        matchConfig.dquote -= 1;
                        str = "";
                    }
                    else {
                        matchConfig.dquote += 1;
                        str += '"';
                    }
                }
                else if (s == ',' && matchConfig.dquote == 0) {
                    function_str += _this.functionCopy(str, mode, step) + ',';
                    str = "";
                }
                else if (s == '&' && matchConfig.dquote == 0) {
                    if (str.length > 0) {
                        function_str += _this.functionCopy(str, mode, step) + "&";
                        str = "";
                    }
                    else {
                        function_str += "&";
                    }
                }
                else if (s in _this.operatorjson && matchConfig.dquote == 0) {
                    let s_next = "";
                    if ((i + 1) < funcstack.length) {
                        s_next = funcstack[i + 1];
                    }
                    let p = i - 1, s_pre = null;
                    if (p >= 0) {
                        do {
                            s_pre = funcstack[p--];
                        } while (p >= 0 && s_pre == " ");
                    }
                    if ((s + s_next) in _this.operatorjson) {
                        if (str.length > 0) {
                            function_str += _this.functionCopy(str, mode, step) + s + s_next;
                            str = "";
                        }
                        else {
                            function_str += s + s_next;
                        }
                        i++;
                    }
                    else if (!(/[^0-9]/.test(s_next)) && s == "-" && (s_pre == "(" || s_pre == null || s_pre == "," || s_pre == " " || s_pre in _this.operatorjson)) {
                        str += s;
                    }
                    else {
                        if (str.length > 0) {
                            function_str += _this.functionCopy(str, mode, step) + s;
                            str = "";
                        }
                        else {
                            function_str += s;
                        }
                    }
                }
                else {
                    str += s;
                }
                if (i == funcstack.length - 1) {
                    if (_this.iscelldata(_this.trim(str))) {
                        if (mode == "down") {
                            function_str += _this.downparam(_this.trim(str), step);
                        }
                        else if (mode == "up") {
                            function_str += _this.upparam(_this.trim(str), step);
                        }
                        else if (mode == "left") {
                            function_str += _this.leftparam(_this.trim(str), step);
                        }
                        else if (mode == "right") {
                            function_str += _this.rightparam(_this.trim(str), step);
                        }
                    }
                    else {
                        function_str += _this.trim(str);
                    }
                }
                i++;
            }
            return function_str;
        }
        static downparam(txt, step) {
            return this.updateparam("d", txt, step);
        }
        static upparam(txt, step) {
            return this.updateparam("u", txt, step);
        }
        static leftparam(txt, step) {
            return this.updateparam("l", txt, step);
        }
        static rightparam(txt, step) {
            return this.updateparam("r", txt, step);
        }
        static updateparam(orient, txt, step) {
            let _this = this;
            let val = txt.split("!"), rangetxt, prefix = "";
            if (val.length > 1) {
                rangetxt = val[1];
                prefix = val[0] + "!";
            }
            else {
                rangetxt = val[0];
            }
            if (rangetxt.indexOf(":") == -1) {
                let row = parseInt(rangetxt.replace(/[^0-9]/g, ""));
                let col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
                let freezonFuc = isfreezonFuc(rangetxt);
                let $row = freezonFuc[0] ? "$" : "", $col = freezonFuc[1] ? "$" : "";
                if (orient == "u" && !freezonFuc[0]) {
                    row -= step;
                }
                else if (orient == "r" && !freezonFuc[1]) {
                    col += step;
                }
                else if (orient == "l" && !freezonFuc[1]) {
                    col -= step;
                }
                else if (!freezonFuc[0]) {
                    row += step;
                }
                if (row < 0 || col < 0) {
                    return _this.error.r;
                }
                if (!isNaN(row) && !isNaN(col)) {
                    return prefix + $col + chatatABC(col) + $row + (row);
                }
                else if (!isNaN(row)) {
                    return prefix + $row + (row);
                }
                else if (!isNaN(col)) {
                    return prefix + $col + chatatABC(col);
                }
                else {
                    return txt;
                }
            }
            else {
                rangetxt = rangetxt.split(":");
                let row = [], col = [];
                row[0] = parseInt(rangetxt[0].replace(/[^0-9]/g, ""));
                row[1] = parseInt(rangetxt[1].replace(/[^0-9]/g, ""));
                if (row[0] > row[1]) {
                    return txt;
                }
                col[0] = ABCatNum(rangetxt[0].replace(/[^A-Za-z]/g, ""));
                col[1] = ABCatNum(rangetxt[1].replace(/[^A-Za-z]/g, ""));
                if (col[0] > col[1]) {
                    return txt;
                }
                let freezonFuc0 = isfreezonFuc(rangetxt[0]);
                let freezonFuc1 = isfreezonFuc(rangetxt[1]);
                let $row0 = freezonFuc0[0] ? "$" : "", $col0 = freezonFuc0[1] ? "$" : "";
                let $row1 = freezonFuc1[0] ? "$" : "", $col1 = freezonFuc1[1] ? "$" : "";
                if (orient == "u") {
                    if (!freezonFuc0[0]) {
                        row[0] -= step;
                    }
                    if (!freezonFuc1[0]) {
                        row[1] -= step;
                    }
                }
                else if (orient == "r") {
                    if (!freezonFuc0[1]) {
                        col[0] += step;
                    }
                    if (!freezonFuc1[1]) {
                        col[1] += step;
                    }
                }
                else if (orient == "l") {
                    if (!freezonFuc0[1]) {
                        col[0] -= step;
                    }
                    if (!freezonFuc1[1]) {
                        col[1] -= step;
                    }
                }
                else {
                    if (!freezonFuc0[0]) {
                        row[0] += step;
                    }
                    if (!freezonFuc1[0]) {
                        row[1] += step;
                    }
                }
                if (row[0] < 0 || col[0] < 0) {
                    return _this.error.r;
                }
                if (isNaN(col[0]) && isNaN(col[1])) {
                    return prefix + $row0 + (row[0]) + ":" + $row1 + (row[1]);
                }
                else if (isNaN(row[0]) && isNaN(row[1])) {
                    return prefix + $col0 + chatatABC(col[0]) + ":" + $col1 + chatatABC(col[1]);
                }
                else {
                    return prefix + $col0 + chatatABC(col[0]) + $row0 + (row[0]) + ":" + $col1 + chatatABC(col[1]) + $row1 + (row[1]);
                }
            }
        }
        static iscelldata(txt) {
            let val = txt.split("!"), rangetxt;
            if (val.length > 1) {
                rangetxt = val[1];
            }
            else {
                rangetxt = val[0];
            }
            let reg_cell = /^(([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+))$/g; //增加正则判断单元格为字母+数字的格式：如 A1:B3
            let reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+))))$/g; //增加正则判断单元格为字母+数字或字母的格式：如 A1:B3，A:A
            if (rangetxt.indexOf(":") == -1) {
                let row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
                let col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
                if (!isNaN(row) && !isNaN(col) && rangetxt.toString().match(reg_cell)) {
                    return true;
                }
                else if (!isNaN(row)) {
                    return false;
                }
                else if (!isNaN(col)) {
                    return false;
                }
                else {
                    return false;
                }
            }
            else {
                reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+)))|((([0-9]+)|([$][0-9]+s))))$/g;
                rangetxt = rangetxt.split(":");
                let row = [], col = [];
                row[0] = parseInt(rangetxt[0].replace(/[^0-9]/g, "")) - 1;
                row[1] = parseInt(rangetxt[1].replace(/[^0-9]/g, "")) - 1;
                if (row[0] > row[1]) {
                    return false;
                }
                col[0] = ABCatNum(rangetxt[0].replace(/[^A-Za-z]/g, ""));
                col[1] = ABCatNum(rangetxt[1].replace(/[^A-Za-z]/g, ""));
                if (col[0] > col[1]) {
                    return false;
                }
                if (rangetxt[0].toString().match(reg_cellRange) && rangetxt[1].toString().match(reg_cellRange)) {
                    return true;
                }
                else {
                    return false;
                }
            }
        }
    }
    fromulaRef.operator = '==|!=|<>|<=|>=|=|+|-|>|<|/|*|%|&|^';
    fromulaRef.error = {
        v: "#VALUE!", //错误的参数或运算符
        n: "#NAME?", //公式名称错误
        na: "#N/A", //函数或公式中没有可用数值
        r: "#REF!", //删除了由其他公式引用的单元格
        d: "#DIV/0!", //除数是0或空单元格
        nm: "#NUM!", //当公式或函数中某个数字有问题时
        nl: "#NULL!", //交叉运算符（空格）使用不正确
        sp: "#SPILL!" //数组范围有其它值
    };
    fromulaRef.operatorjson = null;
    function isChinese(temp) {
        var re = /[^\u4e00-\u9fa5]/;
        var reg = /[\u3002|\uff1f|\uff01|\uff0c|\u3001|\uff1b|\uff1a|\u201c|\u201d|\u2018|\u2019|\uff08|\uff09|\u300a|\u300b|\u3008|\u3009|\u3010|\u3011|\u300e|\u300f|\u300c|\u300d|\ufe43|\ufe44|\u3014|\u3015|\u2026|\u2014|\uff5e|\ufe4f|\uffe5]/;
        if (reg.test(temp))
            return true;
        if (re.test(temp))
            return false;
        return true;
    }
    function isJapanese(temp) {
        var re = /[^\u0800-\u4e00]/;
        if (re.test(temp))
            return false;
        return true;
    }
    function isKoera(chr) {
        if (((chr > 0x3130 && chr < 0x318F) ||
            (chr >= 0xAC00 && chr <= 0xD7A3))) {
            return true;
        }
        return false;
    }
    function getBinaryContent(path, options) {
        let promise, resolve, reject;
        let callback;
        if (!options) {
            options = {};
        }
        // taken from jQuery
        let createStandardXHR = function () {
            try {
                return new window.XMLHttpRequest();
            }
            catch (e) { }
        };
        let createActiveXHR = function () {
            try {
                return new window.ActiveXObject("Microsoft.XMLHTTP");
            }
            catch (e) { }
        };
        // Create the request object
        var createXHR = (typeof window !== "undefined" && window.ActiveXObject) ?
            /* Microsoft failed to properly
            * implement the XMLHttpRequest in IE7 (can't request local files),
            * so we use the ActiveXObject when it is available
            * Additionally XMLHttpRequest can be disabled in IE7/IE8 so
            * we need a fallback.
            */
            function () {
                return createStandardXHR() || createActiveXHR();
            } :
            // For all other browsers, use the standard XMLHttpRequest object
            createStandardXHR;
        // backward compatible callback
        if (typeof options === "function") {
            callback = options;
            options = {};
        }
        else if (typeof options.callback === 'function') {
            // callback inside options object
            callback = options.callback;
        }
        resolve = function (data) { callback(null, data); };
        reject = function (err) { callback(err, null); };
        try {
            var xhr = createXHR();
            xhr.open('GET', path, true);
            // recent browsers
            if ("responseType" in xhr) {
                xhr.responseType = "arraybuffer";
            }
            // older browser
            if (xhr.overrideMimeType) {
                xhr.overrideMimeType("text/plain; charset=x-user-defined");
            }
            xhr.onreadystatechange = function (event) {
                // use `xhr` and not `this`... thanks IE
                if (xhr.readyState === 4) {
                    if (xhr.status === 200 || xhr.status === 0) {
                        try {
                            resolve(function (xhr) {
                                // for xhr.responseText, the 0xFF mask is applied by JSZip
                                return xhr.response || xhr.responseText;
                            }(xhr));
                        }
                        catch (err) {
                            reject(new Error(err));
                        }
                    }
                    else {
                        reject(new Error("Ajax error for " + path + " : " + this.status + " " + this.statusText));
                    }
                }
            };
            if (options.progress) {
                xhr.onprogress = function (e) {
                    options.progress({
                        path: path,
                        originalEvent: e,
                        percent: e.loaded / e.total * 100,
                        loaded: e.loaded,
                        total: e.total
                    });
                };
            }
            xhr.send();
        }
        catch (e) {
            reject(new Error(e), null);
        }
        // returns a promise or undefined depending on whether a callback was
        // provided
        return promise;
    }
    /**
     * multi sequence conversion
     * example:
     *  1、E14 -> 13_4
     *  2、E14 J14 O14 T14 Y14 AD14 AI14 AN14 AS14 AX14 ->
     *     ['13_4', '13_9','13_14', '13_19', '13_24', '13_3', '13_8',  '13_13', '13_18', '13_23']
     *  3、E46:E47 -> ['45_4',  '46_4']
     *
     * @param {string} sqref - before sequence
     * @returns {string[]}
     */
    function getMultiSequenceToNum(sqref) {
        if (!sqref || sqref?.length <= 0)
            return [];
        sqref = sqref.toUpperCase();
        let sqrefRawArr = sqref.split(" ");
        let sqrefArr = sqrefRawArr.filter((e) => e && e.trim());
        let sqrefLastArr = getSqrefRawArrFormat(sqrefArr);
        let resArr = [];
        for (let i = 0; i < sqrefLastArr.length; i++) {
            let _res = getSingleSequenceToNum(sqrefLastArr[i]);
            if (_res)
                resArr.push(_res);
        }
        return resArr;
    }
    /**
     * get region sequence
     * example:
     *  1、[A1:C2'] -> ['A1', 'A2', 'B1', 'B2', 'C1', 'C2']
     *
     * @param {string[]} arr - formats arr
     * @returns {string[]} - after arr
     */
    function getRegionSequence(arr) {
        let formatArr = [];
        const regEn = new RegExp(/[A-Z]+|[0-9]+/g);
        const startArr = arr[0]?.match(regEn);
        const lastArr = arr[1]?.match(regEn);
        const columnMax = Math.max(...[ABCatNum(startArr[0]), ABCatNum(lastArr[0])]);
        const columnMin = Math.min(...[ABCatNum(startArr[0]), ABCatNum(lastArr[0])]);
        const rowMax = Math.max(...[parseInt(startArr[1]), parseInt(lastArr[1])]);
        const rowMin = Math.min(...[parseInt(startArr[1]), parseInt(lastArr[1])]);
        for (let i = columnMin; i <= columnMax; i++) {
            for (let j = rowMin; j <= rowMax; j++) {
                formatArr.push(`${chatatABC(i)}${j}`);
            }
        }
        return formatArr;
    }
    /**
     * unified processing of conversion formats
     * example:
     *  1、['E38', 'A1:C2'] -> ['E38', 'A1', 'A2', 'B1', 'B2', 'C1', 'C2']
     *
     * @param {string[]} arr - formats arr
     * @returns {string[]} - after arr
     */
    function getSqrefRawArrFormat(arr) {
        arr?.map((el) => {
            if (el.includes(":")) {
                let tempArr = el.split(":");
                if (tempArr?.length === 2) {
                    arr = arr.concat(getRegionSequence(tempArr));
                    arr.splice(arr.indexOf(el), 1);
                }
            }
        });
        const resultArr = arr.filter((value, index, array) => array.indexOf(value) === index);
        return resultArr;
    }
    /**
     * single sequence to number
     * example:
     *  1、A1 -> 0_0
     *  2、ES14 -> 13_4
     *
     * @param {string} sqref - before sequence
     * @returns {string} - after sequence
     */
    function getSingleSequenceToNum(sqref) {
        let sqrefArray = sqref.match(/[A-Z]+|[0-9]+/g);
        let sqrefLen = sqrefArray.length;
        let regEn = new RegExp("^[A-Z]+$");
        let ret = "";
        for (let i = sqrefLen - 1; i >= 0; i--) {
            let cur = sqrefArray[i];
            if (regEn.test(cur)) {
                ret += ABCatNum(cur) + "_";
            }
            else {
                ret += parseInt(cur) - 1 + "_";
            }
        }
        return ret.substring(0, ret.length - 1);
    }
    /**
     * R1C1 to Sequence
     * example: sheet2!R1C1 => sheet!A1
     *
     * @param {string} value - R1C1 value
     * @returns
     */
    function getTransR1C1ToSequence(value) {
        if (!value && value?.length <= 0)
            return "";
        const len = value.length;
        const index = value.lastIndexOf("!");
        const valueArr = [value.slice(0, index), value.slice(index + 1, len)];
        const repStr = valueArr[1] || "";
        const indexR = repStr.indexOf("R");
        const indexC = repStr.indexOf("C");
        const row = Number(repStr.slice(indexR + 1, indexC));
        const column = chatatABC(Number(repStr.slice(indexC + 1, repStr?.length)) - 1);
        return `${valueArr[0]}!${column}${row}`;
    }
    /**
     * strip x14 format data
     *
     * @param {string} value
     * @returns {Object} - { formula, sqref }
     */
    function getPeelOffX14(value) {
        if (!value || value?.length <= 0)
            return {};
        // formula
        const formulaReg = new RegExp("</x14:formula[^]>", "g");
        const lastIndex = value.match(formulaReg)?.length;
        const lastValue = `</x14:formula${lastIndex}>`;
        const lastValueEnd = value.indexOf(lastValue);
        let formulaValue = value.substring(0, lastValueEnd + lastValue.length);
        formulaValue = formulaValue
            .replace(/<xm:f>/g, "")
            .replace(/<\/xm:f>/g, "")
            .replace(/x14:/g, "")
            .replace(/\/x14:/g, "");
        const formula = formulaValue;
        // sqref
        const xmSqrefLen = "<xm:sqref>".length;
        const sqrefStart = value.indexOf("<xm:sqref>");
        const sqrefEnd = value.indexOf("</xm:sqref>");
        const sqref = value.substring(sqrefStart + xmSqrefLen, sqrefEnd);
        return {
            formula,
            sqref,
        };
    }
    /**
     * get the value in the formula
     *
     * @param {string} value - extracted value
     * @returns {string[]}
     */
    function getMultiFormulaValue(value) {
        if (!value || value?.length <= 0)
            return [];
        const lenReg = new RegExp("formula", "g");
        const len = (value.match(lenReg)?.length || 0) / 2;
        if (len === 0)
            return [];
        let retArr = [];
        for (let i = 1; i <= len; i++) {
            const startLen = `<formula${i}>`?.length;
            const start = value.indexOf(`<formula${i}>`);
            const end = value.indexOf(`</formula${i}>`);
            const _value = value.substring(start + startLen, end);
            retArr.push(escapeCharacter(_value.replace(/&quot;|^\"|\"$/g, "")));
        }
        return retArr;
    }
    function isfreezonFuc(txt) {
        let row = txt.replace(/[^0-9]/g, "");
        let col = txt.replace(/[^A-Za-z]/g, "");
        let row$ = txt.substr(txt.indexOf(row) - 1, 1);
        let col$ = txt.substr(txt.indexOf(col) - 1, 1);
        let ret = [false, false];
        if (row$ == "$") {
            ret[0] = true;
        }
        if (col$ == "$") {
            ret[1] = true;
        }
        return ret;
    }
    function ABCToNumber(a) {
        if (a == null || a.length === 0) {
            return Number.NaN;
        }
        const str = a.toLowerCase().split('');
        const al = str.length;
        const getCharNumber = (charX) => charX.charCodeAt(0) - 96;
        let numOut = 0;
        let charnum = 0;
        for (let i = 0; i < al; i++) {
            charnum = getCharNumber(str[i]);
            numOut += charnum * 26 ** (al - i - 1);
        }
        if (numOut === 0) {
            return Number.NaN;
        }
        return numOut - 1;
    }
    const orderA = 'A'.charCodeAt(0);
    const orderZ = 'Z'.charCodeAt(0);
    /**
     * column subscript number to letters
     * @param n number
     * @returns
     */
    function numberToABC(n) {
        const len = orderZ - orderA + 1;
        let s = '';
        while (n >= 0) {
            s = String.fromCharCode((n % len) + orderA) + s;
            n = Math.floor(n / len) - 1;
        }
        return s;
    }
    function str2num(val) {
        if (val === undefined || val === null) {
            return val;
        }
        try {
            if (Number.isNaN(Number(val)))
                return val;
            return Number(val);
        }
        catch (error) {
            return val;
        }
    }
    function generateRandomId(n = 21, alphabet) {
        if (alphabet) {
            return nanoid.customAlphabet(alphabet, n)();
        }
        return nanoid.nanoid(n);
    }
    function jsonParse(str) {
        if (!str)
            return {};
        try {
            return JSON.parse(str);
        }
        catch (error) {
            return null;
        }
    }
    function isObject(obj) {
        return Object.prototype.toString.call(obj) === '[object Object]';
    }
    /**
     * 删除对象中含undefined的值
     * @param object
     * @returns
     */
    function removeEmptyAttr$1(object) {
        for (const key in object) {
            if (Object.prototype.hasOwnProperty.call(object, key)) {
                if (object[key] === undefined) {
                    delete object[key]; // 删除值为 undefined 的属性
                }
                else if (isObject(object[key]) && object[key] !== null) {
                    removeEmptyAttr$1(object[key]); // 对子对象递归
                }
            }
        }
        return object;
    }
    function isEmpty(value) {
        if (value === undefined || value === null) {
            return true;
        }
        return false;
    }
    function getRelationShip(params) {
        const { rid, fileName, callback, readXml } = params;
        let Relationships = readXml.getElementsByTagName("Relationships/Relationship", fileName);
        if (Relationships != null && Relationships.length > 0) {
            for (let i = 0; i < Relationships.length; i++) {
                let Relationship = Relationships[i];
                let attrList = Relationship.attributeList;
                let Id = getXmlAttibute(attrList, "Id", null);
                let src = getXmlAttibute(attrList, "Target", null);
                if (Id == rid) {
                    src = src.replace(/\.\.\//g, "");
                    if (callback) {
                        return callback(src);
                    }
                    else {
                        return src;
                    }
                }
            }
        }
        return null;
    }

    class xmloperation {
        /**
        * @param tag Search xml tag name , div,title etc.
        * @param file Xml string
        * @return Xml element string
        */
        getElementsByOneTag(tag, file) {
            //<a:[^/>: ]+?>.*?</a:[^/>: ]+?>
            let readTagReg;
            if (tag.indexOf("|") > -1) {
                let tags = tag.split("|"), tagsRegTxt = "";
                for (let i = 0; i < tags.length; i++) {
                    let t = tags[i];
                    tagsRegTxt += "|<" + t + " [^>]+?[^/]>[\\s\\S]*?</" + t + ">|<" + t + " [^>]+?/>|<" + t + ">[\\s\\S]*?</" + t + ">|<" + t + "/>";
                }
                tagsRegTxt = tagsRegTxt.substr(1, tagsRegTxt.length);
                readTagReg = new RegExp(tagsRegTxt, "g");
            }
            else {
                readTagReg = new RegExp("<" + tag + " [^>]+?[^/]>[\\s\\S]*?</" + tag + ">|<" + tag + " [^>]+?/>|<" + tag + ">[\\s\\S]*?</" + tag + ">|<" + tag + "/>", "g");
            }
            let ret = file.match(readTagReg);
            if (ret == null) {
                return [];
            }
            else {
                return ret;
            }
        }
        getElementByTagLink(tag, file) {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(file, "text/xml");
            let tagVal = Array.from(xmlDoc.children);
            if (tag.indexOf("/") > -1) {
                let tags = tag.split("/");
                for (let index = 0; index < tags.length; index++) {
                    const element = tags[index];
                    const i = tagVal.findIndex(d => Array.from(d.children).findIndex(d => d.tagName === element) > -1);
                    if (i === -1 && index <= tags.length - 1) {
                        return [];
                    }
                    tagVal = Array.from(tagVal[i].children).filter(d => d.tagName === element);
                }
            }
            else {
                tagVal = Array.from(tagVal[0].children).filter(d => d.tagName === tag);
            }
            const serializer = new XMLSerializer();
            return tagVal.map(d => serializer.serializeToString(d));
        }
    }
    class ReadXml extends xmloperation {
        constructor(files) {
            super();
            this.originFile = files;
        }
        /**
        * @param path Search xml tag group , div,title etc.
        * @param fileName One of uploadfileList, uploadfileList is file group, {key:value}
        * @return Xml element calss
        */
        getElementsByTagName(path, fileName, isFile = true) {
            let file = this.getFileByName(fileName);
            if (!isFile)
                file = fileName;
            let pathArr = path.split("/"), ret;
            for (let key in pathArr) {
                let path = pathArr[key];
                if (ret == undefined) {
                    ret = this.getElementsByOneTag(path, file);
                }
                else {
                    if (ret instanceof Array) {
                        let items = [];
                        for (let key in ret) {
                            let item = ret[key];
                            items = items.concat(this.getElementsByOneTag(path, item));
                        }
                        ret = items;
                    }
                    else {
                        ret = this.getElementsByOneTag(path, ret);
                    }
                }
            }
            let elements = [];
            for (let i = 0; i < ret.length; i++) {
                let ele = new Element(ret[i]);
                elements.push(ele);
            }
            return elements;
        }
        getElementsByTagNameLink(path, fileName, isFile = true) {
            let file = this.getFileByName(fileName);
            if (!isFile)
                file = fileName;
            const ret = this.getElementByTagLink(path, file);
            let elements = [];
            for (let i = 0; i < ret.length; i++) {
                let ele = new Element(ret[i]);
                elements.push(ele);
            }
            return elements;
        }
        /**
        * @param name One of uploadfileList's name, search for file by this parameter
        * @retrun Select a file from uploadfileList
        */
        getFileByName(name) {
            for (let fileKey in this.originFile) {
                if (fileKey.indexOf(name) > -1) {
                    return this.originFile[fileKey];
                }
            }
            return "";
        }
    }
    class Element extends xmloperation {
        constructor(str) {
            super();
            this.elementString = str;
            this.setValue();
            const readAttrReg = new RegExp('[a-zA-Z0-9_:]*?=".*?"', "g");
            let attrList = this.container.match(readAttrReg);
            this.attributeList = {};
            if (attrList != null) {
                for (let key in attrList) {
                    let attrFull = attrList[key];
                    // let al= attrFull.split("=");
                    if (attrFull.length == 0) {
                        continue;
                    }
                    let attrKey = attrFull.substr(0, attrFull.indexOf('='));
                    let attrValue = attrFull.substr(attrFull.indexOf('=') + 1);
                    if (attrKey == null || attrValue == null || attrKey.length == 0 || attrValue.length == 0) {
                        continue;
                    }
                    this.attributeList[attrKey] = attrValue.substr(1, attrValue.length - 2);
                }
            }
        }
        /**
        * @param name Get attribute by key in element
        * @return Single attribute
        */
        get(name) {
            return this.attributeList[name];
        }
        /**
        * @param tag Get elements by tag in elementString
        * @return Element group
        */
        getInnerElements(tag) {
            let ret = this.getElementsByOneTag(tag, this.elementString);
            let elements = [];
            for (let i = 0; i < ret.length; i++) {
                let ele = new Element(ret[i]);
                elements.push(ele);
            }
            if (elements.length == 0) {
                return null;
            }
            return elements;
        }
        getInnerElementsTagLink(tag) {
            const ret = this.getElementByTagLink(tag, this.elementString);
            let elements = [];
            for (let i = 0; i < ret.length; i++) {
                let ele = new Element(ret[i]);
                elements.push(ele);
            }
            if (elements.length == 0) {
                return null;
            }
            return elements;
        }
        /**
        * @desc get xml dom value and container, <container>value</container>
        */
        setValue() {
            let str = this.elementString;
            if (str.substr(str.length - 2, 2) == "/>") {
                this.value = "";
                this.container = str;
            }
            else {
                let firstTag = this.getFirstTag();
                const firstTagReg = new RegExp("(<" + firstTag + " [^>]+?[^/]>)([\\s\\S]*?)</" + firstTag + ">|(<" + firstTag + ">)([\\s\\S]*?)</" + firstTag + ">", "g");
                let result = firstTagReg.exec(str);
                if (result != null) {
                    if (result[1] != null) {
                        this.container = result[1];
                        this.value = result[2];
                    }
                    else {
                        this.container = result[3];
                        this.value = result[4];
                    }
                }
            }
        }
        /**
        * @desc get xml dom first tag, <a><b></b></a>, get a
        */
        getFirstTag() {
            let str = this.elementString;
            let firstTag = str.substr(0, str.indexOf(' '));
            if (firstTag == "" || firstTag.indexOf(">") > -1) {
                firstTag = str.substr(0, str.indexOf('>'));
            }
            firstTag = firstTag.substr(1, firstTag.length);
            return firstTag;
        }
    }
    function combineIndexedColor(indexedColorsInner, indexedColors) {
        let ret = {};
        if (indexedColorsInner == null || indexedColorsInner.length == 0) {
            return indexedColors;
        }
        for (let key in indexedColors) {
            let value = indexedColors[key], kn = parseInt(key);
            let inner = indexedColorsInner[kn];
            if (inner == null) {
                ret[key] = value;
            }
            else {
                let rgb = inner.attributeList.rgb;
                ret[key] = rgb;
            }
        }
        return ret;
    }
    //clrScheme:Element[]
    function getColor(color, styles, type = "g") {
        let attrList = color.attributeList;
        let clrScheme = (styles["clrScheme"] ?? []);
        let indexedColorsInner = styles["indexedColors"];
        styles["mruColors"];
        let indexedColorsList = combineIndexedColor(indexedColorsInner, indexedColors);
        let indexed = attrList.indexed, rgb = attrList.rgb, theme = attrList.theme, tint = attrList.tint;
        let bg;
        if (indexed != null) {
            let indexedNum = parseInt(indexed);
            bg = indexedColorsList[indexedNum];
            if (bg != null) {
                bg = bg.substring(bg.length - 6, bg.length);
                bg = "#" + bg;
            }
        }
        else if (rgb != null) {
            rgb = rgb.substring(rgb.length - 6, rgb.length);
            bg = "#" + rgb;
        }
        else if (theme != null) {
            let themeNum = parseInt(theme);
            if (themeNum == 0) {
                themeNum = 1;
            }
            else if (themeNum == 1) {
                themeNum = 0;
            }
            else if (themeNum == 2) {
                themeNum = 3;
            }
            else if (themeNum == 3) {
                themeNum = 2;
            }
            let clrSchemeElement = clrScheme[themeNum];
            if (clrSchemeElement != null) {
                let clrs = clrSchemeElement.getInnerElements("a:sysClr|a:srgbClr");
                if (clrs != null) {
                    let clr = clrs[0];
                    let clrAttrList = clr.attributeList;
                    // console.log(clr.container, );
                    if (clr.container.indexOf("sysClr") > -1) {
                        // if(type=="g" && clrAttrList.val=="windowText"){
                        //     bg = null;
                        // }
                        // else if((type=="t" || type=="b") && clrAttrList.val=="window"){
                        //     bg = null;
                        // }                    
                        // else 
                        if (clrAttrList.lastClr != null) {
                            bg = "#" + clrAttrList.lastClr;
                        }
                        else if (clrAttrList.val != null) {
                            bg = "#" + clrAttrList.val;
                        }
                    }
                    else if (clr.container.indexOf("srgbClr") > -1) {
                        // console.log(clrAttrList.val);
                        bg = "#" + clrAttrList.val;
                    }
                }
            }
        }
        if (tint != null) {
            let tintNum = parseFloat(tint);
            if (bg != null) {
                bg = LightenDarkenColor(bg, tintNum);
            }
        }
        return bg;
    }
    /**
     * @dom xml attribute object
     * @attr attribute name
     * @d if attribute is null, return default value
     * @return attribute value
    */
    function getlineStringAttr(frpr, attr) {
        let attrEle = frpr.getInnerElements(attr), value;
        if (attrEle != null && attrEle.length > 0) {
            if (attr == "b" || attr == "i" || attr == "strike") {
                value = "1";
            }
            else if (attr == "u") {
                let v = attrEle[0].attributeList.val;
                if (v == "double") {
                    value = "2";
                }
                else if (v == "singleAccounting") {
                    value = "3";
                }
                else if (v == "doubleAccounting") {
                    value = "4";
                }
                else {
                    value = "1";
                }
            }
            else if (attr == "vertAlign") {
                let v = attrEle[0].attributeList.val;
                if (v == "subscript") {
                    value = "1";
                }
                else if (v == "superscript") {
                    value = "2";
                }
            }
            else {
                value = attrEle[0].attributeList.val;
            }
        }
        return value;
    }

    class LuckyFileBase {
    }
    class WorkBookInfo {
    }
    class LuckySheetBase {
    }
    class LuckyFileInfo {
    }
    class LuckySheetCelldataBase {
    }
    class LuckySheetCelldataValue {
    }
    class LuckySheetCellFormat {
    }
    class LuckyInlineString {
    }
    class LuckyConfig {
    }
    class LuckySheetborderInfoCellForImp {
    }
    class LuckySheetborderInfoCellValue {
    }
    class LuckySheetborderInfoCellValueStyle {
    }
    class LuckySheetConfigMerge {
    }
    class LuckysheetCalcChain {
    }
    class LuckyImageBase {
    }
    class LuckyChartImageBase {
    }
    class LuckyChart {
    }

    var AbsoluteRefType;
    (function (AbsoluteRefType) {
        AbsoluteRefType[AbsoluteRefType["NONE"] = 0] = "NONE";
        AbsoluteRefType[AbsoluteRefType["ROW"] = 1] = "ROW";
        AbsoluteRefType[AbsoluteRefType["COLUMN"] = 2] = "COLUMN";
        AbsoluteRefType[AbsoluteRefType["ALL"] = 3] = "ALL";
    })(AbsoluteRefType || (AbsoluteRefType = {}));
    function getBorderInfo(borders, styles) {
        if (borders == null) {
            return null;
        }
        let border = borders[0], attrList = border.attributeList;
        let style = attrList.style;
        if (style == null || style == "none") {
            return null;
        }
        let colors = border.getInnerElements("color");
        let colorRet = "#000000";
        if (colors != null) {
            let color = colors[0];
            colorRet = getColor(color, styles, "b");
            if (colorRet == null) {
                colorRet = "#000000";
            }
        }
        let ret = new LuckySheetborderInfoCellValueStyle();
        ret.style = borderTypes[style];
        ret.color = colorRet;
        return ret;
    }
    function handleBorder(border, styles) {
        const borderCellValue = new LuckySheetborderInfoCellValue();
        let isAdd = false;
        if (!border) {
            return {
                borderCellValue,
                isAdd
            };
        }
        let lefts = border.getInnerElements("left");
        let rights = border.getInnerElements("right");
        let tops = border.getInnerElements("top");
        let bottoms = border.getInnerElements("bottom");
        let diagonals = border.getInnerElements("diagonal");
        let starts = border.getInnerElements("start");
        let ends = border.getInnerElements("end");
        let left = getBorderInfo(lefts, styles);
        let right = getBorderInfo(rights, styles);
        let top = getBorderInfo(tops, styles);
        let bottom = getBorderInfo(bottoms, styles);
        let diagonal = getBorderInfo(diagonals, styles);
        let start = getBorderInfo(starts, styles);
        let end = getBorderInfo(ends, styles);
        if (start != null && start.color != null) {
            borderCellValue.l = start;
            isAdd = true;
        }
        if (end != null && end.color != null) {
            borderCellValue.r = end;
            isAdd = true;
        }
        if (left != null && left.color != null) {
            borderCellValue.l = left;
            isAdd = true;
        }
        if (right != null && right.color != null) {
            borderCellValue.r = right;
            isAdd = true;
        }
        if (top != null && top.color != null) {
            borderCellValue.t = top;
            isAdd = true;
        }
        if (bottom != null && bottom.color != null) {
            borderCellValue.b = bottom;
            isAdd = true;
        }
        if (diagonal != null && diagonal.color != null) {
            const diagonalUp = border.attributeList.diagonalUp;
            const diagonalDown = border.attributeList.diagonalDown;
            if (diagonalUp === "1") {
                borderCellValue.bl_tr = diagonal;
            }
            if (diagonalDown === "1") {
                borderCellValue.tl_br = diagonal;
            }
            isAdd = true;
        }
        return {
            borderCellValue,
            isAdd
        };
    }
    function getBackgroundByFill(fill, styles) {
        let patternFills = fill.getInnerElements("patternFill");
        if (patternFills != null) {
            let patternFill = patternFills[0];
            let fgColors = patternFill.getInnerElements("fgColor");
            let bgColors = patternFill.getInnerElements("bgColor");
            let fg, bg;
            if (fgColors != null) {
                let fgColor = fgColors[0];
                fg = getColor(fgColor, styles);
            }
            if (bgColors != null) {
                let bgColor = bgColors[0];
                bg = getColor(bgColor, styles);
            }
            // console.log(fgColors,bgColors,clrScheme);
            if (fg != null) {
                return fg;
            }
            else if (bg != null) {
                return bg;
            }
        }
        else {
            let gradientfills = fill.getInnerElements("gradientFill");
            if (gradientfills != null) {
                //graient color fill handler
                return null;
            }
        }
    }
    function getFontStyle(font, styles) {
        let familyFont = null;
        const cellValue = new LuckySheetCelldataValue();
        let sz = font.getInnerElements("sz"); //font size
        let colors = font.getInnerElements("color"); //font color
        let family = font.getInnerElements("name"); //font family
        let familyOverrides = font.getInnerElements("family"); //font family will be overrided by name
        font.getInnerElements("charset"); //font charset
        let bolds = font.getInnerElements("b"); //font bold
        let italics = font.getInnerElements("i"); //font italic
        let strikes = font.getInnerElements("strike"); //font italic
        let underlines = font.getInnerElements("u"); //font italic
        if (sz != null && sz.length > 0) {
            let fs = sz[0].attributeList.val;
            if (fs != null) {
                cellValue.fs = parseInt(fs);
            }
        }
        if (colors != null && colors.length > 0) {
            let color = colors[0];
            let fc = getColor(color, styles, "t");
            if (fc != null) {
                cellValue.fc = fc;
            }
        }
        if (familyOverrides != null && familyOverrides.length > 0) {
            let val = familyOverrides[0].attributeList.val;
            if (val != null) {
                familyFont = fontFamilys[val];
            }
        }
        if (family != null && family.length > 0) {
            let val = family[0].attributeList.val;
            if (val != null) {
                cellValue.ff = val;
            }
        }
        if (bolds != null && bolds.length > 0) {
            let bold = bolds[0].attributeList.val;
            if (bold == "0") {
                cellValue.bl = 0;
            }
            else {
                cellValue.bl = 1;
            }
        }
        if (italics != null && italics.length > 0) {
            let italic = italics[0].attributeList.val;
            if (italic == "0") {
                cellValue.it = 0;
            }
            else {
                cellValue.it = 1;
            }
        }
        if (strikes != null && strikes.length > 0) {
            let strike = strikes[0].attributeList.val;
            if (strike == "0") {
                cellValue.cl = 0;
            }
            else {
                cellValue.cl = 1;
            }
        }
        if (underlines != null && underlines.length > 0) {
            let underline = underlines[0].attributeList.val;
            if (underline == "single") {
                cellValue.un = 1;
            }
            else if (underline == "double") {
                cellValue.un = 2;
            }
            else if (underline == "singleAccounting") {
                cellValue.un = 3;
            }
            else if (underline == "doubleAccounting") {
                cellValue.un = 4;
            }
            else {
                cellValue.un = 0;
            }
        }
        return {
            cellValue,
            familyFont
        };
    }
    const handleRanges = (sqref) => {
        const list = sqref.split(' ');
        return list.map(d => {
            const rangetxt = d.split(':');
            const startRow = parseInt(rangetxt[0].replace(/[^0-9]/g, "")) - 1;
            const startColumn = ABCToNumber(rangetxt[0].replace(/[^A-Za-z]/g, ""));
            const startFreezon = isfreezonFuc(rangetxt[0]);
            const endRow = rangetxt.length === 1 ? startRow : (parseInt(rangetxt[1].replace(/[^0-9]/g, "")) - 1);
            const endColumn = rangetxt.length === 1 ? startColumn : ABCToNumber(rangetxt[1].replace(/[^A-Za-z]/g, ""));
            const endFreezon = rangetxt.length === 1 ? startFreezon : isfreezonFuc(rangetxt[1]);
            const handleType = (freezon) => {
                if (freezon[0] === true && freezon[1] === true) {
                    return AbsoluteRefType.ALL;
                }
                if (freezon[0] === true) {
                    return AbsoluteRefType.ROW;
                }
                if (freezon[1] === true) {
                    return AbsoluteRefType.COLUMN;
                }
                return AbsoluteRefType.NONE;
            };
            return {
                startRow,
                startColumn,
                endRow,
                endColumn,
                startAbsoluteRefType: handleType(startFreezon),
                endAbsoluteRefType: handleType(endFreezon),
                rangeType: 0
            };
        });
    };

    class LuckySheetCelldata extends LuckySheetCelldataBase {
        constructor(cell, cellSize, styles, sharedStrings, mergeCells, sheetFile, cellImages, imageList, ReadXml) {
            //Private
            super();
            this.cell = cell;
            this.sheetFile = sheetFile;
            this.styles = styles;
            this.sharedStrings = sharedStrings;
            this.readXml = ReadXml;
            this.mergeCells = mergeCells;
            this.cellImages = cellImages;
            this.imageList = imageList;
            this.cellSize = cellSize;
            let attrList = cell.attributeList;
            let r = attrList.r, s = attrList.s, t = attrList.t;
            let range = getcellrange(r);
            this.r = range.row[0];
            this.c = range.column[0];
            this.v = this.generateValue(s, t);
        }
        /**
        * @param s Style index ,start 1
        * @param t Cell type, Optional value is ST_CellType, it's found at constat.ts
        */
        generateValue(s, t) {
            let v = this.cell.getInnerElements("v");
            let f = this.cell.getInnerElements("f");
            if (v == null) {
                v = this.cell.getInnerElements("t");
            }
            let cellXfs = this.styles["cellXfs"];
            let cellStyleXfs = this.styles["cellStyleXfs"];
            this.styles["cellStyles"];
            let fonts = this.styles["fonts"];
            let fills = this.styles["fills"];
            let borders = this.styles["borders"];
            let numfmts = this.styles["numfmts"];
            this.styles["clrScheme"];
            let sharedStrings = this.sharedStrings;
            let cellValue = new LuckySheetCelldataValue();
            if (f != null) {
                let formula = f[0], attrList = formula.attributeList;
                let t = attrList.t, ref = attrList.ref, si = attrList.si;
                let formulaValue = f[0].value;
                if (t == "shared") {
                    this._fomulaRef = ref;
                    this._formulaType = t;
                    this._formulaSi = si;
                }
                // console.log(ref, t, si);
                if (ref != null || (formulaValue != null && formulaValue.length > 0)) {
                    formulaValue = escapeCharacter(formulaValue);
                    cellValue.f = formulaValue[0] === '=' ? formulaValue : "=" + formulaValue;
                }
            }
            let familyFont = null;
            let quotePrefix;
            if (s != null) {
                let sNum = parseInt(s);
                let cellXf = cellXfs[sNum];
                let xfId = cellXf.attributeList.xfId;
                let numFmtId, fontId, fillId, borderId;
                let horizontal, vertical, wrapText, textRotation, indent;
                if (xfId != null) {
                    let cellStyleXf = cellStyleXfs[parseInt(xfId)];
                    let attrList = cellStyleXf.attributeList;
                    let applyNumberFormat = attrList.applyNumberFormat;
                    let applyFont = attrList.applyFont;
                    let applyFill = attrList.applyFill;
                    let applyBorder = attrList.applyBorder;
                    let applyAlignment = attrList.applyAlignment;
                    // let applyProtection = attrList.applyProtection;
                    attrList.applyProtection;
                    quotePrefix = attrList.quotePrefix;
                    if (applyNumberFormat != "0" && attrList.numFmtId != null) {
                        // if(attrList.numFmtId!="0"){
                        numFmtId = attrList.numFmtId;
                        // }
                    }
                    if (applyFont != "0" && attrList.fontId != null) {
                        fontId = attrList.fontId;
                    }
                    if (applyFill != "0" && attrList.fillId != null) {
                        fillId = attrList.fillId;
                    }
                    if (applyBorder != "0" && attrList.borderId != null) {
                        borderId = attrList.borderId;
                    }
                    if (applyAlignment != null && applyAlignment != "0") {
                        let alignment = cellStyleXf.getInnerElements("alignment");
                        if (alignment != null) {
                            let attrList = alignment[0].attributeList;
                            if (attrList.horizontal != null) {
                                horizontal = attrList.horizontal;
                            }
                            if (attrList.vertical != null) {
                                vertical = attrList.vertical;
                            }
                            if (attrList.wrapText != null) {
                                wrapText = attrList.wrapText;
                            }
                            if (attrList.textRotation != null) {
                                textRotation = attrList.textRotation;
                            }
                            if (attrList.shrinkToFit != null) {
                                attrList.shrinkToFit;
                            }
                            if (attrList.indent != null) {
                                indent = attrList.indent;
                            }
                        }
                    }
                }
                let applyNumberFormat = cellXf.attributeList.applyNumberFormat;
                let applyFont = cellXf.attributeList.applyFont;
                let applyFill = cellXf.attributeList.applyFill;
                let applyBorder = cellXf.attributeList.applyBorder;
                let applyAlignment = cellXf.attributeList.applyAlignment;
                if (cellXf.attributeList.applyProtection != null) {
                    cellXf.attributeList.applyProtection;
                }
                if (cellXf.attributeList.quotePrefix != null) {
                    quotePrefix = cellXf.attributeList.quotePrefix;
                }
                if (applyNumberFormat != "0" && cellXf.attributeList.numFmtId != null) {
                    numFmtId = cellXf.attributeList.numFmtId;
                }
                if (applyFont != "0") {
                    fontId = cellXf.attributeList.fontId;
                }
                if (applyFill != "0") {
                    fillId = cellXf.attributeList.fillId;
                }
                if (applyBorder != "0") {
                    borderId = cellXf.attributeList.borderId;
                }
                if (applyAlignment != "0") {
                    let alignment = cellXf.getInnerElements("alignment");
                    if (alignment != null && alignment.length > 0) {
                        let attrList = alignment[0].attributeList;
                        if (attrList.horizontal != null) {
                            horizontal = attrList.horizontal;
                        }
                        if (attrList.vertical != null) {
                            vertical = attrList.vertical;
                        }
                        if (attrList.wrapText != null) {
                            wrapText = attrList.wrapText;
                        }
                        if (attrList.textRotation != null) {
                            textRotation = attrList.textRotation;
                        }
                        if (attrList.shrinkToFit != null) {
                            attrList.shrinkToFit;
                        }
                        if (attrList.indent != null) {
                            indent = attrList.indent;
                        }
                    }
                }
                if (numFmtId != undefined) {
                    let numf = numfmts[parseInt(numFmtId)];
                    let cellFormat = new LuckySheetCellFormat();
                    cellFormat.fa = escapeCharacter(numf);
                    // console.log(numf, numFmtId, this.v, cellFormat);
                    cellFormat.t = t || 'd';
                    cellValue.ct = cellFormat;
                }
                if (fillId != undefined) {
                    let fillIdNum = parseInt(fillId);
                    let fill = fills[fillIdNum];
                    // console.log(cellValue.v);
                    let bg = getBackgroundByFill(fill, this.styles);
                    if (bg != null) {
                        cellValue.bg = bg;
                    }
                }
                if (fontId != undefined) {
                    let fontIdNum = parseInt(fontId);
                    let font = fonts[fontIdNum];
                    if (font != null) {
                        const { cellValue: fontStyle, familyFont: family } = getFontStyle(font, this.styles);
                        cellValue = {
                            ...cellValue,
                            ...fontStyle
                        };
                        familyFont = family;
                    }
                }
                // vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
                // ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
                // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
                // tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment
                if (horizontal != undefined) { //Horizontal alignment
                    if (horizontal == "center") {
                        cellValue.ht = 0;
                    }
                    else if (horizontal == "centerContinuous") {
                        cellValue.ht = 0; //luckysheet unsupport
                    }
                    else if (horizontal == "left") {
                        cellValue.ht = 1;
                    }
                    else if (horizontal == "right") {
                        cellValue.ht = 2;
                    }
                    else if (horizontal == "distributed") {
                        cellValue.ht = 0; //luckysheet unsupport
                    }
                    else if (horizontal == "fill") {
                        cellValue.ht = 1; //luckysheet unsupport
                    }
                    else if (horizontal == "general") {
                        cellValue.ht = 1; //luckysheet unsupport
                    }
                    else if (horizontal == "justify") {
                        cellValue.ht = 0; //luckysheet unsupport
                    }
                    else {
                        cellValue.ht = 1;
                    }
                }
                if (vertical != undefined) { //Vertical alignment
                    if (vertical == "bottom") {
                        cellValue.vt = 2;
                    }
                    else if (vertical == "center") {
                        cellValue.vt = 0;
                    }
                    else if (vertical == "distributed") {
                        cellValue.vt = 0; //luckysheet unsupport
                    }
                    else if (vertical == "justify") {
                        cellValue.vt = 0; //luckysheet unsupport
                    }
                    else if (vertical == "top") {
                        cellValue.vt = 1;
                    }
                    else {
                        cellValue.vt = 1;
                    }
                }
                else {
                    //sometimes bottom style is lost after setting it in excel
                    //when vertical is undefined set it to 2.
                    cellValue.vt = 2;
                }
                if (wrapText != undefined) {
                    if (wrapText == "1") {
                        cellValue.tb = 2;
                    }
                    else {
                        cellValue.tb = 1;
                    }
                }
                else {
                    cellValue.tb = 1;
                }
                if (textRotation != undefined) {
                    // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
                    if (textRotation == "255") {
                        cellValue.tr = 3;
                    }
                    // else if(textRotation=="45"){
                    //     cellValue.tr = 1;
                    // }
                    // else if(textRotation=="90"){
                    //     cellValue.tr = 4;
                    // }
                    // else if(textRotation=="135"){
                    //     cellValue.tr = 2;
                    // }
                    // else if(textRotation=="180"){
                    //     cellValue.tr = 5;
                    // }
                    else {
                        cellValue.tr = 0;
                        cellValue.rt = parseInt(textRotation);
                    }
                }
                if (indent != undefined) { //luckysheet unsupport
                    const result = parseInt(indent);
                    if (!isNaN(result)) {
                        cellValue.ti = result;
                    }
                }
                if (borderId != undefined) {
                    let borderIdNum = parseInt(borderId);
                    let border = borders[borderIdNum];
                    // this._borderId = borderIdNum;
                    let borderObject = new LuckySheetborderInfoCellForImp();
                    borderObject.rangeType = "cell";
                    // borderObject.cells = [];
                    const { isAdd, borderCellValue } = handleBorder(border, this.styles);
                    borderCellValue.row_index = this.r;
                    borderCellValue.col_index = this.c;
                    if (isAdd) {
                        borderObject.value = borderCellValue;
                        // this.config._borderInfo[borderId] = borderObject;
                        this._borderObject = borderObject;
                    }
                }
            }
            else {
                cellValue.tb = 1;
            }
            if (v != null) {
                let value = v[0].value;
                if (/&#\d+;/.test(value)) {
                    value = this.htmlDecode(value);
                }
                if (t == ST_CellType["SharedString"]) {
                    let siIndex = parseInt(v[0].value);
                    let sharedSI = sharedStrings[siIndex];
                    let rFlag = sharedSI.getInnerElements("r");
                    if (rFlag == null) {
                        let tFlag = sharedSI.getInnerElements("t");
                        if (tFlag != null) {
                            let text = "";
                            tFlag.forEach((t) => {
                                text += t.value;
                            });
                            text = escapeCharacter(text);
                            //isContainMultiType(text) &&
                            if (familyFont == "Roman" && text.length > 0) {
                                let textArray = text.split("");
                                let preWordType = null, wordText = "", preWholef = null;
                                let wholef = "Times New Roman";
                                if (cellValue.ff != null) {
                                    wholef = cellValue.ff;
                                }
                                let cellFormat = cellValue.ct;
                                if (cellFormat == null) {
                                    cellFormat = new LuckySheetCellFormat();
                                }
                                if (cellFormat.s == null) {
                                    cellFormat.s = [];
                                }
                                for (let i = 0; i < textArray.length; i++) {
                                    let w = textArray[i];
                                    let type = null, ff = wholef;
                                    if (isChinese(w)) {
                                        type = "c";
                                        ff = "宋体";
                                    }
                                    else if (isJapanese(w)) {
                                        type = "j";
                                        ff = "Yu Gothic";
                                    }
                                    else if (isKoera(w)) {
                                        type = "k";
                                        ff = "Malgun Gothic";
                                    }
                                    else {
                                        type = "e";
                                    }
                                    if ((type != preWordType && preWordType != null) || i == textArray.length - 1) {
                                        let InlineString = new LuckyInlineString();
                                        InlineString.ff = preWholef;
                                        if (cellValue.fc != null) {
                                            InlineString.fc = cellValue.fc;
                                        }
                                        if (cellValue.fs != null) {
                                            InlineString.fs = cellValue.fs;
                                        }
                                        if (cellValue.cl != null) {
                                            InlineString.cl = cellValue.cl;
                                        }
                                        if (cellValue.un != null) {
                                            InlineString.un = cellValue.un;
                                        }
                                        if (cellValue.bl != null) {
                                            InlineString.bl = cellValue.bl;
                                        }
                                        if (cellValue.it != null) {
                                            InlineString.it = cellValue.it;
                                        }
                                        if (i == textArray.length - 1) {
                                            if (type == preWordType) {
                                                InlineString.ff = ff;
                                                InlineString.v = wordText + w;
                                            }
                                            else {
                                                InlineString.ff = preWholef;
                                                InlineString.v = wordText;
                                                cellFormat.s.push(InlineString);
                                                let InlineStringLast = new LuckyInlineString();
                                                InlineStringLast.ff = ff;
                                                InlineStringLast.v = w;
                                                if (cellValue.fc != null) {
                                                    InlineStringLast.fc = cellValue.fc;
                                                }
                                                if (cellValue.fs != null) {
                                                    InlineStringLast.fs = cellValue.fs;
                                                }
                                                if (cellValue.cl != null) {
                                                    InlineStringLast.cl = cellValue.cl;
                                                }
                                                if (cellValue.un != null) {
                                                    InlineStringLast.un = cellValue.un;
                                                }
                                                if (cellValue.bl != null) {
                                                    InlineStringLast.bl = cellValue.bl;
                                                }
                                                if (cellValue.it != null) {
                                                    InlineStringLast.it = cellValue.it;
                                                }
                                                cellFormat.s.push(InlineStringLast);
                                                break;
                                            }
                                        }
                                        else {
                                            InlineString.v = wordText;
                                        }
                                        cellFormat.s.push(InlineString);
                                        wordText = w;
                                    }
                                    else {
                                        wordText += w;
                                    }
                                    preWordType = type;
                                    preWholef = ff;
                                }
                                cellFormat.t = "inlineStr";
                                // cellFormat.s = [InlineString];
                                cellValue.ct = cellFormat;
                                // console.log(cellValue);
                            }
                            else {
                                text = this.replaceSpecialWrap(text);
                                if (text.indexOf("\r\n") > -1 || text.indexOf("\n") > -1) {
                                    let InlineString = new LuckyInlineString();
                                    InlineString.v = text;
                                    let cellFormat = cellValue.ct;
                                    if (cellFormat == null) {
                                        cellFormat = new LuckySheetCellFormat();
                                    }
                                    if (cellValue.ff != null) {
                                        InlineString.ff = cellValue.ff;
                                    }
                                    if (cellValue.fc != null) {
                                        InlineString.fc = cellValue.fc;
                                    }
                                    if (cellValue.fs != null) {
                                        InlineString.fs = cellValue.fs;
                                    }
                                    if (cellValue.cl != null) {
                                        InlineString.cl = cellValue.cl;
                                    }
                                    if (cellValue.un != null) {
                                        InlineString.un = cellValue.un;
                                    }
                                    if (cellValue.bl != null) {
                                        InlineString.bl = cellValue.bl;
                                    }
                                    if (cellValue.it != null) {
                                        InlineString.it = cellValue.it;
                                    }
                                    cellFormat.t = "inlineStr";
                                    cellFormat.s = [InlineString];
                                    cellValue.ct = cellFormat;
                                }
                                else {
                                    cellValue.v = text;
                                    quotePrefix = "1";
                                }
                            }
                        }
                    }
                    else {
                        let styles = [];
                        rFlag.forEach((r) => {
                            let tFlag = r.getInnerElements("t");
                            let rPr = r.getInnerElements("rPr");
                            let InlineString = new LuckyInlineString();
                            if (tFlag != null && tFlag.length > 0) {
                                let text = tFlag[0].value;
                                text = this.replaceSpecialWrap(text);
                                text = escapeCharacter(text);
                                InlineString.v = text;
                            }
                            if (rPr != null && rPr.length > 0) {
                                let frpr = rPr[0];
                                let sz = getlineStringAttr(frpr, "sz"), rFont = getlineStringAttr(frpr, "rFont"); getlineStringAttr(frpr, "family"); getlineStringAttr(frpr, "charset"); getlineStringAttr(frpr, "scheme"); let b = getlineStringAttr(frpr, "b"), i = getlineStringAttr(frpr, "i"), u = getlineStringAttr(frpr, "u"), strike = getlineStringAttr(frpr, "strike"), vertAlign = getlineStringAttr(frpr, "vertAlign"), color;
                                let cEle = frpr.getInnerElements("color");
                                if (cEle != null && cEle.length > 0) {
                                    color = getColor(cEle[0], this.styles, "t");
                                }
                                let ff;
                                // if(family!=null){
                                //     ff = fontFamilys[family];
                                // }
                                if (rFont != null) {
                                    ff = rFont;
                                }
                                if (ff != null) {
                                    InlineString.ff = ff;
                                }
                                else if (cellValue.ff != null) {
                                    InlineString.ff = cellValue.ff;
                                }
                                if (color != null) {
                                    InlineString.fc = color;
                                }
                                // else if(cellValue.fc!=null){
                                //     InlineString.fc = cellValue.fc;
                                // }
                                if (sz != null) {
                                    InlineString.fs = parseInt(sz);
                                }
                                else if (cellValue.fs != null) {
                                    InlineString.fs = cellValue.fs;
                                }
                                if (strike != null) {
                                    InlineString.cl = parseInt(strike);
                                }
                                else if (cellValue.cl != null) {
                                    InlineString.cl = cellValue.cl;
                                }
                                if (u != null) {
                                    InlineString.un = parseInt(u);
                                }
                                else if (cellValue.un != null) {
                                    InlineString.un = cellValue.un;
                                }
                                if (b != null) {
                                    InlineString.bl = parseInt(b);
                                }
                                else if (cellValue.bl != null) {
                                    InlineString.bl = cellValue.bl;
                                }
                                if (i != null) {
                                    InlineString.it = parseInt(i);
                                }
                                else if (cellValue.it != null) {
                                    InlineString.it = cellValue.it;
                                }
                                if (vertAlign != null) {
                                    InlineString.va = parseInt(vertAlign);
                                }
                                // ff:string | undefined //font family
                                // fc:string | undefined//font color
                                // fs:number | undefined//font size
                                // cl:number | undefined//strike
                                // un:number | undefined//underline
                                // bl:number | undefined//blod
                                // it:number | undefined//italic
                                // v:string | undefined
                            }
                            else {
                                if (InlineString.ff == null && cellValue.ff != null) {
                                    InlineString.ff = cellValue.ff;
                                }
                                if (InlineString.fc == null && cellValue.fc != null) {
                                    InlineString.fc = cellValue.fc;
                                }
                                if (InlineString.fs == null && cellValue.fs != null) {
                                    InlineString.fs = cellValue.fs;
                                }
                                if (InlineString.cl == null && cellValue.cl != null) {
                                    InlineString.cl = cellValue.cl;
                                }
                                if (InlineString.un == null && cellValue.un != null) {
                                    InlineString.un = cellValue.un;
                                }
                                if (InlineString.bl == null && cellValue.bl != null) {
                                    InlineString.bl = cellValue.bl;
                                }
                                if (InlineString.it == null && cellValue.it != null) {
                                    InlineString.it = cellValue.it;
                                }
                            }
                            styles.push(InlineString);
                        });
                        let cellFormat = cellValue.ct;
                        if (cellFormat == null) {
                            cellFormat = new LuckySheetCellFormat();
                        }
                        cellFormat.t = "inlineStr";
                        cellFormat.s = styles;
                        cellValue.ct = cellFormat;
                    }
                }
                // to be confirmed
                else if (t == ST_CellType["InlineString"] && v != null) {
                    cellValue.v = "'" + value;
                }
                else if (t == ST_CellType["String"] && value.includes('=DISPIMG')) {
                    let cellFormat = cellValue.ct;
                    if (cellFormat == null) {
                        cellFormat = new LuckySheetCellFormat();
                    }
                    cellFormat.t = "str";
                    cellFormat.ci = this.getCellImage(cellValue, value);
                    cellValue.ct = cellFormat;
                }
                else {
                    value = escapeCharacter(value);
                    cellValue.v = value;
                }
            }
            if (quotePrefix != null) {
                cellValue.qp = parseInt(quotePrefix);
            }
            if (t !== null && !cellValue.ct?.t) {
                let cellFormat = new LuckySheetCellFormat();
                cellFormat.t = t || 'd';
                cellValue.ct = cellFormat;
            }
            return cellValue;
        }
        replaceSpecialWrap(text) {
            text = text.replace(/_x000D_/g, "").replace(/&#13;&#10;/g, "\r\n").replace(/&#13;/g, "\r").replace(/&#10;/g, "\n");
            return text;
        }
        htmlDecode(str) {
            return str.replace(/&#(x)?([^&]{1,5});/g, function ($, $1, $2) {
                return String.fromCharCode(parseInt($2, $1 ? 16 : 10));
            });
        }
        ;
        getCellImage(cellValue, value) {
            const id = this.extractImageId(value);
            let ci = {};
            this.cellImages.forEach((element) => {
                const pic = element.getInnerElements('xdr:pic')[0];
                const picpr = pic.getInnerElements('xdr:nvPicPr')[0];
                const nvpr = picpr.getInnerElements('xdr:cNvPr')[0];
                const blipfill = pic.getInnerElements('xdr:blipFill')[0];
                const blip = blipfill.getInnerElements('a:blip')[0];
                const picId = nvpr.get('name');
                const picRid = blip.get('r:embed');
                if (id == picId) {
                    const obj = this.getBase64ByRid(picRid, cellImagesRels);
                    ci = obj;
                    let x_n = 0, y_n = 0;
                    let cx_n = 0, cy_n = 0;
                    const sp = pic.getInnerElements('xdr:spPr')[0];
                    const xfrm = sp.getInnerElements('a:xfrm')[0];
                    const off = xfrm.getInnerElements('a:off')[0];
                    const ext = xfrm.getInnerElements('a:ext')[0];
                    cx_n = getPxByEMUs(parseInt(ext.get('cx'))), cy_n = getPxByEMUs(parseInt(ext.get('cy')));
                    x_n = getPxByEMUs(parseInt(off.get('x'))), y_n = getPxByEMUs(parseInt(off.get('y')));
                    const rateX = cx_n / this.cellSize.width;
                    const rateY = cy_n / this.cellSize.height;
                    if (rateX > 1 && rateX > rateY) {
                        cy_n = cy_n / rateX;
                        cx_n = this.cellSize.width;
                    }
                    if (rateY > 1 && rateY > rateX) {
                        cx_n = cx_n / rateY;
                        cy_n = this.cellSize.height;
                    }
                    let imageDefault = {
                        height: cy_n,
                        left: x_n,
                        top: y_n,
                        width: cx_n
                    };
                    ci.default = imageDefault;
                    ci.descr = nvpr.get('descr');
                }
            });
            return ci;
        }
        extractImageId(formula) {
            const regex = /ID_[A-Za-z0-9]{32}/;
            const match = formula.match(regex);
            return match ? match[0] : null;
        }
        getBase64ByRid(rid, drawingRelsFile) {
            let Relationships = this.readXml.getElementsByTagName("Relationships/Relationship", drawingRelsFile);
            if (Relationships != null && Relationships.length > 0) {
                for (let i = 0; i < Relationships.length; i++) {
                    let Relationship = Relationships[i];
                    let attrList = Relationship.attributeList;
                    let Id = getXmlAttibute(attrList, "Id", null);
                    let src = getXmlAttibute(attrList, "Target", null);
                    if (Id == rid) {
                        src = src.replace(/\.\.\//g, "");
                        src = "xl/" + src;
                        let imgage = this.imageList.getImageByName(src);
                        return imgage;
                    }
                }
            }
            return {};
        }
    }

    /**:
    * @description: 返回的condition格式为univer使用的格式，但是规则配置中的style是luckyexcel的格式，
    * 后面考虑新增配置用来输出univer所需的数据结构
    * @author: Created by zwight on 2024-09-20 16:13:34
    */
    var CFTextOperator;
    (function (CFTextOperator) {
        CFTextOperator["beginsWith"] = "beginsWith";
        CFTextOperator["endsWith"] = "endsWith";
        CFTextOperator["containsText"] = "containsText";
        CFTextOperator["notContainsText"] = "notContainsText";
        CFTextOperator["equal"] = "equal";
        CFTextOperator["notEqual"] = "notEqual";
        CFTextOperator["containsBlanks"] = "containsBlanks";
        CFTextOperator["notContainsBlanks"] = "notContainsBlanks";
        CFTextOperator["containsErrors"] = "containsErrors";
        CFTextOperator["notContainsErrors"] = "notContainsErrors";
    })(CFTextOperator || (CFTextOperator = {}));
    var CFTimePeriodOperator;
    (function (CFTimePeriodOperator) {
        CFTimePeriodOperator["today"] = "today";
        CFTimePeriodOperator["yesterday"] = "yesterday";
        CFTimePeriodOperator["tomorrow"] = "tomorrow";
        CFTimePeriodOperator["last7Days"] = "last7Days";
        CFTimePeriodOperator["thisMonth"] = "thisMonth";
        CFTimePeriodOperator["lastMonth"] = "lastMonth";
        CFTimePeriodOperator["nextMonth"] = "nextMonth";
        CFTimePeriodOperator["thisWeek"] = "thisWeek";
        CFTimePeriodOperator["lastWeek"] = "lastWeek";
        CFTimePeriodOperator["nextWeek"] = "nextWeek";
    })(CFTimePeriodOperator || (CFTimePeriodOperator = {}));
    var CFNumberOperator;
    (function (CFNumberOperator) {
        CFNumberOperator["greaterThan"] = "greaterThan";
        CFNumberOperator["greaterThanOrEqual"] = "greaterThanOrEqual";
        CFNumberOperator["lessThan"] = "lessThan";
        CFNumberOperator["lessThanOrEqual"] = "lessThanOrEqual";
        CFNumberOperator["notBetween"] = "notBetween";
        CFNumberOperator["between"] = "between";
        CFNumberOperator["equal"] = "equal";
        CFNumberOperator["notEqual"] = "notEqual";
    })(CFNumberOperator || (CFNumberOperator = {}));
    var CFRuleType;
    (function (CFRuleType) {
        CFRuleType["highlightCell"] = "highlightCell";
        CFRuleType["dataBar"] = "dataBar";
        CFRuleType["colorScale"] = "colorScale";
        CFRuleType["iconSet"] = "iconSet";
    })(CFRuleType || (CFRuleType = {}));
    var CFSubRuleType;
    (function (CFSubRuleType) {
        CFSubRuleType["uniqueValues"] = "uniqueValues";
        CFSubRuleType["duplicateValues"] = "duplicateValues";
        CFSubRuleType["rank"] = "rank";
        CFSubRuleType["text"] = "text";
        CFSubRuleType["timePeriod"] = "timePeriod";
        CFSubRuleType["number"] = "number";
        CFSubRuleType["average"] = "average";
        CFSubRuleType["formula"] = "formula";
    })(CFSubRuleType || (CFSubRuleType = {}));
    class LuckyCondition {
        constructor(ele, readXml, styles) {
            this.stopIfTrue = false;
            this.handleRules = (ele, readXml, styles) => {
                const { attributeList, value, extLst, isExtLst } = ele;
                const type = getXmlAttibute(attributeList, 'type', 'expression');
                const rule = {
                    type: CFRuleType.highlightCell
                };
                const operator = getXmlAttibute(attributeList, 'operator', '');
                const rank = getXmlAttibute(attributeList, 'rank', '');
                const formula = readXml.getElementsByTagName("formula", value, false);
                getXmlAttibute(attributeList, 'aboveAverage', '');
                if (operator)
                    rule.operator = operator;
                if (formula[0]?.value || formula[0]?.value == '0')
                    rule.value = str2num(formula[0]?.value);
                switch (type) {
                    case 'expression':
                        rule.subType = CFSubRuleType.formula;
                        break;
                    case 'cellIs':
                        rule.subType = CFSubRuleType.number;
                        break;
                    case 'top10':
                        rule.subType = CFSubRuleType.rank;
                        const percent = getXmlAttibute(attributeList, 'percent', '0');
                        const bottom = getXmlAttibute(attributeList, 'bottom', '0');
                        if (rank)
                            rule.value = str2num(rank);
                        rule.isBottom = bottom === '1' ? true : false;
                        rule.isPercent = percent === '1' ? true : false;
                        break;
                    case 'aboveAverage':
                        rule.subType = CFSubRuleType.average;
                        rule.operator = rule.operator || CFNumberOperator.lessThan;
                        break;
                    case 'timePeriod':
                        rule.subType = CFSubRuleType.timePeriod;
                        rule.operator = getXmlAttibute(attributeList, 'timePeriod', undefined);
                        break;
                    case 'duplicateValues':
                        rule.subType = CFSubRuleType.duplicateValues;
                        break;
                    case 'containsText':
                        rule.subType = CFSubRuleType.text;
                        rule.operator = 'containsText';
                        rule.value = getXmlAttibute(attributeList, 'text', '');
                        break;
                    case 'colorScale':
                        const cfvo = readXml.getElementsByTagName("colorScale/cfvo", value, false);
                        const color = readXml.getElementsByTagName("colorScale/color", value, false);
                        rule.type = CFRuleType.colorScale;
                        rule.config = cfvo.map((d, index) => {
                            const type = getXmlAttibute(d.attributeList, 'type', '');
                            const value = getXmlAttibute(d.attributeList, 'val', undefined);
                            const colorVal = color[index] ? getColor(color[index], styles) : undefined;
                            return {
                                index: 0,
                                color: colorVal,
                                value: {
                                    type,
                                    value: str2num(value)
                                }
                            };
                        });
                        break;
                    case 'dataBar':
                        rule.type = CFRuleType.dataBar;
                        const dataBar = readXml.getElementsByTagName("dataBar", value, false)?.[0];
                        const barCfvo = readXml.getElementsByTagName("cfvo", dataBar.value, false);
                        const barColor = readXml.getElementsByTagName("color", dataBar.value, false);
                        const isShowValue = getXmlAttibute(dataBar.attributeList, 'showValue', '1');
                        rule.isShowValue = isShowValue === '1';
                        let positiveColor = barColor[0] ? getColor(barColor[0], styles) : undefined;
                        let isGradient = true;
                        let nativeColor = '';
                        if (extLst) {
                            const x14DataBar = readXml.getElementsByTagName("x14:dataBar", extLst.value, false)[0];
                            const gradient = getXmlAttibute(x14DataBar.attributeList, 'gradient', null);
                            const negativeFillColor = readXml.getElementsByTagName("x14:dataBar/x14:negativeFillColor", extLst.value, false)?.[0];
                            nativeColor = negativeFillColor ? getColor(negativeFillColor, styles) : undefined;
                            isGradient = gradient !== '0';
                        }
                        rule.config = {
                            min: {
                                type: getXmlAttibute(barCfvo[0]?.attributeList, 'type', 'min'),
                                value: str2num(getXmlAttibute(barCfvo[0]?.attributeList, 'val', undefined))
                            },
                            max: {
                                type: getXmlAttibute(barCfvo[1]?.attributeList, 'type', 'max'),
                                value: str2num(getXmlAttibute(barCfvo[1]?.attributeList, 'val', undefined))
                            },
                            isGradient,
                            positiveColor,
                            nativeColor,
                        };
                        break;
                    case 'iconSet':
                        rule.type = CFRuleType.iconSet;
                        if (!isExtLst) {
                            const iconSet = readXml.getElementsByTagName("iconSet", value, false);
                            const iconCfvo = readXml.getElementsByTagName("iconSet/cfvo", value, false);
                            rule.isShowValue = iconSet[0]?.attributeList?.showValue === '0' ? false : true;
                            rule.config = iconCfvo.map((d, index) => {
                                return {
                                    operator: rule.operator || CFNumberOperator.greaterThanOrEqual,
                                    value: {
                                        type: d.attributeList.type,
                                        value: str2num(d.attributeList.val)
                                    },
                                    iconType: iconSet[0].attributeList.iconSet,
                                    iconId: index
                                };
                            });
                        }
                        else {
                            const iconSet = readXml.getElementsByTagName("x14:iconSet", value, false)[0];
                            const iconCfvo = readXml.getElementsByTagName("x14:iconSet/x14:cfvo", value, false);
                            const cfIcon = readXml.getElementsByTagName("x14:iconSet/x14:cfIcon", value, false);
                            const isCustom = getXmlAttibute(iconSet?.attributeList, 'custom', '0') === '1';
                            rule.isShowValue = iconSet?.attributeList?.showValue === '0' ? false : true;
                            rule.config = iconCfvo.map((d, index) => {
                                const value = readXml.getElementsByTagName("xm:f", d.value, false)[0];
                                let iconInfo = cfIcon[index]?.attributeList;
                                const iconType = isCustom ? iconInfo?.iconSet : iconSet?.attributeList.iconSet;
                                const iconId = str2num(iconType.charAt(0)) - str2num(iconInfo?.iconId) - 1;
                                return {
                                    operator: CFNumberOperator.greaterThanOrEqual,
                                    value: {
                                        type: d.attributeList.type,
                                        value: str2num(value?.value)
                                    },
                                    iconType,
                                    iconId,
                                    // sort: str2num(iconInfo?.iconId)
                                };
                            }).reverse();
                        }
                        break;
                }
                const dxfId = getXmlAttibute(attributeList, 'dxfId', null);
                if (dxfId) {
                    let dxfs = styles["dxfs"];
                    const dxf = dxfId !== null ? dxfs[Number(dxfId)] : undefined;
                    // let numFmt,font,fill,border;
                    const font = dxf.getInnerElements('font')?.[0];
                    const numFmt = dxf.getInnerElements('numFmt')?.[0];
                    const fill = dxf.getInnerElements('fill')?.[0];
                    const border = dxf.getInnerElements('border')?.[0];
                    let numfmts = styles["numfmts"];
                    let style = {};
                    if (border) {
                        const { borderCellValue } = handleBorder(border, styles);
                        style.border = borderCellValue;
                    }
                    if (fill) {
                        const bg = getBackgroundByFill(fill, styles);
                        style.bg = bg;
                    }
                    if (numFmt) {
                        let numf = numfmts[parseInt(numFmt?.attributeList?.numFmtId)];
                        let cellFormat = new LuckySheetCellFormat();
                        cellFormat.fa = numf ? escapeCharacter(numf) : numFmt.attributeList.formatCode;
                        style.ct = cellFormat;
                    }
                    if (font) {
                        const { cellValue } = getFontStyle(font, style);
                        style = { ...style, ...cellValue };
                    }
                    rule.style = style;
                }
                this.rule = rule;
            };
            const { attributeList, parentAttribute } = ele;
            if (parentAttribute?.sqref)
                this.ranges = handleRanges(parentAttribute.sqref);
            this.order = Number(getXmlAttibute(attributeList, 'priority', '1'));
            this.cfId = generateRandomIndex('condition');
            this.stopIfTrue = getXmlAttibute(attributeList, 'stopIfTrue', '0') === '1' ? true : false;
            this.handleRules(ele, readXml, styles);
        }
    }

    /**:
    * @description: 这里处理的dataverification格式是univer用的格式
    * @author: Created by zwight on 2024-09-23 17:50:46
    */
    class LuckyVerification {
        constructor(ele, extLst) {
            let attrList = ele.attributeList;
            let formulaValue = ele.value;
            let type = getXmlAttibute(attrList, "type", undefined);
            if (!type) {
                return;
            }
            this.uid = generateRandomIndex('verification');
            let valueArr = [], sqref = '';
            const operator = getXmlAttibute(attrList, "operator", undefined);
            const allowBlank = getXmlAttibute(attrList, "allowBlank", undefined) !== "1" ? false : true;
            const showInputMessage = getXmlAttibute(attrList, "showInputMessage", undefined) !== "1" ? false : true;
            const showErrorMessage = getXmlAttibute(attrList, "showErrorMessage", undefined) !== "1" ? false : true;
            const prompt = getXmlAttibute(attrList, "prompt", undefined);
            const promptTitle = getXmlAttibute(attrList, "promptTitle", undefined);
            const error = getXmlAttibute(attrList, "error", undefined);
            const errorTitle = getXmlAttibute(attrList, "errorTitle", undefined);
            const errorStyle = getXmlAttibute(attrList, "errorStyle", 'stop');
            const formulaReg = new RegExp(/<x14:formula1>|<xm:sqref>/g);
            if (formulaReg.test(formulaValue) && extLst?.length >= 0) {
                const peelOffData = getPeelOffX14(formulaValue);
                sqref = peelOffData?.sqref;
                valueArr = getMultiFormulaValue(peelOffData?.formula);
            }
            else {
                valueArr = getMultiFormulaValue(formulaValue);
                sqref = getXmlAttibute(attrList, "sqref", null);
            }
            let _value1 = valueArr?.length >= 1 ? valueArr[0] : undefined;
            let _value2 = valueArr?.length === 2 ? valueArr[1] : undefined;
            if (sqref)
                this.ranges = handleRanges(sqref);
            this.type = type;
            this.allowBlank = allowBlank;
            this.operator = operator;
            this.formula1 = _value1;
            this.formula2 = _value2;
            this.showErrorMessage = showErrorMessage;
            this.showInputMessage = showInputMessage;
            this.prompt = prompt;
            this.promptTitle = promptTitle;
            this.error = error;
            this.errorTitle = errorTitle;
            switch (errorStyle) {
                case 'information':
                    this.errorStyle = 0;
                    break;
                case 'warning':
                    this.errorStyle = 2;
                    break;
                case 'stop':
                    this.errorStyle = 1;
                    break;
            }
            // imeMode?: DataValidationImeMode;
            // renderMode?: DataValidationRenderMode;
            // bizInfo?: { showTime?: boolean; };
            this.ranges = handleRanges(sqref);
        }
    }

    /**:
    * @description: 返回的filter属性是univer格式
    * @author: Created by zwight on 2024-09-23 17:50:01
    */
    class LuckFilter {
        constructor(readXml, sheetFile) {
            const autoFilter = readXml.getElementsByTagName('autoFilter', sheetFile)[0];
            if (!autoFilter)
                return;
            this.ref = handleRanges(autoFilter.attributeList.ref)?.[0];
            const filterColumn = autoFilter.getInnerElements('filterColumn') ?? [];
            this.filterColumns = filterColumn.map(d => {
                const filters = d.getInnerElements('filters')?.[0];
                const filter = filters.getInnerElements('filter');
                return {
                    colId: str2num(d.attributeList.colId),
                    filters: {
                        blank: filters.attributeList?.blank === '1',
                        filters: filter.map(d => d.attributeList.val)
                    },
                };
            });
        }
    }

    class LuckyFreezen {
        constructor(pane) {
            const xSplit = getXmlAttibute(pane.attributeList, "xSplit", "0");
            const ySplit = getXmlAttibute(pane.attributeList, "ySplit", "0");
            this.horizen = Number(ySplit);
            this.vertical = Number(xSplit);
        }
    }

    const getChartType = (readXml, chartFile) => {
        const barChart = readXml.getElementsByTagName('c:barChart', chartFile)[0];
        const lineChart = readXml.getElementsByTagName('c:lineChart', chartFile)[0];
        const pieChart = readXml.getElementsByTagName('c:pieChart', chartFile)[0];
        const doughnutChart = readXml.getElementsByTagName('c:doughnutChart', chartFile)[0];
        const areaChart = readXml.getElementsByTagName('c:areaChart', chartFile)[0];
        const radarChart = readXml.getElementsByTagName('c:radarChart', chartFile)[0];
        const scatterChart = readXml.getElementsByTagName('c:scatterChart', chartFile)[0];
        let chartGroup = [];
        let chartEle = barChart;
        let chartType = ChartTypeBits.Column;
        if (barChart) {
            const barDirVal = barChart.getInnerElements('c:barDir')?.[0]?.get('val');
            const groupingVal = barChart.getInnerElements('c:grouping')?.[0]?.get('val');
            if (barDirVal === 'col') {
                chartType = ChartTypeBits.Column;
                if (groupingVal === 'stacked') {
                    chartType = ChartTypeBits.ColumnStacked;
                }
                else if (groupingVal === 'percentStacked') {
                    chartType = ChartTypeBits.ColumnPercentStacked;
                }
            }
            else if (barDirVal === 'bar') {
                chartType = ChartTypeBits.Bar;
                if (groupingVal === 'stacked') {
                    chartType = ChartTypeBits.BarStacked;
                }
                else if (groupingVal === 'percentStacked') {
                    chartType = ChartTypeBits.BarPercentStacked;
                }
            }
            chartEle = barChart;
            chartGroup.push({ chartEle, chartType });
        }
        if (lineChart) {
            chartEle = lineChart;
            chartType = ChartTypeBits.Line;
            chartGroup.push({ chartEle, chartType });
        }
        if (pieChart) {
            chartEle = pieChart;
            chartType = ChartTypeBits.Pie;
            chartGroup.push({ chartEle, chartType });
        }
        if (doughnutChart) {
            chartEle = doughnutChart;
            chartType = ChartTypeBits.Doughnut;
            chartGroup.push({ chartEle, chartType });
        }
        if (areaChart) {
            chartEle = areaChart;
            chartType = ChartTypeBits.Area;
            const groupingVal = areaChart.getInnerElements('c:grouping')?.[0]?.get('val');
            if (groupingVal === 'stacked') {
                chartType = ChartTypeBits.AreaStacked;
            }
            else if (groupingVal === 'percentStacked') {
                chartType = ChartTypeBits.AreaPercentStacked;
            }
            chartGroup.push({ chartEle, chartType });
        }
        if (radarChart) {
            chartEle = radarChart;
            chartType = ChartTypeBits.Radar;
            chartGroup.push({ chartEle, chartType });
        }
        if (scatterChart) {
            chartEle = scatterChart;
            chartType = ChartTypeBits.Scatter;
            chartGroup.push({ chartEle, chartType });
        }
        if (chartGroup.length > 1) {
            chartType = ChartTypeBits.Combination;
            chartEle = chartGroup;
        }
        return {
            chartEle,
            chartType
        };
    };
    class LuckyChartImage extends LuckyChartImageBase {
        constructor(id, xdr_xfrm, range, chartType) {
            super();
            this.id = id;
            this.type = 'chart';
            this.data = {
                chartType,
                range,
                border: '#979DAC',
                background: 'rgba(0,0,0,0)',
                isRowDirection: true,
            };
            let x_n = 0, y_n = 0;
            let cx_n = 0, cy_n = 0;
            const off = xdr_xfrm.getInnerElements('a:off')[0];
            const ext = xdr_xfrm.getInnerElements('a:ext')[0];
            cx_n = getPxByEMUs(parseInt(ext.get('cx'))), cy_n = getPxByEMUs(parseInt(ext.get('cy')));
            x_n = getPxByEMUs(parseInt(off.get('x'))), y_n = getPxByEMUs(parseInt(off.get('y')));
            this.transform = {
                width: cx_n,
                height: cy_n,
                top: y_n,
                left: x_n,
            };
        }
    }
    class ChartImageGroup {
        constructor({ graphicFrame, readXml, drawingRelsFile, styles }) {
            this.getChartRange = (chartEl) => {
                let maxColumn = 0, maxRow = 0, minColumn = 0, minRow = 0;
                if (Array.isArray(chartEl)) {
                    const rangeNumArray = chartEl.map(d => {
                        return this.getChartRef(d.chartEle.value);
                    });
                    rangeNumArray.forEach((d, index) => {
                        if (index === 0) {
                            maxColumn = d.maxColumn;
                            maxRow = d.maxRow;
                            minColumn = d.minColumn;
                            minRow = d.minRow;
                        }
                        else {
                            maxColumn = Math.max(maxColumn, d.maxColumn);
                            maxRow = Math.max(maxRow, d.maxRow);
                            minColumn = Math.min(minColumn, d.minColumn);
                            minRow = Math.min(minRow, d.minRow);
                        }
                    });
                }
                else {
                    const rangeNum = this.getChartRef(chartEl.value);
                    maxColumn = rangeNum.maxColumn;
                    maxRow = rangeNum.maxRow;
                    minColumn = rangeNum.minColumn;
                    minRow = rangeNum.minRow;
                }
                const maxRef = core.numberToABC(maxColumn) + (maxRow + 1);
                const minRef = core.numberToABC(minColumn) + (minRow + 1);
                return minRef + ':' + maxRef;
            };
            this.getChartRef = (chart) => {
                const catNumRef = this.readXml.getElementsByTagName('c:ser/c:cat/c:numRef/c:f', chart, false)?.[0];
                const catStrRef = this.readXml.getElementsByTagName('c:ser/c:cat/c:strRef/c:f', chart, false)?.[0];
                const cXvalStrRef = this.readXml.getElementsByTagName('c:ser/c:xVal/c:strRef/c:f', chart, false)?.[0];
                const cXvalNumRef = this.readXml.getElementsByTagName('c:ser/c:xVal/c:numRef/c:f', chart, false)?.[0];
                const catRef = catNumRef || catStrRef || cXvalStrRef || cXvalNumRef;
                const strRef = this.readXml.getElementsByTagName('c:ser/c:tx/c:strRef/c:f', chart, false);
                const x = catRef, y = strRef[strRef.length - 1];
                const xRange = getcellrange(x.value), yRange = getcellrange(y.value);
                const column = [...xRange.column, ...yRange.column];
                const row = [...xRange.row, ...yRange.row];
                const maxColumn = Math.max(...column), maxRow = Math.max(...row);
                const minColumn = Math.min(...column), minRow = Math.min(...row);
                return {
                    maxColumn,
                    maxRow,
                    minColumn,
                    minRow
                };
            };
            this.readXml = readXml;
            const xdr_xfrm = graphicFrame.getInnerElements('xdr:xfrm')[0];
            const cChart = readXml.getElementsByTagName('a:graphic/a:graphicData/c:chart', graphicFrame.value, false)[0];
            const chartRid = cChart.get('r:id');
            const chartFile = getRelationShip({
                rid: chartRid,
                fileName: drawingRelsFile,
                readXml
            });
            const { chartEle, chartType } = getChartType(readXml, chartFile);
            const range = this.getChartRange(chartEle);
            const id = generateRandomId();
            this.image = new LuckyChartImage(id, xdr_xfrm, range, chartType);
            const chart = new Chart({
                id,
                range,
                chartType,
                chartFile,
                readXml,
                image: this.image,
                styles,
            });
            this.chart = chart.model;
        }
    }
    class Chart extends LuckyChart {
        constructor(args) {
            super();
            this.getStyle = (image) => {
                const bodySpPr = this.readXml.getElementsByTagNameLink('c:spPr', this.chartFile)[0];
                const fill = bodySpPr.getInnerElements('a:solidFill')[0];
                const ln = bodySpPr.getInnerElements('a:ln')[0];
                const backgroundColor = fill ? this.getColor(fill.getInnerElements('a:schemeClr')[0]) : undefined;
                const borderColor = ln ? this.getColor(ln.getInnerElements('a:schemeClr')[0]) : undefined;
                const allTitle = this.getAllTitle();
                const autoTitleDeleted = this.readXml.getElementsByTagNameLink('c:chart/c:autoTitleDeleted', this.chartFile)[0];
                const plotArea = this.readXml.getElementsByTagNameLink('c:chart/c:plotArea', this.chartFile)[0];
                const xAxis = this.getAxis(plotArea?.getInnerElements('c:catAx')?.[0]);
                const yAxis = this.getAxis(plotArea?.getInnerElements('c:valAx')?.[0]);
                const cLegend = this.readXml.getElementsByTagNameLink('c:chart/c:legend', this.chartFile)[0];
                const legend = this.getLegend(cLegend);
                return {
                    titles: {
                        ...allTitle,
                        titlePosition: autoTitleDeleted?.get('val') === '1' ? 'hide' : 'top',
                    },
                    runtime: {},
                    width: image.transform.width,
                    height: image.transform.height,
                    backgroundColor,
                    borderColor,
                    xAxis,
                    yAxis,
                    legend,
                    ...this.getChartSeries()
                };
            };
            this.getChartSeries = () => {
                const { chartEle, chartType } = getChartType(this.readXml, this.chartFile);
                if (!chartEle)
                    return {};
                if (Array.isArray(chartEle)) {
                    const seriesStyleArray = chartEle.map(d => {
                        return this.getChartSeriesBase(d.chartEle, d.chartType, true);
                    });
                    const seriesStyleMap = seriesStyleArray.reduce((pre, cur) => {
                        return Object.assign(pre, cur);
                    }, {});
                    return {
                        seriesStyleMap
                    };
                }
                const allDlbls = chartEle.getInnerElementsTagLink('c:dLbls')[0];
                const show = allDlbls.getInnerElements('c:showVal')[0];
                const allSeriesStyle = {
                    label: {
                        visible: show ? show.get('val') === '1' : false
                    }
                };
                const seriesStyleMap = this.getChartSeriesBase(chartEle, chartType);
                return {
                    allSeriesStyle,
                    seriesStyleMap
                };
            };
            this.getChartSeriesBase = (chartEle, chartType, isGroup) => {
                const cSer = chartEle.getInnerElements('c:ser');
                const seriesStyleMap = {};
                cSer.forEach((element) => {
                    const idx = element.getInnerElements('c:idx')[0];
                    const idxVal = parseInt(idx.get('val')) + 1;
                    const spPr = element.getInnerElements('c:spPr')[0];
                    const fill = spPr?.getInnerElementsTagLink('a:solidFill')?.[0];
                    const ln = spPr?.getInnerElements('a:ln')?.[0];
                    const barColor = this.getColor(fill?.getInnerElements('a:schemeClr')?.[0]);
                    const fillOpacity = fill?.getInnerElements('a:alpha')?.[0]?.get('val');
                    const border = ln ? this.getLine(ln) : undefined;
                    const dLbls = element.getInnerElements('c:dLbls')[0];
                    const showValue = dLbls.getInnerElements('c:showVal')?.[0]?.get('val');
                    const showCatName = dLbls.getInnerElements('c:showCatName')?.[0]?.get('val');
                    const showSerName = dLbls.getInnerElements('c:showSerName')?.[0]?.get('val');
                    const showPercent = dLbls.getInnerElements('c:showPercent')?.[0]?.get('val');
                    const labelStyle = this.getBaseStyle(dLbls.getInnerElements('a:defRPr')?.[0]);
                    const contentType = 0 |
                        (showValue === '1' ? LabelContentType.Value : 0) |
                        (showCatName === '1' ? LabelContentType.CategoryName : 0) |
                        (showSerName === '1' ? LabelContentType.SeriesName : 0) |
                        (showPercent === '1' ? LabelContentType.Percentage : 0);
                    const base = {
                        border,
                        label: {
                            visible: !!contentType,
                            contentType,
                            ...labelStyle,
                        },
                        color: barColor,
                        fillOpacity: fillOpacity ? parseInt(fillOpacity) / 100000 : 1,
                        chartType: isGroup ? chartType : undefined,
                    };
                    let serConf = this.getExtraSerise(base, element, chartType);
                    seriesStyleMap[idxVal] = serConf;
                });
                return seriesStyleMap;
            };
            this.getExtraSerise = (baseObj, series, chartType) => {
                if (chartType === ChartTypeBits.Line) {
                    if (baseObj.border?.color) {
                        baseObj.color = baseObj.border.color;
                    }
                    const marker = series.getInnerElements('c:marker')[0];
                    const symbol = marker.getInnerElements('c:symbol')?.[0]?.get('val');
                    const size = marker.getInnerElements('c:size')?.[0]?.get('val');
                    const schemeClr = marker.getInnerElementsTagLink('c:spPr/a:solidFill/a:schemeClr')?.[0];
                    const color = this.getColor(schemeClr);
                    baseObj['point'] = {
                        color,
                        size: size ? parseInt(size) : undefined,
                        shape: symbol === 'none' && baseObj.label.visible ? 'circle' : symbol
                    };
                }
                return baseObj;
            };
            this.getLegend = (legend) => {
                if (!legend)
                    return {
                        position: 'hide'
                    };
                const positionMap = {
                    t: 'top',
                    b: 'bottom',
                    l: 'left',
                    r: 'right',
                };
                const position = legend.getInnerElements('c:legendPos')[0]?.get('val');
                const txPr = legend.getInnerElements('c:txPr')[0];
                const pPr = txPr.getInnerElements('a:pPr')[0];
                const labelStyle = this.getBaseStyle(pPr.getInnerElements('a:defRPr')[0]);
                return {
                    position: position ? positionMap[position] : 'bottom',
                    label: labelStyle
                };
            };
            this.getAxis = (axis) => {
                if (!axis)
                    return undefined;
                const visible = !axis.getInnerElementsTagLink('c:spPr/a:ln/a:noFill')?.length;
                // const visible = axis.getInnerElements('c:delete')[0]?.get('val') === '0';
                const scaling = axis.getInnerElements('c:scaling')[0];
                const reverse = scaling?.getInnerElements('c:orientation')[0]?.get('val') === 'maxMin';
                const max = scaling?.getInnerElements('c:max')?.[0]?.get('val');
                const min = scaling?.getInnerElements('c:min')?.[0]?.get('val');
                const titleAlignMap = {
                    l: 'start',
                    r: 'end',
                    ctr: 'center'
                };
                const axisTitleAlign = axis.getInnerElements('c:lblAlgn')?.[0]?.get('val');
                const txPr = axis.getInnerElements('c:txPr')[0];
                const pPr = txPr.getInnerElements('a:pPr')[0];
                const rotate = txPr?.getInnerElements('a:bodyPr')?.[0]?.get('rot');
                const labelStyle = this.getBaseStyle(pPr.getInnerElements('a:defRPr')[0]);
                const majorGridlines = axis.getInnerElements('c:majorGridlines')?.[0];
                const ln = majorGridlines?.getInnerElements('a:ln')[0];
                const gridLineWidth = ln?.get('w') ? getPxByEMUs(parseInt(ln.get('w'))) : undefined;
                const majorTickMark = axis.getInnerElements('c:majorTickMark')?.[0];
                const gridLineColor = this.getColor(ln?.getInnerElements('a:schemeClr')?.[0]);
                return {
                    lineVisible: visible,
                    reverse,
                    max: max ? parseInt(max) : undefined,
                    min: min ? parseInt(min) : undefined,
                    label: {
                        axisTitleAlign: axisTitleAlign ? titleAlignMap[axisTitleAlign] : 'center',
                        rotate: rotate && parseInt(rotate) > 0 ? parseInt(rotate) / 60000 : 0,
                        ...labelStyle
                    },
                    gridLine: {
                        visible: !!majorGridlines?.value,
                        width: gridLineWidth,
                        color: gridLineColor,
                    },
                    tick: {
                        visible: majorTickMark?.get('val') !== 'none',
                        position: majorTickMark?.get('val') === 'out' ? 'outside' : 'inside',
                        lineWidth: gridLineWidth
                    }
                };
            };
            this.getLine = (ln) => {
                const borderWidth = ln?.get('w') ? getPxByEMUs(parseInt(ln.get('w'))) : 0;
                const borderColor = this.getColor(ln?.getInnerElements('a:schemeClr')?.[0]);
                const borderType = ln.getInnerElements('a:prstDash')?.[0]?.get('val');
                const borderMap = {
                    solid: 'solid',
                    dot: 'dotted',
                    dash: 'dashed',
                    sysDot: 'dotted',
                    sysDash: 'dashed',
                    lgDash: 'dashed',
                };
                return {
                    dashType: borderType && borderMap[borderType] ? borderMap[borderType] : 'solid',
                    color: borderColor,
                    width: borderWidth
                };
            };
            this.getAllTitle = () => {
                const mainTitle = this.readXml.getElementsByTagNameLink('c:chart/c:title', this.chartFile)[0];
                const catTitle = this.readXml.getElementsByTagNameLink('c:chart/c:plotArea/c:catAx/c:title', this.chartFile)[0];
                const valTitle = this.readXml.getElementsByTagNameLink('c:chart/c:plotArea/c:valAx/c:title', this.chartFile)[0];
                const title = this.getTitle(mainTitle);
                const xTitle = this.getTitle(catTitle);
                const yTitle = this.getTitle(valTitle);
                return {
                    title,
                    xAxisTitle: xTitle,
                    yAxisTitle: yTitle,
                };
            };
            /**
             * 标题
             * @returns
             */
            this.getTitle = (title) => {
                if (!title)
                    return;
                const ar = title.getInnerElements('a:r');
                const pPr = title.getInnerElements('a:pPr')[0];
                const titleBase = this.getBaseStyle(pPr.getInnerElements('a:defRPr')[0]);
                const titleStyle = ar?.map(d => {
                    return this.getBaseStyle(d.getInnerElements('a:rPr')?.[0]);
                }) || [];
                return {
                    content: ar?.map(d => d.getInnerElements('a:t')[0].value).join(''),
                    ...titleBase,
                    ...(titleStyle.length ? titleStyle[0] : {}),
                };
            };
            this.getBaseStyle = (style) => {
                if (!style)
                    return {};
                const schemaClr = style.getInnerElements('a:schemeClr')?.[0];
                const solidFill = style.getInnerElements('a:solidFill')?.[0];
                const color = (!schemaClr && solidFill) ? this.getThemColor(solidFill) : this.getColor(schemaClr);
                const obj = {};
                if (style.get('sz')) {
                    obj.fontSize = parseInt(style.get('sz')) / 100;
                }
                if (style.get('b')) {
                    obj.bold = style.get('b') === '1';
                }
                if (style.get('i')) {
                    obj.italic = style.get('i') === '1';
                }
                if (color) {
                    obj.color = color;
                }
                return obj;
            };
            this.getColor = (ele) => {
                if (!ele)
                    return undefined;
                const val = ele.get('val');
                const clrScheme = this.styles['clrScheme'];
                const schema = clrScheme.find((item) => {
                    const clrs = item.getInnerElements("a:sysClr|a:srgbClr")[0];
                    if (item.container.includes(val))
                        return true;
                    if (val === 'tx1')
                        return clrs.get('val') === 'windowText';
                    if (val === 'bg1')
                        return clrs.get('val') === 'window';
                    return false;
                });
                let color = '#000000';
                if (schema) {
                    color = this.getThemColor(schema);
                }
                const mod = ele.getInnerElements('a:lumMod')?.[0]?.get('val');
                const lum = ele.getInnerElements('a:lumOff')?.[0]?.get('val');
                const rgbArray = hexToRgbArray(color).map(d => {
                    const sumMod = mod ? d * parseInt(mod) / 100000 : d;
                    const sumLum = lum ? Math.round(255 * (parseInt(lum) / 100000)) : 0;
                    return Math.round(sumMod + sumLum);
                });
                return `rgb(${rgbArray.join(',')})`;
            };
            this.getThemColor = (schema) => {
                let color = '#000000';
                const clrs = schema.getInnerElements("a:sysClr|a:srgbClr");
                if (clrs != null) {
                    const clr = clrs[0];
                    const clrAttrList = clr.attributeList;
                    if (clr.container.indexOf("sysClr") > -1) {
                        if (clrAttrList.lastClr != null) {
                            color = "#" + clrAttrList.lastClr;
                        }
                        else if (clrAttrList.val != null) {
                            color = "#" + clrAttrList.val;
                        }
                    }
                    else if (clr.container.indexOf("srgbClr") > -1) {
                        color = "#" + clrAttrList.val;
                    }
                }
                return color;
            };
            this.getContext = () => {
                const cSer = this.readXml.getElementsByTagName('c:ser', this.chartFile);
                const indexs = cSer.map((d) => {
                    const idx = d.getInnerElements('c:idx')[0];
                    const idxVal = parseInt(idx.get('val')) + 1;
                    return idxVal;
                });
                // console.log(cSer, indexs)
                return {
                    categoryIndex: 0,
                    seriesIndexes: indexs
                };
            };
            const { id, range, chartType, chartFile, readXml, image, styles } = args;
            this.styles = styles;
            this.range = range;
            this.chartType = chartType;
            this.readXml = readXml;
            this.chartFile = chartFile;
            this.isRowDirection = true;
            this.id = id;
            this.context = this.getContext();
            this.style = this.getStyle(image);
        }
        get model() {
            return {
                id: this.id,
                range: this.range,
                chartType: this.chartType,
                context: this.context,
                style: this.style,
                isRowDirection: this.isRowDirection,
            };
        }
    }

    class LuckySheet extends LuckySheetBase {
        constructor(sheetName, sheetId, sheetOrder, isInitialCell = false, allFileOption) {
            //Private
            super();
            this.getImageBaseInfo = (drawingFile, drawingRelsFile) => {
                let twoCellAnchors = this.readXml.getElementsByTagName('xdr:twoCellAnchor', drawingFile);
                let oneCellAnchors = this.readXml.getElementsByTagName('xdr:oneCellAnchor', drawingFile);
                twoCellAnchors = [...twoCellAnchors, ...oneCellAnchors];
                if (twoCellAnchors != null && twoCellAnchors.length > 0) {
                    for (let i = 0; i < twoCellAnchors.length; i++) {
                        let twoCellAnchor = twoCellAnchors[i];
                        let xdrFroms = twoCellAnchor.getInnerElements('xdr:from'), xdrTos = twoCellAnchor.getInnerElements('xdr:to');
                        if (xdrFroms != null && xdrFroms.length > 0) {
                            let xdrFrom = xdrFroms[0], xdrTo, xdrExt;
                            if (xdrTos) {
                                xdrTo = xdrTos[0];
                            }
                            else {
                                xdrExt = twoCellAnchor.getInnerElements('xdr:ext')[0];
                            }
                            let xdr_graphicFrame = twoCellAnchor.getInnerElements('xdr:graphicFrame');
                            let imageObject = xdr_graphicFrame
                                ? this.getGraphic(twoCellAnchor, drawingRelsFile)
                                : this.getImage(twoCellAnchor, drawingRelsFile);
                            let x_n = 0, y_n = 0;
                            let cx_n = 0, cy_n = 0;
                            imageObject.fromCol = this.getXdrValue(xdrFrom.getInnerElements('xdr:col'));
                            imageObject.fromColOff = getPxByEMUs(this.getXdrValue(xdrFrom.getInnerElements('xdr:colOff')));
                            imageObject.fromRow = this.getXdrValue(xdrFrom.getInnerElements('xdr:row'));
                            imageObject.fromRowOff = getPxByEMUs(this.getXdrValue(xdrFrom.getInnerElements('xdr:rowOff')));
                            if (xdrTo) {
                                imageObject.toCol = this.getXdrValue(xdrTo.getInnerElements('xdr:col'));
                                imageObject.toColOff = getPxByEMUs(this.getXdrValue(xdrTo.getInnerElements('xdr:colOff')));
                                imageObject.toRow = this.getXdrValue(xdrTo.getInnerElements('xdr:row'));
                                imageObject.toRowOff = getPxByEMUs(this.getXdrValue(xdrTo.getInnerElements('xdr:rowOff')));
                            }
                            else {
                                let a = xdrExt.attributeList;
                                (cx_n = getPxByEMUs(parseInt(a.cx))),
                                    (cy_n = getPxByEMUs(parseInt(a.cy)));
                                imageObject.toCol = imageObject.fromCol;
                                imageObject.toColOff = Number(imageObject.fromColOff) + cx_n;
                                imageObject.toRow = imageObject.fromRow;
                                imageObject.toRowOff = Number(imageObject.fromRowOff) + cy_n;
                            }
                            imageObject.originWidth = cx_n;
                            imageObject.originHeight = cy_n;
                            imageObject.isFixedPos = false;
                            imageObject.fixedLeft = 0;
                            imageObject.fixedTop = 0;
                            let imageBorder = {
                                color: '#000',
                                radius: 0,
                                style: 'solid',
                                width: 0,
                            };
                            imageObject.border = imageBorder;
                            let imageCrop = {
                                height: cy_n,
                                offsetLeft: 0,
                                offsetTop: 0,
                                width: cx_n,
                            };
                            imageObject.crop = imageCrop;
                            let imageDefault = {
                                height: cy_n,
                                left: x_n,
                                top: y_n,
                                width: cx_n,
                            };
                            imageObject.default = imageDefault;
                            if (this.images == null) {
                                this.images = {};
                            }
                            this.images[imageObject.id || generateRandomIndex('image')] =
                                imageObject;
                        }
                    }
                }
                return null;
            };
            this.getImage = (twoCellAnchor, drawingRelsFile) => {
                let xdr_blipfills = twoCellAnchor.getInnerElements('a:blip');
                let editAs = getXmlAttibute(twoCellAnchor.attributeList, 'editAs', 'twoCell');
                if (xdr_blipfills != null && xdr_blipfills.length > 0) {
                    var xdr_blipfill = xdr_blipfills[0];
                    let rembed = getXmlAttibute(xdr_blipfill.attributeList, 'r:embed', null);
                    let imageObject = this.getBase64ByRid(rembed, drawingRelsFile);
                    if (editAs == 'absolute') {
                        imageObject.type = '3';
                    }
                    else if (editAs == 'oneCell') {
                        imageObject.type = '2';
                    }
                    else {
                        imageObject.type = '1';
                    }
                    return imageObject;
                }
                return {};
            };
            this.getGraphic = (twoCellAnchor, drawingRelsFile) => {
                const xdr_graphicFrames = twoCellAnchor.getInnerElements('xdr:graphicFrame');
                if (xdr_graphicFrames.length) {
                    const xdr_graphicFrame = xdr_graphicFrames[0];
                    const chartImageGroup = new ChartImageGroup({
                        graphicFrame: xdr_graphicFrame,
                        readXml: this.readXml,
                        drawingRelsFile,
                        styles: this.styles,
                    });
                    const imageObject = chartImageGroup.image;
                    if (chartImageGroup.chart) {
                        if (this.charts == null) {
                            this.charts = [];
                        }
                        this.charts.push(chartImageGroup.chart);
                    }
                    return imageObject;
                }
                return {};
            };
            // private getBorderInfo(borders:Element[]):LuckySheetborderInfoCellValueStyle{
            //     if(borders==null){
            //         return null;
            //     }
            //     let border = borders[0], attrList = border.attributeList;
            //     let clrScheme = this.styles["clrScheme"] as Element[];
            //     let style:string = attrList.style;
            //     if(style==null || style=="none"){
            //         return null;
            //     }
            //     let colors = border.getInnerElements("color");
            //     let colorRet = "#000000";
            //     if(colors!=null){
            //         let color = colors[0];
            //         colorRet = getColor(color, clrScheme);
            //     }
            //     let ret = new LuckySheetborderInfoCellValueStyle();
            //     ret.style = borderTypes[style];
            //     ret.color = colorRet;
            //     return ret;
            // }
            this.getCellSize = (cell) => {
                let attrList = cell.attributeList;
                let r = attrList.r; attrList.s; attrList.t;
                let range = getcellrange(r);
                const row = range.row[0];
                const col = range.column[0];
                const width = this.config.columnlen && this.config.columnlen[col]
                    ? this.config.columnlen[col]
                    : this.defaultColWidth;
                const height = this.config.rowlen && this.config.rowlen[row]
                    ? this.config.rowlen[row]
                    : this.defaultRowHeight;
                return {
                    width,
                    height,
                };
            };
            this.isInitialCell = isInitialCell;
            this.readXml = allFileOption.readXml;
            this.sheetFile = allFileOption.sheetFile;
            this.styles = allFileOption.styles;
            this.sharedStrings = allFileOption.sharedStrings;
            this.calcChainEles = allFileOption.calcChain;
            this.sheetList = allFileOption.sheetList;
            this.imageList = allFileOption.imageList;
            this.hide = allFileOption.hide;
            this.cellImages = allFileOption.cellImages;
            //Output
            this.name = sheetName;
            this.index = sheetId;
            this.order = sheetOrder.toString();
            this.config = new LuckyConfig();
            this.celldata = [];
            this.mergeCells = this.readXml.getElementsByTagName('mergeCells/mergeCell', this.sheetFile);
            this.styles['clrScheme'];
            let sheetView = this.readXml.getElementsByTagName('sheetViews/sheetView', this.sheetFile);
            let showGridLines = '1', tabSelected = '0', zoomScale = '100', activeCell = 'A1';
            if (sheetView.length > 0) {
                let attrList = sheetView[0].attributeList;
                showGridLines = getXmlAttibute(attrList, 'showGridLines', '1');
                tabSelected = getXmlAttibute(attrList, 'tabSelected', '0');
                zoomScale = getXmlAttibute(attrList, 'zoomScale', '100');
                // let colorId = getXmlAttibute(attrList, "colorId", "0");
                let selections = sheetView[0].getInnerElements('selection');
                if (selections != null && selections.length > 0) {
                    activeCell = getXmlAttibute(selections[0].attributeList, 'activeCell', 'A1');
                    let range = getcellrange(activeCell);
                    this.luckysheet_select_save = [];
                    this.luckysheet_select_save.push(range);
                }
                let pane = sheetView[0].getInnerElements('pane');
                if (pane?.length > 0) {
                    this.freezen = new LuckyFreezen(pane[0]);
                }
            }
            this.showGridLines = showGridLines;
            this.status = tabSelected;
            this.zoomRatio = parseInt(zoomScale) / 100;
            let tabColors = this.readXml.getElementsByTagName('sheetPr/tabColor', this.sheetFile);
            if (tabColors != null && tabColors.length > 0) {
                let tabColor = tabColors[0]; tabColor.attributeList;
                // if(attrList.rgb!=null){
                let tc = getColor(tabColor, this.styles, 'b');
                this.color = tc;
                // }
            }
            let sheetFormatPr = this.readXml.getElementsByTagName('sheetFormatPr', this.sheetFile);
            let defaultColWidth, defaultRowHeight;
            if (sheetFormatPr.length > 0) {
                let attrList = sheetFormatPr[0].attributeList;
                defaultColWidth = getXmlAttibute(attrList, 'defaultColWidth', '9.21');
                defaultRowHeight = getXmlAttibute(attrList, 'defaultRowHeight', '19');
            }
            this.defaultColWidth = getColumnWidthPixel(parseFloat(defaultColWidth));
            this.defaultRowHeight = getRowHeightPixel(parseFloat(defaultRowHeight));
            this.generateConfigColumnLenAndHidden();
            let cellOtherInfo = this.generateConfigRowLenAndHiddenAddCell();
            if (this.calcChain == null) {
                this.calcChain = [];
            }
            let formulaListExist = {};
            for (let c = 0; c < this.calcChainEles.length; c++) {
                let calcChainEle = this.calcChainEles[c], attrList = calcChainEle.attributeList;
                if (attrList.i != sheetId) {
                    continue;
                }
                let r = attrList.r; attrList.i; attrList.l; attrList.s; attrList.a; attrList.t;
                let range = getcellrange(r);
                let chain = new LuckysheetCalcChain();
                chain.r = range.row[0];
                chain.c = range.column[0];
                chain.index = this.index;
                this.calcChain.push(chain);
                formulaListExist['r' + r + 'c' + c] = null;
            }
            if (this.formulaRefList != null) {
                for (let key in this.formulaRefList) {
                    let funclist = this.formulaRefList[key];
                    let mainFunc = funclist['mainRef'], mainCellValue = mainFunc.cellValue;
                    let formulaTxt = mainFunc.fv;
                    let mainR = mainCellValue.r, mainC = mainCellValue.c;
                    // let refRange = getcellrange(ref);
                    for (let name in funclist) {
                        if (name == 'mainRef') {
                            continue;
                        }
                        let funcValue = funclist[name], cellValue = funcValue.cellValue;
                        if (cellValue == null) {
                            continue;
                        }
                        let r = cellValue.r, c = cellValue.c;
                        let func = formulaTxt;
                        let offsetRow = r - mainR, offsetCol = c - mainC;
                        if (offsetRow > 0) {
                            func = '=' + fromulaRef.functionCopy(func, 'down', offsetRow);
                        }
                        else if (offsetRow < 0) {
                            func =
                                '=' + fromulaRef.functionCopy(func, 'up', Math.abs(offsetRow));
                        }
                        if (offsetCol > 0) {
                            func = '=' + fromulaRef.functionCopy(func, 'right', offsetCol);
                        }
                        else if (offsetCol < 0) {
                            func =
                                '=' + fromulaRef.functionCopy(func, 'left', Math.abs(offsetCol));
                        }
                        // console.log(offsetRow, offsetCol, func);
                        cellValue.v.f = func;
                        //添加共享公式链
                        let chain = new LuckysheetCalcChain();
                        chain.r = cellValue.r;
                        chain.c = cellValue.c;
                        chain.index = this.index;
                        this.calcChain.push(chain);
                    }
                }
            }
            //There may be formulas that do not appear in calcChain
            for (let key in cellOtherInfo.formulaList) {
                if (!(key in formulaListExist)) {
                    let formulaListItem = cellOtherInfo.formulaList[key];
                    let chain = new LuckysheetCalcChain();
                    chain.r = formulaListItem.r;
                    chain.c = formulaListItem.c;
                    chain.index = this.index;
                    this.calcChain.push(chain);
                }
            }
            const conditionList = this.readXml.getElementsByTagName('conditionalFormatting', this.sheetFile);
            const extLstCondition = this.readXml.getElementsByTagName('extLst/ext/x14:conditionalFormattings/x14:conditionalFormatting', this.sheetFile) || [];
            const extLstRule = extLstCondition
                ?.map((condition) => {
                const sqref = this.readXml.getElementsByTagName('xm:sqref', condition.value, false)?.[0];
                return this.readXml
                    .getElementsByTagName('x14:cfRule', condition.value, false)
                    .map((d) => ({
                    ...d,
                    parentAttribute: { sqref: sqref?.value },
                    isExtLst: true,
                    extLst: undefined,
                }));
            })
                ?.flat() || [];
            if (conditionList?.length) {
                const ruleList = conditionList
                    .map((condition) => {
                    return this.readXml
                        .getElementsByTagName('cfRule', condition.value, false)
                        ?.map((d) => ({
                        ...d,
                        parentAttribute: condition.attributeList,
                        extLst: extLstRule.find((d) => d.parentAttribute.sqref === condition.attributeList?.sqref),
                    }));
                })
                    ?.flat()
                    .filter(Boolean)
                    .concat(extLstRule?.filter((d) => conditionList.findIndex((condition) => condition.attributeList.sqref === d.parentAttribute.sqref) === -1)) || [];
                this.conditionalFormatting = ruleList.map((d) => new LuckyCondition(d, this.readXml, this.styles));
                // console.log(ruleList, allFileOption, this.conditionalFormatting)
            }
            // console.log(allFileOption)
            const filter = new LuckFilter(this.readXml, this.sheetFile);
            if (filter.ref)
                this.filter = filter;
            // dataVerification config
            this.dataVerification = this.generateConfigDataValidations();
            this.dataVerificationList = this.generateConfigDataValidationsList();
            // console.log('dataVerificationList ---->', this.dataVerificationList)
            // hyperlink config
            this.hyperlink = this.generateConfigHyperlinks();
            // sheet hide
            this.hide = this.hide;
            if (this.mergeCells != null) {
                for (let i = 0; i < this.mergeCells.length; i++) {
                    let merge = this.mergeCells[i], attrList = merge.attributeList;
                    let ref = attrList.ref;
                    if (ref == null) {
                        continue;
                    }
                    let range = getcellrange(ref);
                    let mergeValue = new LuckySheetConfigMerge();
                    mergeValue.r = range.row[0];
                    mergeValue.c = range.column[0];
                    mergeValue.rs = range.row[1] - range.row[0] + 1;
                    mergeValue.cs = range.column[1] - range.column[0] + 1;
                    if (this.config.merge == null) {
                        this.config.merge = {};
                    }
                    this.config.merge[range.row[0] + '_' + range.column[0]] = mergeValue;
                }
            }
            let drawingFile = allFileOption.drawingFile, drawingRelsFile = allFileOption.drawingRelsFile;
            if (drawingFile != null && drawingRelsFile != null) {
                this.getImageBaseInfo(drawingFile, drawingRelsFile);
            }
        }
        getXdrValue(ele) {
            if (ele == null || ele.length == 0) {
                return null;
            }
            return parseInt(ele[0].value);
        }
        getBase64ByRid(rid, drawingRelsFile) {
            let Relationships = this.readXml.getElementsByTagName('Relationships/Relationship', drawingRelsFile);
            if (Relationships != null && Relationships.length > 0) {
                for (let i = 0; i < Relationships.length; i++) {
                    let Relationship = Relationships[i];
                    let attrList = Relationship.attributeList;
                    let Id = getXmlAttibute(attrList, 'Id', null);
                    let src = getXmlAttibute(attrList, 'Target', null);
                    if (Id == rid) {
                        src = src.replace(/\.\.\//g, '');
                        src = 'xl/' + src;
                        let imgage = this.imageList.getImageByName(src);
                        return imgage;
                    }
                }
            }
            return {};
        }
        /**
         * @desc This will convert cols/col to luckysheet config of column'width
         */
        generateConfigColumnLenAndHidden() {
            let cols = this.readXml.getElementsByTagName('cols/col', this.sheetFile);
            for (let i = 0; i < cols.length; i++) {
                let col = cols[i], attrList = col.attributeList;
                let min = getXmlAttibute(attrList, 'min', null);
                let max = getXmlAttibute(attrList, 'max', null);
                let width = getXmlAttibute(attrList, 'width', null);
                let hidden = getXmlAttibute(attrList, 'hidden', null);
                let customWidth = getXmlAttibute(attrList, 'customWidth', null);
                if (min == null || max == null) {
                    continue;
                }
                let minNum = parseInt(min) - 1, maxNum = parseInt(max) - 1, widthNum = parseFloat(width);
                for (let m = minNum; m <= maxNum; m++) {
                    if (width != null) {
                        if (this.config.columnlen == null) {
                            this.config.columnlen = {};
                        }
                        this.config.columnlen[m] = getColumnWidthPixel(widthNum);
                    }
                    if (hidden == '1') {
                        if (this.config.colhidden == null) {
                            this.config.colhidden = {};
                        }
                        this.config.colhidden[m] = 0;
                        if (this.config.columnlen) {
                            delete this.config.columnlen[m];
                        }
                    }
                    if (customWidth != null) {
                        if (this.config.customWidth == null) {
                            this.config.customWidth = {};
                        }
                        this.config.customWidth[m] = 1;
                    }
                }
            }
        }
        /**
         * @desc This will convert cols/col to luckysheet config of column'width
         */
        generateConfigRowLenAndHiddenAddCell() {
            let rows = this.readXml.getElementsByTagName('sheetData/row', this.sheetFile);
            let cellOtherInfo = {};
            let formulaList = {};
            cellOtherInfo.formulaList = formulaList;
            for (let i = 0; i < rows.length; i++) {
                let row = rows[i], attrList = row.attributeList;
                let rowNo = getXmlAttibute(attrList, 'r', null);
                let height = getXmlAttibute(attrList, 'ht', null);
                let hidden = getXmlAttibute(attrList, 'hidden', null);
                let customHeight = getXmlAttibute(attrList, 'customHeight', null);
                if (rowNo == null) {
                    continue;
                }
                let rowNoNum = parseInt(rowNo) - 1;
                if (height != null) {
                    let heightNum = parseFloat(height);
                    if (this.config.rowlen == null) {
                        this.config.rowlen = {};
                    }
                    this.config.rowlen[rowNoNum] = getRowHeightPixel(heightNum);
                }
                if (hidden == '1') {
                    if (this.config.rowhidden == null) {
                        this.config.rowhidden = {};
                    }
                    this.config.rowhidden[rowNoNum] = 0;
                    if (this.config.rowlen) {
                        delete this.config.rowlen[rowNoNum];
                    }
                }
                if (customHeight != null) {
                    if (this.config.customHeight == null) {
                        this.config.customHeight = {};
                    }
                    this.config.customHeight[rowNoNum] = 1;
                }
                if (this.isInitialCell) {
                    let cells = row.getInnerElements('c');
                    for (let key in cells) {
                        let cell = cells[key];
                        const cellSize = this.getCellSize(cell);
                        let cellValue = new LuckySheetCelldata(cell, cellSize, this.styles, this.sharedStrings, this.mergeCells, this.sheetFile, this.cellImages, this.imageList, this.readXml);
                        if (cellValue._borderObject != null) {
                            if (this.config.borderInfo == null) {
                                this.config.borderInfo = [];
                            }
                            this.config.borderInfo.push(cellValue._borderObject);
                            delete cellValue._borderObject;
                        }
                        // let borderId = cellValue._borderId;
                        // if(borderId!=null){
                        //     let borders = this.styles["borders"] as Element[];
                        //     if(this.config._borderInfo==null){
                        //         this.config._borderInfo = {};
                        //     }
                        //     if( borderId in this.config._borderInfo){
                        //         this.config._borderInfo[borderId].cells.push(cellValue.r + "_" + cellValue.c);
                        //     }
                        //     else{
                        //         let border = borders[borderId];
                        //         let borderObject = new LuckySheetborderInfoCellForImp();
                        //         borderObject.rangeType = "cellGroup";
                        //         borderObject.cells = [];
                        //         let borderCellValue = new LuckySheetborderInfoCellValue();
                        //         let lefts = border.getInnerElements("left");
                        //         let rights = border.getInnerElements("right");
                        //         let tops = border.getInnerElements("top");
                        //         let bottoms = border.getInnerElements("bottom");
                        //         let diagonals = border.getInnerElements("diagonal");
                        //         let left = this.getBorderInfo(lefts);
                        //         let right = this.getBorderInfo(rights);
                        //         let top = this.getBorderInfo(tops);
                        //         let bottom = this.getBorderInfo(bottoms);
                        //         let diagonal = this.getBorderInfo(diagonals);
                        //         let isAdd = false;
                        //         if(left!=null && left.color!=null){
                        //             borderCellValue.l = left;
                        //             isAdd = true;
                        //         }
                        //         if(right!=null && right.color!=null){
                        //             borderCellValue.r = right;
                        //             isAdd = true;
                        //         }
                        //         if(top!=null && top.color!=null){
                        //             borderCellValue.t = top;
                        //             isAdd = true;
                        //         }
                        //         if(bottom!=null && bottom.color!=null){
                        //             borderCellValue.b = bottom;
                        //             isAdd = true;
                        //         }
                        //         if(isAdd){
                        //             borderObject.value = borderCellValue;
                        //             this.config._borderInfo[borderId] = borderObject;
                        //         }
                        //     }
                        // }
                        if (cellValue._formulaType == 'shared') {
                            if (this.formulaRefList == null) {
                                this.formulaRefList = {};
                            }
                            if (this.formulaRefList[cellValue._formulaSi] == null) {
                                this.formulaRefList[cellValue._formulaSi] = {};
                            }
                            let fv;
                            if (cellValue.v != null) {
                                fv = cellValue.v.f;
                            }
                            let refValue = {
                                t: cellValue._formulaType,
                                ref: cellValue._fomulaRef,
                                si: cellValue._formulaSi,
                                fv: fv,
                                cellValue: cellValue,
                            };
                            if (cellValue._fomulaRef != null) {
                                this.formulaRefList[cellValue._formulaSi]['mainRef'] = refValue;
                            }
                            else {
                                this.formulaRefList[cellValue._formulaSi][cellValue.r + '_' + cellValue.c] = refValue;
                            }
                            // console.log(refValue, this.formulaRefList);
                        }
                        //There may be formulas that do not appear in calcChain
                        if (cellValue.v != null &&
                            cellValue.v.f != null) {
                            let formulaCell = {
                                r: cellValue.r,
                                c: cellValue.c,
                            };
                            cellOtherInfo.formulaList['r' + cellValue.r + 'c' + cellValue.c] =
                                formulaCell;
                        }
                        this.celldata.push(cellValue);
                    }
                }
            }
            return cellOtherInfo;
        }
        generateConfigDataValidationsList() {
            let rows = this.readXml.getElementsByTagName('dataValidations/dataValidation', this.sheetFile);
            let extLst = this.readXml.getElementsByTagName('extLst/ext/x14:dataValidations/x14:dataValidation', this.sheetFile) || [];
            rows = rows.concat(extLst);
            return rows
                .map((d) => new LuckyVerification(d, extLst))
                .filter((d) => d.uid);
        }
        /**
         * luckysheet config of dataValidations
         *
         * @returns {IluckysheetDataVerification} - dataValidations config
         */
        generateConfigDataValidations() {
            let rows = this.readXml.getElementsByTagName('dataValidations/dataValidation', this.sheetFile);
            let extLst = this.readXml.getElementsByTagName('extLst/ext/x14:dataValidations/x14:dataValidation', this.sheetFile) || [];
            rows = rows.concat(extLst);
            let dataVerification = {};
            for (let i = 0; i < rows.length; i++) {
                let row = rows[i];
                let attrList = row.attributeList;
                let formulaValue = row.value;
                let type = getXmlAttibute(attrList, 'type', null);
                if (!type) {
                    continue;
                }
                let operator = '', sqref = '', sqrefIndexArr = [], valueArr = [];
                let _prohibitInput = getXmlAttibute(attrList, 'allowBlank', null) !== '1' ? false : true;
                // x14 processing
                const formulaReg = new RegExp(/<x14:formula1>|<xm:sqref>/g);
                if (formulaReg.test(formulaValue) && extLst?.length >= 0) {
                    operator = getXmlAttibute(attrList, 'operator', null);
                    const peelOffData = getPeelOffX14(formulaValue);
                    sqref = peelOffData?.sqref;
                    sqrefIndexArr = getMultiSequenceToNum(sqref);
                    valueArr = getMultiFormulaValue(peelOffData?.formula);
                }
                else {
                    operator = getXmlAttibute(attrList, 'operator', null);
                    sqref = getXmlAttibute(attrList, 'sqref', null);
                    sqrefIndexArr = getMultiSequenceToNum(sqref);
                    valueArr = getMultiFormulaValue(formulaValue);
                }
                let _type = DATA_VERIFICATION_MAP[type];
                let _type2 = null;
                let _value1 = valueArr?.length >= 1 ? valueArr[0] : '';
                let _value2 = valueArr?.length === 2 ? valueArr[1] : '';
                let _hint = getXmlAttibute(attrList, 'prompt', null);
                let _hintShow = _hint ? true : false;
                const matchType = COMMON_TYPE2.includes(_type) || !DATA_VERIFICATION_TYPE2_MAP[_type]
                    ? 'common'
                    : _type;
                _type2 = operator
                    ? DATA_VERIFICATION_TYPE2_MAP[matchType][operator]
                    : 'bw';
                // mobile phone number processing
                if (_type === 'text_content' &&
                    (_value1?.includes('LEN') || _value1?.includes('len')) &&
                    _value1?.includes('=11')) {
                    _type = 'validity';
                    _type2 = 'phone';
                }
                // date processing
                if (_type === 'date') {
                    const D1900 = new Date(1899, 11, 30, 0, 0, 0);
                    _value1 = dayjs__default["default"](D1900)
                        .clone()
                        .add(Number(_value1), 'day')
                        .format('YYYY-MM-DD');
                    _value2 = dayjs__default["default"](D1900)
                        .clone()
                        .add(Number(_value2), 'day')
                        .format('YYYY-MM-DD');
                }
                // checkbox and dropdown processing
                if (_type === 'checkbox' || _type === 'dropdown') {
                    _type2 = null;
                }
                // dynamically add dataVerifications
                for (const ref of sqrefIndexArr) {
                    dataVerification[ref] = {
                        type: _type,
                        type2: _type2,
                        value1: _value1,
                        value2: _value2,
                        checked: false,
                        remote: false,
                        prohibitInput: _prohibitInput,
                        hintShow: _hintShow,
                        hintText: _hint,
                    };
                }
            }
            return dataVerification;
        }
        /**
         * luckysheet config of hyperlink
         *
         * @returns {IluckysheetHyperlink} - hyperlink config
         */
        generateConfigHyperlinks() {
            let rows = this.readXml.getElementsByTagName('hyperlinks/hyperlink', this.sheetFile);
            let hyperlink = {};
            for (let i = 0; i < rows.length; i++) {
                let row = rows[i];
                let attrList = row.attributeList;
                let ref = getXmlAttibute(attrList, 'ref', null), refArr = getMultiSequenceToNum(ref), _display = getXmlAttibute(attrList, 'display', null), _address = getXmlAttibute(attrList, 'location', null), _tooltip = getXmlAttibute(attrList, 'tooltip', null);
                let _type = _address ? 'internal' : 'external';
                // external hyperlink
                if (!_address) {
                    let rid = attrList['r:id'];
                    let sheetFile = this.sheetFile;
                    let relationshipList = this.readXml.getElementsByTagName('Relationships/Relationship', `xl/worksheets/_rels/${sheetFile.replace(worksheetFilePath, '')}.rels`);
                    const findRid = relationshipList?.find((e) => e.attributeList['Id'] === rid);
                    if (findRid) {
                        _address = findRid.attributeList['Target'];
                        _type = findRid.attributeList['TargetMode']?.toLocaleLowerCase();
                    }
                }
                // match R1C1
                const addressReg = new RegExp(/^.*!R([\d$])+C([\d$])*$/g);
                if (addressReg.test(_address)) {
                    _address = getTransR1C1ToSequence(_address);
                }
                // dynamically add hyperlinks
                for (const ref of refArr) {
                    hyperlink[ref] = {
                        linkAddress: _address,
                        linkTooltip: _tooltip || '',
                        linkType: _type,
                        display: _display || '',
                    };
                }
            }
            return hyperlink;
        }
    }

    let UDOC = {};
    UDOC.G = {
        concat: function (p, r) {
            for (var i = 0; i < r.cmds.length; i++)
                p.cmds.push(r.cmds[i]);
            for (var i = 0; i < r.crds.length; i++)
                p.crds.push(r.crds[i]);
        },
        getBB: function (ps) {
            var x0 = 1e99, y0 = 1e99, x1 = -x0, y1 = -y0;
            for (var i = 0; i < ps.length; i += 2) {
                var x = ps[i], y = ps[i + 1];
                if (x < x0)
                    x0 = x;
                else if (x > x1)
                    x1 = x;
                if (y < y0)
                    y0 = y;
                else if (y > y1)
                    y1 = y;
            }
            return [x0, y0, x1, y1];
        },
        rectToPath: function (r) { return { cmds: ["M", "L", "L", "L", "Z"], crds: [r[0], r[1], r[2], r[1], r[2], r[3], r[0], r[3]] }; },
        // a inside b
        insideBox: function (a, b) { return b[0] <= a[0] && b[1] <= a[1] && a[2] <= b[2] && a[3] <= b[3]; },
        isBox: function (p, bb) {
            var sameCrd8 = function (pcrd, crds) {
                for (var o = 0; o < 8; o += 2) {
                    var eq = true;
                    for (var j = 0; j < 8; j++)
                        if (Math.abs(crds[j] - pcrd[(j + o) & 7]) >= 2) {
                            eq = false;
                            break;
                        }
                    if (eq)
                        return true;
                }
                return false;
            };
            if (p.cmds.length > 10)
                return false;
            var cmds = p.cmds.join(""), crds = p.crds;
            var sameRect = false;
            if ((cmds == "MLLLZ" && crds.length == 8)
                || (cmds == "MLLLLZ" && crds.length == 10)) {
                if (crds.length == 10)
                    crds = crds.slice(0, 8);
                var x0 = bb[0], y0 = bb[1], x1 = bb[2], y1 = bb[3];
                if (!sameRect)
                    sameRect = sameCrd8(crds, [x0, y0, x1, y0, x1, y1, x0, y1]);
                if (!sameRect)
                    sameRect = sameCrd8(crds, [x0, y1, x1, y1, x1, y0, x0, y0]);
            }
            return sameRect;
        },
        boxArea: function (a) { var w = a[2] - a[0], h = a[3] - a[1]; return w * h; },
        newPath: function (gst) { gst.pth = { cmds: [], crds: [] }; },
        moveTo: function (gst, x, y) {
            var p = UDOC.M.multPoint(gst.ctm, [x, y]); //if(gst.cpos[0]==p[0] && gst.cpos[1]==p[1]) return;
            gst.pth.cmds.push("M");
            gst.pth.crds.push(p[0], p[1]);
            gst.cpos = p;
        },
        lineTo: function (gst, x, y) {
            var p = UDOC.M.multPoint(gst.ctm, [x, y]);
            if (gst.cpos[0] == p[0] && gst.cpos[1] == p[1])
                return;
            gst.pth.cmds.push("L");
            gst.pth.crds.push(p[0], p[1]);
            gst.cpos = p;
        },
        curveTo: function (gst, x1, y1, x2, y2, x3, y3) {
            var p;
            p = UDOC.M.multPoint(gst.ctm, [x1, y1]);
            x1 = p[0];
            y1 = p[1];
            p = UDOC.M.multPoint(gst.ctm, [x2, y2]);
            x2 = p[0];
            y2 = p[1];
            p = UDOC.M.multPoint(gst.ctm, [x3, y3]);
            x3 = p[0];
            y3 = p[1];
            gst.cpos = p;
            gst.pth.cmds.push("C");
            gst.pth.crds.push(x1, y1, x2, y2, x3, y3);
        },
        closePath: function (gst) { gst.pth.cmds.push("Z"); },
        arc: function (gst, x, y, r, a0, a1, neg) {
            // circle from a0 counter-clock-wise to a1
            if (neg)
                while (a1 > a0)
                    a1 -= 2 * Math.PI;
            else
                while (a1 < a0)
                    a1 += 2 * Math.PI;
            var th = (a1 - a0) / 4;
            var x0 = Math.cos(th / 2), y0 = -Math.sin(th / 2);
            var x1 = (4 - x0) / 3, y1 = y0 == 0 ? y0 : (1 - x0) * (3 - x0) / (3 * y0);
            var x2 = x1, y2 = -y1;
            var x3 = x0, y3 = -y0;
            var p1 = [x1, y1], p2 = [x2, y2], p3 = [x3, y3];
            var pth = { cmds: [(gst.pth.cmds.length == 0) ? "M" : "L", "C", "C", "C", "C"], crds: [x0, y0, x1, y1, x2, y2, x3, y3] };
            var rot = [1, 0, 0, 1, 0, 0];
            UDOC.M.rotate(rot, -th);
            for (var i = 0; i < 3; i++) {
                p1 = UDOC.M.multPoint(rot, p1);
                p2 = UDOC.M.multPoint(rot, p2);
                p3 = UDOC.M.multPoint(rot, p3);
                pth.crds.push(p1[0], p1[1], p2[0], p2[1], p3[0], p3[1]);
            }
            var sc = [r, 0, 0, r, x, y];
            UDOC.M.rotate(rot, -a0 + th / 2);
            UDOC.M.concat(rot, sc);
            UDOC.M.multArray(rot, pth.crds);
            UDOC.M.multArray(gst.ctm, pth.crds);
            UDOC.G.concat(gst.pth, pth);
            var y = pth.crds.pop();
            x = pth.crds.pop();
            gst.cpos = [x, y];
        },
        toPoly: function (p) {
            if (p.cmds[0] != "M" || p.cmds[p.cmds.length - 1] != "Z")
                return null;
            for (var i = 1; i < p.cmds.length - 1; i++)
                if (p.cmds[i] != "L")
                    return null;
            var out = [], cl = p.crds.length;
            if (p.crds[0] == p.crds[cl - 2] && p.crds[1] == p.crds[cl - 1])
                cl -= 2;
            for (var i = 0; i < cl; i += 2)
                out.push([p.crds[i], p.crds[i + 1]]);
            if (UDOC.G.polyArea(p.crds) < 0)
                out.reverse();
            return out;
        },
        fromPoly: function (p) {
            var o = { cmds: [], crds: [] };
            for (var i = 0; i < p.length; i++) {
                o.crds.push(p[i][0], p[i][1]);
                o.cmds.push(i == 0 ? "M" : "L");
            }
            o.cmds.push("Z");
            return o;
        },
        polyArea: function (p) {
            if (p.length < 6)
                return 0;
            var l = p.length - 2;
            var sum = (p[0] - p[l]) * (p[l + 1] + p[1]);
            for (var i = 0; i < l; i += 2)
                sum += (p[i + 2] - p[i]) * (p[i + 1] + p[i + 3]);
            return -sum * 0.5;
        },
        polyClip: function (p0, p1) {
            var cp1, cp2, s, e;
            var inside = function (p) {
                return (cp2[0] - cp1[0]) * (p[1] - cp1[1]) > (cp2[1] - cp1[1]) * (p[0] - cp1[0]);
            };
            var isc = function () {
                var dc = [cp1[0] - cp2[0], cp1[1] - cp2[1]], dp = [s[0] - e[0], s[1] - e[1]], n1 = cp1[0] * cp2[1] - cp1[1] * cp2[0], n2 = s[0] * e[1] - s[1] * e[0], n3 = 1.0 / (dc[0] * dp[1] - dc[1] * dp[0]);
                return [(n1 * dp[0] - n2 * dc[0]) * n3, (n1 * dp[1] - n2 * dc[1]) * n3];
            };
            var out = p0;
            cp1 = p1[p1.length - 1];
            for (let j in p1) {
                var cp2 = p1[j];
                var inp = out;
                out = [];
                s = inp[inp.length - 1]; //last on the input list
                for (let i in inp) {
                    var e = inp[i];
                    if (inside(e)) {
                        if (!inside(s)) {
                            out.push(isc());
                        }
                        out.push(e);
                    }
                    else if (inside(s)) {
                        out.push(isc());
                    }
                    s = e;
                }
                cp1 = cp2;
            }
            return out;
        }
    };
    UDOC.M = {
        getScale: function (m) { return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])); },
        translate: function (m, x, y) { UDOC.M.concat(m, [1, 0, 0, 1, x, y]); },
        rotate: function (m, a) { UDOC.M.concat(m, [Math.cos(a), -Math.sin(a), Math.sin(a), Math.cos(a), 0, 0]); },
        scale: function (m, x, y) { UDOC.M.concat(m, [x, 0, 0, y, 0, 0]); },
        concat: function (m, w) {
            var a = m[0], b = m[1], c = m[2], d = m[3], tx = m[4], ty = m[5];
            m[0] = (a * w[0]) + (b * w[2]);
            m[1] = (a * w[1]) + (b * w[3]);
            m[2] = (c * w[0]) + (d * w[2]);
            m[3] = (c * w[1]) + (d * w[3]);
            m[4] = (tx * w[0]) + (ty * w[2]) + w[4];
            m[5] = (tx * w[1]) + (ty * w[3]) + w[5];
        },
        invert: function (m) {
            var a = m[0], b = m[1], c = m[2], d = m[3], tx = m[4], ty = m[5], adbc = a * d - b * c;
            m[0] = d / adbc;
            m[1] = -b / adbc;
            m[2] = -c / adbc;
            m[3] = a / adbc;
            m[4] = (c * ty - d * tx) / adbc;
            m[5] = (b * tx - a * ty) / adbc;
        },
        multPoint: function (m, p) { var x = p[0], y = p[1]; return [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]]; },
        multArray: function (m, a) { for (var i = 0; i < a.length; i += 2) {
            var x = a[i], y = a[i + 1];
            a[i] = x * m[0] + y * m[2] + m[4];
            a[i + 1] = x * m[1] + y * m[3] + m[5];
        } }
    };
    UDOC.C = {
        srgbGamma: function (x) { return x < 0.0031308 ? 12.92 * x : 1.055 * Math.pow(x, 1.0 / 2.4) - 0.055; },
        cmykToRgb: function (clr) {
            var c = clr[0], m = clr[1], y = clr[2], k = clr[3];
            // return [1-Math.min(1,c+k), 1-Math.min(1, m+k), 1-Math.min(1,y+k)];
            var r = 255
                + c * (-4.387332384609988 * c + 54.48615194189176 * m + 18.82290502165302 * y + 212.25662451639585 * k + -285.2331026137004)
                + m * (1.7149763477362134 * m - 5.6096736904047315 * y + -17.873870861415444 * k - 5.497006427196366)
                + y * (-2.5217340131683033 * y - 21.248923337353073 * k + 17.5119270841813)
                + k * (-21.86122147463605 * k - 189.48180835922747);
            var g = 255
                + c * (8.841041422036149 * c + 60.118027045597366 * m + 6.871425592049007 * y + 31.159100130055922 * k + -79.2970844816548)
                + m * (-15.310361306967817 * m + 17.575251261109482 * y + 131.35250912493976 * k - 190.9453302588951)
                + y * (4.444339102852739 * y + 9.8632861493405 * k - 24.86741582555878)
                + k * (-20.737325471181034 * k - 187.80453709719578);
            var b = 255
                + c * (0.8842522430003296 * c + 8.078677503112928 * m + 30.89978309703729 * y - 0.23883238689178934 * k + -14.183576799673286)
                + m * (10.49593273432072 * m + 63.02378494754052 * y + 50.606957656360734 * k - 112.23884253719248)
                + y * (0.03296041114873217 * y + 115.60384449646641 * k + -193.58209356861505)
                + k * (-22.33816807309886 * k - 180.12613974708367);
            return [Math.max(0, Math.min(1, r / 255)), Math.max(0, Math.min(1, g / 255)), Math.max(0, Math.min(1, b / 255))];
            //var iK = 1-c[3];  
            //return [(1-c[0])*iK, (1-c[1])*iK, (1-c[2])*iK];  
        },
        labToRgb: function (lab) {
            var k = 903.3, e = 0.008856, L = lab[0], a = lab[1], b = lab[2];
            var fy = (L + 16) / 116, fy3 = fy * fy * fy;
            var fz = fy - b / 200, fz3 = fz * fz * fz;
            var fx = a / 500 + fy, fx3 = fx * fx * fx;
            var zr = fz3 > e ? fz3 : (116 * fz - 16) / k;
            var yr = fy3 > e ? fy3 : (116 * fy - 16) / k;
            var xr = fx3 > e ? fx3 : (116 * fx - 16) / k;
            var X = xr * 96.72, Y = yr * 100, Z = zr * 81.427, xyz = [X / 100, Y / 100, Z / 100];
            var x2s = [3.1338561, -1.6168667, -0.4906146, -0.9787684, 1.9161415, 0.0334540, 0.0719453, -0.2289914, 1.4052427];
            var rgb = [x2s[0] * xyz[0] + x2s[1] * xyz[1] + x2s[2] * xyz[2],
                x2s[3] * xyz[0] + x2s[4] * xyz[1] + x2s[5] * xyz[2],
                x2s[6] * xyz[0] + x2s[7] * xyz[1] + x2s[8] * xyz[2]];
            for (var i = 0; i < 3; i++)
                rgb[i] = Math.max(0, Math.min(1, UDOC.C.srgbGamma(rgb[i])));
            return rgb;
        }
    };
    UDOC.getState = function (crds) {
        return {
            font: UDOC.getFont(),
            dd: { flat: 1 }, // device-dependent
            space: "/DeviceGray",
            // fill
            ca: 1,
            colr: [0, 0, 0],
            sspace: "/DeviceGray",
            // stroke
            CA: 1,
            COLR: [0, 0, 0],
            bmode: "/Normal",
            SA: false, OPM: 0, AIS: false, OP: false, op: false, SMask: "/None",
            lwidth: 1,
            lcap: 0,
            ljoin: 0,
            mlimit: 10,
            SM: 0.1,
            doff: 0,
            dash: [],
            ctm: [1, 0, 0, 1, 0, 0],
            cpos: [0, 0],
            pth: { cmds: [], crds: [] },
            cpth: crds ? UDOC.G.rectToPath(crds) : null // clipping path
        };
    };
    UDOC.getFont = function () {
        return {
            Tc: 0, // character spacing
            Tw: 0, // word spacing
            Th: 100, // horizontal scale
            Tl: 0, // leading
            Tf: "Helvetica-Bold",
            Tfs: 1, // font size
            Tmode: 0, // rendering mode
            Trise: 0, // rise
            Tk: 0, // knockout
            Tal: 0, // align, 0: left, 1: right, 2: center
            Tun: 0, // 0: no, 1: underline
            Tm: [1, 0, 0, 1, 0, 0],
            Tlm: [1, 0, 0, 1, 0, 0],
            Trm: [1, 0, 0, 1, 0, 0]
        };
    };
    let FromEMF = function () {
    };
    FromEMF.Parse = function (buff, genv) {
        buff = new Uint8Array(buff);
        var off = 0;
        //console.log(buff.slice(0,32));
        var prms = { fill: false, strk: false, bb: [0, 0, 1, 1], wbb: [0, 0, 1, 1], fnt: { nam: "Arial", hgh: 25, und: false, orn: 0 }, tclr: [0, 0, 0], talg: 0 }, gst, tab = [], sts = [];
        var rI = FromEMF.B.readShort, rU = FromEMF.B.readUshort, rI32 = FromEMF.B.readInt, rU32 = FromEMF.B.readUint, rF32 = FromEMF.B.readFloat;
        while (true) {
            var fnc = rU32(buff, off);
            off += 4;
            var fnm = FromEMF.K[fnc];
            var siz = rU32(buff, off);
            off += 4;
            //if(gst && isNaN(gst.ctm[0])) throw "e";
            //console.log(fnc,fnm,siz);
            var loff = off;
            //if(opn++==253) break;
            var obj = null, oid = 0;
            //console.log(fnm, siz);
            if (fnm == "EOF") {
                break;
            }
            else if (fnm == "HEADER") {
                prms.bb = FromEMF._readBox(buff, loff);
                loff += 16; //console.log(fnm, prms.bb);
                genv.StartPage(prms.bb[0], prms.bb[1], prms.bb[2], prms.bb[3]);
                gst = UDOC.getState(prms.bb);
            }
            else if (fnm == "SAVEDC")
                sts.push(JSON.stringify(gst), JSON.stringify(prms));
            else if (fnm == "RESTOREDC") {
                var dif = rI32(buff, loff);
                loff += 4;
                while (dif < -1) {
                    sts.pop();
                    sts.pop();
                }
                prms = JSON.parse(sts.pop());
                gst = JSON.parse(sts.pop());
            }
            else if (fnm == "SELECTCLIPPATH") {
                gst.cpth = JSON.parse(JSON.stringify(gst.pth));
            }
            else if (["SETMAPMODE", "SETPOLYFILLMODE", "SETBKMODE" /*,"SETVIEWPORTEXTEX"*/, "SETICMMODE", "SETROP2", "EXTSELECTCLIPRGN"].indexOf(fnm) != -1) ;
            //else if(fnm=="INTERSECTCLIPRECT") {  var r=prms.crct=FromEMF._readBox(buff, loff);  /*var y0=r[1],y1=r[3]; if(y0>y1){r[1]=y1; r[3]=y0;}*/ console.log(prms.crct);  }
            else if (fnm == "SETMITERLIMIT")
                gst.mlimit = rU32(buff, loff);
            else if (fnm == "SETTEXTCOLOR")
                prms.tclr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
            else if (fnm == "SETTEXTALIGN")
                prms.talg = rU32(buff, loff);
            else if (fnm == "SETVIEWPORTEXTEX" || fnm == "SETVIEWPORTORGEX") {
                if (prms.vbb == null)
                    prms.vbb = [];
                var coff = fnm == "SETVIEWPORTORGEX" ? 0 : 2;
                prms.vbb[coff] = rI32(buff, loff);
                loff += 4;
                prms.vbb[coff + 1] = rI32(buff, loff);
                loff += 4;
                //console.log(prms.vbb);
                if (fnm == "SETVIEWPORTEXTEX")
                    FromEMF._updateCtm(prms, gst);
            }
            else if (fnm == "SETWINDOWEXTEX" || fnm == "SETWINDOWORGEX") {
                var coff = fnm == "SETWINDOWORGEX" ? 0 : 2;
                prms.wbb[coff] = rI32(buff, loff);
                loff += 4;
                prms.wbb[coff + 1] = rI32(buff, loff);
                loff += 4;
                if (fnm == "SETWINDOWEXTEX")
                    FromEMF._updateCtm(prms, gst);
            }
            //else if(fnm=="SETMETARGN") {}
            else if (fnm == "COMMENT") {
                rU32(buff, loff);
                loff += 4;
            }
            else if (fnm == "SELECTOBJECT") {
                var ind = rU32(buff, loff);
                loff += 4;
                //console.log(ind.toString(16), tab, tab[ind]);
                if (ind == 0x80000000) {
                    prms.fill = true;
                    gst.colr = [1, 1, 1];
                } // white brush
                else if (ind == 0x80000005) {
                    prms.fill = false;
                } // null brush
                else if (ind == 0x80000007) {
                    prms.strk = true;
                    prms.lwidth = 1;
                    gst.COLR = [0, 0, 0];
                } // black pen
                else if (ind == 0x80000008) {
                    prms.strk = false;
                } // null  pen
                else if (ind == 0x8000000d) ; // system font
                else if (ind == 0x8000000e) ; // device default font
                else {
                    var co = tab[ind]; //console.log(ind, co);
                    if (co.t == "b") {
                        prms.fill = co.stl != 1;
                        if (co.stl == 0) ;
                        else if (co.stl == 1) ;
                        else
                            throw co.stl + " e";
                        gst.colr = co.clr;
                    }
                    else if (co.t == "p") {
                        prms.strk = co.stl != 5;
                        gst.lwidth = co.wid;
                        gst.COLR = co.clr;
                    }
                    else if (co.t == "f") {
                        prms.fnt = co;
                        gst.font.Tf = co.nam;
                        gst.font.Tfs = Math.abs(co.hgh);
                        gst.font.Tun = co.und;
                    }
                    else
                        throw "e";
                }
            }
            else if (fnm == "DELETEOBJECT") {
                var ind = rU32(buff, loff);
                loff += 4;
                if (tab[ind] != null)
                    tab[ind] = null;
                else
                    throw "e";
            }
            else if (fnm == "CREATEBRUSHINDIRECT") {
                oid = rU32(buff, loff);
                loff += 4;
                obj = { t: "b" };
                obj.stl = rU32(buff, loff);
                loff += 4;
                obj.clr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
                loff += 4;
                obj.htc = rU32(buff, loff);
                loff += 4;
                //console.log(oid, obj);
            }
            else if (fnm == "CREATEPEN" || fnm == "EXTCREATEPEN") {
                oid = rU32(buff, loff);
                loff += 4;
                obj = { t: "p" };
                if (fnm == "EXTCREATEPEN") {
                    loff += 16;
                    obj.stl = rU32(buff, loff);
                    loff += 4;
                    obj.wid = rU32(buff, loff);
                    loff += 4;
                    //obj.stl = rU32(buff, loff);  
                    loff += 4;
                }
                else {
                    obj.stl = rU32(buff, loff);
                    loff += 4;
                    obj.wid = rU32(buff, loff);
                    loff += 4;
                    loff += 4;
                }
                obj.clr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
                loff += 4;
            }
            else if (fnm == "EXTCREATEFONTINDIRECTW") {
                oid = rU32(buff, loff);
                loff += 4;
                obj = { t: "f", nam: "" };
                obj.hgh = rI32(buff, loff);
                loff += 4;
                loff += 4 * 2;
                obj.orn = rI32(buff, loff) / 10;
                loff += 4;
                var wgh = rU32(buff, loff);
                loff += 4; //console.log(fnm, obj.orn, wgh);
                //console.log(rU32(buff,loff), rU32(buff,loff+4), buff.slice(loff,loff+8));
                obj.und = buff[loff + 1];
                obj.stk = buff[loff + 2];
                loff += 4 * 2;
                while (rU(buff, loff) != 0) {
                    obj.nam += String.fromCharCode(rU(buff, loff));
                    loff += 2;
                }
                if (wgh > 500)
                    obj.nam += "-Bold";
                //console.log(wgh, obj.nam);
            }
            else if (fnm == "EXTTEXTOUTW") {
                //console.log(buff.slice(loff-8, loff-8+siz));
                loff += 16;
                var mod = rU32(buff, loff);
                loff += 4; //console.log(mod);
                rF32(buff, loff);
                loff += 4;
                rF32(buff, loff);
                loff += 4;
                var rfx = rI32(buff, loff);
                loff += 4;
                var rfy = rI32(buff, loff);
                loff += 4;
                //console.log(mod, scx, scy,rfx,rfy);
                gst.font.Tm = [1, 0, 0, -1, 0, 0];
                UDOC.M.rotate(gst.font.Tm, prms.fnt.orn * Math.PI / 180);
                UDOC.M.translate(gst.font.Tm, rfx, rfy);
                var alg = prms.talg; //console.log(alg.toString(2));
                if ((alg & 6) == 6)
                    gst.font.Tal = 2;
                else if ((alg & 7) == 0)
                    gst.font.Tal = 0;
                else
                    throw alg + " e";
                if ((alg & 24) == 24) ; // baseline
                else if ((alg & 24) == 0)
                    UDOC.M.translate(gst.font.Tm, 0, gst.font.Tfs);
                else
                    throw "e";
                var crs = rU32(buff, loff);
                loff += 4;
                var ofs = rU32(buff, loff);
                loff += 4;
                rU32(buff, loff);
                loff += 4; //if(ops!=0) throw "e";
                //console.log(ofs,ops,crs);
                loff += 16;
                rU32(buff, loff);
                loff += 4; //console.log(ops, ofD, loff, ofs+off-8);
                ofs += off - 8; //console.log(crs, ops);
                var str = "";
                for (var i = 0; i < crs; i++) {
                    var cc = rU(buff, ofs + i * 2);
                    str += String.fromCharCode(cc);
                }
                var oclr = gst.colr;
                gst.colr = prms.tclr;
                //console.log(str, gst.colr, gst.font.Tm);
                //var otfs = gst.font.Tfs;  gst.font.Tfs *= 1/gst.ctm[0];
                genv.PutText(gst, str, str.length * gst.font.Tfs * 0.5);
                gst.colr = oclr;
                //gst.font.Tfs = otfs;
                //console.log(rfx, rfy, scx, ops, rcX, rcY, rcW, rcH, offDx, str);
            }
            else if (fnm == "BEGINPATH") {
                UDOC.G.newPath(gst);
            }
            else if (fnm == "ENDPATH") ;
            else if (fnm == "CLOSEFIGURE")
                UDOC.G.closePath(gst);
            else if (fnm == "MOVETOEX") {
                UDOC.G.moveTo(gst, rI32(buff, loff), rI32(buff, loff + 4));
            }
            else if (fnm == "LINETO") {
                if (gst.pth.cmds.length == 0) {
                    var im = gst.ctm.slice(0);
                    UDOC.M.invert(im);
                    var p = UDOC.M.multPoint(im, gst.cpos);
                    UDOC.G.moveTo(gst, p[0], p[1]);
                }
                UDOC.G.lineTo(gst, rI32(buff, loff), rI32(buff, loff + 4));
            }
            else if (fnm == "POLYGON" || fnm == "POLYGON16" || fnm == "POLYLINE" || fnm == "POLYLINE16" || fnm == "POLYLINETO" || fnm == "POLYLINETO16") {
                loff += 16;
                var ndf = fnm.startsWith("POLYGON"), isTo = fnm.indexOf("TO") != -1;
                var cnt = rU32(buff, loff);
                loff += 4;
                if (!isTo)
                    UDOC.G.newPath(gst);
                loff = FromEMF._drawPoly(buff, loff, cnt, gst, fnm.endsWith("16") ? 2 : 4, ndf, isTo);
                if (!isTo)
                    FromEMF._draw(genv, gst, prms, ndf);
                //console.log(prms, gst.lwidth);
                //console.log(JSON.parse(JSON.stringify(gst.pth)));
            }
            else if (fnm == "POLYPOLYGON16") {
                loff += 16;
                var ndf = fnm.startsWith("POLYPOLYGON"), isTo = fnm.indexOf("TO") != -1;
                var nop = rU32(buff, loff);
                loff += 4;
                loff += 4;
                var pi = loff;
                loff += nop * 4;
                if (!isTo)
                    UDOC.G.newPath(gst);
                for (var i = 0; i < nop; i++) {
                    var ppp = rU(buff, pi + i * 4);
                    loff = FromEMF._drawPoly(buff, loff, ppp, gst, fnm.endsWith("16") ? 2 : 4, ndf, isTo);
                }
                if (!isTo)
                    FromEMF._draw(genv, gst, prms, ndf);
            }
            else if (fnm == "POLYBEZIER" || fnm == "POLYBEZIER16" || fnm == "POLYBEZIERTO" || fnm == "POLYBEZIERTO16") {
                loff += 16;
                var is16 = fnm.endsWith("16"), rC = is16 ? rI : rI32, nl = is16 ? 2 : 4;
                var cnt = rU32(buff, loff);
                loff += 4;
                if (fnm.indexOf("TO") == -1) {
                    UDOC.G.moveTo(gst, rC(buff, loff), rC(buff, loff + nl));
                    loff += 2 * nl;
                    cnt--;
                }
                while (cnt > 0) {
                    UDOC.G.curveTo(gst, rC(buff, loff), rC(buff, loff + nl), rC(buff, loff + 2 * nl), rC(buff, loff + 3 * nl), rC(buff, loff + 4 * nl), rC(buff, loff + 5 * nl));
                    loff += 6 * nl;
                    cnt -= 3;
                }
                //console.log(JSON.parse(JSON.stringify(gst.pth)));
            }
            else if (fnm == "RECTANGLE" || fnm == "ELLIPSE") {
                UDOC.G.newPath(gst);
                var bx = FromEMF._readBox(buff, loff);
                if (fnm == "RECTANGLE") {
                    UDOC.G.moveTo(gst, bx[0], bx[1]);
                    UDOC.G.lineTo(gst, bx[2], bx[1]);
                    UDOC.G.lineTo(gst, bx[2], bx[3]);
                    UDOC.G.lineTo(gst, bx[0], bx[3]);
                }
                else {
                    var x = (bx[0] + bx[2]) / 2, y = (bx[1] + bx[3]) / 2;
                    UDOC.G.arc(gst, x, y, (bx[2] - bx[0]) / 2, 0, 2 * Math.PI, false);
                }
                UDOC.G.closePath(gst);
                FromEMF._draw(genv, gst, prms, true);
                //console.log(prms, gst.lwidth);
            }
            else if (fnm == "FILLPATH")
                genv.Fill(gst, false);
            else if (fnm == "STROKEPATH")
                genv.Stroke(gst);
            else if (fnm == "STROKEANDFILLPATH") {
                genv.Fill(gst, false);
                genv.Stroke(gst);
            }
            else if (fnm == "SETWORLDTRANSFORM" || fnm == "MODIFYWORLDTRANSFORM") {
                var mat = [];
                for (var i = 0; i < 6; i++)
                    mat.push(rF32(buff, loff + i * 4));
                loff += 24;
                //console.log(fnm, gst.ctm.slice(0), mat);
                if (fnm == "SETWORLDTRANSFORM")
                    gst.ctm = mat;
                else {
                    var mod = rU32(buff, loff);
                    loff += 4;
                    if (mod == 2) {
                        var om = gst.ctm;
                        gst.ctm = mat;
                        UDOC.M.concat(gst.ctm, om);
                    }
                    else
                        throw "e";
                }
            }
            else if (fnm == "SETSTRETCHBLTMODE") {
                rU32(buff, loff);
                loff += 4;
            }
            else if (fnm == "STRETCHDIBITS") {
                var bx = FromEMF._readBox(buff, loff);
                loff += 16;
                var xD = rI32(buff, loff);
                loff += 4;
                var yD = rI32(buff, loff);
                loff += 4;
                rI32(buff, loff);
                loff += 4;
                rI32(buff, loff);
                loff += 4;
                var wS = rI32(buff, loff);
                loff += 4;
                var hS = rI32(buff, loff);
                loff += 4;
                var ofH = rU32(buff, loff) + off - 8;
                loff += 4;
                rU32(buff, loff);
                loff += 4;
                var ofB = rU32(buff, loff) + off - 8;
                loff += 4;
                rU32(buff, loff);
                loff += 4;
                var usg = rU32(buff, loff);
                loff += 4;
                if (usg != 0)
                    throw "e";
                rU32(buff, loff);
                loff += 4;
                var wD = rI32(buff, loff);
                loff += 4;
                var hD = rI32(buff, loff);
                loff += 4; //console.log(bop, wD, hD);
                //console.log(ofH, szH, ofB, szB, ofH+40);
                //console.log(bx, xD,yD,wD,hD);
                //console.log(xS,yS,wS,hS);
                //console.log(ofH,szH,ofB,szB,usg,bop);
                rU32(buff, ofH);
                ofH += 4;
                var w = rU32(buff, ofH);
                ofH += 4;
                var h = rU32(buff, ofH);
                ofH += 4;
                if (w != wS || h != hS)
                    throw "e";
                var ps = rU(buff, ofH);
                ofH += 2;
                var bc = rU(buff, ofH);
                ofH += 2;
                if (bc != 8 && bc != 24 && bc != 32)
                    throw bc + " e";
                var cpr = rU32(buff, ofH);
                ofH += 4;
                if (cpr != 0)
                    throw cpr + " e";
                rU32(buff, ofH);
                ofH += 4;
                rU32(buff, ofH);
                ofH += 4;
                rU32(buff, ofH);
                ofH += 4;
                rU32(buff, ofH);
                ofH += 4;
                rU32(buff, ofH);
                ofH += 4; //console.log(hl, w, h, ps, bc, cpr, sz, xpm, ypm, cu, ci);
                //console.log(hl,w,h,",",xS,yS,wS,hS,",",xD,yD,wD,hD,",",xpm,ypm);
                var rl = Math.floor(((w * ps * bc + 31) & ~31) / 8);
                var img = new Uint8Array(w * h * 4);
                if (bc == 8) {
                    for (var y = 0; y < h; y++)
                        for (var x = 0; x < w; x++) {
                            var qi = (y * w + x) << 2, ind = buff[ofB + (h - 1 - y) * rl + x] << 2;
                            img[qi] = buff[ofH + ind + 2];
                            img[qi + 1] = buff[ofH + ind + 1];
                            img[qi + 2] = buff[ofH + ind + 0];
                            img[qi + 3] = 255;
                        }
                }
                if (bc == 24) {
                    for (var y = 0; y < h; y++)
                        for (var x = 0; x < w; x++) {
                            var qi = (y * w + x) << 2, ti = ofB + (h - 1 - y) * rl + x * 3;
                            img[qi] = buff[ti + 2];
                            img[qi + 1] = buff[ti + 1];
                            img[qi + 2] = buff[ti + 0];
                            img[qi + 3] = 255;
                        }
                }
                if (bc == 32) {
                    for (var y = 0; y < h; y++)
                        for (var x = 0; x < w; x++) {
                            var qi = (y * w + x) << 2, ti = ofB + (h - 1 - y) * rl + x * 4;
                            img[qi] = buff[ti + 2];
                            img[qi + 1] = buff[ti + 1];
                            img[qi + 2] = buff[ti + 0];
                            img[qi + 3] = buff[ti + 3];
                        }
                }
                var ctm = gst.ctm.slice(0);
                gst.ctm = [1, 0, 0, 1, 0, 0];
                UDOC.M.scale(gst.ctm, wD, -hD);
                UDOC.M.translate(gst.ctm, xD, yD + hD);
                UDOC.M.concat(gst.ctm, ctm);
                genv.PutImage(gst, img, w, h);
                gst.ctm = ctm;
            }
            else {
                console.log(fnm, siz);
            }
            if (obj != null)
                tab[oid] = obj;
            off += siz - 8;
        }
        //genv.Stroke(gst);
        genv.ShowPage();
        genv.Done();
    };
    FromEMF._readBox = function (buff, off) { var b = []; for (var i = 0; i < 4; i++)
        b[i] = FromEMF.B.readInt(buff, off + i * 4); return b; };
    FromEMF._updateCtm = function (prms, gst) {
        var mat = [1, 0, 0, 1, 0, 0];
        var wbb = prms.wbb; prms.bb; var vbb = (prms.vbb && prms.vbb.length == 4) ? prms.vbb : prms.bb;
        //var y0 = bb[1], y1 = bb[3];  bb[1]=Math.min(y0,y1);  bb[3]=Math.max(y0,y1);
        UDOC.M.translate(mat, -wbb[0], -wbb[1]);
        UDOC.M.scale(mat, 1 / wbb[2], 1 / wbb[3]);
        UDOC.M.scale(mat, vbb[2], vbb[3]);
        //UDOC.M.scale(mat, vbb[2]/(bb[2]-bb[0]), vbb[3]/(bb[3]-bb[1]));
        //UDOC.M.scale(mat, bb[2]-bb[0],bb[3]-bb[1]);
        gst.ctm = mat;
    };
    FromEMF._draw = function (genv, gst, prms, needFill) {
        if (prms.fill && needFill)
            genv.Fill(gst, false);
        if (prms.strk && gst.lwidth != 0)
            genv.Stroke(gst);
    };
    FromEMF._drawPoly = function (buff, off, ppp, gst, nl, clos, justLine) {
        var rS = nl == 2 ? FromEMF.B.readShort : FromEMF.B.readInt;
        for (var j = 0; j < ppp; j++) {
            var px = rS(buff, off);
            off += nl;
            var py = rS(buff, off);
            off += nl;
            if (j == 0 && !justLine)
                UDOC.G.moveTo(gst, px, py);
            else
                UDOC.G.lineTo(gst, px, py);
        }
        if (clos)
            UDOC.G.closePath(gst);
        return off;
    };
    FromEMF.B = {
        uint8: new Uint8Array(4),
        readShort: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; return FromEMF.B.int16[0]; },
        readUshort: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; return FromEMF.B.uint16[0]; },
        readInt: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; u8[2] = buff[p + 2]; u8[3] = buff[p + 3]; return FromEMF.B.int32[0]; },
        readUint: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; u8[2] = buff[p + 2]; u8[3] = buff[p + 3]; return FromEMF.B.uint32[0]; },
        readFloat: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; u8[2] = buff[p + 2]; u8[3] = buff[p + 3]; return FromEMF.B.flot32[0]; },
        readASCII: function (buff, p, l) { var s = ""; for (var i = 0; i < l; i++)
            s += String.fromCharCode(buff[p + i]); return s; }
    };
    FromEMF.B.int16 = new Int16Array(FromEMF.B.uint8.buffer);
    FromEMF.B.uint16 = new Uint16Array(FromEMF.B.uint8.buffer);
    FromEMF.B.int32 = new Int32Array(FromEMF.B.uint8.buffer);
    FromEMF.B.uint32 = new Uint32Array(FromEMF.B.uint8.buffer);
    FromEMF.B.flot32 = new Float32Array(FromEMF.B.uint8.buffer);
    FromEMF.C = {
        EMR_HEADER: 0x00000001,
        EMR_POLYBEZIER: 0x00000002,
        EMR_POLYGON: 0x00000003,
        EMR_POLYLINE: 0x00000004,
        EMR_POLYBEZIERTO: 0x00000005,
        EMR_POLYLINETO: 0x00000006,
        EMR_POLYPOLYLINE: 0x00000007,
        EMR_POLYPOLYGON: 0x00000008,
        EMR_SETWINDOWEXTEX: 0x00000009,
        EMR_SETWINDOWORGEX: 0x0000000A,
        EMR_SETVIEWPORTEXTEX: 0x0000000B,
        EMR_SETVIEWPORTORGEX: 0x0000000C,
        EMR_SETBRUSHORGEX: 0x0000000D,
        EMR_EOF: 0x0000000E,
        EMR_SETPIXELV: 0x0000000F,
        EMR_SETMAPPERFLAGS: 0x00000010,
        EMR_SETMAPMODE: 0x00000011,
        EMR_SETBKMODE: 0x00000012,
        EMR_SETPOLYFILLMODE: 0x00000013,
        EMR_SETROP2: 0x00000014,
        EMR_SETSTRETCHBLTMODE: 0x00000015,
        EMR_SETTEXTALIGN: 0x00000016,
        EMR_SETCOLORADJUSTMENT: 0x00000017,
        EMR_SETTEXTCOLOR: 0x00000018,
        EMR_SETBKCOLOR: 0x00000019,
        EMR_OFFSETCLIPRGN: 0x0000001A,
        EMR_MOVETOEX: 0x0000001B,
        EMR_SETMETARGN: 0x0000001C,
        EMR_EXCLUDECLIPRECT: 0x0000001D,
        EMR_INTERSECTCLIPRECT: 0x0000001E,
        EMR_SCALEVIEWPORTEXTEX: 0x0000001F,
        EMR_SCALEWINDOWEXTEX: 0x00000020,
        EMR_SAVEDC: 0x00000021,
        EMR_RESTOREDC: 0x00000022,
        EMR_SETWORLDTRANSFORM: 0x00000023,
        EMR_MODIFYWORLDTRANSFORM: 0x00000024,
        EMR_SELECTOBJECT: 0x00000025,
        EMR_CREATEPEN: 0x00000026,
        EMR_CREATEBRUSHINDIRECT: 0x00000027,
        EMR_DELETEOBJECT: 0x00000028,
        EMR_ANGLEARC: 0x00000029,
        EMR_ELLIPSE: 0x0000002A,
        EMR_RECTANGLE: 0x0000002B,
        EMR_ROUNDRECT: 0x0000002C,
        EMR_ARC: 0x0000002D,
        EMR_CHORD: 0x0000002E,
        EMR_PIE: 0x0000002F,
        EMR_SELECTPALETTE: 0x00000030,
        EMR_CREATEPALETTE: 0x00000031,
        EMR_SETPALETTEENTRIES: 0x00000032,
        EMR_RESIZEPALETTE: 0x00000033,
        EMR_REALIZEPALETTE: 0x00000034,
        EMR_EXTFLOODFILL: 0x00000035,
        EMR_LINETO: 0x00000036,
        EMR_ARCTO: 0x00000037,
        EMR_POLYDRAW: 0x00000038,
        EMR_SETARCDIRECTION: 0x00000039,
        EMR_SETMITERLIMIT: 0x0000003A,
        EMR_BEGINPATH: 0x0000003B,
        EMR_ENDPATH: 0x0000003C,
        EMR_CLOSEFIGURE: 0x0000003D,
        EMR_FILLPATH: 0x0000003E,
        EMR_STROKEANDFILLPATH: 0x0000003F,
        EMR_STROKEPATH: 0x00000040,
        EMR_FLATTENPATH: 0x00000041,
        EMR_WIDENPATH: 0x00000042,
        EMR_SELECTCLIPPATH: 0x00000043,
        EMR_ABORTPATH: 0x00000044,
        EMR_COMMENT: 0x00000046,
        EMR_FILLRGN: 0x00000047,
        EMR_FRAMERGN: 0x00000048,
        EMR_INVERTRGN: 0x00000049,
        EMR_PAINTRGN: 0x0000004A,
        EMR_EXTSELECTCLIPRGN: 0x0000004B,
        EMR_BITBLT: 0x0000004C,
        EMR_STRETCHBLT: 0x0000004D,
        EMR_MASKBLT: 0x0000004E,
        EMR_PLGBLT: 0x0000004F,
        EMR_SETDIBITSTODEVICE: 0x00000050,
        EMR_STRETCHDIBITS: 0x00000051,
        EMR_EXTCREATEFONTINDIRECTW: 0x00000052,
        EMR_EXTTEXTOUTA: 0x00000053,
        EMR_EXTTEXTOUTW: 0x00000054,
        EMR_POLYBEZIER16: 0x00000055,
        EMR_POLYGON16: 0x00000056,
        EMR_POLYLINE16: 0x00000057,
        EMR_POLYBEZIERTO16: 0x00000058,
        EMR_POLYLINETO16: 0x00000059,
        EMR_POLYPOLYLINE16: 0x0000005A,
        EMR_POLYPOLYGON16: 0x0000005B,
        EMR_POLYDRAW16: 0x0000005C,
        EMR_CREATEMONOBRUSH: 0x0000005D,
        EMR_CREATEDIBPATTERNBRUSHPT: 0x0000005E,
        EMR_EXTCREATEPEN: 0x0000005F,
        EMR_POLYTEXTOUTA: 0x00000060,
        EMR_POLYTEXTOUTW: 0x00000061,
        EMR_SETICMMODE: 0x00000062,
        EMR_CREATECOLORSPACE: 0x00000063,
        EMR_SETCOLORSPACE: 0x00000064,
        EMR_DELETECOLORSPACE: 0x00000065,
        EMR_GLSRECORD: 0x00000066,
        EMR_GLSBOUNDEDRECORD: 0x00000067,
        EMR_PIXELFORMAT: 0x00000068,
        EMR_DRAWESCAPE: 0x00000069,
        EMR_EXTESCAPE: 0x0000006A,
        EMR_SMALLTEXTOUT: 0x0000006C,
        EMR_FORCEUFIMAPPING: 0x0000006D,
        EMR_NAMEDESCAPE: 0x0000006E,
        EMR_COLORCORRECTPALETTE: 0x0000006F,
        EMR_SETICMPROFILEA: 0x00000070,
        EMR_SETICMPROFILEW: 0x00000071,
        EMR_ALPHABLEND: 0x00000072,
        EMR_SETLAYOUT: 0x00000073,
        EMR_TRANSPARENTBLT: 0x00000074,
        EMR_GRADIENTFILL: 0x00000076,
        EMR_SETLINKEDUFIS: 0x00000077,
        EMR_SETTEXTJUSTIFICATION: 0x00000078,
        EMR_COLORMATCHTOTARGETW: 0x00000079,
        EMR_CREATECOLORSPACEW: 0x0000007A
    };
    FromEMF.K = [];
    // (function() {
    //     var inp, out, stt;
    //     inp = FromEMF.C;   out = FromEMF.K;   stt=4;
    //     for(var p in inp) out[inp[p]] = p.slice(stt);
    // }  )();
    let ToContext2D = function (needPage, scale) {
        this.canvas = document.createElement("canvas");
        this.ctx = this.canvas.getContext("2d");
        this.bb = null;
        this.currPage = 0;
        this.needPage = needPage;
        this.scale = scale;
    };
    ToContext2D.prototype.StartPage = function (x, y, w, h) {
        if (this.currPage != this.needPage)
            return;
        this.bb = [x, y, w, h];
        var scl = this.scale, dpr = window.devicePixelRatio;
        var cnv = this.canvas, ctx = this.ctx;
        cnv.width = Math.round(w * scl);
        cnv.height = Math.round(h * scl);
        ctx.translate(0, h * scl);
        ctx.scale(scl, -scl);
        cnv.setAttribute("style", "border:1px solid; width:" + (cnv.width / dpr) + "px; height:" + (cnv.height / dpr) + "px");
    };
    ToContext2D.prototype.Fill = function (gst, evenOdd) {
        if (this.currPage != this.needPage)
            return;
        var ctx = this.ctx;
        ctx.beginPath();
        this._setStyle(gst, ctx);
        this._draw(gst.pth, ctx);
        ctx.fill();
    };
    ToContext2D.prototype.Stroke = function (gst) {
        if (this.currPage != this.needPage)
            return;
        var ctx = this.ctx;
        ctx.beginPath();
        this._setStyle(gst, ctx);
        this._draw(gst.pth, ctx);
        ctx.stroke();
    };
    ToContext2D.prototype.PutText = function (gst, str, stw) {
        if (this.currPage != this.needPage)
            return;
        this._scale(gst.ctm);
        var ctx = this.ctx;
        this._setStyle(gst, ctx);
        ctx.save();
        var m = [1, 0, 0, -1, 0, 0];
        this._concat(m, gst.font.Tm);
        this._concat(m, gst.ctm);
        //console.log(str, m, gst);  throw "e";
        ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
        ctx.fillText(str, 0, 0);
        ctx.restore();
    };
    ToContext2D.prototype.PutImage = function (gst, buff, w, h, msk) {
        if (this.currPage != this.needPage)
            return;
        var ctx = this.ctx;
        if (buff.length == w * h * 4) {
            buff = buff.slice(0);
            if (msk && msk.length == w * h * 4)
                for (var i = 0; i < buff.length; i += 4)
                    buff[i + 3] = msk[i + 1];
            var cnv = document.createElement("canvas"), cctx = cnv.getContext("2d");
            cnv.width = w;
            cnv.height = h;
            var imgd = cctx.createImageData(w, h);
            for (var i = 0; i < buff.length; i++)
                imgd.data[i] = buff[i];
            cctx.putImageData(imgd, 0, 0);
            ctx.save();
            var m = [1, 0, 0, 1, 0, 0];
            this._concat(m, [1 / w, 0, 0, -1 / h, 0, 1]);
            this._concat(m, gst.ctm);
            ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
            ctx.drawImage(cnv, 0, 0);
            ctx.restore();
        }
    };
    ToContext2D.prototype.ShowPage = function () { this.currPage++; };
    ToContext2D.prototype.Done = function () { };
    function _flt(n) { return "" + parseFloat(n.toFixed(2)); }
    ToContext2D.prototype._setStyle = function (gst, ctx) {
        var scl = this._scale(gst.ctm);
        ctx.fillStyle = this._getFill(gst.colr, gst.ca, ctx);
        ctx.strokeStyle = this._getFill(gst.COLR, gst.CA, ctx);
        ctx.lineCap = ["butt", "round", "square"][gst.lcap];
        ctx.lineJoin = ["miter", "round", "bevel"][gst.ljoin];
        ctx.lineWidth = gst.lwidth * scl;
        var dsh = gst.dash.slice(0);
        for (var i = 0; i < dsh.length; i++)
            dsh[i] = _flt(dsh[i] * scl);
        ctx.setLineDash(dsh);
        ctx.miterLimit = gst.mlimit * scl;
        var fn = gst.font.Tf, ln = fn.toLowerCase();
        var p0 = ln.indexOf("bold") != -1 ? "bold " : "";
        var p1 = (ln.indexOf("italic") != -1 || ln.indexOf("oblique") != -1) ? "italic " : "";
        ctx.font = p0 + p1 + gst.font.Tfs + "px \"" + fn + "\"";
    };
    ToContext2D.prototype._getFill = function (colr, ca, ctx) {
        if (colr.typ == null)
            return this._colr(colr, ca);
        else {
            var grd = colr, crd = grd.crds, mat = grd.mat, scl = this._scale(mat), gf;
            if (grd.typ == "lin") {
                var p0 = this._multPoint(mat, crd.slice(0, 2)), p1 = this._multPoint(mat, crd.slice(2));
                gf = ctx.createLinearGradient(p0[0], p0[1], p1[0], p1[1]);
            }
            else if (grd.typ == "rad") {
                var p0 = this._multPoint(mat, crd.slice(0, 2)), p1 = this._multPoint(mat, crd.slice(3));
                gf = ctx.createRadialGradient(p0[0], p0[1], crd[2] * scl, p1[0], p1[1], crd[5] * scl);
            }
            for (var i = 0; i < grd.grad.length; i++)
                gf.addColorStop(grd.grad[i][0], this._colr(grd.grad[i][1], ca));
            return gf;
        }
    };
    ToContext2D.prototype._colr = function (c, a) { return "rgba(" + Math.round(c[0] * 255) + "," + Math.round(c[1] * 255) + "," + Math.round(c[2] * 255) + "," + a + ")"; };
    ToContext2D.prototype._scale = function (m) { return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])); };
    ToContext2D.prototype._concat = function (m, w) {
        var a = m[0], b = m[1], c = m[2], d = m[3], tx = m[4], ty = m[5];
        m[0] = (a * w[0]) + (b * w[2]);
        m[1] = (a * w[1]) + (b * w[3]);
        m[2] = (c * w[0]) + (d * w[2]);
        m[3] = (c * w[1]) + (d * w[3]);
        m[4] = (tx * w[0]) + (ty * w[2]) + w[4];
        m[5] = (tx * w[1]) + (ty * w[3]) + w[5];
    };
    ToContext2D.prototype._multPoint = function (m, p) { var x = p[0], y = p[1]; return [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]]; },
        ToContext2D.prototype._draw = function (path, ctx) {
            var c = 0, crds = path.crds;
            for (var j = 0; j < path.cmds.length; j++) {
                var cmd = path.cmds[j];
                if (cmd == "M") {
                    ctx.moveTo(crds[c], crds[c + 1]);
                    c += 2;
                }
                else if (cmd == "L") {
                    ctx.lineTo(crds[c], crds[c + 1]);
                    c += 2;
                }
                else if (cmd == "C") {
                    ctx.bezierCurveTo(crds[c], crds[c + 1], crds[c + 2], crds[c + 3], crds[c + 4], crds[c + 5]);
                    c += 6;
                }
                else if (cmd == "Q") {
                    ctx.quadraticCurveTo(crds[c], crds[c + 1], crds[c + 2], crds[c + 3]);
                    c += 4;
                }
                else if (cmd == "Z") {
                    ctx.closePath();
                }
            }
        };

    class ImageList {
        constructor(files) {
            if (files == null) {
                return;
            }
            this.images = {};
            for (let fileKey in files) {
                // let reg = new RegExp("xl/media/image1.png", "g");
                if (fileKey.indexOf("xl/media/") > -1) {
                    let fileNameArr = fileKey.split(".");
                    let suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                    if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1, "emf": 1 }) {
                        if (suffix == "emf") {
                            var pNum = 0; // number of the page, that you want to render
                            var scale = 1; // the scale of the document
                            var wrt = new ToContext2D(pNum, scale);
                            var inp, out, stt;
                            FromEMF.K = [];
                            inp = FromEMF.C;
                            out = FromEMF.K;
                            stt = 4;
                            for (var p in inp)
                                out[inp[p]] = p.slice(stt);
                            FromEMF.Parse(files[fileKey], wrt);
                            this.images[fileKey] = wrt.canvas.toDataURL("image/png");
                        }
                        else {
                            this.images[fileKey] = files[fileKey];
                        }
                    }
                }
            }
        }
        getImageByName(pathName) {
            if (pathName in this.images) {
                let base64 = this.images[pathName];
                return new Image(pathName, base64);
            }
            return null;
        }
    }
    class Image extends LuckyImageBase {
        constructor(pathName, base64) {
            super();
            this.src = base64;
        }
        setDefault() {
        }
    }

    class LuckyDefineNames {
        constructor(readXml) {
            let definedNames = readXml.getElementsByTagName("definedNames/definedName", workBookFile);
            const obj = {};
            definedNames.forEach(d => {
                const definedName = new LuckyDefineName(d);
                obj[definedName.id] = definedName;
            });
            this.defineNames = obj;
        }
    }
    class LuckyDefineName {
        constructor(ele) {
            this.id = generateRandomId(6);
            this.name = ele.get('name');
            this.formulaOrRefString = escapeCharacter(ele.value);
            this.comment = ele.get('comment');
            this.localSheetId = ele.get('localSheetId');
            this.hidden = ele.get('hidden') === '1';
        }
    }

    class LuckyFile extends LuckyFileBase {
        constructor(files, fileName) {
            super();
            this.columnWidthSet = [];
            this.rowHeightSet = [];
            this.handleWorkBookInfo = () => {
                this.workbook = new WorkBookInfo();
                const defineNames = new LuckyDefineNames(this.readXml);
                if (defineNames?.defineNames)
                    this.workbook.defineNames = defineNames.defineNames;
            };
            this.files = files;
            this.fileName = fileName;
            this.readXml = new ReadXml(files);
            this.getSheetNameList();
            this.sharedStrings = this.readXml.getElementsByTagName("sst/si", sharedStringsFile);
            this.calcChain = this.readXml.getElementsByTagName("calcChain/c", calcChainFile);
            this.styles = {};
            this.styles["cellXfs"] = this.readXml.getElementsByTagName("cellXfs/xf", stylesFile);
            this.styles["cellStyleXfs"] = this.readXml.getElementsByTagName("cellStyleXfs/xf", stylesFile);
            this.styles["cellStyles"] = this.readXml.getElementsByTagName("cellStyles/cellStyle", stylesFile);
            this.styles["fonts"] = this.readXml.getElementsByTagName("fonts/font", stylesFile);
            this.styles["fills"] = this.readXml.getElementsByTagName("fills/fill", stylesFile);
            this.styles["borders"] = this.readXml.getElementsByTagName("borders/border", stylesFile);
            this.styles["clrScheme"] = this.readXml.getElementsByTagName("a:clrScheme/a:dk1|a:lt1|a:dk2|a:lt2|a:accent1|a:accent2|a:accent3|a:accent4|a:accent5|a:accent6|a:hlink|a:folHlink", theme1File);
            this.styles["indexedColors"] = this.readXml.getElementsByTagName("colors/indexedColors/rgbColor", stylesFile);
            this.styles["mruColors"] = this.readXml.getElementsByTagName("colors/mruColors/color", stylesFile);
            this.styles['dxfs'] = this.readXml.getElementsByTagName("dxfs/dxf", stylesFile);
            this.imageList = new ImageList(files);
            this.cellImages = this.readXml.getElementsByTagName('etc:cellImages/etc:cellImage', cellImages);
            let numfmts = this.readXml.getElementsByTagName("numFmt/numFmt", stylesFile);
            let numFmtDefaultC = JSON.parse(JSON.stringify(numFmtDefault));
            for (let i = 0; i < numfmts.length; i++) {
                let attrList = numfmts[i].attributeList;
                let numfmtid = getXmlAttibute(attrList, "numFmtId", "49");
                let formatcode = getXmlAttibute(attrList, "formatCode", "@");
                // console.log(numfmtid, formatcode);
                if (!(numfmtid in numFmtDefault)) {
                    numFmtDefaultC[numfmtid] = numFmtDefaultMap[formatcode] || formatcode;
                }
            }
            // console.log(JSON.stringify(numFmtDefaultC), numfmts);
            this.styles["numfmts"] = numFmtDefaultC;
        }
        /**
        * @return All sheet name of workbook
        */
        getSheetNameList() {
            let workbookRelList = this.readXml.getElementsByTagName("Relationships/Relationship", workbookRels);
            if (workbookRelList == null) {
                return;
            }
            let regex = new RegExp("worksheets/[^/]*?.xml");
            let sheetNames = {};
            for (let i = 0; i < workbookRelList.length; i++) {
                let rel = workbookRelList[i], attrList = rel.attributeList;
                let id = attrList["Id"], target = attrList["Target"];
                if (regex.test(target)) {
                    if (target.indexOf('/xl') === 0) {
                        sheetNames[id] = target.substr(1);
                    }
                    else {
                        sheetNames[id] = "xl/" + target;
                    }
                }
            }
            this.sheetNameList = sheetNames;
        }
        /**
        * @param sheetName WorkSheet'name
        * @return sheet file name and path in zip
        */
        getSheetFileBysheetId(sheetId) {
            // for(let i=0;i<this.sheetNameList.length;i++){
            //     let sheetFileName = this.sheetNameList[i];
            //     if(sheetFileName.indexOf("sheet"+sheetId)>-1){
            //         return sheetFileName;
            //     }
            // }
            return this.sheetNameList[sheetId];
        }
        /**
        * @return workBook information
        */
        getWorkBookInfo() {
            let Company = this.readXml.getElementsByTagName("Company", appFile);
            let AppVersion = this.readXml.getElementsByTagName("AppVersion", appFile);
            let creator = this.readXml.getElementsByTagName("dc:creator", coreFile);
            let lastModifiedBy = this.readXml.getElementsByTagName("cp:lastModifiedBy", coreFile);
            let created = this.readXml.getElementsByTagName("dcterms:created", coreFile);
            let modified = this.readXml.getElementsByTagName("dcterms:modified", coreFile);
            this.info = new LuckyFileInfo();
            this.info.name = this.fileName;
            this.info.creator = creator.length > 0 ? creator[0].value : "";
            this.info.lastmodifiedby = lastModifiedBy.length > 0 ? lastModifiedBy[0].value : "";
            this.info.createdTime = created.length > 0 ? created[0].value : "";
            this.info.modifiedTime = modified.length > 0 ? modified[0].value : "";
            this.info.company = Company.length > 0 ? Company[0].value : "";
            this.info.appversion = AppVersion.length > 0 ? AppVersion[0].value : "";
        }
        /**
        * @return All sheet , include whole information
        */
        getSheetsFull(isInitialCell = true) {
            let sheets = this.readXml.getElementsByTagName("sheets/sheet", workBookFile);
            let sheetList = {};
            for (let key in sheets) {
                let sheet = sheets[key];
                sheetList[sheet.attributeList.name] = sheet.attributeList["sheetId"];
            }
            this.sheets = [];
            let order = 0;
            for (let key in sheets) {
                let sheet = sheets[key];
                let sheetName = sheet.attributeList.name;
                let sheetId = sheet.attributeList["sheetId"];
                let rid = sheet.attributeList["r:id"];
                let sheetFile = this.getSheetFileBysheetId(rid);
                let hide = sheet.attributeList.state === "hidden" ? 1 : 0;
                let drawing = this.readXml.getElementsByTagName("worksheet/drawing", sheetFile), drawingFile, drawingRelsFile;
                if (drawing != null && drawing.length > 0) {
                    let attrList = drawing[0].attributeList;
                    let rid = getXmlAttibute(attrList, "r:id", null);
                    if (rid != null) {
                        drawingFile = this.getDrawingFile(rid, sheetFile);
                        drawingRelsFile = this.getDrawingRelsFile(drawingFile);
                    }
                }
                if (sheetFile != null) {
                    let sheet = new LuckySheet(sheetName, sheetId, order, isInitialCell, {
                        sheetFile: sheetFile,
                        readXml: this.readXml,
                        sheetList: sheetList,
                        styles: this.styles,
                        sharedStrings: this.sharedStrings,
                        calcChain: this.calcChain,
                        imageList: this.imageList,
                        drawingFile: drawingFile,
                        drawingRelsFile: drawingRelsFile,
                        hide: hide,
                        cellImages: this.cellImages
                    });
                    this.columnWidthSet = [];
                    this.rowHeightSet = [];
                    this.imagePositionCaculation(sheet);
                    this.sheets.push(sheet);
                    order++;
                }
            }
        }
        extendArray(index, sets, def, hidden, lens) {
            if (index < sets.length) {
                return;
            }
            let startIndex = sets.length, endIndex = index;
            let allGap = 0;
            if (startIndex > 0) {
                allGap = sets[startIndex - 1];
            }
            // else{
            //     sets.push(0);
            // }
            for (let i = startIndex; i <= endIndex; i++) {
                let gap = def, istring = i.toString();
                if (istring in hidden) {
                    gap = 0;
                }
                else if (istring in lens) {
                    gap = lens[istring];
                }
                allGap += Math.round(gap + 1);
                sets.push(allGap);
            }
        }
        imagePositionCaculation(sheet) {
            let images = sheet.images, defaultColWidth = sheet.defaultColWidth, defaultRowHeight = sheet.defaultRowHeight;
            let colhidden = {};
            if (sheet.config.colhidden) {
                colhidden = sheet.config.colhidden;
            }
            let columnlen = {};
            if (sheet.config.columnlen) {
                columnlen = sheet.config.columnlen;
            }
            let rowhidden = {};
            if (sheet.config.rowhidden) {
                rowhidden = sheet.config.rowhidden;
            }
            let rowlen = {};
            if (sheet.config.rowlen) {
                rowlen = sheet.config.rowlen;
            }
            for (let key in images) {
                let imageObject = images[key]; //Image, luckyImage
                let fromCol = imageObject.fromCol;
                let fromColOff = imageObject.fromColOff;
                let fromRow = imageObject.fromRow;
                let fromRowOff = imageObject.fromRowOff;
                let toCol = imageObject.toCol;
                let toColOff = imageObject.toColOff;
                let toRow = imageObject.toRow;
                let toRowOff = imageObject.toRowOff;
                let x_n = 0, y_n = 0;
                let cx_n = 0, cy_n = 0;
                if (fromCol >= this.columnWidthSet.length) {
                    this.extendArray(fromCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
                }
                if (fromCol == 0) {
                    x_n = 0;
                }
                else {
                    x_n = this.columnWidthSet[fromCol - 1];
                }
                x_n = x_n + fromColOff;
                if (fromRow >= this.rowHeightSet.length) {
                    this.extendArray(fromRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
                }
                if (fromRow == 0) {
                    y_n = 0;
                }
                else {
                    y_n = this.rowHeightSet[fromRow - 1];
                }
                y_n = y_n + fromRowOff;
                if (toCol >= this.columnWidthSet.length) {
                    this.extendArray(toCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
                }
                if (toCol == 0) {
                    cx_n = 0;
                }
                else {
                    cx_n = this.columnWidthSet[toCol - 1];
                }
                cx_n = cx_n + toColOff - x_n;
                if (toRow >= this.rowHeightSet.length) {
                    this.extendArray(toRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
                }
                if (toRow == 0) {
                    cy_n = 0;
                }
                else {
                    cy_n = this.rowHeightSet[toRow - 1];
                }
                cy_n = cy_n + toRowOff - y_n;
                // console.log(defaultColWidth, colhidden , columnlen);
                // console.log(fromCol, this.columnWidthSet[fromCol] , fromColOff);
                // console.log(toCol, this.columnWidthSet[toCol] , toColOff, JSON.stringify(this.columnWidthSet));
                imageObject.originWidth = cx_n;
                imageObject.originHeight = cy_n;
                imageObject.crop.height = cy_n;
                imageObject.crop.width = cx_n;
                imageObject.default.height = cy_n;
                imageObject.default.left = x_n;
                imageObject.default.top = y_n;
                imageObject.default.width = cx_n;
            }
            //console.log(this.columnWidthSet, this.rowHeightSet);
        }
        /**
        * @return drawing file string
        */
        getDrawingFile(rid, sheetFile) {
            let sheetRelsPath = "xl/worksheets/_rels/";
            let sheetFileArr = sheetFile.split("/");
            let sheetRelsName = sheetFileArr[sheetFileArr.length - 1];
            let sheetRelsFile = sheetRelsPath + sheetRelsName + ".rels";
            let drawing = this.readXml.getElementsByTagName("Relationships/Relationship", sheetRelsFile);
            if (drawing.length > 0) {
                for (let i = 0; i < drawing.length; i++) {
                    let relationship = drawing[i];
                    let attrList = relationship.attributeList;
                    let relationshipId = getXmlAttibute(attrList, "Id", null);
                    if (relationshipId == rid) {
                        let target = getXmlAttibute(attrList, "Target", null);
                        if (target != null) {
                            return target.replace(/\.\.\//g, "");
                        }
                    }
                }
            }
            return null;
        }
        getDrawingRelsFile(drawingFile) {
            let drawingRelsPath = "xl/drawings/_rels/";
            let drawingFileArr = drawingFile.split("/");
            let drawingRelsName = drawingFileArr[drawingFileArr.length - 1];
            let drawingRelsFile = drawingRelsPath + drawingRelsName + ".rels";
            return drawingRelsFile;
        }
        /**
        * @return All sheet base information widthout cell and config
        */
        getSheetsWithoutCell() {
            this.getSheetsFull(false);
        }
        /**
        * @return LuckySheet file json
        */
        Parse() {
            // let xml = this.readXml;
            // for(let key in this.sheetNameList){
            //     let sheetName=this.sheetNameList[key];
            //     let sheetColumns = xml.getElementsByTagName("row/c/f", sheetName);
            //     console.log(sheetColumns);
            // }
            // return "";
            this.getWorkBookInfo();
            this.handleWorkBookInfo();
            this.getSheetsFull();
            // for(let i=0;i<this.sheets.length;i++){
            //     let sheet = this.sheets[i];
            //     let _borderInfo = sheet.config._borderInfo;
            //     if(_borderInfo==null){
            //         continue;
            //     }
            //     let _borderInfoKeys = Object.keys(_borderInfo);
            //     _borderInfoKeys.sort();
            //     for(let a=0;a<_borderInfoKeys.length;a++){
            //         let key = parseInt(_borderInfoKeys[a]);
            //         let b = _borderInfo[key];
            //         if(b.cells.length==0){
            //             continue;
            //         }
            //         if(sheet.config.borderInfo==null){
            //             sheet.config.borderInfo = [];
            //         }
            //         sheet.config.borderInfo.push(b);
            //     }
            // }
            return this.toJsonString(this);
        }
        toJsonString(file) {
            let LuckyOutPutFile = new LuckyFileBase();
            LuckyOutPutFile.info = file.info;
            LuckyOutPutFile.workbook = file.workbook;
            LuckyOutPutFile.sheets = [];
            file.sheets.forEach((sheet) => {
                let sheetout = new LuckySheetBase();
                //let attrName = ["name","color","config","index","status","order","row","column","luckysheet_select_save","scrollLeft","scrollTop","zoomRatio","showGridLines","defaultColWidth","defaultRowHeight","celldata","chart","isPivotTable","pivotTable","luckysheet_conditionformat_save","freezen","calcChain"];
                if (sheet.name != null) {
                    sheetout.name = sheet.name;
                }
                if (sheet.color != null) {
                    sheetout.color = sheet.color;
                }
                if (sheet.config != null) {
                    sheetout.config = sheet.config;
                    // if(sheetout.config._borderInfo!=null){
                    //     delete sheetout.config._borderInfo;
                    // }
                }
                if (sheet.index != null) {
                    sheetout.index = sheet.index;
                }
                if (sheet.status != null) {
                    sheetout.status = sheet.status;
                }
                if (sheet.order != null) {
                    sheetout.order = sheet.order;
                }
                if (sheet.row != null) {
                    sheetout.row = sheet.row;
                }
                if (sheet.column != null) {
                    sheetout.column = sheet.column;
                }
                if (sheet.luckysheet_select_save != null) {
                    sheetout.luckysheet_select_save = sheet.luckysheet_select_save;
                }
                if (sheet.scrollLeft != null) {
                    sheetout.scrollLeft = sheet.scrollLeft;
                }
                if (sheet.scrollTop != null) {
                    sheetout.scrollTop = sheet.scrollTop;
                }
                if (sheet.zoomRatio != null) {
                    sheetout.zoomRatio = sheet.zoomRatio;
                }
                if (sheet.showGridLines != null) {
                    sheetout.showGridLines = sheet.showGridLines;
                }
                if (sheet.defaultColWidth != null) {
                    sheetout.defaultColWidth = sheet.defaultColWidth;
                }
                if (sheet.defaultRowHeight != null) {
                    sheetout.defaultRowHeight = sheet.defaultRowHeight;
                }
                if (sheet.celldata != null) {
                    // sheetout.celldata = sheet.celldata;
                    sheetout.celldata = [];
                    sheet.celldata.forEach((cell) => {
                        let cellout = new LuckySheetCelldataBase();
                        cellout.r = cell.r;
                        cellout.c = cell.c;
                        cellout.v = cell.v;
                        sheetout.celldata.push(cellout);
                    });
                }
                if (sheet.chart != null) {
                    sheetout.chart = sheet.chart;
                }
                if (sheet.isPivotTable != null) {
                    sheetout.isPivotTable = sheet.isPivotTable;
                }
                if (sheet.pivotTable != null) {
                    sheetout.pivotTable = sheet.pivotTable;
                }
                if (sheet.luckysheet_conditionformat_save != null) {
                    sheetout.luckysheet_conditionformat_save = sheet.luckysheet_conditionformat_save;
                }
                if (sheet.freezen != null) {
                    sheetout.freezen = sheet.freezen;
                }
                if (sheet.calcChain != null) {
                    sheetout.calcChain = sheet.calcChain;
                }
                if (sheet.images != null) {
                    sheetout.images = sheet.images;
                }
                if (sheet.charts != null) {
                    sheetout.charts = sheet.charts;
                }
                if (sheet.dataVerification != null) {
                    sheetout.dataVerification = sheet.dataVerification;
                }
                if (sheet.hyperlink != null) {
                    sheetout.hyperlink = sheet.hyperlink;
                }
                if (sheet.hide != null) {
                    sheetout.hide = sheet.hide;
                }
                if (sheet.conditionalFormatting != null && sheet.conditionalFormatting.length) {
                    sheetout.conditionalFormatting = sheet.conditionalFormatting;
                }
                if (sheet.dataVerificationList != null && sheet.dataVerificationList.length) {
                    sheetout.dataVerificationList = sheet.dataVerificationList;
                }
                if (sheet.filter != null) {
                    sheetout.filter = sheet.filter;
                }
                LuckyOutPutFile.sheets.push(sheetout);
            });
            return JSON.stringify(LuckyOutPutFile);
        }
    }

    class HandleZip {
        constructor(file) {
            // Support nodejs fs to read files
            // if(file instanceof File){
            this.uploadFile = file;
            // }
        }
        unzipFile(successFunc, errorFunc) {
            var new_zip = new JSZip__default["default"]();
            new_zip.loadAsync(this.uploadFile) // 1) read the Blob
                .then(function (zip) {
                let fileList = {}, lastIndex = Object.keys(zip.files).length, index = 0;
                zip.forEach(function (relativePath, zipEntry) {
                    let fileName = zipEntry.name;
                    let fileNameArr = fileName.split(".");
                    let suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                    let fileType = "string";
                    if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1, }) {
                        fileType = "base64";
                    }
                    else if (suffix == "emf") {
                        fileType = "arraybuffer";
                    }
                    zipEntry.async(fileType).then(function (data) {
                        if (fileType == "base64") {
                            data = "data:image/" + suffix + ";base64," + data;
                        }
                        fileList[zipEntry.name] = data;
                        // console.log(lastIndex, index);
                        if (lastIndex == index + 1) {
                            successFunc(fileList);
                        }
                        index++;
                    });
                });
            }, function (e) {
                errorFunc(e);
            });
        }
        unzipFileByUrl(url, successFunc, errorFunc) {
            var new_zip = new JSZip__default["default"]();
            getBinaryContent(url, function (err, data) {
                if (err) {
                    throw err; // or handle err
                }
                new_zip.loadAsync(data).then(function (zip) {
                    let fileList = {}, lastIndex = Object.keys(zip.files).length, index = 0;
                    zip.forEach(function (relativePath, zipEntry) {
                        let fileName = zipEntry.name;
                        let fileNameArr = fileName.split(".");
                        let suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                        let fileType = "string";
                        if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1, }) {
                            fileType = "base64";
                        }
                        else if (suffix == "emf") {
                            fileType = "arraybuffer";
                        }
                        zipEntry.async(fileType).then(function (data) {
                            if (fileType == "base64") {
                                data = "data:image/" + suffix + ";base64," + data;
                            }
                            fileList[zipEntry.name] = data;
                            // console.log(lastIndex, index);
                            if (lastIndex == index + 1) {
                                successFunc(fileList);
                            }
                            index++;
                        });
                    });
                }, function (e) {
                    errorFunc(e);
                });
            });
        }
        newZipFile() {
            var zip = new JSZip__default["default"]();
            this.workBook = zip;
        }
        //title:"nested/hello.txt", content:"Hello Worldasdfasfasdfasfasfasfasfasdfas"
        addToZipFile(title, content) {
            if (this.workBook == null) {
                var zip = new JSZip__default["default"]();
                this.workBook = zip;
            }
            this.workBook.file(title, content);
        }
    }

    function hex2argb(rgb, prefix = '') {
        if (!rgb)
            return rgb;
        const colorName = colors[rgb];
        if (colorName)
            rgb = colorName;
        return rgb.toLocaleUpperCase().replace('#', prefix);
    }
    const colors = {
        "aliceblue": "#f0f8ff",
        "antiquewhite": "#faebd7",
        "aqua": "#00ffff",
        "aquamarine": "#7fffd4",
        "azure": "#f0ffff",
        "beige": "#f5f5dc",
        "bisque": "#ffe4c4",
        "black": "#000000",
        "blanchedalmond": "#ffebcd",
        "blue": "#0000ff",
        "blueviolet": "#8a2be2",
        "brown": "#a52a2a",
        "burlywood": "#deb887",
        "cadetblue": "#5f9ea0",
        "chartreuse": "#7fff00",
        "chocolate": "#d2691e",
        "coral": "#ff7f50",
        "cornflowerblue": "#6495ed",
        "cornsilk": "#fff8dc",
        "crimson": "#dc143c",
        "cyan": "#00ffff",
        "darkblue": "#00008b",
        "darkcyan": "#008b8b",
        "darkgoldenrod": "#b8860b",
        "darkgray": "#a9a9a9",
        "darkgreen": "#006400",
        "darkkhaki": "#bdb76b",
        "darkmagenta": "#8b008b",
        "darkolivegreen": "#556b2f",
        "darkorange": "#ff8c00",
        "darkorchid": "#9932cc",
        "darkred": "#8b0000",
        "darksalmon": "#e9967a",
        "darkseagreen": "#8fbc8f",
        "darkslateblue": "#483d8b",
        "darkslategray": "#2f4f4f",
        "darkturquoise": "#00ced1",
        "darkviolet": "#9400d3",
        "deeppink": "#ff1493",
        "deepskyblue": "#00bfff",
        "dimgray": "#696969",
        "dodgerblue": "#1e90ff",
        "firebrick": "#b22222",
        "floralwhite": "#fffaf0",
        "forestgreen": "#228b22",
        "fuchsia": "#ff00ff",
        "gainsboro": "#dcdcdc",
        "ghostwhite": "#f8f8ff",
        "gold": "#ffd700",
        "goldenrod": "#daa520",
        "gray": "#808080",
        "green": "#008000",
        "greenyellow": "#adff2f",
        "honeydew": "#f0fff0",
        "hotpink": "#ff69b4",
        "indianred ": "#cd5c5c",
        "indigo": "#4b0082",
        "ivory": "#fffff0",
        "khaki": "#f0e68c",
        "lavender": "#e6e6fa",
        "lavenderblush": "#fff0f5",
        "lawngreen": "#7cfc00",
        "lemonchiffon": "#fffacd",
        "lightblue": "#add8e6",
        "lightcoral": "#f08080",
        "lightcyan": "#e0ffff",
        "lightgoldenrodyellow": "#fafad2",
        "lightgrey": "#d3d3d3",
        "lightgreen": "#90ee90",
        "lightpink": "#ffb6c1",
        "lightsalmon": "#ffa07a",
        "lightseagreen": "#20b2aa",
        "lightskyblue": "#87cefa",
        "lightslategray": "#778899",
        "lightsteelblue": "#b0c4de",
        "lightyellow": "#ffffe0",
        "lime": "#00ff00",
        "limegreen": "#32cd32",
        "linen": "#faf0e6",
        "magenta": "#ff00ff",
        "maroon": "#800000",
        "mediumaquamarine": "#66cdaa",
        "mediumblue": "#0000cd",
        "mediumorchid": "#ba55d3",
        "mediumpurple": "#9370d8",
        "mediumseagreen": "#3cb371",
        "mediumslateblue": "#7b68ee",
        "mediumspringgreen": "#00fa9a",
        "mediumturquoise": "#48d1cc",
        "mediumvioletred": "#c71585",
        "midnightblue": "#191970",
        "mintcream": "#f5fffa",
        "mistyrose": "#ffe4e1",
        "moccasin": "#ffe4b5",
        "navajowhite": "#ffdead",
        "navy": "#000080",
        "oldlace": "#fdf5e6",
        "olive": "#808000",
        "olivedrab": "#6b8e23",
        "orange": "#ffa500",
        "orangered": "#ff4500",
        "orchid": "#da70d6",
        "palegoldenrod": "#eee8aa",
        "palegreen": "#98fb98",
        "paleturquoise": "#afeeee",
        "palevioletred": "#d87093",
        "papayawhip": "#ffefd5",
        "peachpuff": "#ffdab9",
        "peru": "#cd853f",
        "pink": "#ffc0cb",
        "plum": "#dda0dd",
        "powderblue": "#b0e0e6",
        "purple": "#800080",
        "rebeccapurple": "#663399",
        "red": "#ff0000",
        "rosybrown": "#bc8f8f",
        "royalblue": "#4169e1",
        "saddlebrown": "#8b4513",
        "salmon": "#fa8072",
        "sandybrown": "#f4a460",
        "seagreen": "#2e8b57",
        "seashell": "#fff5ee",
        "sienna": "#a0522d",
        "silver": "#c0c0c0",
        "skyblue": "#87ceeb",
        "slateblue": "#6a5acd",
        "slategray": "#708090",
        "snow": "#fffafa",
        "springgreen": "#00ff7f",
        "steelblue": "#4682b4",
        "tan": "#d2b48c",
        "teal": "#008080",
        "thistle": "#d8bfd8",
        "tomato": "#ff6347",
        "turquoise": "#40e0d0",
        "violet": "#ee82ee",
        "wheat": "#f5deb3",
        "white": "#ffffff",
        "whitesmoke": "#f5f5f5",
        "yellow": "#ffff00",
        "yellowgreen": "#9acd32"
    };
    function heightConvert(height) {
        return (height / 96) * 72;
    }
    function wdithConvert(width) {
        return (width - 10) / 8 + 0.83 + 0.64;
    }
    function isUndefined(value, result) {
        if (value === undefined)
            return undefined;
        return result === undefined ? value : result;
    }
    function convertSheetIdToName(sheets, id) {
        return sheets[id]?.name;
    }

    function cellStyle(style = {}, numFmt, isCondition = false) {
        return {
            numFmt: numFmt,
            font: fontConvert(style),
            alignment: alignmentConvert(style),
            protection: null,
            border: borderConvert(style.bd),
            fill: fillConvert(style.bg?.rgb, isCondition)
        };
    }
    function fontConvert(style) {
        const univerToExcel = {
            underline: {
                10: 'double',
                12: 'single',
            },
            vertAlign: {
                1: undefined,
                2: 'subscript',
                3: 'superscript'
            }
        };
        return {
            name: style.ff,
            size: style.fs,
            family: 1,
            color: isUndefined(style.cl?.rgb, { argb: hex2argb(style.cl?.rgb) }),
            bold: isUndefined(style.bl, style.bl === 1),
            italic: isUndefined(style.it, style.it === 1),
            underline: isUndefined(style.ul?.s, (style.ul?.s === 1 ? (univerToExcel.underline[style.ul.t] || true) : false)),
            vertAlign: isUndefined(style.va, univerToExcel.vertAlign[style.va]),
            strike: isUndefined(style.st?.s, style.st?.s === 1),
            outline: isUndefined(style.ol?.s, style.ol?.s === 1),
            charset: 134
        };
    }
    function borderConvert(border) {
        if (!border) {
            return null;
        }
        const borderStyle = {
            0: 'none',
            1: 'thin',
            2: 'hair',
            3: 'dotted',
            4: 'dashDot', // 'Dashed',
            5: 'dashDot',
            6: 'dashDotDot',
            7: 'double',
            8: 'medium',
            9: 'mediumDashed',
            10: 'mediumDashDot',
            11: 'mediumDashDotDot',
            12: 'slantDashDot',
            13: 'thick'
        };
        const template = (bd) => {
            if (!bd)
                return undefined;
            const st = {
                style: borderStyle[bd?.s || 1],
                color: {
                    argb: hex2argb(bd?.cl?.rgb || '#d9d9d9')
                }
            };
            return st;
        };
        const diagonal = template(border.bl_tr || border.tl_br) || {};
        return {
            top: template(border.t),
            right: template(border.r),
            bottom: template(border.b),
            left: template(border.l),
            diagonal: {
                up: border.bl_tr ? true : false,
                down: border.tl_br ? true : false,
                ...diagonal
            }
        };
    }
    function alignmentConvert(style) {
        const univerToExcel = {
            horizontal: {
                0: 'left',
                1: 'left',
                2: 'center',
                3: 'right',
            },
            vertical: {
                1: 'top',
                2: 'middle',
                3: 'bottom',
            },
            wrapText: {
                3: true
            }
        };
        return {
            horizontal: isUndefined(style.ht, univerToExcel.horizontal[style.ht]),
            vertical: isUndefined(style.vt, univerToExcel.vertical[style.vt]),
            wrapText: isUndefined(style.tb, style.tb === 3),
            textRotation: isUndefined(style.tr?.a)
        };
    }
    function fillConvert(bg, isCondition = false) {
        if (!bg)
            return null;
        if (!bg)
            return {
                type: 'pattern',
                pattern: 'none',
            };
        const fill = {
            type: 'pattern',
            pattern: 'solid',
        };
        if (isCondition) {
            fill.bgColor = { argb: hex2argb(bg, 'FF') };
        }
        else {
            fill.fgColor = { argb: hex2argb(bg, 'FF') };
        }
        return fill;
    }

    class Resource {
        constructor(sheetId, workbook, worksheet, resources) {
            this.sheetId = sheetId;
            this.workbook = workbook;
            this.worksheet = worksheet;
            this.resources = resources;
            this.setImages();
            this.setConditional();
            this.setDataValidation();
            this.setFilter();
        }
        handleRang(range) {
            const { startRow, startColumn, endRow, endColumn } = range;
            return getRangetxt({
                row: [startRow, endRow],
                column: [startColumn, endColumn],
            }, '');
        }
        // private setRangeProtection() {
        //     const rangeProtection = this.getSheetResource('SHEET_RANGE_PROTECTION_PLUGIN');
        // }
        setFilter() {
            const filters = this.getSheetResource('SHEET_FILTER_PLUGIN');
            if (!filters)
                return;
            this.worksheet.autoFilter = this.handleRang(filters.ref);
        }
        setConditional() {
            const conditionals = this.getSheetResource('SHEET_CONDITIONAL_FORMATTING_PLUGIN');
            const ruleList = [];
            if (!conditionals)
                return;
            conditionals.forEach((conditional) => {
                const { ranges, rule, stopIfTrue } = conditional;
                ranges.forEach((range) => {
                    const ref = this.handleRang(range);
                    const index = ruleList.findIndex(d => d.ref === ref);
                    const ruleValue = this.handleRule(conditional);
                    if (index > -1) {
                        ruleList[index].rules.push(ruleValue);
                        return;
                    }
                    ruleList.push({
                        ref,
                        rules: [ruleValue]
                    });
                });
            });
            // console.log(this.worksheet.name, ruleList)
            ruleList.forEach(d => {
                this.worksheet.addConditionalFormatting(d);
            });
        }
        setDataValidation() {
            const datavalidations = this.getSheetResource('SHEET_DATA_VALIDATION_PLUGIN');
            const excelToDate = (v, date1904 = false) => {
                // eslint-disable-next-line no-mixed-operators
                const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
                return new Date(millisecondSinceEpoch);
            };
            datavalidations?.forEach((validate) => {
                const { ranges = [], type, allowBlank, operator, formula1, formula2, showErrorMessage, showInputMessage, prompt, promptTitle, error, errorTitle, errorStyle } = validate || {};
                const styleMap = ['information', 'stop', 'warning'];
                let formulae = [formula1];
                if (!isEmpty(formula2))
                    formulae.push(formula2);
                if (type === 'list') {
                    formulae = [`"${formula1}"`];
                    if (!isEmpty(formula2))
                        formulae.push(`"${formula2}"`);
                }
                if (type === 'date') {
                    formulae = [excelToDate(formula1)];
                    if (!isEmpty(formula2))
                        formulae.push(excelToDate(formula2));
                }
                const list = ranges.map((range) => this.handleRang(range));
                const valid = {
                    type,
                    allowBlank,
                    operator,
                    formulae,
                    showErrorMessage,
                    showInputMessage,
                    prompt,
                    promptTitle,
                    error,
                    errorTitle,
                    errorStyle: styleMap[errorStyle]
                };
                // this.worksheet.dataValidations.add(list.join(' '), removeEmptyAttr(valid))
                // console.log(this.sheetId, this.worksheet.name, list, valid);
                list.forEach((address) => {
                    this.worksheet.dataValidations.add(address, removeEmptyAttr$1(valid));
                });
            });
        }
        setImages() {
            const images = this.getSheetResource('SHEET_DRAWING_PLUGIN');
            const sheetIamges = images?.data;
            if (!sheetIamges)
                return;
            for (const key in sheetIamges) {
                const element = sheetIamges[key];
                const images = this.workbook.getImages();
                let imageId = images.findIndex(d => d.extension === 'png' && d.base64 === element.source);
                if (imageId === -1) {
                    imageId = this.workbook.addImage({
                        base64: element.source,
                        extension: 'png'
                    });
                }
                const handlePosition = (position) => {
                    return {
                        nativeCol: position.column,
                        nativeColOff: getEmusByPx(position.columnOffset),
                        nativeRow: position.row,
                        nativeRowOff: getEmusByPx(position.rowOffset),
                    };
                };
                // console.log(handlePosition(element.sheetTransform.from), handlePosition(element.sheetTransform.to))
                this.worksheet.addImage(imageId, {
                    tl: handlePosition(element.sheetTransform.from),
                    br: handlePosition(element.sheetTransform.to)
                });
            }
        }
        getSheetResource(name) {
            const resources = jsonParse(this.resources.find((d) => d.name === name)?.data);
            const sheetResources = resources[this.sheetId];
            return sheetResources;
        }
        handleRule(conditional) {
            const { rule, stopIfTrue, order } = conditional;
            const ruleValue = {};
            if (stopIfTrue)
                ruleValue.stopIfTrue = 1;
            ruleValue.priority = order;
            if (rule.style) {
                // const formatCode = rule.style.n?.pattern ? { formatCode: rule.style.n?.pattern } : undefined;
                ruleValue.style = cellStyle(rule.style, rule.style.n?.pattern, true);
            }
            if (rule.operator)
                ruleValue.operator = rule.operator;
            switch (rule.type) {
                case CFRuleType.colorScale:
                    ruleValue.type = CFRuleType.colorScale;
                    ruleValue.cfvo = rule.config?.map((d) => {
                        return {
                            type: d.value.type,
                            value: d.value.value
                        };
                    });
                    ruleValue.color = rule.config?.map((d) => {
                        return {
                            argb: hex2argb(d.color)
                        };
                    });
                    break;
                case CFRuleType.dataBar:
                    ruleValue.type = CFRuleType.dataBar;
                    ruleValue.showValue = rule.isShowValue;
                    ruleValue.gradient = rule.config.isGradient;
                    ruleValue.cfvo = [{
                            type: rule.config.min.type,
                            value: rule.config.min.value
                        }, {
                            type: rule.config.max.type,
                            value: rule.config.max.value
                        }];
                    ruleValue.negativeFillColor = { argb: hex2argb(rule.config.nativeColor) };
                    ruleValue.color = { argb: hex2argb(rule.config.positiveColor) };
                    // ruleValue.border = false;
                    ruleValue.axisPosition = 'auto';
                    ruleValue.direction = 'leftToRight';
                    ruleValue.minLength = 0;
                    ruleValue.maxLength = 100;
                    ruleValue.negativeBarColorSameAsPositive = true;
                    ruleValue.negativeBarBorderColorSameAsPositive = true;
                    break;
                case CFRuleType.iconSet:
                    ruleValue.type = CFRuleType.iconSet;
                    ruleValue.reverse = false;
                    ruleValue.showValue = rule.isShowValue;
                    ruleValue.icons = rule.config?.map((d) => {
                        const iconId = str2num(d.iconType.charAt(0)) - str2num(d?.iconId) - 1;
                        return {
                            iconId,
                            iconSet: d.iconType
                        };
                    }).reverse();
                    ruleValue.custom = true;
                    // ruleValue.iconSet = rule.config[0]?.iconType;
                    ruleValue.cfvo = rule.config?.map((d) => {
                        return {
                            type: d.value.type,
                            value: d.value.value
                        };
                    }).reverse();
                    break;
                case CFRuleType.highlightCell:
                    switch (rule.subType) {
                        case CFSubRuleType.average:
                            ruleValue.type = 'aboveAverage';
                            ruleValue.aboveAverage = false;
                            break;
                        case CFSubRuleType.duplicateValues:
                            ruleValue.type = 'duplicateValues';
                            break;
                        case CFSubRuleType.formula:
                            ruleValue.type = 'expression';
                            ruleValue.formulae = [escapeCharacter(rule.value)];
                            break;
                        case CFSubRuleType.number:
                            ruleValue.type = 'cellIs';
                            ruleValue.formulae = [rule.value];
                            break;
                        case CFSubRuleType.rank:
                            ruleValue.type = 'top10';
                            ruleValue.rank = rule.value;
                            ruleValue.percent = rule.isPercent;
                            ruleValue.bottom = rule.isBottom;
                            break;
                        case CFSubRuleType.text:
                            ruleValue.type = 'containsText';
                            ruleValue.text = rule.value;
                            break;
                        case CFSubRuleType.timePeriod:
                            ruleValue.type = 'timePeriod';
                            ruleValue.timePeriod = rule.operator;
                            break;
                        // case CFSubRuleType.uniqueValues:
                        //     ruleValue.type = 'uniqueValues';
                        //     break;
                    }
                    break;
            }
            return ruleValue;
        }
    }

    class ViewCommon {
    }
    class FrozenView {
    }
    function ExcelWorkSheet(workbook, snapshot) {
        const { sheetOrder, sheets, styles, resources } = snapshot;
        sheetOrder.forEach((sheetId) => {
            const sheet = sheets[sheetId];
            const { id, name, tabColor, defaultRowHeight, defaultColumnWidth, hidden, rightToLeft, showGridlines, freeze, mergeData } = sheet;
            const commonView = new ViewCommon();
            commonView.rightToLeft = rightToLeft === 1;
            commonView.showGridLines = showGridlines === 1;
            const frozenView = new FrozenView();
            if (freeze.xSplit > 0 || freeze.ySplit > 0) {
                frozenView.state = 'frozen';
                frozenView.xSplit = freeze.xSplit;
                frozenView.ySplit = freeze.ySplit;
            }
            const views = Object.assign(commonView, frozenView);
            const defaultColWidth = wdithConvert(defaultColumnWidth);
            const defaultRowHeightR = heightConvert(defaultRowHeight);
            const worksheet = workbook.addWorksheet(name, {
                views: [views],
                state: hidden === 1 ? 'hidden' : 'visible',
                properties: {
                    tabColor: tabColor ? { argb: hex2argb(tabColor) } : undefined,
                    defaultColWidth: defaultColWidth,
                    defaultRowHeight: defaultRowHeightR,
                    dyDescent: 0
                }
            });
            setColumns(worksheet, sheet.columnData, defaultColWidth);
            setRows(worksheet, sheet.rowData, defaultRowHeightR);
            setCell(worksheet, sheet, styles, snapshot, workbook);
            setMerges(worksheet, mergeData);
            new Resource(id, workbook, worksheet, resources);
        });
    }
    function setMerges(worksheet, mergeData) {
        mergeData.forEach(d => {
            worksheet.mergeCells(d.startRow + 1, d.startColumn + 1, d.endRow + 1, d.endColumn + 1);
        });
    }
    function setCell(worksheet, sheet, styles, snapshot, workbook) {
        const { resources, sheets } = snapshot;
        const { cellData, id } = sheet;
        for (const rowid in cellData) {
            const row = cellData[rowid];
            for (const columnid in row) {
                const cell = row[columnid];
                if (!cell)
                    continue;
                // console.log(rowid + 1, columnid + 1)
                const target = worksheet.getCell(Number(rowid) + 1, Number(columnid) + 1);
                target.value = handleValue(cell, {
                    resources,
                    sheetId: id,
                    rowId: rowid,
                    columnId: columnid,
                    sheets
                }, workbook);
                let originStyle = cell.s;
                if (typeof cell.s === 'string') {
                    originStyle = styles[cell.s];
                }
                const style = removeEmptyAttr$1(cellStyle(originStyle, originStyle?.n?.pattern || cell.f));
                Object.assign(target, style);
                // console.log(target)
            }
        }
    }
    function getHyperLink(cellSource) {
        const { resources, sheetId, rowId, columnId } = cellSource;
        const hyperlinks = jsonParse(resources.find((d) => d.name === 'SHEET_HYPER_LINK_PLUGIN')?.data);
        const list = hyperlinks?.[sheetId] || [];
        const hyperlink = list.find((d) => d.row === Number(rowId) && d.column === Number(columnId));
        return hyperlink;
    }
    function handleHyperLink(hyperlink, sheets) {
        let hyperlinks;
        if (hyperlink) {
            const { payload } = hyperlink;
            let link = '';
            let model = '';
            if (payload.includes('#gid=') || payload.includes('range=')) {
                const str = payload.replace('#', '');
                const arr = str.split('&');
                link += '';
                if (arr.length === 1 && arr[0].includes('range=')) {
                    link += arr[0].replace('range=');
                }
                if (arr.length === 2) {
                    link += `\'${convertSheetIdToName(sheets, arr[0].replace('gid=', ''))}\'`;
                    link += `!${arr[1].replace('range=', '')}`;
                }
            }
            else {
                link = payload;
                model = 'External';
            }
            if (link)
                hyperlinks = {
                    hyperlink: link,
                    hyperlinkModel: model
                };
        }
        return hyperlinks;
    }
    function handleValue(cell, cellSource, workbook) {
        const { sheets } = cellSource;
        const hyperlink = getHyperLink(cellSource);
        const hyperlinks = handleHyperLink(hyperlink, sheets);
        let value;
        if (cell.p) {
            const body = cell.p?.body;
            if (cell.p.drawingsOrder?.length) {
                const image = cell.p.drawings[cell.p.drawingsOrder[0]];
                const { id, value: imgId } = workbook.addCellImage({
                    base64: image.source,
                    extension: 'png',
                    descr: image.description,
                    ext: {
                        width: image.transform.width,
                        height: image.transform.height,
                    }
                });
                value = { id, cellImageId: imgId, ...(hyperlinks || {}) };
                return value;
            }
            else {
                value = {
                    richText: body?.textRuns.map((d) => {
                        return {
                            text: body.dataStream.substring(d.st, d.ed),
                            font: fontConvert(d.ts)
                        };
                    })
                };
            }
        }
        else if (cell.si) {
            value = { formula: cell.si, result: cell.v };
        }
        else {
            value = cell.v;
        }
        if (hyperlinks) {
            const text = value?.richText?.map?.((d) => d.text)?.join('') || value?.result || value;
            value = {
                text: text,
                ...hyperlinks
            };
        }
        return value;
    }
    function setColumns(worksheet, columnData = {}, defaultColumnWidth) {
        for (const key in columnData) {
            if (Object.prototype.hasOwnProperty.call(columnData, key)) {
                const element = columnData[key];
                const column = worksheet.getColumn(Number(key) + 1);
                column.width = element.w ? wdithConvert(element.w) : defaultColumnWidth;
                column.hidden = element.hd === 1;
            }
        }
    }
    function setRows(worksheet, rowData = {}, defaultRowHeight) {
        for (const key in rowData) {
            if (Object.prototype.hasOwnProperty.call(rowData, key)) {
                const element = rowData[key];
                const row = worksheet.getRow(Number(key) + 1);
                row.height = element.h ? heightConvert(element.h) : defaultRowHeight;
                row.hidden = element.hd === 1;
            }
        }
    }

    // SHEET_HYPER_LINK_PLUGIN
    // SHEET_DRAWING_PLUGIN
    // SHEET_DEFINED_NAME_PLUGIN
    // SHEET_CONDITIONAL_FORMATTING_PLUGIN
    // SHEET_DATA_VALIDATION_PLUGIN
    // SHEET_FILTER_PLUGIN
    const Workbook = exceljs__default["default"].Workbook;
    class WorkBook extends Workbook {
        constructor(snapshot) {
            super();
            this.init(snapshot);
        }
        init(snapshot) {
            // this.properties.date1904 = true;
            this.calcProperties.fullCalcOnLoad = true;
            this.setDefineNames(snapshot.resources);
            ExcelWorkSheet(this, snapshot);
        }
        setDefineNames(resources) {
            const definedNames = jsonParse(resources.find(d => d.name === 'SHEET_DEFINED_NAME_PLUGIN')?.data);
            for (const key in definedNames) {
                const element = definedNames[key];
                this.definedNames.add(element.formulaOrRefString, element.name);
            }
        }
    }

    class CSV {
        constructor(snapshot) {
            this.csvList = {};
            this.csvContent = {};
            this.init(snapshot);
        }
        init(snapshot) {
            if (!snapshot) {
                return;
            }
            const { sheetOrder, sheets } = snapshot;
            const data = {};
            sheetOrder.forEach((d) => {
                const sheet = sheets[d];
                if (!sheet)
                    return;
                const { cellData, name } = sheet;
                const list = [];
                for (const key in cellData) {
                    const rows = cellData[key];
                    for (const key in rows) {
                        const row = Number(key);
                        const col = Number(key);
                        if (!list[row])
                            list[row] = [];
                        list[row][col] = rows[key]?.v;
                    }
                }
                data[name] = list;
            });
            this.csvList = data;
            this.handleCsvContent();
        }
        handleCsvContent() {
            const data = {};
            for (const key in this.csvList) {
                const csv = this.csvList[key];
                let csvContent = "data:text/csv;charset=utf-8,";
                // 拼接csv数据
                csv.forEach(row => {
                    csvContent += row.join(",") + "\r\n";
                });
                data[key] = csvContent;
            }
            this.csvContent = data;
        }
    }

    class UniverSheetBase {
        constructor(params) {
            this.type = core.SheetTypes.GRID;
            this.tabColor = '';
            this.hidden = 0;
            this.freeze = {
                xSplit: 0,
                ySplit: 0,
                startRow: -1,
                startColumn: -1,
            };
            this.rowCount = 100;
            this.columnCount = 20;
            this.zoomRatio = 1;
            this.scrollTop = 0;
            this.scrollLeft = 0;
            this.defaultColumnWidth = 93;
            this.defaultRowHeight = 27;
            this.mergeData = [];
            this.cellData = {};
            this.rowData = {};
            this.columnData = {};
            this.rowHeader = {
                width: 46,
                hidden: 0,
            };
            this.columnHeader = {
                height: 20,
                hidden: 0,
            };
            this.showGridlines = 1;
            this.rightToLeft = 0;
            this.selections = [];
            const { id, name, cellData, rowCount = 0, colCount = 0 } = params || {};
            this.id = id || '';
            this.name = name || '';
            this.cellData = cellData || {};
            this.rowCount = Math.max(this.rowCount, rowCount);
            this.columnCount = Math.max(this.columnCount, colCount);
        }
    }

    /**
     * 删除对象中含undefined的值
     * @param object
     * @returns
     */
    function removeEmptyAttr(object) {
        for (const key in object) {
            if (Object.prototype.hasOwnProperty.call(object, key)) {
                if (object[key] === undefined) {
                    delete object[key]; // 删除值为 undefined 的属性
                }
                else if (isObject(object[key]) && object[key] !== null) {
                    removeEmptyAttr(object[key]); // 对子对象递归
                }
            }
        }
        return object;
    }
    const handleStyle = (row, borderInfo, domContent = false) => {
        const { v } = row;
        if (typeof v === 'string' || v === null || v === undefined) {
            return undefined;
        }
        // 0 middle, 1 up, 2 down
        const VerticalAlignMap = {
            0: core.VerticalAlign.MIDDLE,
            1: core.VerticalAlign.TOP,
            2: core.VerticalAlign.BOTTOM,
        };
        let border = undefined;
        if (borderInfo?.value && !domContent) {
            const handleBorder = (con) => {
                if (!con)
                    return null;
                return {
                    s: con.style,
                    cl: { rgb: con.color, th: core.ThemeColorType.DARK1 },
                };
            };
            border = {
                t: handleBorder(borderInfo.value?.t),
                r: handleBorder(borderInfo.value?.r),
                b: handleBorder(borderInfo.value?.b),
                l: handleBorder(borderInfo.value?.l),
                bl_tr: handleBorder(borderInfo.value?.bl_tr),
                tl_br: handleBorder(borderInfo.value?.tl_br),
            };
        }
        const TextWrap = {
            0: core.WrapStrategy.CLIP,
            1: core.WrapStrategy.OVERFLOW,
            2: core.WrapStrategy.WRAP,
        };
        let angle = undefined;
        const vtMap = {
            1: 45,
            2: 135,
            3: 255,
            4: 90,
            5: 180,
        };
        if (v.tr)
            angle = vtMap[v.tr];
        if (v.rt)
            angle = v.rt;
        const unMap = {
            0: core.TextDecoration.DASH,
            1: core.TextDecoration.SINGLE,
            2: core.TextDecoration.DOUBLE,
            3: core.TextDecoration.SINGLE,
            4: core.TextDecoration.DOUBLE,
        };
        // 0 center, 1 left, 2 right
        const htMap = {
            0: core.HorizontalAlign.CENTER,
            1: core.HorizontalAlign.LEFT,
            2: core.HorizontalAlign.RIGHT,
        };
        return {
            // bbl: , // bottomBorerLine
            bd: border, // border
            bg: v.bg !== undefined ? { rgb: v.bg, th: core.ThemeColorType.DARK1 } : undefined, // background
            bl: v.bl, // bold 0: false 1: true
            cl: v.fc !== undefined ? { rgb: v.fc, th: core.ThemeColorType.DARK1 } : undefined, // foreground
            ff: v.ff, // fontFamily
            fs: v.fs, // fontSize
            ht: v.ht !== undefined ? htMap[v.ht] : undefined, // horizontalAlignment
            it: v.it, // italic 0: false 1: true
            n: v.ct?.fa !== undefined ? { pattern: v.ct.fa } : undefined, //Numfmt pattern
            // ol: { s: v.cl === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE}, // overline
            // pd: , // padding
            st: v.cl !== undefined
                ? { s: v.cl === 1 ? core.BooleanNumber.TRUE : core.BooleanNumber.FALSE }
                : undefined, // strikethrough
            tb: v.tb !== undefined ? TextWrap[v.tb] : undefined, // wrapStrategy
            // td: , // textDirection
            tr: angle !== undefined
                ? {
                    a: angle,
                    v: v.tr || v.rt ? core.BooleanNumber.TRUE : core.BooleanNumber.FALSE,
                }
                : undefined, // textRotation
            ul: v.un !== undefined
                ? {
                    s: v.un === undefined ? core.BooleanNumber.FALSE : core.BooleanNumber.TRUE,
                    t: v.un ? unMap[v.un] : core.TextDecoration.DASH,
                }
                : undefined, // underline
            // va: , // (Subscript 下标 /Superscript上标 Text)
            vt: v.vt !== undefined ? VerticalAlignMap[v.vt] : undefined, // verticalAlignment
        };
    };

    var ImageSourceType;
    (function (ImageSourceType) {
        ImageSourceType["URL"] = "URL";
        ImageSourceType["UUID"] = "UUID";
        ImageSourceType["BASE64"] = "BASE64";
    })(ImageSourceType || (ImageSourceType = {}));

    class UniverSheet extends UniverSheetBase {
        constructor(sheetData) {
            super();
            this.hyperLink = [];
            this.handleMerge = (config) => {
                const merges = config.merge;
                if (!merges)
                    return [];
                return Object.values(merges).map((merge) => {
                    return {
                        startRow: merge.r,
                        endRow: merge.r + merge.rs - 1,
                        startColumn: merge.c,
                        endColumn: merge.c + merge.cs - 1,
                    };
                });
            };
            this.handleCellData = (celldata, config) => {
                const handleCell = (row) => {
                    const { v } = row;
                    if (typeof v === 'string' || v === null || v === undefined) {
                        return { v: v };
                    }
                    const tMap = {
                        s: core.CellValueType.STRING,
                        n: core.CellValueType.NUMBER,
                        b: core.CellValueType.BOOLEAN,
                        str: core.CellValueType.STRING,
                    };
                    const borderConf = config.borderInfo?.find((d) => d.value.col_index === row.c && d.value.row_index === row.r);
                    let cellType = v.ct?.t && tMap[v.ct?.t] ? tMap[v.ct?.t] : core.CellValueType.NUMBER;
                    let val = cellType === core.CellValueType.NUMBER ? str2num(v.v) : v.v;
                    if (cellType === core.CellValueType.BOOLEAN)
                        val = v.v == '1' ? 1 : 0;
                    if (Number.isNaN(Number(val)) && cellType === core.CellValueType.NUMBER)
                        cellType = core.CellValueType.STRING;
                    if (this.hyperLink.findIndex((d) => d.column === row.c && d.row === row.r) > -1)
                        cellType = core.CellValueType.STRING;
                    const f = v.f?.replace(/=_xlfn./g, '=');
                    const cell = {
                        // custom: v., // User stored custom fields
                        f,
                        // p: , // The unique key, a random string, is used for the plug-in to associate the cell. When the cell information changes, the plug-in does not need to change the data, reducing the pressure on the back-end interface id?: string.
                        s: handleStyle(row, borderConf),
                        // si: f, // Id of the formula.
                        t: cellType,
                        v: val,
                    };
                    const pVal = this.handleDocument(row, config);
                    if (pVal)
                        cell.p = pVal;
                    const pValImg = this.handleCellImage(row, config);
                    if (pValImg) {
                        cell.p = pValImg;
                        cell.f = undefined;
                        cell.v = undefined;
                    }
                    return removeEmptyAttr(cell);
                };
                let row = undefined;
                let colCount = 0;
                const rowData = celldata.reduce((pre, cur) => {
                    if (row === cur.r) {
                        pre[cur.r].push(cur);
                    }
                    else {
                        row = cur.r;
                        pre[row] = [cur];
                    }
                    if (cur.c > colCount)
                        colCount = cur.c;
                    return pre;
                }, []);
                const cell = {};
                // console.log(rowData, celldata, colCount)
                rowData.forEach((row, rowIndex) => {
                    for (let index = 0; index < colCount + 1; index++) {
                        const element = row.find((d) => d.c === index) || {
                            r: rowIndex,
                            c: index,
                            v: null,
                        };
                        if (!cell[element.r])
                            cell[element.r] = {};
                        cell[element.r][element.c] = handleCell(element);
                    }
                });
                return {
                    cellData: cell,
                    rowCount: rowData.length,
                    colCount,
                };
            };
            this.handleDocument = (row, config) => {
                const matchArray = (str, charToFind) => {
                    const regex = new RegExp(charToFind, 'g');
                    let match;
                    const indices = [];
                    while ((match = regex.exec(str))) {
                        indices.push(match.index);
                    }
                    return indices;
                };
                const removeLastChar = (str, charToRemove) => {
                    const regex = new RegExp(`${charToRemove}`, 'g');
                    return str.replace(regex, '\r');
                };
                let pVlaue = null;
                const { v } = row;
                if (typeof v === 'string' || v === null || v === undefined) {
                    return undefined;
                }
                if (v.ct && v.ct.t === 'inlineStr') {
                    v.ct.s = v.ct.s.map(d => {
                        d.v = removeLastChar(d.v || '', '\r\n');
                        return d;
                    });
                    let dataStream = v.ct.s.reduce((prev, cur) => {
                        return prev + cur.v;
                    }, '');
                    dataStream = dataStream?.replace(/\n/g, '\r') + '\r\n';
                    const matchChart = {
                        r: '\r', // PARAGRAPH
                        n: '\n', // SECTION_BREAK
                        v: '\v', // COLUMN_BREAK
                        f: '\f', // PAGE_BREAK
                        '0': '\0', // DOCS_END
                        t: '\t', // TAB
                        b: '\b', // customBlock
                        x1A: '\x1A', // table start
                        x1B: '\x1B', // table row start
                        x1C: '\x1C', // table cell start
                        x1D: '\x1D', // table cell end
                        x1E: '\x1E', // table row end || customRange end
                        x1F: '\x1F', // table end || customRange start
                    };
                    const paragraphs = matchArray(dataStream, matchChart.r).map((d) => {
                        return {
                            startIndex: d,
                        };
                    });
                    const sectionBreaks = matchArray(dataStream, matchChart.n).map((d) => {
                        return {
                            startIndex: d,
                        };
                    });
                    const textRuns = v.ct.s?.map((d, index) => {
                        const start = v.ct.s?.reduce((prev, cur, curi) => {
                            if (curi < index)
                                return prev + (cur.v?.length || 0);
                            return prev;
                        }, 0);
                        const end = start + (v.ct.s?.[index]?.v?.length || 0);
                        const borderConf = config.borderInfo?.find((d) => d.value.col_index === row.c && d.value.row_index === row.r);
                        return {
                            st: start,
                            ed: end,
                            ts: handleStyle({
                                v: (v.ct.s[index] || v.ct.s[0]),
                                r: row.r,
                                c: row.c,
                            }, borderConf, true),
                        };
                    });
                    pVlaue = {
                        id: generateRandomId(6),
                        documentStyle: {
                            documentFlavor: 0,
                            pageSize: { width: 0, height: 0 },
                            renderConfig: {},
                            textStyle: {},
                        },
                        body: {
                            dataStream,
                            paragraphs,
                            sectionBreaks,
                            textRuns,
                        },
                        drawings: {},
                    };
                }
                return pVlaue;
            };
            this.handleCellImage = (row, config) => {
                let pVlaue = null;
                const { v } = row;
                if (typeof v === 'string' || v === null || v === undefined) {
                    return undefined;
                }
                if (v.ct && v.ct.t === 'str' && v.ct.ci) {
                    const blockId = generateRandomId(6);
                    const valueId = generateRandomId(6);
                    const { default: defaultData, src, descr } = v.ct.ci || {};
                    const borderConf = config.borderInfo?.find((d) => d.value.col_index === row.c && d.value.row_index === row.r);
                    pVlaue = {
                        id: valueId,
                        documentStyle: {
                            documentFlavor: 0,
                            pageSize: { width: 0, height: 0 },
                            renderConfig: {},
                            textStyle: {},
                        },
                        body: {
                            dataStream: '\b\r\n',
                            paragraphs: [{
                                    startIndex: 1,
                                    paragraphStyle: { horizontalAlign: v.ht }
                                }],
                            sectionBreaks: [{ startIndex: 2 }],
                            textRuns: [{
                                    ed: 1,
                                    st: 0,
                                    ts: handleStyle({
                                        v: v,
                                        r: row.r,
                                        c: row.c,
                                    }, borderConf, true),
                                }],
                            customBlocks: [{ startIndex: 0, blockId }]
                        },
                        drawings: {
                            [blockId]: {
                                unitId: valueId,
                                subUnitId: valueId,
                                drawingId: blockId,
                                layoutType: core.PositionedObjectLayoutType.INLINE,
                                title: '',
                                description: descr,
                                docTransform: {
                                    size: {
                                        width: defaultData.width,
                                        height: defaultData.height
                                    },
                                    positionH: {
                                        relativeFrom: 0,
                                        posOffset: 0
                                    },
                                    positionV: {
                                        relativeFrom: 1,
                                        posOffset: 0
                                    },
                                    angle: 0
                                },
                                drawingType: core.DrawingTypeEnum.DRAWING_IMAGE,
                                imageSourceType: ImageSourceType.BASE64,
                                source: src,
                                transform: defaultData
                            }
                        },
                        drawingsOrder: [blockId]
                    };
                }
                return pVlaue;
            };
            this.handleRowAndColumnData = (config) => {
                const columnData = {};
                const rowData = {};
                for (let index = 0; index < this.rowCount; index++) {
                    rowData[index] = {
                        h: config.rowlen?.[index] || this.defaultRowHeight,
                        ia: !config.rowlen?.[index] ? core.BooleanNumber.TRUE : core.BooleanNumber.FALSE,
                        ah: this.defaultRowHeight,
                        hd: config.rowhidden?.[index] === 0 ? core.BooleanNumber.TRUE : core.BooleanNumber.FALSE,
                    };
                }
                for (let index = 0; index < this.columnCount; index++) {
                    columnData[index] = {
                        w: config.columnlen?.[index] || this.defaultColumnWidth,
                        hd: config.colhidden?.[index] === 0 ? core.BooleanNumber.TRUE : core.BooleanNumber.FALSE,
                    };
                }
                this.rowData = rowData;
                this.columnData = columnData;
            };
            /**
             * 处理链接
             * @param sheetName IluckysheetHyperlink
             */
            this.handleSheetLink = (hyperlinks) => {
                if (!hyperlinks)
                    return;
                const links = Object.keys(hyperlinks).map((d) => {
                    const row = d.split('_')[0], column = d.split('_')[1];
                    const item = hyperlinks[d];
                    let payload = item.linkAddress;
                    if (item.linkType === 'internal') {
                        const locationList = item.linkAddress.split('!');
                        payload = {};
                        if (locationList[0])
                            payload['gid'] = locationList[0];
                        if (locationList[1])
                            payload['range'] = locationList[1];
                    }
                    return {
                        id: generateRandomId(6),
                        row: Number(row),
                        column: Number(column),
                        payload,
                    };
                });
                this.hyperLink = links;
            };
            this.handleFreeze = (freeze) => {
                this.freeze = {
                    xSplit: freeze.vertical,
                    ySplit: freeze.horizen,
                    startColumn: freeze.vertical,
                    startRow: freeze.horizen,
                };
            };
            const { color, zoomRatio, celldata, config = {}, showGridLines, defaultColWidth, defaultRowHeight, hide, } = sheetData || {};
            this.name = sheetData.name;
            this.id = `sheet-${sheetData.index}`;
            if (sheetData) {
                this.tabColor = color;
                this.zoomRatio = zoomRatio;
                this.showGridlines = Number(showGridLines);
                this.defaultColumnWidth = defaultColWidth;
                this.defaultRowHeight = defaultRowHeight;
                this.hidden = hide;
                this.handleSheetLink(sheetData.hyperlink);
                if (config.merge)
                    this.mergeData = this.handleMerge(config);
                if (celldata?.length) {
                    const { cellData, rowCount, colCount } = this.handleCellData(celldata, config);
                    this.cellData = cellData;
                    this.rowCount = this.rowCount > rowCount ? this.rowCount : rowCount + 1;
                    this.columnCount = this.columnCount > colCount ? this.columnCount : colCount + 1;
                }
                console.log(this.rowCount, this.columnCount);
                this.handleRowAndColumnData(config);
                if (sheetData.freezen)
                    this.handleFreeze(sheetData.freezen);
            }
        }
        get mode() {
            return {
                id: this.id,
                name: this.name,
                type: this.type,
                tabColor: this.tabColor,
                hidden: this.hidden,
                freeze: this.freeze,
                rowCount: this.rowCount,
                columnCount: this.columnCount,
                zoomRatio: this.zoomRatio,
                scrollTop: this.scrollTop,
                scrollLeft: this.scrollLeft,
                defaultColumnWidth: this.defaultColumnWidth,
                defaultRowHeight: this.defaultRowHeight,
                mergeData: this.mergeData,
                cellData: this.cellData,
                rowData: this.rowData,
                columnData: this.columnData,
                rowHeader: this.rowHeader,
                columnHeader: this.columnHeader,
                showGridlines: this.showGridlines,
                rightToLeft: this.rightToLeft,
                selections: this.selections,
                hyperLink: this.hyperLink,
            };
        }
    }

    class UniverWorkBook {
        constructor(file) {
            this.resources = [];
            this.handleHyperLinks = (workSheets) => {
                const hyperLinks = {};
                for (const key in workSheets) {
                    const link = workSheets[key].hyperLink;
                    if (!link?.length)
                        continue;
                    hyperLinks[key] = link.map((d) => {
                        let payload = d.payload;
                        if (typeof d.payload !== 'string') {
                            payload = '#';
                            const gid = d.payload.gid.replace(/'|"/g, '');
                            const sheetId = Object.values(workSheets).find((sheet) => sheet.name === gid)?.id;
                            if (gid && sheetId) {
                                payload += `gid=${sheetId}`;
                            }
                            if (gid && sheetId && d.payload.range)
                                payload += '&';
                            if (d.payload.range)
                                payload += `range=${d.payload.range}`;
                        }
                        return {
                            ...d,
                            payload,
                        };
                    });
                }
                // console.log(workSheets, hyperLinks)
                this.resources?.push({
                    name: 'SHEET_HYPER_LINK_PLUGIN',
                    data: JSON.stringify(hyperLinks),
                });
            };
            this.handleImage = (workSheets, sheets) => {
                const drawerList = {};
                Object.values(workSheets).forEach((sheet) => {
                    const images = sheets.find((d) => d.name === sheet.name)?.images;
                    if (!images)
                        return;
                    const order = Object.keys(images);
                    const data = {};
                    order.forEach((key) => {
                        const image = images[key];
                        if (sheet.columnCount < image.toCol) {
                            sheet.columnCount = image.toCol + 1;
                        }
                        if (sheet.rowCount < image.toRow) {
                            sheet.rowCount = image.toRow + 1;
                        }
                        let imageObj = {
                            unitId: this.id,
                            subUnitId: sheet.id || '',
                            drawingId: key,
                            transform: {
                                width: 0,
                                height: 0,
                                scaleX: 0,
                                scaleY: 0,
                                left: 0,
                                top: 0,
                                angle: 0,
                                skewX: 0,
                                skewY: 0,
                                flipX: false,
                                flipY: false,
                                ...(image.transform || {}),
                            },
                            sheetTransform: {
                                angle: 0,
                                skewX: 0,
                                skewY: 0,
                                flipX: false,
                                flipY: false,
                                from: {
                                    column: image.fromCol,
                                    columnOffset: image.fromColOff,
                                    row: image.fromRow,
                                    rowOffset: image.fromRowOff,
                                },
                                to: {
                                    column: image.toCol,
                                    columnOffset: image.toColOff,
                                    row: image.toRow,
                                    rowOffset: image.toRowOff,
                                },
                            },
                        };
                        if (image.type === 'chart') {
                            imageObj = {
                                ...imageObj,
                                drawingType: core.DrawingTypeEnum.DRAWING_CHART,
                                componentKey: 'Chart',
                                data: {
                                    ...(image.data || {}),
                                    range: `${sheet.name}!${image.data.range}`
                                },
                                allowTransform: true
                            };
                        }
                        else {
                            imageObj = {
                                ...imageObj,
                                drawingType: core.DrawingTypeEnum.DRAWING_IMAGE,
                                imageSourceType: ImageSourceType.BASE64,
                                source: image.src,
                                prstGeom: 'rect',
                                anchorType: '1',
                            };
                        }
                        data[key] = imageObj;
                    });
                    drawerList[sheet.id] = {
                        data,
                        order,
                    };
                });
                this.resources?.push({
                    name: 'SHEET_DRAWING_PLUGIN',
                    data: JSON.stringify(drawerList),
                });
            };
            this.handleChart = (workSheets, sheets) => {
                const chartList = {};
                Object.values(workSheets).forEach((sheet) => {
                    const charts = sheets.find((d) => d.name === sheet.name)?.charts;
                    if (!charts)
                        return;
                    charts.forEach((chart) => {
                        if (!chartList[sheet.id]) {
                            chartList[sheet.id] = [];
                        }
                        chartList[sheet.id].push({
                            rangeInfo: {
                                isRowDirection: chart.isRowDirection,
                                rangeInfo: {
                                    unitId: this.id,
                                    subUnitId: sheet.id || '',
                                    range: handleRanges(chart.range)[0]
                                }
                            },
                            id: chart.id,
                            chartType: chart.chartType,
                            context: chart.context,
                            style: chart.style,
                            dataAggregation: {}
                        });
                    });
                });
                // console.log('chartList', chartList)
                this.resources?.push({
                    name: 'SHEET_CHART_PLUGIN',
                    data: JSON.stringify(chartList),
                });
            };
            this.handleNames = (workbook) => {
                this.resources?.push({
                    name: 'SHEET_DEFINED_NAME_PLUGIN',
                    data: JSON.stringify(workbook.defineNames),
                });
            };
            this.handleCondition = (sheets) => {
                const obj = {};
                Object.keys(sheets).forEach((d) => {
                    const condition = sheets[d].conditionalFormatting?.map((d) => {
                        if (d.rule?.style) {
                            d.rule.style = handleStyle({ v: d.rule.style, r: 0, c: 0 }, { value: d.rule?.style?.border, rangeType: '' });
                        }
                        return d;
                    });
                    obj[d] = condition;
                });
                this.resources?.push({
                    name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
                    data: JSON.stringify(obj),
                });
            };
            this.handleVerification = (sheets) => {
                const obj = {};
                Object.keys(sheets).forEach((d) => {
                    obj[d] = sheets[d].dataVerificationList;
                });
                this.resources?.push({
                    name: 'SHEET_DATA_VALIDATION_PLUGIN',
                    data: JSON.stringify(obj),
                });
            };
            this.handleFilter = (sheets) => {
                const obj = {};
                Object.keys(sheets).forEach((d) => {
                    obj[d] = sheets[d].filter;
                });
                this.resources?.push({
                    name: 'SHEET_FILTER_PLUGIN',
                    data: JSON.stringify(obj),
                });
            };
            const { info, sheets, workbook } = file;
            this.id = generateRandomId(6);
            this.name = info.name;
            this.appVersion = info.appversion;
            this.locale = core.LocaleType.ZH_CN;
            const workSheets = {}, order = [], sheetsObj = {};
            sheets
                .sort((a, b) => Number(a.order) - Number(b.order))
                .forEach((d) => {
                const sheet = new UniverSheet(d);
                workSheets[sheet.id] = sheet.mode;
                sheetsObj[sheet.id] = d;
                order.push(sheet.id);
            });
            // console.log(workSheets,sheets)
            this.handleHyperLinks(workSheets);
            this.handleImage(workSheets, sheets);
            this.handleChart(workSheets, sheets);
            this.handleNames(workbook);
            this.handleCondition(sheetsObj);
            this.handleVerification(sheetsObj);
            this.handleFilter(sheetsObj);
            this.sheetOrder = order;
            this.sheets = workSheets;
        }
        get mode() {
            return {
                id: this.id,
                rev: this.rev,
                name: this.name,
                appVersion: this.appVersion,
                locale: this.locale,
                styles: this.styles,
                sheetOrder: this.sheetOrder,
                sheets: this.sheets,
                resources: this.resources,
            };
        }
    }

    var CHARSET_TYPE;
    (function (CHARSET_TYPE) {
        CHARSET_TYPE["UTF8"] = "UTF-8";
        CHARSET_TYPE["GBK"] = "GBK";
        CHARSET_TYPE["CP936"] = "CP936";
        CHARSET_TYPE["ISO8859"] = "ISO-8859";
    })(CHARSET_TYPE || (CHARSET_TYPE = {}));
    const getDataByFile = ({ file, charset = CHARSET_TYPE.UTF8 }) => {
        return new Promise((resolve, reject) => {
            if (!(file instanceof File))
                resolve('');
            const fileReader = new FileReader();
            fileReader.onload = (e) => {
                try {
                    const { result } = e.target;
                    // 老的逻辑，.txt、.log、.csv 直接用 readAsText 读取后返回即可
                    // 新的 xlsx、xls 用 readAsBinaryString 读取处理数据
                    resolve(result);
                }
                catch (err) {
                    reject(err);
                }
            };
            // 二进制文件通过 readAsBinaryString 读取
            fileReader.readAsText(file, charset);
        });
    };
    const formatSheetData = (sheetData, file) => {
        const splitData = sheetData.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
        let arr = [];
        if (file.name.endsWith('.csv') || file.type === 'text/csv') {
            const csvData = Papa__default["default"].parse(sheetData, {
                delimiter: ',',
                skipEmptyLines: true,
            });
            arr = csvData.data;
        }
        else {
            for (let i = 0; i < splitData.length; i++) {
                const str = splitData[i].replace(/\r/, ''); // 清除无用\r字符
                if (str) {
                    arr.push(str.split(','));
                }
            }
        }
        return arr;
    };

    class UniverCsvWorkBook {
        constructor(data) {
            console.log(data);
            const cellData = {};
            let rowCount = 0, colCount = 0;
            data.forEach((row, rowIndex) => {
                if (rowIndex + 1 > rowCount)
                    rowCount = rowIndex + 1;
                row.forEach((col, colIndex) => {
                    if (colIndex + 1 > colCount)
                        colCount = colIndex + 1;
                    if (!cellData[rowIndex])
                        cellData[rowIndex] = {};
                    cellData[rowIndex][colIndex] = { v: col || '' };
                });
            });
            const sheetId = `sheet1`;
            const sheet = new UniverSheetBase({ id: sheetId, name: sheetId, cellData, rowCount, colCount });
            this.sheets = { [sheetId]: sheet };
            this.sheetOrder = [sheetId];
            this.id = generateRandomId(6);
            this.name = this.id;
        }
        get mode() {
            return {
                id: this.id,
                rev: this.rev,
                name: this.name,
                appVersion: this.appVersion,
                locale: this.locale,
                styles: this.styles,
                sheetOrder: this.sheetOrder,
                sheets: this.sheets,
                resources: this.resources,
            };
        }
    }

    class LuckyExcel {
        constructor() { }
        static transformExcelToLucky(excelFile, callback, errorHandler) {
            let handleZip = new HandleZip(excelFile);
            // const fileReader = new FileReader();
            // fileReader.onload = async (e) => {
            //     const { result } = e.target as any;
            //     const workbook = new exceljs.Workbook();
            //     const data = await workbook.xlsx.load(result);
            //     // console.log('exceljs', data)
            // }
            // fileReader.readAsArrayBuffer(excelFile)
            handleZip.unzipFile(function (files) {
                let luckyFile = new LuckyFile(files, excelFile.name);
                let luckysheetfile = luckyFile.Parse();
                let exportJson = JSON.parse(luckysheetfile);
                // console.log('output---->', exportJson)
                if (callback != undefined) {
                    callback(exportJson, luckysheetfile);
                }
            }, function (err) {
                if (errorHandler) {
                    errorHandler(err);
                }
                else {
                    console.error(err);
                }
            });
        }
        static transformExcelToLuckyByUrl(url, name, callBack, errorHandler) {
            let handleZip = new HandleZip();
            handleZip.unzipFileByUrl(url, function (files) {
                let luckyFile = new LuckyFile(files, name);
                let luckysheetfile = luckyFile.Parse();
                let exportJson = JSON.parse(luckysheetfile);
                if (callBack != undefined) {
                    callBack(exportJson, luckysheetfile);
                }
            }, function (err) {
                if (errorHandler) {
                    errorHandler(err);
                }
                else {
                    console.error(err);
                }
            });
        }
        static transformExcelToUniver(excelFile, callback, errorHandler) {
            let handleZip = new HandleZip(excelFile);
            handleZip.unzipFile(function (files) {
                let luckyFile = new LuckyFile(files, excelFile.name);
                let luckysheetfile = luckyFile.Parse();
                let exportJson = JSON.parse(luckysheetfile);
                console.log('output---->', exportJson, files);
                if (callback != undefined) {
                    const univerData = new UniverWorkBook(exportJson);
                    callback(univerData.mode, luckysheetfile);
                }
            }, function (err) {
                if (errorHandler) {
                    errorHandler(err);
                }
                else {
                    console.error(err);
                }
            });
        }
        static transformCsvToUniver(file, callback, errorHandler) {
            try {
                getDataByFile({ file }).then((source) => {
                    const sheetData = formatSheetData(source, file);
                    const univerData = new UniverCsvWorkBook(sheetData || []);
                    callback?.(univerData.mode, sheetData);
                });
            }
            catch (error) {
                errorHandler(error);
            }
        }
        static async transformUniverToExcel(params) {
            const { snapshot, fileName = `excel_${(new Date).getTime()}.xlsx`, getBuffer = false, success, error } = params;
            try {
                // console.log(1, new Date())
                const workbook = new WorkBook(snapshot);
                // console.log(snapshot, workbook)
                // console.log(2, new Date())
                const buffer = await workbook.xlsx.writeBuffer();
                // console.log(3, new Date())
                if (getBuffer) {
                    success?.(buffer);
                }
                else {
                    this.downloadFile(fileName, buffer);
                    success?.();
                }
            }
            catch (err) {
                error?.(err);
            }
        }
        static async transformUniverToCsv(params) {
            const { snapshot, fileName = `csv_${(new Date).getTime()}.csv`, getBuffer = false, success, error, sheetName } = params;
            try {
                const csv = new CSV(snapshot);
                console.log(csv);
                let contents;
                if (sheetName) {
                    contents = csv.csvContent[sheetName];
                }
                else {
                    contents = csv.csvContent;
                }
                if (getBuffer) {
                    success?.(contents);
                }
                else {
                    if (isObject(contents)) {
                        for (const key in contents) {
                            if (Object.prototype.hasOwnProperty.call(contents, key)) {
                                const element = contents[key];
                                this.downloadFile(`${fileName}_${key}`, element);
                            }
                        }
                    }
                    else {
                        this.downloadFile(fileName, contents);
                    }
                    success?.();
                }
            }
            catch (err) {
                error(err);
            }
        }
        static downloadFile(fileName, buffer) {
            const link = document.createElement('a');
            let blob;
            if (typeof buffer === 'string') {
                blob = new Blob([buffer], { type: "text/csv;charset=utf-8;" });
            }
            else {
                blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
            }
            const url = URL.createObjectURL(blob);
            link.href = url;
            link.download = fileName;
            document.body.appendChild(link);
            link.click();
            link.addEventListener('click', () => {
                link.remove();
                setTimeout(() => {
                    URL.revokeObjectURL(url);
                }, 200);
            });
        }
    }

    return LuckyExcel;

}));
