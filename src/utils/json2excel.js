import { saveAs } from "file-saver";
import XLSX from "xlsx";

function datenum(v, date1904) {
    if (date1904) v += 1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data) {
    let ws = {};
    let range = {
        s: {
            c: 10000000,
            r: 10000000,
        },
        e: {
            c: 0,
            r: 0,
        },
    };
    for (let R = 0; R != data.length; ++R) {
        for (let C = 0; C != data[R].length; ++C) {
            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;
            let cell = {
                v: data[R][C],
            };
            if (cell.v == null) continue;
            let cell_ref = XLSX.utils.encode_cell({
                c: C,
                r: R,
            });

            if (typeof cell.v === "number") cell.t = "n";
            else if (typeof cell.v === "boolean") cell.t = "b";
            else if (cell.v instanceof Date) {
                cell.t = "n";
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            } else cell.t = "s";

            ws[cell_ref] = cell;
        }
    }
    if (range.s.c < 10000000) ws["!ref"] = XLSX.utils.encode_range(range);
    return ws;
}

function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    let view = new Uint8Array(buf);
    for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
}

export function export_json_to_excel(opts = {}) {
    /* 拼接数据 */
    //{ multiHeader = [], header, data, filename, merges = [], autoWidth = true, bookType = "xlsx" }
    const config = {
        header: [],
        data: [],
        multiHeader: [],
        merges: [],
        filename: "excel-list",
        autoWidth: true,
        bookType: "xlsx",
    };
    const conf = Object.assign(config, opts);
    const filename = conf.filename;
    let data = [...conf.data];
    data.unshift(conf.header);

    for (let i = conf.multiHeader.length - 1; i > -1; i--) {
        data.unshift(conf.multiHeader[i]);
    }

    const ws_name = "SheetJS";
    let wb = new Workbook(),
        ws = sheet_from_array_of_arrays(data);

    if (conf.merges.length > 0) {
        if (!ws["!merges"]) ws["!merges"] = [];
        conf.merges.forEach(item => {
            ws["!merges"].push(XLSX.utils.decode_range(item));
        });
    }

    if (conf.autoWidth) {
        /*设置worksheet每列的最大宽度*/
        const colWidth = data.map(row => {
            return row.map(val => {
                /*先判断是否为null/undefined*/
                if (val == null) {
                    return {
                        wch: 10,
                    };
                } else if (val.toString().charCodeAt(0) > 255) {
                    /*再判断是否为中文*/
                    return {
                        wch: val.toString().length * 2,
                    };
                } else {
                    return {
                        wch: val.toString().length,
                    };
                }
            });
        });
        /*以第一行为初始值*/
        let result = colWidth[0];
        for (let i = 1; i < colWidth.length; i++) {
            for (let j = 0; j < colWidth[i].length; j++) {
                if (result[j]["wch"] < colWidth[i][j]["wch"]) {
                    result[j]["wch"] = colWidth[i][j]["wch"];
                }
            }
        }
        ws["!cols"] = result;
    }

    /* 添加sheet */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    let wbout = XLSX.write(wb, {
        bookType: conf.bookType,
        bookSST: false,
        type: "binary",
    });
    saveAs(
        new Blob([s2ab(wbout)], {
            type: "application/octet-stream",
        }),
        `${filename}.${conf.bookType}`
    );
}

/**
 *
 * @param {*} element 元素选择器
 * @param {*} opts 参数
 */
export function table2excel(element, opts = {}) {
    const config = {
        filename: "excel-list",
        bookType: "xlsx",
    };
    const conf = { ...config, ...opts };
    const bookType = conf.bookType;
    const filename = conf.filename;

    var wb = XLSX.utils.table_to_book(document.querySelector(element));
    var wbout = XLSX.write(wb, { bookType: bookType, bookSST: true, type: "array" });
    try {
        saveAs(new Blob([wbout], { type: "application/octet-stream" }), `${filename}.${bookType}`);
    } catch (e) {
        if (typeof console !== "undefined") console.log(e, wbout);
    }
    return wbout;
}

export function array2excel(element, opts = {}) {
    const config = {
        filename: "excel-list",
        bookType: "xlsx",
    };
    const conf = { ...config, ...opts };
    const bookType = conf.bookType;
    const filename = conf.filename;

    var wb = XLSX.utils.table_to_book(document.querySelector(element));
    var wbout = XLSX.write(wb, { bookType: bookType, bookSST: true, type: "array" });
    try {
        saveAs(new Blob([wbout], { type: "application/octet-stream" }), `${filename}.${bookType}`);
    } catch (e) {
        if (typeof console !== "undefined") console.log(e, wbout);
    }
    return wbout;
}
