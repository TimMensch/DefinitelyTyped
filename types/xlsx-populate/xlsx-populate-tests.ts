import * as XlsxPopulate from "./index";

const workbook = XlsxPopulate.fromFileAsync("foo"); // $ExpectType Promise<Workbook>

workbook.then(wb => {
    const sheet = wb.sheet(0);
    if (!sheet) return;
    sheet; // $ExpectType Sheet
});
