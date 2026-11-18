import { chromium } from 'playwright';
import ExcelJS from 'exceljs';

interface Company {
    page: number;
    stt: string;
    name: string;
    industry: string;
    address: string;
    phone: string;
    hotline: string;
    lastUpdate: string;
}


(async () => {
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();

    const baseUrl = 'https://trangvangvietnam.com/cateprovinces/81960/c%C3%B4ng-ty-s%E1%BA%A3n-xu%E1%BA%A5t-c%E1%BB%ADa-%E1%BB%9F-t%E1%BA%A1i-b%C3%ACnh-d%C6%B0%C6%A1ng.html?page=';

    const companies: Company[] = [];

    // Lấy số trang cuối thực sự
    await page.goto(baseUrl + '1');
    await page.waitForSelector('#paging', { state: 'attached' });

    const pageLinks = await page.$$eval('#paging a', links =>
        links
            .map(l => l.textContent?.trim())
            .filter(t => t && !isNaN(parseInt(t)))
            .map(t => parseInt(t!))
    );
    const totalPages = pageLinks.length > 0 ? Math.max(...pageLinks) : 1;
    console.log('Tổng số trang:', totalPages);

    for (let p = 1; p <= totalPages; p++) {
        console.log(`Đang scrape trang ${p}...`);
        await page.goto(baseUrl + p);
        await page.waitForSelector('xpath=/html/body/div[2]/div/div[5]/div[2]');

        const companyElements = await page.$$('xpath=/html/body/div[2]/div/div[5]/div[2]/div');

        for (const el of companyElements) {
            try {
                // Bỏ qua công ty có cảnh báo "không chính xác"
                const invalidInfo = await el.$('xpath=.//small/span[contains(text(), "Thông tin này đã không còn chính xác")]');
                if (invalidInfo) continue;

                // STT trên web
                const sttEl = await el.$('xpath=.//div[contains(@class,"stt_txt")]');
                const stt = sttEl ? (await sttEl.textContent())?.trim() || '' : '';

                // Tên
                const name = (await el.$eval('xpath=.//h2/a', node => node.textContent?.trim())) || '';

                // Ngành
                const industryEl = await el.$('xpath=.//span[contains(@class,"nganh_listing_txt")]');
                const industry = industryEl ? (await industryEl.textContent())?.trim() || '' : '';

                // Địa chỉ
                let address = '';
                const addressEl1 = await el.$('xpath=.//div[contains(@class,"pt-0 pb-2 ps-3 pe-4")]/small');
                if (addressEl1) {
                    address = (await addressEl1.textContent())?.trim() || '';
                }
                if (!address) {
                    const addressEl2 = await el.$('xpath=.//div[contains(@class,"listing_diachi_nologo")]//small');
                    if (addressEl2) {
                        address = (await addressEl2.textContent())?.trim() || '';
                    }
                }

                // Phone
                const phoneEl = await el.$('xpath=.//i[contains(@class,"fa-phone-volume")]/following-sibling::a');
                const phone = phoneEl ? (await phoneEl.textContent())?.trim() || '' : '';

                // Hotline
                const hotlineEl = await el.$('xpath=.//i[contains(@class,"fa-mobile-screen-button")]/following-sibling::span/a');
                const hotline = hotlineEl ? (await hotlineEl.textContent())?.trim() || '' : '';

                // Ngày cập nhật gần nhất (nếu có)
                const lastUpdateEl = await el.$('xpath=.//small[contains(text(), "Ngày cập nhật gần nhất")]');
                const lastUpdate = lastUpdateEl ? (await lastUpdateEl.textContent())?.replace('Ngày cập nhật gần nhất:', '').trim() || '' : '';

                if (name) companies.push({ page: p, stt, name, industry, address, phone, hotline, lastUpdate });

            } catch (err) {
                continue;
            }
        }
    }

    // Xuất Excel với style đẹp
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Companies');

    sheet.columns = [
        { header: 'Page', key: 'page', width: 10 },
        { header: 'STT', key: 'stt', width: 10 },
        { header: 'Tên công ty', key: 'name', width: 40 },
        { header: 'Ngành nghề', key: 'industry', width: 30 },
        { header: 'Địa chỉ', key: 'address', width: 50 },
        { header: 'SĐT', key: 'phone', width: 20 },
        { header: 'Hotline', key: 'hotline', width: 20 },
        { header: 'Ngày cập nhật gần nhất', key: 'lastUpdate', width: 25 },
    ];


    // Style header
    sheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1F4E78' },
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
        };
    });

    // Thêm dữ liệu với border & wrapText
    companies.forEach(c => {
        const row = sheet.addRow(c);
        row.eachCell(cell => {
            cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
            };
        });
    });

    await workbook.xlsx.writeFile('companies.xlsx');
    console.log('Đã lưu companies.xlsx với', companies.length, 'công ty');

    await browser.close();
})();
