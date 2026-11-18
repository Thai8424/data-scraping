import axios, { AxiosResponse } from 'axios';
import * as cheerio from 'cheerio';
import PQueue from 'p-queue';
import ExcelJS from 'exceljs';

/* ================= CONFIG ================= */
const BASE = 'https://hsctvn.com';
const LIST_PATH = '/nam-2025/page-';
const TOTAL_PAGES = 18619;

const LIST_CONCURRENCY = 10;
const DETAIL_CONCURRENCY = 20;
const BATCH_SIZE = 2000;

/* ================= TYPES ================= */
interface Company {
    page: number;
    name: string;
    shortName: string;
    taxCode: string;
    address: string;
    legalRepresentative: string;
    phone: string;
    issueDate: string;
    status: string;
    lastUpdate: string;
    detailUrl: string;
}

/* ================= QUEUES ================= */
const listQueue = new PQueue({ concurrency: LIST_CONCURRENCY });
const detailQueue = new PQueue({ concurrency: DETAIL_CONCURRENCY });

/* ================= EXCEL STREAM ================= */
let fileIndex = 1;
let rowCount = 0;
let workbook: ExcelJS.stream.xlsx.WorkbookWriter | undefined;
let sheet: ExcelJS.Worksheet | undefined;

function createNewExcel() {
    workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: `companies_part_${fileIndex}.xlsx`,
    });

    sheet = workbook.addWorksheet('Companies');

    sheet.columns = [
        { header: 'Page', key: 'page', width: 8 },
        { header: 'Tên công ty', key: 'name', width: 40 },
        { header: 'Tên viết tắt', key: 'shortName', width: 30 },
        { header: 'Mã số thuế', key: 'taxCode', width: 20 },
        { header: 'Địa chỉ', key: 'address', width: 50 },
        { header: 'Đại diện pháp luật', key: 'legalRepresentative', width: 25 },
        { header: 'SĐT', key: 'phone', width: 20 },
        { header: 'Ngày cấp', key: 'issueDate', width: 15 },
        { header: 'Trạng thái', key: 'status', width: 20 },
        { header: 'Cập nhật lần cuối', key: 'lastUpdate', width: 20 },
        { header: 'Detail URL', key: 'detailUrl', width: 40 },
    ];

    rowCount = 0;
}

async function saveRow(company: Company) {
    if (!sheet || !workbook) createNewExcel();

    sheet!.addRow(company).commit();
    rowCount++;

    if (rowCount >= BATCH_SIZE) {
        await sheet!.commit();
        await workbook!.commit();
        console.log(`Saved companies_part_${fileIndex}.xlsx`);
        fileIndex++;
        createNewExcel();
    }
}

/* ================= UTILS ================= */
async function fetchHTML(url: string): Promise<string> {
    const res: AxiosResponse<string> = await axios.get(url, {
        timeout: 30000,
        headers: { 'User-Agent': 'Mozilla/5.0 (fast-scraper)' },
    });
    return res.data;
}

function cleanText(text?: string | null): string {
    return text?.replace(/\s+/g, ' ').trim() ?? '';
}

/* ================= DETAIL PARSER ================= */
async function parseDetail(url: string, pageNumber: number): Promise<void> {
    try {
        const html = await fetchHTML(url);
        const $ = cheerio.load(html);

        const items = $('.box_content ul.hsct li');

        const getByIcon = (icon: string): string => {
            const el = items.filter((_, li) => $(li).find(`i.${icon}`).length > 0);
            return cleanText(el.text().replace(/^[^:]+:/, ''));
        };

        const company: Company = {
            page: pageNumber,
            detailUrl: url,
            name: cleanText($('.box_content ul.hsct li h1').text()),
            shortName: getByIcon('fa-question-circle'),
            taxCode: getByIcon('fa-hashtag'),
            address: getByIcon('fa-map-marker'),
            legalRepresentative: cleanText(
                items.find('i.fa-user-o').parent().find('a').text()
            ),
            phone: getByIcon('fa-phone'),
            issueDate: cleanText(
                items.find('i.fa-calendar').parent().find('a').text()
            ),
            status: getByIcon('fa-info'),
            lastUpdate: cleanText(
                items.find('i.fa-clock-o').parent().find('i').text()
            ),
        };

        await saveRow(company);

    } catch (err) {
        console.error('Detail error:', url, err);
    }
}

/* ================= LIST PARSER ================= */
async function parseList(pageNumber: number): Promise<void> {
    try {
        const html = await fetchHTML(`${BASE}${LIST_PATH}${pageNumber}`);
        const $ = cheerio.load(html);

        $('div > ul > li').each((_, li) => {
            const href = $(li).find('h3 a').attr('href');
            if (!href) return;

            const detailUrl = href.startsWith('http') ? href : `${BASE}/${href}`;
            detailQueue.add(() => parseDetail(detailUrl, pageNumber));
        });

        console.log(`Parsed list page ${pageNumber}`);
    } catch (err) {
        console.error(`List error page ${pageNumber}`, err);
    }
}

/* ================= MAIN ================= */
async function run() {
    console.log('Start scraping...');

    for (let page = 1; page <= TOTAL_PAGES; page++) {
        listQueue.add(() => parseList(page));
    }

    await listQueue.onIdle();
    await detailQueue.onIdle();

    if (sheet && workbook && rowCount > 0) {
        await sheet.commit();
        await workbook.commit();
        console.log(`Saved companies_part_${fileIndex}.xlsx (last)`);
    }

    console.log('DONE.');
}

run().catch(console.error);
