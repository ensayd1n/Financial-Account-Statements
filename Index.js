const express = require('express');
const { engine } = require('express-handlebars');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const fs = require('fs');

const app = express();
const port = 3000;

// Handlebars motorunun ayarlanması
app.engine('hbs', engine({ extname: '.hbs', defaultLayout: false }));
app.set('view engine', 'hbs');
app.set('views', path.join(__dirname, 'views'));

// Statik dosyaların sunulması
app.use(express.static(path.join(__dirname, 'public')));

// Multer yapılandırmaları
const upload = multer({
    storage: multer.diskStorage({
        destination: (req, file, cb) => {
            if (file.fieldname === 'logo') {
                cb(null, 'public/image_datas');
            } else if (file.fieldname === 'excel_file') {
                cb(null, 'public/excel_datas');
            } else {
                cb(new Error('Unknown field'));
            }
        },
        filename: (req, file, cb) => {
            cb(null, Date.now() + path.extname(file.originalname));
        }
    }),
}).fields([
    { name: 'logo', maxCount: 1 },
    { name: 'excel_file', maxCount: 1 },
]);

const excelDatasDir = path.join(__dirname, 'public', 'excel_datas');

const getLatestExcelFileSync = () => {
    const files = fs.readdirSync(excelDatasDir);
    let latestFile = null;
    let latestMtime = 0;

    files.forEach(file => {
        const filePath = path.join(excelDatasDir, file);
        const stats = fs.statSync(filePath);

        if (stats.mtimeMs > latestMtime) {
            latestMtime = stats.mtimeMs;
            latestFile = file;
        }
    });

    return latestFile ? path.join(excelDatasDir, latestFile) : null;
};

const imageDatasDir = path.join(__dirname, 'public', 'image_datas');
const getLatestImageFileSync = () => {
    const files = fs.readdirSync(imageDatasDir);
    let latestFile = null;
    let latestMtime = 0;

    files.forEach(file => {
        const filePath = path.join(imageDatasDir, file);
        const stats = fs.statSync(filePath);

        if (stats.mtimeMs > latestMtime) {
            latestMtime = stats.mtimeMs;
            latestFile = file;
        }
    });

    return latestFile ? path.join(imageDatasDir, latestFile) : null;
};

const pdfOutputDir = path.join(__dirname, 'public', 'pdf_output');

function getFormattedDate() {
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    return `${day}/${month}/${year}`;
}
const inputDataDir=path.join(__dirname,'public','input_datas');
const getLatestInputFileSync = () => {
    const files = fs.readdirSync(inputDataDir);
    let latestFile = null;
    let latestMtime = 0;

    files.forEach(file => {
        const filePath = path.join(inputDataDir, file);
        const stats = fs.statSync(filePath);

        if (stats.mtimeMs > latestMtime) {
            latestMtime = stats.mtimeMs;
            latestFile = file;
        }
    });

    return latestFile ? path.join(inputDataDir, latestFile) : null;
};
const readJsonFileSync = () => {
    try {
        const jsonData = fs.readFileSync(getLatestInputFileSync(), 'utf-8');
        return JSON.parse(jsonData);
    } catch (error) {
        console.error('JSON dosyası okuma hatası:', error);
        return null;
    }
};
async function convertExcelToPdf(filePath) {
    try {

        const latestInputFile = getLatestInputFileSync();
        const jsonData = readJsonFileSync(latestInputFile);


        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.worksheets[0];

        const pdf = new PDFDocument({ size: 'A4', margin: 50 });
        const outputPath = path.join(pdfOutputDir, `${Date.now()}.pdf`);
        const writeStream = fs.createWriteStream(outputPath);

        pdf.pipe(writeStream);


        pdf.moveDown(3)
        .fontSize(10).text(jsonData.companyName, 100, pdf.y, { align: 'left' })
        .fontSize(8).text(jsonData.companyAddress, 100, pdf.y + 10, { align: 'left' });
        
        
        pdf.moveDown(3)
            .fontSize(12).text('Sayın', { align: 'left' })
            .fontSize(14).font('Helvetica-Bold').text(jsonData.targetCompanyName)
            .fontSize(10).font('Helvetica').text(jsonData.targetCompanyAddress);

        pdf.moveDown(1)
            .fontSize(12).fillColor('#000000').text(`${getFormattedDate()} tarihi hesap özeti ektedir`, { align: 'left' });


        const tableTop = 300;
        pdf.font('Helvetica-Bold').fontSize(10);
        pdf.text('Tarih', 55, tableTop);
        pdf.text('Evrak No', 200, tableTop);
        pdf.text('Aciklama', 275, tableTop);
        pdf.text('Borç', 350, tableTop);
        pdf.text('Alacak', 400, tableTop);
        pdf.text('Bakiye', 500, tableTop);


        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            if (rowNumber > 1) {
                const rowY = tableTop + ((rowNumber - 1) * 20);
                pdf.font('Helvetica').fontSize(6);
        
                pdf.text(row.getCell(1).value , 55, rowY);
                pdf.text(row.getCell(2).value , 200, rowY);
                pdf.text(row.getCell(3).value , 275, rowY);
                pdf.text(row.getCell(4).value , 350, rowY);
                pdf.text(row.getCell(5).value , 400, rowY);
                pdf.text(row.getCell(6).value , 500, rowY);
            }
        });

        const totalPrice = () => {
            let total = 0;
        
            worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                if (rowNumber > 1) {
                    const cellValue = row.getCell(4).value;
                    total += cellValue || 0;
                }
            });
        
            return total;
        }


        pdf.moveDown(2)
            .font('Helvetica-Bold').fontSize(12).text('TOPLAM', 400, pdf.y, { align: 'right' })
            .font('Helvetica').fontSize(12).text(totalPrice() , 400, pdf.y, { align: 'right' });

        pdf.end();
       
    } catch (error) {
        console.error('Hata oluştu:', error);
    }
}


app.post('/document', upload, async (req, res) => {
    const latestFile = getLatestExcelFileSync();
    if (latestFile) {
        console.log('Son eklenen dosya:', latestFile);

        const {'company-name':companyName,'company-address': companyAddress, 'target-company-name': targetCompanyName, 'target-company-address': targetCompanyAddress } = req.body;
        const jsonData = {
            companyName,
            companyAddress,
            targetCompanyName,
            targetCompanyAddress
        };

        const jsonOutputDir = path.join(__dirname, 'public', 'input_datas');
        const jsonFilePath = path.join(jsonOutputDir, `${Date.now()}.json`);
        fs.writeFileSync(jsonFilePath, JSON.stringify(jsonData, null, 2), 'utf-8');

        await convertExcelToPdf(latestFile);

        res.render('document');
    } else {
        console.log('Klasörde dosya bulunamadı.');
        res.status(404).send('Dosya bulunamadı.');
    }
});

app.get('/', (req, res) => {
    res.render('home');
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});