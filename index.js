const fs = require('fs');
const qrCode = require('qrcode');
const Jimp = require('jimp');
const XLSX = require('xlsx');

function toInches(cm) {
    return cm * 0.393701;
}

function genQRCode(qrCodeImg, text, callback, textData) {
    qrCode.toFile(qrCodeImg, text, {
        color: {
            dark: '#000000',  // Black dots
            light: '#FFF' // Transparent background
        },
        margin: 0,
        width: qrCodeImgWidth
    }, function (err) {
        if (err) throw err
        console.log('QRcode Generated');
        callback(qrCodeImg, textData);
    })
}
function createWholeImg(qrCodeImg, textData) {
    new Jimp(bgSize.width, bgSize.height, (err, bg) => {
        Jimp.read(qrCodeImg)
            .then(qrCodeImg => {
                return bg.composite(qrCodeImg, (bgSize.width - qrCodeImgWidth) / 2, qrCodeImgPositionFromTop);
            })

            //load font	
            .then(wholeImg => (
                Jimp.loadFont(defaultFont).then(font => ([wholeImg, font]))
            ))

            //add footer text
            .then(data => {

                let wholeImg = data[0];
                let font = data[1];
                let extraLineSpacing = 0;


                wholeImg.print(font, textData.placementX, textData.placementY, {
                    text: textData.textLine1,
                    alignmentX: Jimp.HORIZONTAL_ALIGN_CENTER
                }, textData.maxWidth, textData.maxHeight,
                );

                if (textData.textLine1.length >= 20) {
                    extraLineSpacing += lineSpacing;
                }

                wholeImg.print(font, textData.placementX, textData.placementY + lineSpacing + extraLineSpacing, {
                    text: textData.textLine2,
                    alignmentX: Jimp.HORIZONTAL_ALIGN_CENTER
                }, textData.maxWidth, textData.maxHeight,
                );
                return wholeImg.print(font, textData.placementX, textData.placementY + lineSpacing * 2 + extraLineSpacing, {
                    text: textData.textLine3,
                    alignmentX: Jimp.HORIZONTAL_ALIGN_CENTER
                }, textData.maxWidth, textData.maxHeight,
                );
            })

            .then(wholeImg => (wholeImg.quality(100).write(qrCodeImg)))
            // //export image

            .then(wholeImg => {
                //log exported filename
                console.log('exported file: ' + qrCodeImg);
            })

            //catch errors
            .catch(err => {
                console.error(err);
            });


    });
}

//Convert Excel Sheet to JSON format
function readExcel(xlsx) {
    const wb = XLSX.readFile(xlsx);
    let ws1 = wb.Sheets[wb.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(ws1);
    return json;
}


let qrCodeImg = { width: 400, PositionFromTop: 100 };
let qrCodeImgWidth = 400;
let qrCodeImgPositionFromTop = 100;
let lineSpacing = 40;
let defaultFont = Jimp.FONT_SANS_32_BLACK;
let numOfLines = 3;

//read the first sheet in the excel file
let data = readExcel('text.xlsx');
console.log(data.length);

//calculate the background image size
let totalTextHeight = lineSpacing * numOfLines;

let bgSize = { height: totalTextHeight + qrCodeImgWidth + qrCodeImgPositionFromTop + 50, width: qrCodeImgWidth * 1.5 };

//generate QR code and write text as callback
for (let i = 0; i < data.length; i++) {
    let folder = 'QR Code';
    let dirName = `${folder}/G${data[i]['Group']}/`;
    if (!fs.existsSync(folder)) {
        fs.mkdirSync(folder);
    }
    if (!fs.existsSync(dirName)) {
        fs.mkdirSync(dirName);
    }
    let nameOnly = data[i]['Name'].split('\r\n');
    let qrCodeImg = `${dirName}` + nameOnly[0] + '.png';

    //textData.textLineX is the text to be rendered on the image
    //X is index of the line that the text will be rendered on

    let textData = {
        textLine1: `${nameOnly[0]}`, // Name
        textLine2: `${data[i]['Matric No']}`, //Matric No.
        textLine3: `Group ${data[i]['Group']}`, //Group No.
        placementY: qrCodeImgWidth + qrCodeImgPositionFromTop + 10,
        maxWidth: qrCodeImgWidth
    };
    let textLines = [textData.textLine1, textData.textLine2, textData.textLine3];

    textData['placementX'] = (bgSize.width - qrCodeImgWidth) / 2;
    textData['maxHeight'] = bgSize.height - (qrCodeImgWidth + qrCodeImgPositionFromTop),

    //Generate QR Code based on the URL given on the sheet
    genQRCode(qrCodeImg, data[i]['URL'], createWholeImg, textData);
}
