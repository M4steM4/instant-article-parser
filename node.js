var fs = require('fs');
var cheerio = require('cheerio');
var XLSX = require('xlsx');
var workbook = XLSX.readFile('index.xlsx');
var firstSheetName = workbook.SheetNames[0];
var firstSheet = workbook.Sheets[firstSheetName];

var list = [1391317824317039, 1306277106153724, 1881802405475619, 472608033087930, 114765269149546]


function autoTitle(title) {
    for(var i = 0; i < title.length; i++) {
        var a = String.fromCharCode(65 + i);
        firstSheet[a + 3].v = title[i];
    }
}

function autoContent(content) {
    var change = 0;
    for(var i = 0; i < content.length; i++) {
        if(i % 9 == 0 && i != 0) {
            change += 1;
            var a = String.fromCharCode(65 + i - (9 * change));
            if(a == "A") {
                firstSheet[a + (4 + change)].v = content[i].replace(/[^\d]+/g, '');
            } else {
                firstSheet[a + (4 + change)].v = content[i];
            }
        } else {
            var a = String.fromCharCode(65 + i - (9 * change));
            if(a == "A" || a == "B" || a == "C" || a == "D") {
                firstSheet[a + (4 + change)].v = content[i].replace(/[^\d]+/g, '');
            } else if(a == "H" || a == "I") {
                firstSheet[a + (4 + change)].v = content[i].replace(/[^\d]+/g, '');
            } else {
                firstSheet[a + (4 + change)].v = content[i];
            }
        }
    }
}

function file(id) {
    var title = [];
    var content = [];
    fs.readFile(id + '.htm', 'utf8', function(err, data) {
        var $ = cheerio.load(data);

        $('th').each(function() {
            title.push($(this).text());
        });

        $('td').each(function() {
            content.push($(this).text());
        });

        var filename = $('._5ykk').text().split('-');
        firstSheet['A1'].v = filename[0];

        autoTitle(title);
        autoContent(content);
        XLSX.writeFile(workbook, filename[0] + '.xlsx');
    });
}

for(k = 0; k < list.length; k++) {
    file(list[k]);
}
