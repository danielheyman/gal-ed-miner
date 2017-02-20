var fs = require('fs');
var cheerio = require('cheerio');
var request = require('request');
var async = require('async');
var excelbuilder = require('msexcel-builder');

var jsonfile = require('jsonfile');
var resultsFile = function(id) { 
    return './results/' + id + '.json'; 
};

function scraper(counter, callback) {
    var path = resultsFile(counter);
    if (!fs.existsSync(path)) {
        request('http://www.gal-ed.co.il/nachal/info/n_show.aspx?id=' + counter, function (error, response, html) {
            if (error || response.statusCode !== 200) return callback('Error: ' + error, null);
            
            var $ = cheerio.load(html);
            
            var data = null;
            
            if(html.slice(0, 5) !== 'Nofel') {
                data = {
                    name: $("#TD0").text().trim().replace(/\n|\r|\t/g,''),
                    date: $("#TD2").text().match(/[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}/g),
                    age: $("#TD2 td[align=center]").text().match(/\s[0-9]{1,3}\s/g)
                };
                data.date = data.date ? data.date.slice(-1)[0] : null;
                data.age = data.age ? data.age[0].trim() : null;
            }
            
            console.log(counter);
            
            jsonfile.writeFile(resultsFile(counter), data, function (err) { });
            
            setTimeout(function() {
                callback(null, null);
            }, 300);
        });
    }
    else callback(null, null);
}

var batch = 50;


function run(counter) {
    if(counter > 42500) return count();
    
    var range = Array.apply(null, {length: batch}).map(Number.call, function(n) { return Number(n) + counter; });
    async.map(range, scraper,
        function(err, results) {
            if(err) {
                console.log(err);
                console.log('Delaying 60 seconds.');
                setTimeout(function() {
                    run(counter);
                }, 60 * 1000);
            } 
            else run(counter + batch);
        }
    );

}

function count() {
    const fs = require('fs');
    fs.readdir('./results/', (err, files) => {
        var results = [];
        files.forEach(file => {
            try {
                var obj = jsonfile.readFileSync('./results/' + file);
                if(obj !== null && obj.name && obj.date && obj.age) results.push(obj);
            } catch(e) {}
        });
        
        results = results.sort(function(a, b) {
            if(a.date === null) return 1;
            if(b.date === null) return -1;
            return Date.parse(a.date.split('/').reverse().join('/')) - Date.parse(b.date.split('/').reverse().join('/'));
        });
        
        var workbook = excelbuilder.createWorkbook('./', 'output.xlsx');
        var sheet1 = workbook.createSheet('sheet1', 3, results.length + 1);
        var row = 2;
        
        sheet1.set(1, 1, 'Name');
        sheet1.set(2, 1, 'Date');
        sheet1.set(3, 1, 'Age');

        results.forEach(res => {
            sheet1.set(1, row, res.name);
            sheet1.set(2, row, res.date);
            sheet1.set(3, row++, res.age);
        });
            
        workbook.save(function(ok) {
            console.log('Workbook saved with ' + results.length + ' results!');
        });
    });
}

run(1);
