var Excel = require('exceljs');
var _ = require('underscore');
//var _SaveExcelPath = "C:\\Users\\611352789\\desktop\\some.xlsx";
var find = require('find');
var pkgInfo =[];
var pkgJsonPath = [];
var argslist =[];
var corruptedCount = 0;
process.argv.forEach(function(val,index){
    argslist.push(val);
});
var __dirname = argslist[2];
var _SaveExcelPath  = argslist[3];
find.file(/\\package.json$/, __dirname, function(files) {
    _.each(files,function(item,index){
        files[index].toString();
        var t = files[index].split('\\');
        var s = '';
        _.each(t,function(oitem,oindex){
             s = s + oitem + "/" ;           
        });
        s = s.slice(0,-1);
        pkgJsonPath[index] = s;
        try{
            var pkgdata = require(s);
            if(pkgdata != null)
            {
                var obj = new Object();
                obj["AuthorName"] = (pkgdata.author && pkgdata.author.name) != null ? pkgdata.author.name : '';
                obj["Version"] =  pkgdata.version != null ? pkgdata.version: 'NotFound';
                obj["homepage"] =  pkgdata.homepage != null ? pkgdata.homepage: s;
                obj["Componentname"] =  pkgdata.name != null ? pkgdata.name: '';
                pkgInfo.push(obj);              
            }
         }
        catch(err){
           corruptedCount++;
           console.log(err.message);
        }
    });
    console.log("Total File Found is :  "+pkgJsonPath.length);
    console.log("No. Of file corrupt is :  "+ (pkgJsonPath.length - pkgInfo.length));   
    CreateWorkBook(pkgInfo);
  });
function CreateWorkBook(pkgInfo){
    console.log("******************Start Writing to Excel****************");
     var workbook = new Excel.Workbook();
     workbook.creator = 'Anshul Rana'; 
     var worksheet = workbook.addWorksheet('Reoprt1', { pageSetup:{paperSize: 9, orientation:'landscape'}});
     worksheet.columns = [
        { header: 'Component  Name', key: 'Componentname', width: 30 },
        { header: 'Version', key: 'Version', width: 30 },
        { header: 'Author Name', key: 'AuthorName', width: 30},
        { header: 'Homepage', key: 'homepage', width: 70 }
    ];
    worksheet.getRow(1).font = { name: 'Arial Black', bold: true };
    _.each(pkgInfo,function(item,index){
        worksheet.addRow({Componentname:item.Componentname, Version: item.Version, AuthorName: item.AuthorName, homepage: item.homepage});           
        workbook.xlsx.writeFile(_SaveExcelPath).then(function() {   });  
     }); 
     console.log("No. of item added to Excel is : "+ pkgInfo.length);
     console.log("******************End Writing to Excel****************");
    }









