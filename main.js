let puppeteer=require("puppeteer");
let fs=require("fs");
let path=require("path");
let link="https://ipuranklist.com/";
let xlsx=require("xlsx");

let browserStartPromise=puppeteer.launch({
    headless:false,
    defaultViewport:null,
    args:["--start-maximized","--disable-notifications"]
});
let page;
(async function fn() {

    try{
        let browserObj=await browserStartPromise;
    console.log("Browser opened");
    page= await browserObj.newPage();
    await page.goto(link);
    console.log("page opened");
    await waitAndClick("a[href='/university-ranklist/btech'].nav-link",page);
    console.log("ranklist page opened");

    let pFolderLocation=path.join(process.cwd()," IPU University Rankings");

    if(fs.existsSync(pFolderLocation)==false){
        fs.mkdirSync(pFolderLocation);
    }

    let yearOptionArr=["15","16","17","18","19","20"];

    for(let i=0;i<yearOptionArr.length;i++){

        await page.select("#batch",yearOptionArr[i]);
        let year=Number(yearOptionArr[i])+4;
        console.log("20"+yearOptionArr[i]+"-"+"20"+year+ " batch selected");
        
        let folderLocation=path.join(pFolderLocation,"20"+yearOptionArr[i]+"-"+"20"+year+" Batch Results");
    
        if(fs.existsSync(folderLocation)==false){
            fs.mkdirSync(folderLocation);
        }
    
        let branchOptionArr=["CSE","CIV","ECE","EE","EEE","ICE","IT","MAE","ME",];
    
    
        for(let i=0;i<branchOptionArr.length;i++){
            let branchFolderLocation=path.join(folderLocation,branchOptionArr[i]);
            if(fs.existsSync(branchFolderLocation)==false){
                fs.mkdirSync(branchFolderLocation);
            }
            await page.select("#branch",branchOptionArr[i]);
            console.log(branchOptionArr[i]);
            await rankCalculation(page,branchOptionArr[i],branchFolderLocation);
    
        }

    }
    
    }
    catch(err){
        console.log(err);
    }
})();

//rank clculation
function rankCalculation(page,branchName,branchLocation){
    return new Promise(function(resolve,reject){
        (async function fn() {
          try{
            await waitAndClick(".btn.btn-dark",page);
            await page.waitForSelector("tr.ng-star-inserted");
            await page.waitFor(5000);
        
            let data = await page.$$eval('table tr td', tds => tds.map((td) => {
                return td.innerText;
              }));
              
            
            let names=[];
            for(let i=1;i<data.length;i+=5){
                names.push(data[i]);
            }
            // console.table(names);
            
            let gpa=[];
            for(let i=3;i<data.length;i+=5){
                gpa.push(data[i]);
            }
            // console.table(gpa);
        
            let ranks=[];
            for(let i=4;i<data.length;i+=5){
                ranks.push(data[i]);
            }
            // console.table(ranks);

            let fileLocation=path.join(branchLocation,"University Ranklist.xlsx");

            for(let i=0;i<names.length;i++){
                let content=excelReader(fileLocation,branchName);
                let dataObj=contentRead(names[i],gpa[i],ranks[i]);
                content.push(dataObj);
                console.log(content);
                excelWriter(content,fileLocation,branchName);
            }

              resolve();
          }
          catch(err){
              reject(err);
          }
  
        })();    
      })
}

//wait and click
function waitAndClick(selector,cPage){
    return new Promise(function(resolve,reject){
      (async function fn() {
        try{
            await cPage.waitForSelector(selector,{visible:true});
            await cPage.click(selector,{delay:1000});
            resolve();
        }
        catch(err){
            reject(err);
        }

      })();  
        
    })
}

function contentRead(NameOfStudent,GPA,Rank){
    let data={
        NameOfStudent,
        GPA,
        Rank
    }
    
    return data;
}

//excel writer
function excelWriter(jsonData,location,branchName){
    let newWb=xlsx.utils.book_new();
    let newWS=xlsx.utils.json_to_sheet(jsonData);
    xlsx.utils.book_append_sheet(newWb,newWS,branchName);
    xlsx.writeFileSync(newWb,location);
}
//excel Reader
function excelReader(location,sheetName){
    if(fs.existsSync(location)==false){
        return [];
    }
    let wb=xlsx.readFile(location);
    let excelData=wb.Sheets[sheetName];
    let ans=xlsx.utils.sheet_to_json(excelData);
    return ans;
}