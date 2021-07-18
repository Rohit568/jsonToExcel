    const PORT = 8000;

    const app = require('express')();
    const file = require('./data.json');
    const xl = require('excel4node');
    const { alignment } = require('excel4node/distribution/lib/types');
    var data = JSON.parse(JSON.stringify(file));

    app.get('/', (req, res)=>{
        res.json(file);
    })


    var ws = new xl.Workbook();
        var myStyle = ws.createStyle({
            alignment:{
                horizontal : 'center',
                vertical : 'justify',
                wrapText : true
                }
            });
        const bgStyle = ws.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#FFFF00',
            fgColor:'#FFFF00'
        },
        alignment:{
            horizontal : 'center'
        }
        });
        const bgStyle1 = ws.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            horizontal:'center',
            bgColor: '#FFCCBB',
            fgColor:'#FFCCBB'
        },
        alignment:{
            horizontal : 'center'
        }
        });
    
        var headStyle = ws.createStyle({
        font:{
            bold: true,
            },
        alignment: {
            
            horizontal: 'center',
        },
        });

    app.get('/download',(req, res)=>{

        
        var wb = ws.addWorksheet('Sheet 1');
            wb.cell(1, 1).string('Trade No.').style(headStyle);
            wb.cell(1, 2).string('Lots').style(headStyle);
            wb.cell(1, 3).string('Legs').style(headStyle);
            wb.cell(1, 4).string('Entry Date').style(headStyle);
            wb.cell(1, 5).string('Strike').style(headStyle);
            wb.cell(1, 6).string('B/S').style(headStyle);
            wb.cell(1, 7).string('Options').style(headStyle);
            wb.cell(1, 8).string('Entry Price').style(headStyle);
            wb.cell(1, 9).string('Exit Date').style(headStyle);
            wb.cell(1, 10).string('Exit Price').style(headStyle);
            wb.cell(1, 11).string('Days').style(headStyle);
            wb.cell(1, 12).string('Profit').style(headStyle);
            wb.cell(1, 13).string('Total Profit').style(headStyle);
            let x = 1;
            let y = 1;
            let trade = 1;
            var curData = data.data;
            var totalProfit = 0;
            var maxY = 0;
            for(let i = 0; i<12; i++)
            {  
                ++x;
                y=1;
                wb.cell(x,y).number(trade++);
                
                y++;
                totalProfit = 0;

                for(let j = 0; j<6; j++)
                {
                    y = 2;
                    if(j!=0)
                    ++x;
                    if(typeof curData[i].legs[j] !== 'undefined'){
                        var k = curData[i].legs[j];
                        let lots  = k.lots;
                        let ent_date = k.entryDate;
                        let strike = k.strikePrice;
                        let bos = k.buyOrSell;
                        let options = k.futuresOrOptions;
                        let entry_value = k.entryValue;
                        let ex_date = k.exitDate;
                        let exit_value = k.exitValue;
                        let entry_date = formateDate(ent_date);
                        let exit_date = formateDate(ex_date);
                        let profit = calculateProfit(exit_value,entry_value,lots);
                        let days = calculateDays(ex_date, ent_date);
                        totalProfit =profit +totalProfit;
                        wb.cell(x, y++).number(lots).style(myStyle);
                        wb.cell(x, y++).number(j+1).style(myStyle);
                        wb.cell(x, y++).string(entry_date).style(myStyle);
                        wb.cell(x, y++).number(strike).style(myStyle);
                        wb.cell(x, y++).string(bos).style(myStyle);
                        wb.cell(x, y++).string(options).style(myStyle);
                        wb.cell(x, y++).number(entry_value).style(bgStyle);
                        wb.cell(x, y++).string(exit_date).style(myStyle);
                        wb.cell(x, y++).number(exit_value).style(bgStyle);
                        wb.cell(x, y++).number(days).style(bgStyle);
                        wb.cell(x, y++).number(profit).style(bgStyle1); 
                        if(maxY < y)
                        maxY = y;
                }
                
                }
                wb.cell(x, maxY).number(totalProfit).style(bgStyle1);
        
            }
            ws.write('excel.xlsx', res);
    });


    function calculateProfit(x, y, z){
        return (x - y) * z * 75;
    }
    function calculateDays(exit_date, entry_date)
    {
            var date1 = new Date(entry_date);
            var date2 = new Date(exit_date);
            var diffDays = date2.getDate() - date1.getDate(); 
           // console.log("*****"+ diffDays);
            return diffDays;
    }

    function formateDate(date)
    {
        var d = new Date(date);
        d= ""+d;
        //console.log(d);
        const day = {
                        'Mon' : 'Monday',
                        'Tue' : 'Tuesday',
                        'Wed':'Wednesday',
                        'Thu':'Thursday',
                        'Fri':'Friday',
                        'Sat':'Saturday',
                        'Sun':'Sunday',
                        };
        const month = {'Jan': 'January', 'Feb':'February',
                        'Mar':'March', 'Apr':'April',
                        'May':'May', 'Jun':'June',
                        'Jul':'July', 'Aug':'August',
                        'Sept':'September', 'Oct': 'October',
                        'Nov':'November', 'Dec':'December'
                    };
        var s = d.split(" ");
        let o = day[s[0]];
        let m = month[s[1]];
        let l  = s[2];
        let y = s[3]

        var z = o + ", " + m + " "+l+ ", " + y;
        return z;  
    }

    app.listen(PORT);