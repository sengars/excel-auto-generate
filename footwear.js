const express=require('express');
const xl =require('excel4node');
const app=express();

var wb=new xl.Workbook();

const arr=[{ "displayName": "Merchant_Sku_id", "color": "#00BFFF","type": "text"},
{ "displayName": "Seller Name", "color": "#00BFFF","type": "text"},
{ "displayName": "Brand Name", "color": "#00BFFF","type": "text"},
{ "displayName": "Product Title", "color": "#00BFFF","type": "text"},
{ "displayName": "MRP", "color": "#00BFFF","type": "text"},
{ "displayName": "Tp with GST", "color": "#00BFFF","type": "text"},
{ "displayName": "8 Digit HSN", "color": "#00BFFF","type": "text"},
{ "displayName": "Size", "color": "#00BFFF","type": "keyword", 
"values": [ "8", "9", "10" ]},
{ "displayName": "Color", "color": "#00BFFF","type": "keyword", 
"values": [ "Black", "White", "Blue" ]},
{ "displayName": "P_Sub_type", "color": "#00BFFF","type": "keyword", 
"values": [ "X", "Y", "Z" ]}];



var ws1=wb.addWorksheet('Sheet1');
var ws2=wb.addWorksheet('Sheet2');
var style = wb.createStyle({
    font: {
      bold: true,
      color: '#000000',
      size: 12,
    },
  });

  var colNumber=1;
  for(let x of arr)
  {
    ws1.cell(1,colNumber).string(x["displayName"])
    .style(style);
    
    let colName=String.fromCharCode(65+colNumber-1);
    let sqreference=colName+'1:'+colName+'100';
    
    if(x["color"])
    {
      let myStyle = wb.createStyle({
        font: {
          "color": x["color"],
        },
      });
      ws1.addConditionalFormattingRule(sqreference, {
        type: 'expression', // the conditional formatting type
        priority: 1, // rule priority order (required)
        formula: 1, // formula that returns nonzero or 0
        style: myStyle
      });
    }
    if(x["type"]=="keyword")
    {
        let formula='';
        for(let y of x["values"])
        {
          formula+=y+',';
        }
        formula=formula.substring(0,formula.length-1);
        ws1.addDataValidation({
        type: 'list',
        operator: 'equal',
        allowBlank: 0,
        showDropDown: true,
        sqref: sqreference,
        formulas:[formula]
      });
    }
    else
    {
      if(x["isRequired"]=="yes")
      {
        ws1.addDataValidation({
          errorStyle: 'warning', // One of 'stop', 'warning', 'information'. You must specify an error string for this to take effect
	        error:"This field is required",
          allowBlank: 0,
          sqref: sqreference,
        });
      }
      else
      {
        ws1.addDataValidation({
          allowBlank: 1,
          sqref: sqreference,
        });
      }
    }
    colNumber++;
  }
  wb.write('Excel1.xlsx');

