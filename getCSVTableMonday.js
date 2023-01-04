const https = require('https')
const formulajs = require('@formulajs/formulajs')
const fs = require('fs').promises;


function doRequest(options, data) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      res.setEncoding('utf8');
      let responseBody = '';
      res.on('data', (chunk) => {
        responseBody += chunk;
      });

      res.on('end', () => {
        resolve(responseBody);
      });
    });

    req.on('error', (err) => {
      reject(err);
    });

    req.write(data)
    req.end();
  });
}


var apiKey = ""
var apiUrl = "https://api.monday.com/v2";
var apiHostname = "api.monday.com";
var headers = { "Authorization": apiKey }
var bulk = 50;

async function makeQuery(query) {
  var data = { 'query': query };
  data = JSON.stringify(data);
  var headers = {
    "Authorization": apiKey,
    'Content-Type': 'application/json',
    'Content-Length': data.length
  }
  var options = {
    hostname: apiHostname,
    port: 443,
    path: '/v2',
    method: 'POST',
    headers: headers,
  };
  return await doRequest(options, data);
}

async function getItemData(itemid) {
  var query = 'query {items (ids: [' + itemid + ']) {name group {title} column_values{ id value }}}'
  return makeQuery(query);
}

async function retrieve_column_data_and_write_to_json_file(boardid) {
  query = 'query {boards (ids:[' + boardid + ']){id name description columns {id title type description settings_str}}}'
  return makeQuery(query)
}

function retrieve_item_ids_from_board(boardid) {
  var query = '{boards(ids:[' + boardid + ']){items{id}}}';
  return makeQuery(query);
}

function findColumnValue(id, data) {
  for (var item in data) {
    if (item['id'] == id) {
      if (item['value'] != undefined) {
        return item['value'];
      }
      else {
        return "";
      }
    }
  }
  return "";
}

function getItemData(itemid) {
  query = 'query {items (ids: [' + itemid + ']) {name group {title} column_values{ id text title value type }}}'
  return makeQuery(query);
}

function evalformula(formula, cvs,tablerow,fastcolumn) {
  const regex = /#(.*?)#/g;
  const list = formula.match(regex);
  var ilist = 0
 
  while (list !== undefined && ilist < list.length) {
    list[ilist] = list[ilist].replaceAll("#","");
    ilist++;
  } 
  var ilist = 0;
  while (ilist < list.length) {
    if(fastcolumn[list[ilist]]!==undefined)
    {
      if(fastcolumn[list[ilist]]["type"]=="formula")
      {
        formula = formula.replace("#"+list[ilist]+"#","("+evalformula(fastcolumn[list[ilist]]["evilFormula"], cvs,tablerow,fastcolumn)+")");
      }
      else
      {
        if(tablerow[list[ilist]]=="")
        {
          tablerow[list[ilist]] = 0;
        }
        var str =tablerow[list[ilist]]; 
        if(isNaN(tablerow[list[ilist]]))
        {
          str = "\""+str+"\"";
        }
        formula = formula.replace("#"+list[ilist]+"#",str);
      }
    }
    else
    {
      formula = formula.replace("#"+list[ilist]+"#","0");
    }
    
    ilist++;
  }
  return formula;
}





function parseNumber(string) {
  if (string instanceof Error) {
    return string
  }

  if (string === undefined || string === null || string === '') {
    return 0
  }

  if (typeof string === 'boolean') {
    string = +string
  }

  if (!isNaN(string)) {
    return parseFloat(string)
  }

  return error.value
}

async function getBoardData(boardid) {
  //console.log(formulajs);
  var itemids = await retrieve_item_ids_from_board(boardid);
  //await fs.writeFile('board_' + boardid + '_items.txt', itemids);
  var columns = await retrieve_column_data_and_write_to_json_file(boardid);
  columns = JSON.parse(columns);
  columnsobj = columns["data"]["boards"][0]["columns"];

  var icolumns = 0;
  var columnslength = columnsobj.length;

  var fastcolumn = {};
  while (icolumns < columnslength) {
    var columnsid = columnsobj[icolumns]["id"];
    fastcolumn[columnsid] = {};
    fastcolumn[columnsid]["title"] = columnsobj[icolumns]["title"];
    fastcolumn[columnsid]["type"] = columnsobj[icolumns]["type"];
    fastcolumn[columnsid]["settings_str"] = columnsobj[icolumns]["settings_str"];
    fastcolumn[columnsid]["settings_obj"] = JSON.parse(fastcolumn[columnsid]["settings_str"]);

    if (fastcolumn[columnsid]["settings_obj"].labels !== undefined
      && Object.keys(fastcolumn[columnsid]["settings_obj"].labels).length > 0) {
      var labellength = Object.keys(fastcolumn[columnsid]["settings_obj"].labels).length;
      fastcolumn[columnsid]["labels"] = {};
      var ilabels = 0;
      while (ilabels < labellength) {
        fastcolumn[columnsid]["labels"][ilabels] = fastcolumn[columnsid]["settings_obj"].labels[ilabels];
        ilabels++;
      }
    }
    if (fastcolumn[columnsid]["settings_obj"]["formula"] !== undefined) {
      fastcolumn[columnsid]["formula"] = fastcolumn[columnsid]["settings_obj"]["formula"];
      var formula = fastcolumn[columnsid]["settings_obj"]["formula"];
      formula = formula.replace(/{(.*?)}/g, "#$1#");
      
      //Formulas taken from https://github.com/formulajs/formulajs
      let regex = /[a-zA-Z]*\(/gm;
      const found = formula.match(regex);
      
      if (found!==null)
      {
        var ifound = 0; 
        while(ifound < found.length){
          var fPart = found[ifound].toUpperCase();
          fPart = fPart.replaceAll("(","");
          if (typeof formulajs[fPart] == 'function')
          {
            formula = formula.replaceAll(found[ifound],"formulajs." + fPart + "~("); 
          }
          ifound++;
        }
      }
      
      formula = formula.replaceAll("~", "");
      formula = formula.replaceAll("=", "==");
      formula = formula.replaceAll("<>", "!=");
      fastcolumn[columnsid]["evilFormula"] = formula;
    }
    icolumns++;
  }
  var csvdata = '';
  icolumns = 0;
  csvdata = "Name;Group";
  while (icolumns < columnslength) {
    //console.log(columnsobj[icolumns]["title"]);
    if(columnsobj[icolumns]["title"]!="Name")
    {
      csvdata += ";" + columnsobj[icolumns]["title"];
    }
    
    icolumns++;
  }
  
  csvdata += "\n";

  //await fs.writeFile('board_' + boardid + '_columns.txt', JSON.stringify(columns));

  var board = JSON.parse(itemids)["data"]["boards"][0];
  var itemlength = board.items.length;
  var iitems = 0;
  //itemlength = 200;
  while (iitems < itemlength) {
    var ibulk = 0;
    var queryitem = '';
    var first = true;

    while (ibulk < bulk) {
      if (first == false) {
        if (iitems < itemlength) {
          queryitem += "," + board.items[iitems].id;
        }
      }
      else {
        first = false;
        queryitem += "" + board.items[iitems].id;
      }

      iitems++;
      ibulk++;
    }

    var items = await getItemData(queryitem);

    var itemobject = JSON.parse(items);
    var arrItems = itemobject.data.items;
    var arrItemslength = arrItems.length;
    var iarritems = 0;
    var table = [];
    while (iarritems < arrItemslength) {
      var columnvalues = [];
      table[iarritems] = {};
      columnvalues = arrItems[iarritems]["column_values"]
      var icolumnvalues = 0;
      var columnvalueslength = columnvalues.length;
      var name = arrItems[iarritems]["name"];
      var group = arrItems[iarritems]["group"]["title"];

      csvdata += name + ";" + group;
      var first = false;
      while (icolumnvalues < columnvalueslength) {

        var id = columnvalues[icolumnvalues]["id"];
        table[iarritems][id] = columnvalues[icolumnvalues]["text"];
        icolumnvalues++;
      }
      icolumnvalues = 0;
      while (icolumnvalues < columnvalueslength) {

        var id = columnvalues[icolumnvalues]["id"];
        var firststr = "";
        if (!first) {
          firststr = ";";
        }
        if (fastcolumn[columnvalues[icolumnvalues]["id"]]["type"] == "formula") {

          
          var evilformula = evalformula(fastcolumn[columnvalues[icolumnvalues]["id"]]["evilFormula"], columnvalues,table[iarritems],fastcolumn);
          //console.log(evilformula);
          var value = firststr + "n.a.";
          try
          {
            value = firststr + eval(evilformula);
            if(value == "NaN")
            {
              console.log(evilformula)
              value = 0;
            }
          }
          catch{
            value = firststr + "ERROR!";
            console.log(evilformula);
          }

          csvdata += (value+'').replace(".",","); 
        }
        else {
          if (columnvalues[icolumnvalues]["text"] !== null) {
            if(isNaN(columnvalues[icolumnvalues]["text"]))
            {
              csvdata += firststr + cleanTextforCSV(columnvalues[icolumnvalues]["text"]);
            }
            else
            {
              csvdata += firststr + cleanTextforCSV(columnvalues[icolumnvalues]["text"]);
            }
          }
          else {
            csvdata += firststr + "";
          }
          table[iarritems][id] = cleanTextforCSV(columnvalues[icolumnvalues]["text"]);
        }
        icolumnvalues++;
      }
      csvdata += "\n";
      iarritems++;
    }
    iitems++;
  }
  await fs.writeFile('board_' + boardid + '_export.csv', csvdata, "latin1");
}

function cleanTextforCSV(txt)
{
  if(txt!==null&&txt!==undefined)
  {
    txt = txt.replace(".",",");
    txt = txt.replace("\n","");
    txt = txt.replace(";",",");
    return txt
  }
  else
  {
    return ""
  }
}

var table = process.argv[2];

getBoardData(table);
