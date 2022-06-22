var express = require("express");
var bodyParser = require("body-parser");
var sql = require("mssql");
var cors = require('cors');
const { json } = require("body-parser");
var app = express(); 

// Body Parser Middleware
app.use(bodyParser.json()); 
app.use(cors());

//Setting up server
 var server = app.listen(process.env.PORT || 8080, function () {
    var port = server.address().port;
    console.log("App now running on port", port);
 });

 //Initializing connection string
var sqlconnection = sql.connect({
   user: 'sa',
    password: '123$abc',
    server: 'COMP-038', 
    database: 'EquipmentTrial', 
    Options:{
      trustedconnection: true,
      enableArithAbort: true,
      instancename: '',
    },
   port: 50023

});
sqlconnection.connect((err)=> {

    if (!err) 
    console.log('DB connection succedded');
    else
    console.log('DB connection fail\n Error:'+json.stringify(err,undefined,2));
    // create Request object
 
 });
