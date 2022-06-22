var express = require('express');
var app = express();

app.get('/', function (req, res) {
   
    var sql = require("mssql");

    // config for your database
    var config = {
      user: "sa",
      password: "123$abc",
      server: "COMP-038", 
      database: "EquipmentTrial",
      options:{
       trustServerCertificate: true,
       trustedconnection: true,
       enableArithAbort: true,
       instancename: "COMP-038\\SQLSERVER2019",
  
      },
      port: 50023,

    };
   
     
    // connect to your database
    sql.connect(config, function (err) {
    
        if (err) 
        console.log(err);
       
      //create Request object
        var request = new sql.Request();
           
       //query to the database and get the records
     request.query('select * from departments', function (err, recordset) {
            
         if (err) console.log(err)

          //send records as a response
        res.send(recordset);
            
        });
    });
});

var server = app.listen(5000, function () {
    console.log('Server is running..');
});