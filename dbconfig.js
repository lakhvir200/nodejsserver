

const config = {
    user: "sa",
    password: "123$abc",
    server: "COMP-038", 
    database: "Equipment",
    options:{
     trustServerCertificate: true,
     trustedconnection: true,
     enableArithAbort: true,
     instancename: "COMP-038\\SQLSERVER2019",

    },
    port: 50023,
};

module.exports= config;
/*
//var mysql = require('mysql');
//get an instance of sqlserver 
var sql = require('mssql');

//set up a sql server credentials
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


function con() {
var dbConn = new sql.ConnectionPool(config)
dbConn.connect().then(function(){
    console.log("connected")
}).catch(function (err) {
    console.log(err);
})
}

con();
*/