var config= require("../dbconfig");

const sql =require("mssql");

//equipments
async function getpurchase()
{
    try
    {
        let pool =await sql.connect(config);
        let purchase= await pool.request().query
            ("select * from purchase ORDER BY DOC_DT DESC " );
        return purchase.recordsets;
    }   

     catch(error)
    {
        console.log(error);
    }
}

module.exports=
{
   
   
    getpurchase: getpurchase,
    
}
    