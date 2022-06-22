var config= require("../dbconfig");

const sql =require("mssql");

//equipments
async function getconsumption()
{
    try
    {
        let pool =await sql.connect(config);
        let consumption= await pool.request().query
            ("select * from consumption ORDER BY DOC_DT DESC" );
        return consumption.recordsets;
    }   

     catch(error)
    {
        console.log(error);
    }
}

module.exports=
{
   
   
    getconsumption: getconsumption,
    
}
    