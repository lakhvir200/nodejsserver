var config= require("../dbconfig");

const sql =require("mssql");

//equipments
async function getrepair()
{
    try
    {
        let pool =await sql.connect(config);
        let repair= await pool.request().query
            ("sp_repair" );
        return repair.recordsets;
    }   

     catch(error)
    {
        console.log(error);
    }
}

module.exports=
{
   
   
    getrepair: getrepair,
    
}
    