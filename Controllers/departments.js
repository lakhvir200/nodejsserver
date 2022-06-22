var config= require("../dbconfig");
const sql =require("mssql");

async function getdepts()
{
    try
    {
        let pool =await sql.connect(config);
        let departments= await pool.request().query
        ("select department from Dept_Details ORDER BY ID DESC")    
        return departments;
    }        
     catch(error){
        console.log(error);
    }
}

//departments
async function getDeptById(deptId)
{
    try
    {
        let pool =await sql.connect(config);
        let departments= await pool.request()
        .input('input_parameter',sql.Int,deptId).query
        ("select * from Dept_Details where Id =@input_parameter" );

        return departments.recordsets;
    }        
     catch(error)
     {
        console.log(error);
    }
}
async function deleteDepts(deptId)
{
    try
    {
        let pool =await sql.connect(config);
        let departments= await pool.request()
        .input('input_parameter',sql.Int,deptId).query
        ("Delete from Dept_Details where id =@input_parameter" )
    
        return departments.recordsets;
    }        
     catch(error){
        console.log(error);
    }
    
}

    async function addDept(departments)
    {
        try
        {
            let pool = await sql.connect(config)    
            let insertDepartment =await pool.request()            
            .input('DepartmentID',sql.NVarChar,departments.DepartmentID)
            .input('DepartmentName',sql.NVarChar,departments.DepartmentName)
            .execute('sp_Insert_Department');
            return insertDepartment.recordsets;
        }
      catch (error)
    
        {
            console.log(error)
        }
}



module.exports=
{
    getdepts: getdepts,
    getDeptById :getDeptById,
   
    addDept: addDept,
    deleteDepts:deleteDepts
}
    