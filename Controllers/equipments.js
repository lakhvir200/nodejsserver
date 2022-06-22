var config= require("../dbconfig");

const sql =require("mssql");

//equipments
async function getequipments()
{
    try
    {
        let pool =await sql.connect(config);
        let equipments= await pool.request()
        .input('action','select')       
        .execute('sp_equipment_curd'); 
        return equipments.recordsets;
    }   

     catch(error)
    {
        console.log(error);
    }
}
async function getequipmentsById(id)
{
    //console.log(id)
    try
    {
        let pool =await sql.connect(config);
        let equipments= await pool.request()
        .input('action','select')
        .input('ID',id)
        .execute('sp_equipment_curd'); 
        return equipments.recordsets;
    }   

     catch(error)
    {
        console.log(error);
    }
}
async function addequipment(equip)
{
    
    try
    {
        let pool = await sql.connect(config)    
        let insertEquipment =await pool.request() 
        .input('action','INSERT')           
        .input('EQUIPMENT_ID',sql.NVarChar,equip.EQUIPMENT_ID)
        .input('EQUIPMENT_NAME',sql.NVarChar,equip.EQUIPMENT_NAME)
        .input('DEPARTMENT',sql.NVarChar,equip.DEPARTMENT)
        .input('COMPLETE_SPECIFICATION',sql.NVarChar,equip.COMPLETE_SPECIFICATION)
        .input('DATE_OF_PURCHASE',sql.DATE,equip.DATE_OF_PURCHASE)
        .input('BILL_DATE',sql.DATE,equip.BILL_DATE)
        .input('COST_OF_EQUIPMENT',sql.INT,equip.COST_OF_EQUIPMENT)
        .input('CATEGORY',sql.NVarChar,equip.CATEGORY)      
        .input('SUBCATEGORY',sql.NVarChar,equip.SUBCATEGORY)
        .input('MODEL',sql.NVarChar,equip.MODEL)
        .input('UNIT_NAME',sql.NVarChar,equip.UNIT_NAME)
        .input('WARRANTY',sql.INT,equip.WARRANTY)
        .input('BUDGET_YEAR',sql.NVarChar,equip.BUDGET_YEAR)
        .input('ISACTIVE',sql.NVarChar,1)
        .input('EQUIP_STATUS',sql.NVarChar,equip.EQUIP_STATUS)
        .input('EQUIP_REMARKS',sql.NVarChar,equip.EQUIP_REMARKS)
        .input('Photo',sql.Image,equip.Photo)
        .input('SUPPLIER',sql.NVarChar,equip.SUPPLIER)
        .execute('sp_equipment_curd');
        return insertEquipment.recordsets;
    }
  catch (error)
    {
        console.log(error)
    }
}
async function editequipment(equip)
{
    
    try
    {
        let pool = await sql.connect(config)    
        let EditEquipment =await pool.request() 
        .input('action','Update')  
        .input('ID',sql.NVarChar,equip.ID)         
        .input('EQUIPMENT_ID',equip.EQUIPMENT_ID)
        .input('EQUIPMENT_NAME',equip.EQUIPMENT_NAME)
        .input('DEPARTMENT',equip.DEPARTMENT)
        .input('COMPLETE_SPECIFICATION',equip.COMPLETE_SPECIFICATION)
        .input('DATE_OF_PURCHASE',equip.DATE_OF_PURCHASE)
        .input('BILL_DATE',equip.BILL_DATE)
        .input('COST_OF_EQUIPMENT',equip.COST_OF_EQUIPMENT)
        .input('CATEGORY',equip.CATEGORY)
        .input('MAINT_PERIODICITY',equip.MAINT_PERIODICITY)
        .input('UNIT_NAME',equip.UNIT_NAME)
        .input('WARRANTY',equip.WARRANTY)
        .input('ISACTIVE',equip.ISACTIVE)
        .input('EQUIP_STATUS',equip.EQUIP_STATUS)
        .input('EQUIP_REMARKS',equip.EQUIP_REMARKS)
        .input('Photo',sql.Image,equip.Photo)
        .input('SUPPLIER',equip.SUPPLIER)
        .execute('sp_equipment_curd');
        return EditEquipment.recordsets;
    }
  catch (error)
    {
        console.log(error)
    }
}
async function deleteequipment(id)
{
    console.log(id)
    try
    {
        let pool =await sql.connect(config);
        let equipments= await pool.request()
        .input('action','delete')
        .input('ID',id)
        .execute('sp_equipment_curd'); 
        return equipments.recordsets;
    }   

     catch(error)
    {
        console.log(error);
    }
}

module.exports=
{     
    getequipments: getequipments,
    getequipmentsById:getequipmentsById,
    addequipment:addequipment,
    editequipment:editequipment,
    deleteequipment:deleteequipment    
}
    