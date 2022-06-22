var departments = require('./Controllers/departments');
 const equipments= require('./Controllers/equipments')
 const path = require("path");
 const hbs =require("hbs");
 var express = require('express');
 var bodyParser=require('body-parser');
  var cors= require('cors');
const { request } = require('express');
const Purchase = require('./Controllers/Purchase');
const consumption = require('./Controllers/consumption');
const repair = require('./Controllers/repair');
  var app = express();
 var router = express.Router(); 

 app.use(bodyParser.urlencoded({extended:true}));
 app.use(bodyParser.json());
 app.use(cors()); 
 //app.use(express.json());
 app.use('/', router);

 router.use((Request,Response,next)=>{
    console.log('middleware');
    next();
})
const template_path = path.join(__dirname, "../nodejsserver/templates/views")
const partials_path = path.join(__dirname, "../nodejsserver/templates/partials")

 console.log (template_path);

app.set("view engine", "ejs");
app.set("views", template_path);
hbs.registerPartials(partials_path);

router.get("/", (req, res) => {
    res.render("index");
  }); 
    
//departments
router.route('/departments').get((req,res)=>{
    departments.getdepts()
    .then(result=>{       
       res.json(result);   
        //var data= (result); 
        //console.log(data);
       // res.render({data:result}) 
        }) 

})

router.route('/departmentsById/:id').get((req,res)=>{
   departments.getDeptById(req.params.id).then(result=>{
        //console.log(result);
         res.json(result[0]);  
        }) 

})
router.route('/Deletedepartments/:id').delete((req,res)=>{
    departments.deleteDepts(req.params.id).then(result=>{
        //console.log(result);
         res.json(result[0]);  
        }) 

})
//equipments

router.get("/equipment/add", function(req,res, next){    
  res.render("equipments/equip", {title:'Add Equipment', action:'add'});
 })
router.route('/insertdepartments').post((req,res)=>{
    let dept = {...req.body}
   departments.addDept(dept).then(result=>{

       // console.log(result);
         res.status(201).json(result);  
        }) 
      })

router.route('/equipments/list').get((req,res)=>{
    equipments.getequipments().then(result=>{       
       res.json(result);
       // res.render('equipments/equip',{title:'Equipments List',action:'list',equipments:data})
       //  response.render('equipments', {title:'Node.js MySQL CRUD Application', action:'list', sampleData:result[0]});
        }) 
})
router.route('/equipments/byId/:id').get((req,res)=>{
    equipments.getequipmentsById(req.params.id).then(result=>{
        //console.log(result);
         //res.json(result[0]);  
         var data= (result[0]); 
        // res.render('equipments/equip',{title:'Equipments List',action:'list',equipments:data})
         res.render('equipments/equip',{title:'Equipments List',action:'edit',equipments:data})
        })})
router.route('/equipment/save').post((req,res)=>{
    let equip = {...req.body}
   equipments.addequipment(equip).then(result=>{
   // C:\Users\lakhvir\Desktop\NodeJsServer\templates\views\equipments\addequipment.hbs
       // console.log(result);
        // res.status(201).json(result);
        //var data= (result); 
       // res.render('equipments/equipmentsList',{title:'Equipments List',action:'list',equipments:data}) 
       res.redirect("/equipments/list"); 
        })})
router.route('/equipment/edit/:id').put((req,res)=>{
    let equip = {...req.body}
   equipments.editequipment(req.params.id).then(result=>{
       // console.log(result);
         res.status(201).json(result);  
        }) 
})
router.route('/equipment/delete/:id').get((req,res)=>{  
 equipments.deleteequipment(req.params.id).then(result=>{

   // console.log(req.params.id);
      // res.status(201).json(result); 
      //var data= (result[0]); 
      res.redirect('/equipments/list') 
      }) 

})
//purchase

router.route('/purchase').get((req,res)=>{
    Purchase.getpurchase().then(result=>{
        //console.log(result);
         res.json(result[0]);  
  }) 

})
//consumption
router.route('/consumption').get((req,res)=>{
    consumption.getconsumption().then(result=>{
        //console.log(result);
         res.json(result[0]);  
  }) 
  
})
//consumption
router.route('/repair').get((req,res)=>{
   repair.getrepair().then(result=>{
        //console.log(result);
         res.json(result[0]);  
  }) 
  
})

 var port = process.env.PORT || 3000;
app.listen(port);
console.log('server is running at  ' + port);
