const express = require('express');
const router = express.Router();
const path = require('path');
const multer = require('multer');
const indexController = require('../controllers/indexController');
const loginController = require('../controllers/loginController');

const storageNewExcel = multer.diskStorage({
    destination: path.join(__dirname,'..','excels'),
    filename: (req,file,cb)=>{
        if(file.mimetype=="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
            const ext = "xlsx";
            const date = new Date().toLocaleString({ timeZone: "America/Lima" });
            const newDate = date.split(' ');
            const fecha = newDate[0].replaceAll('/','');
            const hora = newDate[1].replaceAll(':','');
            const filename = fecha + '-' + hora + '.' + ext;
            req.filename = filename;
            cb(null,filename);
        }

        
    }
   
});
const multerNewExcel = multer({
    storage: storageNewExcel,
    dest: path.join(__dirname,'..','excels')
});



router.get('/',(req,res)=>{
    res.send("Servidor nodejs");
    
})


router.post('/login',loginController.login);
router.post('/cargarExcel',multerNewExcel.single("excelFile"),indexController.cargarExcel);
router.post('/enviarCorreos',indexController.enviarCorreos)




router.get('/obtenerArchivos',indexController.obtenerArchivos);
router.get('/descargarArchivo/:fileName',indexController.descargarArchivo);
/* router.get('/obtenerPDF',indexController.obtenerPDF); */

module.exports = router;