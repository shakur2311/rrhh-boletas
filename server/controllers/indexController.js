const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const pdf = require('html-pdf');




//Variables
let excelFileSheets = {};
let excelCargado = false;

const cargarExcel = (req,res)=>{
    try {
        //Obtenemos el archivo excel almacenado en la carpeta excels
        const excelFile = xlsx.readFile(path.join(__dirname, '..','excels',req.filename));
        
        for (const sheetName of excelFile.SheetNames) {
            excelFileSheets[sheetName] = xlsx.utils.sheet_to_json(excelFile.Sheets[sheetName]);
        }
        res.json({"message":"Excel cargado"});
        excelCargado=true;
    } catch (error) {
        res.json({'error':error.message})
    }
    
}

const enviarCorreos = (req,res)=>{
    try {
        
        if(excelCargado){
            
            //Capturo valor del mes de la boleta a emitir
            let mesPago = req.body.mesPago;

            //Config de nodemailer
            let transporter = nodemailer.createTransport({
                host: "smtp.gmail.com",
                port: 465,
                secure: true, // true for 465, false for other ports
                auth: {
                user: 'shakuriwo23@gmail.com', // user gmail acc
                pass: 'eymebzzhvbcddzyk', // contraseña de aplicaciones generadas en gmail acc
                },
            });
            //

            for(var i = 0;i<(excelFileSheets.Hoja1).length;i++){
                //Elaborando la boleta pdf

                //Template boleta html
                let ubicacionPlantilla = require.resolve("../boletas/templateMail/index.html");
                let contenidoHtml = fs.readFileSync(ubicacionPlantilla, 'utf8');

                let codigo = excelFileSheets.Hoja1[i]['CODIGO'];
                let apenom = excelFileSheets.Hoja1[i]['APENOM'];
                let dni = excelFileSheets.Hoja1[i]['DNI/C.E'];
                let dependencia = excelFileSheets.Hoja1[i]['DEPENDENCIA'];
                let carnetEssalud = excelFileSheets.Hoja1[i]['CARNET ESSALUD'];
                let afp = excelFileSheets.Hoja1[i]['AFP'];
                let tpers = excelFileSheets.Hoja1[i]["T.PERS/PLAZA MGRH"];
                let diasLaborados = excelFileSheets.Hoja1[i]['DIAS LABORADOS'];
                let fechaIngreso = excelFileSheets.Hoja1[i]["FECHA DE INGRESO"];
                let nivelRem = excelFileSheets.Hoja1[i]["NIVEL REM"];
                let nroCuenta = excelFileSheets.Hoja1[i]["NRO CUENTA"];
                let cargoEstructural = excelFileSheets.Hoja1[i]["CARGO ESTRUCTURAL"];
                let condLaboral = excelFileSheets.Hoja1[i]["COND. LABORAL"];
                let tServicios = excelFileSheets.Hoja1[i]["T.SERVICIOS"];
                let cuspp = excelFileSheets.Hoja1[i]["CUSPP"];
                let horasLaboradas = excelFileSheets.Hoja1[i]["HORAS LABORADOS"];
                let ingresos = excelFileSheets.Hoja1[i]["INGRESOS"];
                let egresos = excelFileSheets.Hoja1[i]["EGRESOS"];
                let aportes = excelFileSheets.Hoja1[i]["APORTES"];
                let total = excelFileSheets.Hoja1[i]["TOTAL"];
                let correo = excelFileSheets.Hoja1[i]["CORREO"];

                //Guarda boleta pdf dentro de la carpeta boletas
                contenidoHtml = contenidoHtml.replace('{{mesPago}}',mesPago);
                contenidoHtml = contenidoHtml.replace('{{codigo}}', codigo);
                contenidoHtml = contenidoHtml.replace('{{apenom}}', apenom);
                contenidoHtml = contenidoHtml.replace('{{dni}}', dni);
                contenidoHtml = contenidoHtml.replace('{{dependencia}}', dependencia);
                contenidoHtml = contenidoHtml.replace('{{carnetEssalud}}', carnetEssalud);
                contenidoHtml = contenidoHtml.replace('{{afp}}', afp);
                contenidoHtml = contenidoHtml.replace('{{tpers}}', tpers);
                contenidoHtml = contenidoHtml.replace('{{diasLaborados}}', diasLaborados);           
                contenidoHtml = contenidoHtml.replace('{{fechaIngreso}}', fechaIngreso);
                contenidoHtml = contenidoHtml.replace('{{nivelRem}}', nivelRem);
                contenidoHtml = contenidoHtml.replace('{{nroCuenta}}', nroCuenta);
                contenidoHtml = contenidoHtml.replace('{{cargoEstructural}}', cargoEstructural);
                contenidoHtml = contenidoHtml.replace('{{condLaboral}}', condLaboral);
                contenidoHtml = contenidoHtml.replace('{{tServicios}}', tServicios);
                contenidoHtml = contenidoHtml.replace('{{cuspp}}', cuspp);
                contenidoHtml = contenidoHtml.replace('{{horasLaboradas}}', horasLaboradas);
                contenidoHtml = contenidoHtml.replace('{{ingresos}}', ingresos);
                contenidoHtml = contenidoHtml.replace('{{egresos}}', egresos);
                contenidoHtml = contenidoHtml.replace('{{aportes}}', aportes);
                contenidoHtml = contenidoHtml.replace('{{totalIngresos}}', ingresos);
                contenidoHtml = contenidoHtml.replace('{{totalEgresos}}', egresos);
                contenidoHtml = contenidoHtml.replace('{{totalAportes}}', aportes);
                contenidoHtml = contenidoHtml.replace('{{total}}', total);
  
                //Obteniendo hora y fecha actual para guardar nombre boleta PDF
                const date = new Date().toLocaleString({ timeZone: "America/Lima" });
                const newDate = date.split(' ');
                const fecha = newDate[0].replaceAll('/','');
                const hora = newDate[1].replaceAll(':','');
                const filename = codigo+'-'+fecha+'-'+hora+'.pdf';

                //Creando boleta pdf con valores reemplazados y guardandolo
                pdf.create(contenidoHtml).toFile(`./boletas/emitidas/${filename}`, function(err,res) {
                    if (err) {
                        console.log("error al guardar boleta")
                    
                    }else{

                        console.log("pdf guardado en servidor!");
                        //Enviando email
                        transporter.sendMail({
                            from: '<shakuriwo23@gmail.com>', // sender address
                            to: correo, // list of receivers
                            subject: "Prueba", // Subject line
                            html: `<p>Saludos cordiales <strong>${apenom}</strong> se hace envio de la boleta 
                            correspondiente al mes de <strong>${mesPago}</strong></p>`,
                            attachments:[{
                                filename:filename,
                                path:path.join(__dirname,'..','boletas/emitidas',filename)
                            }]
                        });
                    }
                });

                
            }

            //Devuelve correos enviados luego de recorrer todo el for
            res.json({'message':'correos enviados!'});
            excelCargado = false;
            excelFileSheets = {};
        }else{
            res.json({'message':'Debe cargar un excel'});
        }
        

    } catch (error) {
        res.json({'error':error.message})
    }
}

const obtenerArchivos = (req,res)=>{
    try {
        //Obtener listado de archivos pdf en el servidor
        const boletasFolder = path.join(__dirname,'..','boletas/emitidas'); //Obtener path donde estan los excels      
        fs.readdir(boletasFolder, async (err,files)=>{
            
            var filesOfPath = [];
            
            
            for (const [index,file] of files.entries()) {
                const filePath = path.join(boletasFolder,file)
                const stat = await fs.promises.stat(filePath);
                filesOfPath[index]={
                                        fileName: file,
                                        fileSize: stat.size
                                    }
                
            }

            res.json(filesOfPath)
                      
        })
        
    } catch (error) {
        res.json({
            'error':error.message
        })
    }
}

const descargarArchivo = (req,res)=>{
    try {
        const downloadFile = path.join(__dirname, '..','boletas/emitidas',req.params.fileName);
        res.download(downloadFile);
    } catch (error) {
        res.send("No se pudo descargar el archivo!");
    }
}
module.exports = {enviarCorreos,obtenerArchivos,descargarArchivo,cargarExcel};