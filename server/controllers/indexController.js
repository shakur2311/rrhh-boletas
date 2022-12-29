const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const pdf = require('html-pdf');
const ejs = require('ejs');
const { type } = require('os');




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
        res.json({'error':error.message});
        excelCargado=false;
    }
    
}

const enviarCorreos = async (req,res)=>{
    try {
        
        if(excelCargado){
            
            //Capturo valor del mes de la boleta a emitir y tipo de boleta
            let mesPago = req.body.mesPago;
            let tipoBoleta;
            


            switch(req.body.tipoBoleta){
                case "planillaCas":
                    tipoBoleta = "PLANILLA - CAS";
                    break;
                case "planillaHaberes":
                    tipoBoleta = "PLANILLA - HABERES";
                    break;
                case "planillaPensiones":
                    tipoBoleta = "PLANILLA - PENSIONES";
                    break;
                default:
                    tipoBoleta = "Indefinido";
            }

            //Config de nodemailer
            let transporter = nodemailer.createTransport({
                host: "smtp.gmail.com",
                port: 465,
                pool:true,
                secure: true, // true for 465, false for other ports
                auth: {
                user: 'shakuriwo23@gmail.com', // user gmail acc
                pass: 'eymebzzhvbcddzyk', // contraseña de aplicaciones generadas en gmail acc
                },
            });
            //
            //Convertir createPDF de html-pdf a promesa
            const createPDF = (html, options) => new Promise(((resolve, reject) => {
                pdf.create(html, options).toBuffer((err, buffer) => {
                    if (err !== null) {reject(err);}
                    else {resolve(buffer);}
                });
            }));
            
            for(let i = 0;i<(excelFileSheets.Hoja1).length;i++){
                //Elaborando la boleta pdf
                
                //Datos extraidos del excel
                //#region 
                //Info de empleado
                let codigo = excelFileSheets.Hoja1[i]["CODIGO"];
                let apenom = excelFileSheets.Hoja1[i]["APELLIDOS Y NOMBRES"];
                let dni = excelFileSheets.Hoja1[i]["DNI/C.E"];
                let dependencia = excelFileSheets.Hoja1[i]['DEPENDENCIA'];
                let carnetEssalud = excelFileSheets.Hoja1[i]["CARNET ESSALUD"];
                let afp = excelFileSheets.Hoja1[i]['AFP'];
                let tpers = excelFileSheets.Hoja1[i]["T.PERS/PLAZA MGRH"];
                let diasLaborados = excelFileSheets.Hoja1[i]['DIAS LABORADOS'];    
                let fechaIngreso;
                if(excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().length==6){
                    fechaIngreso = excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().substr(0,2) 
                    + " Años " + excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().substr(2,2) + " Meses "
                    + excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().substr(4,2) + " Dias ";
                }else if(excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().length==5){
                    fechaIngreso = excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().substr(0,1) 
                    + " Años " + excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().substr(1,2) + " Meses "
                    + excelFileSheets.Hoja1[i]["FECHA DE INGRESO"].toString().substr(3,2) + " Dias ";
                }
                let nivelRem = excelFileSheets.Hoja1[i]["NIVEL REM"];
                let nroCuenta = excelFileSheets.Hoja1[i]["NRO CUENTA"];
                let cargoEstructural = excelFileSheets.Hoja1[i]["CARGO ESTRUCTURAL"];
                let condLaboral = excelFileSheets.Hoja1[i]["COND. LABORAL"];
                let tServicios = excelFileSheets.Hoja1[i]["T.SERVICIOS"];
                let cuspp = excelFileSheets.Hoja1[i]["CUSPP"];
                let horasLaboradas = excelFileSheets.Hoja1[i]["HORAS LABORADOS"];
                let correo = excelFileSheets.Hoja1[i]["CORREO"];
                
                //INGRESOS
                let pensiones = excelFileSheets.Hoja1[i]["PENSIONES"];
                let DS062020ef = excelFileSheets.Hoja1[i]["S *, 110-06, 39-07,6-2020-EF"];
                let DS170519ef = excelFileSheets.Hoja1[i]["D.S. N* 17-2005-EF-09-2019-EF"];
                let Asig27691 = excelFileSheets.Hoja1[i]["ASIGNACION 276-91-EF"];
                let DS1118ef = excelFileSheets.Hoja1[i]["LEY 28449--28789--D.S.11-18 EF"];
                let DS0621ef = excelFileSheets.Hoja1[i]["D.S.006-2021-EF"];
                let DS01422ef = excelFileSheets.Hoja1[i]["D.S.014-2022-EF"];
                let grati = excelFileSheets.Hoja1[i]["GRATIFICACIÓN"];
                let docentescontratados = excelFileSheets.Hoja1[i]["D.S 418"];
                let autoridades = excelFileSheets.Hoja1[i]["D.S 313"];
                let docentesnombrados = excelFileSheets.Hoja1[i]["MUC 58"];
                let administrativos1 = excelFileSheets.Hoja1[i]["MUC DU 38"];
                let administrativos2 = excelFileSheets.Hoja1[i]["BDP"];
                let administrativos3 = excelFileSheets.Hoja1[i]["BET"];
                let aguinaldo = excelFileSheets.Hoja1[i]["AGUINALDO"];
                let cas = excelFileSheets.Hoja1[i]["CAS"];
                let reintegro = excelFileSheets.Hoja1[i]["REINTEGRO"];
                let totalingresos = parseFloat(excelFileSheets.Hoja1[i]["TOTAL DE ING."]).toFixed(2);

                let ingresosArray = [];

                if(typeof pensiones!='undefined'){
                    ingresosArray.push({"texto":"Pensiones","valor":pensiones});
                }
                if(typeof DS062020ef!='undefined'){
                    ingresosArray.push({"texto":"D.S. 06-2020-EF","valor":DS062020ef});
                }
                if(typeof DS170519ef!='undefined'){
                    ingresosArray.push({"texto":"D.S. 17-05.19-EF","valor":DS170519ef});
                }
                if(typeof Asig27691!='undefined'){
                    ingresosArray.push({"texto":"ASIG-276-91-EF","valor":Asig27691});
                }
                if(typeof DS1118ef!='undefined'){
                    ingresosArray.push({"texto":"D.S. 11-18-EF","valor":DS1118ef});
                }
                if(typeof DS0621ef!='undefined'){
                    ingresosArray.push({"texto":"D.S. 006-21-EF","valor":DS0621ef});
                }
                if(typeof DS01422ef!='undefined'){
                    ingresosArray.push({"texto":"D.S. 014-2022-EF","valor":DS01422ef});
                }
                if(typeof grati!='undefined'){
                    ingresosArray.push({"texto":"GRATIFIC.","valor":grati});
                }
                if(typeof docentescontratados!='undefined'){
                    ingresosArray.push({"texto":"D.S 418","valor":docentescontratados});
                }
                if(typeof autoridades!='undefined'){
                    ingresosArray.push({"texto":"D.S 313","valor":autoridades});
                }
                if(typeof docentesnombrados!='undefined'){
                    ingresosArray.push({"texto":"MUC 58","valor":docentesnombrados});
                }
                if(typeof administrativos1!='undefined'){
                    ingresosArray.push({"texto":"MUC DU 38","valor":administrativos1});
                }
                if(typeof administrativos2!='undefined'){
                    ingresosArray.push({"texto":"BDP","valor":administrativos2});
                }
                if(typeof administrativos3!='undefined'){
                    ingresosArray.push({"texto":"BET","valor":administrativos3});
                }
                if(typeof aguinaldo!='undefined'){
                    ingresosArray.push({"texto":"Aguinaldo","valor":aguinaldo});
                }
                if(typeof cas!= 'undefined'){
                    ingresosArray.push({"texto":"CAS","valor":cas});
                }
                if(typeof reintegro!= 'undefined'){
                    ingresosArray.push({"texto":"reintegro","valor":reintegro});
                }
                
        
                //EGRESOS
                let faltasyotardanzas = excelFileSheets.Hoja1[i]["FALTAS Y/O TARDANZAS"];
                let sudunaccp = excelFileSheets.Hoja1[i]["SUDUNAC (CASTILLO PRADO)"];
                let tespublico = excelFileSheets.Hoja1[i]["RESPONS. FISCAL (TESORO PUBLI)"];
                let cajamunareq = excelFileSheets.Hoja1[i]["CAJA MUNICIPAL DE AREQUIPA"];
                let cooplaunion = excelFileSheets.Hoja1[i]["COOPERATIVA LA UNION"];
                let coopsanmiguel = excelFileSheets.Hoja1[i]["SAN MIGUEL EX COOP-PONDEROSA"];
                let otrossudunac = excelFileSheets.Hoja1[i]["OTROS(SUTUNAC)"];
                let bancognb = excelFileSheets.Hoja1[i]["BANCO GNB PERU S.A."];
                let coopeltumi = excelFileSheets.Hoja1[i]["COOPERATIVO EL TUMI"];
                let asocjubunac = excelFileSheets.Hoja1[i]["ASOC. JUB.UNAC"];
                let bancoscotiabank = excelFileSheets.Hoja1[i]["SCOTIABANK PERU S.A.A."];
                let ceuunac = excelFileSheets.Hoja1[i]["CEU-UNAC"];
                let sutunacfall = excelFileSheets.Hoja1[i]["SUTUNAC (FALL. CESE)"];
                let sudunacfall = excelFileSheets.Hoja1[i]["FALLECIMIENTO (SUDUNAC)JCASTIL"];
                let cajachica = excelFileSheets.Hoja1[i]["CAJA CHICA O.TES."];
                let regularonp = excelFileSheets.Hoja1[i]["REGULAR ONP"];
                let regularabono = excelFileSheets.Hoja1[i]["REGULAR ABONO"];
                let omc = excelFileSheets.Hoja1[i]["OMC"];
                let sutunac = excelFileSheets.Hoja1[i]["SUTUNAC"];
                let segmasvida = excelFileSheets.Hoja1[i]["+VIDA SEGURO DE ACCIDENTES"];
                let segmasvidapension = excelFileSheets.Hoja1[i]["+VIDA SEG.ACC.PENSION"];
                let segrimac = excelFileSheets.Hoja1[i]["RIMAC INTERNAC.CIA SEG."];
                let sudunacjc = excelFileSheets.Hoja1[i]["SUDUNAC( JORGE CASTILLO P)"];
                let seginterseguro = excelFileSheets.Hoja1[i]["INTERSEGURO"];
                let segmapfre = excelFileSheets.Hoja1[i]["MAPFRE-PERU"];
                let colenfermeros = excelFileSheets.Hoja1[i]["COLEGIO ENFERMEROS"];
                let seglapositiva = excelFileSheets.Hoja1[i]["LA POSITIVA VIDA"];
                let dsctojudicial = excelFileSheets.Hoja1[i]["DESCUENTO JUDICIAL"];
                let onpcas = excelFileSheets.Hoja1[i]["ONP CAS"];
                let afpaporteoblig = excelFileSheets.Hoja1[i]["AFP APORTE OBLIGATORIO"];
                let afptasaefectiva = excelFileSheets.Hoja1[i]["AFP TASA EFECTIVA"];
                let afpcom = excelFileSheets.Hoja1[i]["AFP COM.PORC.CAS"];
                let comisporcent = excelFileSheets.Hoja1[i]["COMISION PORCENTUAL"];
                let segprima = excelFileSheets.Hoja1[i]["PRIMA SEGURO"];
                let onp = excelFileSheets.Hoja1[i]["ONP"];
                let dl2530 = excelFileSheets.Hoja1[i]["D.L. 20530"];
                let cuartacat = excelFileSheets.Hoja1[i]["4TA CATEGORIA"];
                let quintacat = excelFileSheets.Hoja1[i]["5TA CATEGORIA"];
                let essaludPensionistas = excelFileSheets.Hoja1[i]["ESSALUD PENS."];
                let totaldscts = parseFloat(excelFileSheets.Hoja1[i]["TOTAL DESCUENTOS"]).toFixed(2);
                
                
                let egresosArray = [];
                if(typeof faltasyotardanzas!='undefined'){
                    egresosArray.push({"texto":"Faltas y/o tardanzas","valor":faltasyotardanzas});
                }
                if(typeof sudunaccp!='undefined'){
                    egresosArray.push({"texto":"Sudunac","valor":sudunaccp});
                }
                if(typeof tespublico!='undefined'){
                    egresosArray.push({"texto":"Tes Publ.","valor":tespublico});
                }
                if(typeof cajamunareq!='undefined'){
                    egresosArray.push({"texto":"Caja Mun. de Arequ.","valor":cajamunareq});
                }
                if(typeof cooplaunion!='undefined'){
                    egresosArray.push({"texto":"Coop. la Union","valor":cooplaunion});
                }
                if(typeof coopsanmiguel!='undefined'){
                    egresosArray.push({"texto":"San Miguel Ex Coop. Pond.","valor":coopsanmiguel});
                }
                if(typeof otrossudunac!='undefined'){
                    egresosArray.push({"texto":"Otros","valor":otrossudunac});
                }
                if(typeof bancognb!='undefined'){
                    egresosArray.push({"texto":"Banco GNB","valor":bancognb});
                }
                if(typeof coopeltumi!='undefined'){
                    egresosArray.push({"texto":"Coop. 'El Tumi'","valor":coopeltumi});
                }
                if(typeof asocjubunac!='undefined'){
                    egresosArray.push({"texto":"ASOC. JUB.UNAC","valor":asocjubunac});
                }
                if(typeof bancoscotiabank!='undefined'){
                    egresosArray.push({"texto":"Scotiabank Perú","valor":bancoscotiabank});
                }
                if(typeof ceuunac!='undefined'){
                    egresosArray.push({"texto":"CEU-UNAC","valor":ceuunac});
                }
                if(typeof sutunacfall!='undefined'){
                    egresosArray.push({"texto":"Sutunac","valor":sutunacfall});
                }
                if(typeof sudunacfall!='undefined'){
                    egresosArray.push({"texto":"Fallec. (Sudunac)","valor":sudunacfall});
                }
                if(typeof cajachica!='undefined'){
                    egresosArray.push({"texto":"Caja chica o Tes.","valor":cajachica});
                }
                if(typeof regularonp!='undefined'){
                    egresosArray.push({"texto":"Regular ONP","valor":regularonp});
                }
                if(typeof regularabono!='undefined'){
                    egresosArray.push({"texto":"Regular abono","valor":regularabono});
                }
                if(typeof omc!='undefined'){
                    egresosArray.push({"texto":"OMC","valor":omc});
                }
                if(typeof sutunac!='undefined'){
                    egresosArray.push({"texto":"Sutunac","valor":sutunac});
                }
                if(typeof segmasvida!='undefined'){
                    egresosArray.push({"texto":"+Vida Seguro","valor":segmasvida});
                }
                if(typeof segmasvidapension!='undefined'){
                    egresosArray.push({"texto":"+Vida.Seg.Acc.Pens","valor":segmasvidapension});
                }
                if(typeof segrimac!='undefined'){
                    egresosArray.push({"texto":"Seg. Rimac Inter.","valor":segrimac});
                }
                if(typeof sudunacjc!='undefined'){
                    egresosArray.push({"texto":"Sudunac (Jorge Castillo P)","valor":sudunacjc});
                }
                if(typeof seginterseguro!='undefined'){
                    egresosArray.push({"texto":"Seg. Interseguro","valor":seginterseguro});
                }
                if(typeof segmapfre!='undefined'){
                    egresosArray.push({"texto":"Mapfre Perú","valor":segmapfre});
                }
                if(typeof colenfermeros!='undefined'){
                    egresosArray.push({"texto":"Coleg. Enfermeros","valor":colenfermeros});
                }
                if(typeof seglapositiva!='undefined'){
                    egresosArray.push({"texto":"La positiva Vida","valor":seglapositiva});
                }
                if(typeof dsctojudicial!='undefined'){
                    egresosArray.push({"texto":"Des. Judicial","valor":dsctojudicial});
                }
                if(typeof onpcas!='undefined'){
                    egresosArray.push({"texto":"ONP CAS","valor":onpcas});
                }
                if(typeof afpaporteoblig!='undefined'){
                    egresosArray.push({"texto":"AFP Aporte Obl.","valor":afpaporteoblig});
                }
                if(typeof afptasaefectiva!='undefined'){
                    egresosArray.push({"texto":"AFP Tasa Ef.","valor":afptasaefectiva});
                }
                if(typeof afpcom!='undefined'){
                    egresosArray.push({"texto":"AFP Com. Porc. CAS","valor":afpcom});
                }
                if(typeof comisporcent!='undefined'){
                    egresosArray.push({"texto":"Comisión Porc.","valor":comisporcent});
                }
                if(typeof segprima!='undefined'){
                    egresosArray.push({"texto":"Prima Seguro","valor":segprima});
                }
                if(typeof onp!='undefined'){
                    egresosArray.push({"texto":"ONP","valor":onp});
                }
                if(typeof dl2530!='undefined'){
                    egresosArray.push({"texto":"D.L. 20530","valor":dl2530});
                }
                if(typeof cuartacat!='undefined'){
                    egresosArray.push({"texto":"4ta Categoria","valor":cuartacat});
                }
                if(typeof quintacat!='undefined'){
                    egresosArray.push({"texto":"5ta Categoria","valor":quintacat});
                }
                if(typeof essaludPensionistas!='undefined'){
                    egresosArray.push({"texto":"Essalud Pens.","valor":essaludPensionistas});
                }
                

                //ESSALUD
                let aportes = excelFileSheets.Hoja1[i]["APORTES"];

                let aportesInfo;
                let totalAportes;
                if(tipoBoleta=="PLANILLA - PENSIONES"){
                    aportesInfo={"texto":"","valor":""};
                    totalAportes = 0;
                }else{
                    aportesInfo = {"texto":"ESSALUD","valor":aportes};
                    totalAportes = aportes;
                }

                //TOTAL LIQUIDO
                let totalLiquido = parseFloat(excelFileSheets.Hoja1[i]["TOTAL LÍQUIDO"]).toFixed(2);
                
                //#endregion

                
                //Obteniendo hora y fecha actual para guardar nombre boleta PDF
                const date = new Date().toLocaleString({ timeZone: "America/Lima" });
                const newDate = date.split(' ');
                const fecha = newDate[0].replaceAll('/','');
                const hora = newDate[1].replaceAll(':','');
                const filename = codigo+'-'+fecha+'-'+hora+'.pdf';


                let fileRendered = await ejs.renderFile(path.join(__dirname,'..','boletas/templateMail/index.ejs'),{
                    tipoBoleta,
                    mesPago,
                    codigo,
                    apenom,
                    dni,
                    dependencia,
                    carnetEssalud,
                    afp,
                    tpers,
                    diasLaborados,
                    fechaIngreso,
                    nivelRem,
                    nroCuenta,
                    cargoEstructural,
                    condLaboral,
                    tServicios,
                    cuspp,
                    horasLaboradas,
                    correo,
                    //Ingresos
                    ingresosArray,
                    totalingresos,
                    //Egresos
                    egresosArray,
                    totaldscts,
                    //ESSALUD
                    aportesInfo,
                    totalAportes,
                    //TOTALLIQUIDO
                    totalLiquido                  
                })  
                
                let pdfCreado = await createPDF(fileRendered,{timeout: '540000'});
                await fs.writeFileSync(`./boletas/emitidas/${filename}`,pdfCreado);
                //Enviando email
                await transporter.sendMail({
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
                console.log("PDF GUARDADO Y ENVIADO POR CORREO");

     
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
        console.log(error);
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