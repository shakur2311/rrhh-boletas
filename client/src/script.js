const URI = "http://localhost:8000/";
const {ipcRenderer} = require('electron');


//Open emitir y buscar boletas ventana
const emitirBoletas = ()=>{
    ipcRenderer.send("open:emitirBoletasWindow")
}

const buscarBoletas = ()=>{
    ipcRenderer.send("open:buscarBoletasWindow");
}



//Esto lo comenté porque ya no fue necesario hacer un metodo para descargar un archivo, lo puedo realizar desde
// el mismo backend con res.download()
/* const descargarArchivo = (nombreArchivo)=>{
    try {
        fetch(`${URI}descargarArchivo/${nombreArchivo}`).then(res=>res.json()).then(data=>{
            console.log(data);
        });
    } catch (error) {
        console.log(error)
    }
} */


