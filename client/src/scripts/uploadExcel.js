const URI = "http://localhost:8000/";
const {ipcRenderer} = require('electron');

const cargarExcel = (e)=>{
    e.preventDefault();
    let excelFile = document.getElementById("excelFile").files[0];
    let formularioCargarExcel = document.getElementById("cargarExcelForm");
    const formdata = new FormData();
    formdata.append('excelFile',excelFile);

    fetch(`${URI}cargarExcel`,{
        method: 'POST',
        body: formdata
        
    }).then(res=>res.json()).then(data=>{
        if(data.message=="Excel cargado"){
            Swal.fire({
                icon: 'success',
                title:'Éxito',
                text: 'Excel cargado!',
            })
        }else{
            Swal.fire({
                icon: 'error',
                title:'Ocurrio un error',
                text: 'Error al cargar excel!',
            })
        }
    });
    formularioCargarExcel.reset();

}

const enviarCorreos = (e)=>{
    e.preventDefault();
    let mesPago = document.getElementById("mesPago").value;
    let tipoBoleta = document.getElementById("tipoBoleta").value;
    let enviandoCorreosModal = bootstrap.Modal.getOrCreateInstance(document.getElementById('enviandoCorreosModal'));
    enviandoCorreosModal.show();
    fetch(`${URI}enviarCorreos`,{
        method:'POST',
        headers:{
            'Content-type':'application/json'
        },
        body:JSON.stringify(
            {"mesPago": mesPago,
             "tipoBoleta":tipoBoleta})
    }).then(res=>res.json()).then(data=>{
        if(data.message=="Debe cargar un excel"){
            Swal.fire({
                icon: 'warning',
                title:'Advertencia!',
                text: 'Debe cargar un excel primero!',
            })
        }else if(data.message="correos enviados!"){
            Swal.fire({
                icon: 'success',
                title:'Éxito',
                text: 'Correos enviados!',
            })
        }else{
            Swal.fire({
                icon: 'error',
                title:'Ocurrio un error',
                text: 'Error desconocido!',
            })
        }
        enviandoCorreosModal.hide();
    })
}

