const URI = "http://localhost:8000/";


const obtenerArchivos = ()=>{
    const filesTable = document.getElementById('filesTable'); 
    filesTable.innerHTML = '';
    fetch(`${URI}obtenerArchivos`).then(res=>res.json()).then(data=>{
       for(i in data){
        filesTable.innerHTML+=  `<tr>
                                    <td>${data[i].fileName}</td>
                                    <td>pdf</td>
                                    <td>${data[i].fileSize} bytes</td>
                                    <td><a href="${URI}descargarArchivo/${data[i].fileName}" class="btn btn-info">Descargar</a></td>
                                  </tr>`
       }
       $('#tableSubidos').DataTable({
            pageLength:5,
            destroy:true,
            dom: 'Bfrtip',
            buttons: [
                'copy', 'csv', 'excel', 'pdf', 'print'
            ],
            "language": {
                "lengthMenu": "Display _MENU_ records per page",
                "zeroRecords": "No hay datos disponibles",
                "info": "Mostrando páginas _PAGE_ de _PAGES_",
                "infoEmpty": "No hay datos disponibles",
                "infoFiltered": "(filtered from _MAX_ total records)",
                "paginate":{
                    "previous":"Anterior",
                    "next":"Siguiente"
                },
                "search":"Buscar"
            }
        });
    })
}


obtenerArchivos();
