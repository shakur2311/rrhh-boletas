const URI = "http://localhost:8000/";
const {ipcRenderer} = require('electron');

const login = ()=>{
    let user = document.querySelector("#userInput").value;
    let pass = document.querySelector("#passInput").value;
    
    if(user!="" && pass!=""){

        fetch(`${URI}login`,{
            method:'POST',
            body:JSON.stringify({
                user,
                pass
            }),
            headers: { "Content-Type": "application/json" }

        }).then(res=>res.json()).then(data=>{
            if(data.mensaje=="Usuario no existe!"){
                Swal.fire({
                    icon: 'warning',
                    title:'Ocurrio un error',
                    text: 'El usuario es incorrecto!',
                })
            }else if(data.mensaje=="Contraseña incorrecta!"){
                Swal.fire({
                    icon: 'warning',
                    title:'Ocurrio un error',
                    text: 'La contraseña es incorrecta!',
                })
            }else if(data.mensaje=="ingresaste"){
                ipcRenderer.send("login");
            }else{
                Swal.fire({
                    icon: 'warning',
                    title: 'Ocurrio un error',
                    text: 'Error desconocido!',
                })
            }
        });
    }else{
        Swal.fire({
            icon: 'error',
            title: 'Importante',
            text: 'Llenar todos los campos!',
          })
    }
}

