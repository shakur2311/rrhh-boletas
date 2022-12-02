const {connection} = require('../database/connection');


const login = (req,res)=>{
    try {
        if(req.body.user!="" && req.body.user!=null && req.body.pass!="" && req.body.pass!=null){
            connection.query(
                'SELECT * FROM usuarios WHERE userUsuario = ?',[req.body.user],
                function(err, results, fields) {
                  if(results.length>0){
                    let usuario = results[0].userUsuario;
                    connection.query(
                        'SELECT passUsuario FROM usuarios WHERE userUsuario = ?',[usuario],
                        function(err,results,fields){
                            let pass = results[0].passUsuario;
                            if(pass==req.body.pass){
                                res.json({"mensaje":"ingresaste"});
                            }else{
                                res.json({"mensaje":"Contraseña incorrecta!"});
                            }
                        }
                    )
                  }else{
                    res.json({"mensaje":"Usuario no existe!"});
                  }
                }
              );
        }
        
    } catch (error) {
        res.json({mensaje:'Error al hacer login'})
    }
}

module.exports = {login};