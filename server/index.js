const express = require('express');
const morgan = require('morgan');
const cors = require('cors');
const indexRoutes = require('./routes/indexRoutes')
const bodyParser = require('body-parser');
const app = express();

//Settings
app.set('PORT',8000||process.env.PORT);
app.set('name','Correos Masivos App');
app.set('view engine', 'ejs');

//Middlewares
app.use(morgan('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended:true}));
app.use(cors());

//Routes
app.use('/',indexRoutes);

app.get('*',(req,res)=>{
    res.send("No se encontró la ruta");
})






const server = app.listen(app.get('PORT'),()=>{
    console.log("Corriendo en puerto: "+app.get('PORT'));
})