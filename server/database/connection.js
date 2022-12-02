const mysql2 = require("mysql2");

 const connection = mysql2.createConnection({
    host:'localhost',
    user:'root',
    password:'',
    database:'boletasdb'
 })
 

 module.exports = {connection};