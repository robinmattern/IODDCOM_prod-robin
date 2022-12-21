                           require( 'dotenv' ).config( ); 

       var  mysql       =  require( 'mysql'  );

//     var  connection;
       var  config      =
             { host     :  process.env.DBHOST
             , port     :  process.env.DBPORT || 3306  
             , user     :  process.env.DBUSERNAME
             , password :  process.env.DBPASSWORD 
             , database :  process.env.DBNAME
               };
//          console.log( config ); process.exit()

       var  connection  =  mysql.createConnection( config )
            connection.connect( onConnect ) 

  function  onConnect( error, okPacket ) {
       var  config2     =  connection._protocol._config
       var  connectedDB = `${config2.host}:${config2.port}, User:${config2.user}@${config2.password}`      
        if (error) {
       var  errorno     =  error.errorno ? error.errorno : error.errno
            console.log(`\n  ** MySQL Server error ${errorno}: ${error.code}\n     ${error.sqlMessage}` );
            console.log(`     Trying to connect to ${connectedDB}\n` )  
                           process.exit()     
        } else {
            console.log(`  MySQL Server is connected to ${connectedDB}\n` )
            console.log(`  Press CTRL-C to stop server\n`)
            }

          }

      module.exports  = { connection }

