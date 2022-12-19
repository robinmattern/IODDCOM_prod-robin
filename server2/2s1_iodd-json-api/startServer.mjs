#! /usr/bin/env node

    import  shelljs         from 'shelljs' 
      var   aURI        =   import.meta.url
      var __filename    =   aURI.replace( /^.+\//, "" )
      var __dirname     =   aURI.replace( `/${__filename}`, "" ).replace( "file:///", "" )

//                          myExec( `../node_modules/.bin/json-server --watch api/models/db.json` )
//     var  aFileList    =  await myExec( `rdir -r 9` )
//     var  aFileList    =  await myExec( `ls -l` )
//          console.log( `  aFileList:\n${aFileList}` )

            shelljs.cd( __dirname )
            console.log(  shelljs.pwd().stdout ); // shelljs.exit()  

       var  aResult      =  await myExec( `npm start` )
        if (aResult) {
            console.log( `  aResult:\n${aResult}` )
            }
//--------  ------------ =  ------------------------------------------

async  function myExec( aCmd, aOpts )  { 

       var  shelljs_     =  ( cmd, opts ) => new Promise( resolve => {
            shelljs.exec(     cmd, opts, (code, stdout, stderr) => { resolve( { code, stdout, stderr } ); } )
            } )
       var  pResult      =    await shelljs_( aCmd, aOpts );
        if (pResult.stderr) { console.log( "  Shell Error:", pResult.stderr ); return "" }
        if (pResult.stdout) {                        return  pResult.stdout.replace( /\n$/, "" ) }
            } 
//--------  ------------ =  ------------------------------------------
