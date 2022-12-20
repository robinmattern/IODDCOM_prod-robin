#! /usr/bin/env node

   import   dotenv          from 'dotenv'
   import   shell           from 'shelljs'
   import   util            from 'util'
   import   fs              from 'fs'

//--------  ----------  =  -------------------------------------------------------------

      var   aURI        =   import.meta.url
      var __filename    =   aURI.replace( /^.+\//, "" )
      var __dirname     =   aURI.replace( `/${__filename}`, "" ).replace( "file:///", "" )

            dotenv.config()

      var   aDBName     =  process.env.DBNAME || ''
      var   aDBHost     =  process.env.DBHOST || '127.0.0.1'
      var   aDBUser     =  process.env.DBUSER || ''
      var   aDBPass     =  process.env.DBPASS || ''
      var   aReNameEm   =  process.env.RENAME_EM || ''
      var   mDBSQLs     =  [], i = 0

    while ( process.env[  `DBSQL${++i}` ] ) {
            mDBSQLs.push(  process.env[ `DBSQL${i}` ] )
            }
//          ----------  =  --------------------------------

       var  pJSON       = { }

            await savEm( mDBSQLs )

//--------  ----------  =  -------------------------------------------------------------

async function savEm( mDBSQLs ) {

//          mDBSQLs.forEach(           aSQL   => await getJSON( aSQL )   )
//          mDBSQLs.forEach(         ( aSQL ) => await getJSON( aSQL )   )
//          mDBSQLs.forEach( function( aSQL )  { await getJSON( aSQL ) } )
//          mDBSQLs.forEach( function( aSQL )  { return await getJSON( aSQL ) } )
//     for (i =0; i < mDBSQLs.length; i++) {     await getJSON( mDBSQLs[i] ) }

       for (i in mDBSQLs) {                      await getJSON( mDBSQLs[i] ) }

       var  ifProcess   = "if ( typeof( process ) != 'undefined' ) {"                  // .(21218.01.1 Was "if (process) { "
//     var  aJSON       = "var pJSON =\n" + inspect( pJSON ) + "\n\n  if (process) { var util = await import('util'); console.log( util.inspect( pJSON, { depth: 99 } ) ) }"
//     var  aJSON       = "var pJSON =\n" + inspect( pJSON ) + "\n\n  if (process) { console.log(    require('util'               ).inspect( pJSON, { depth: 99 } ) ) }"
//     var  aJSON       = "var pJSON =\n" + inspect( pJSON ) + "\n\n  if (process) {  var util = require('util'); console.log( util.inspect( pJSON, { depth: 99 } ) ) }"
       var  aJSON       = "var pJSON =\n" + inspect( pJSON ) + "\n\n  "+ ifProcess +" var util = require('util'); console.log( util.inspect( pJSON, { depth: 99 } ) ) }"  // .(21218.01.2)

//          ----------  =  --------------------------------

        if (aReNameEm  == 'true' || aReNameEm.match( /toold|tonew/i) ) {

       var  bOldTS      = aReNameEm.match( /tonew/i ) == null; console.log( "" )
            backup( `${__dirname}/db.json.js`, bOldTS )
            backup( `${__dirname}/db.json`,    bOldTS ); // process.exit()

       var  aTS         =  ''
        } else {
       var  aTS         =  `_v${fmtTS( 4 ) }`
            }
//          ----------  =  --------------------------------

       var  aFile1      = `${__dirname}/db${aTS}.json.js`
       var  aFile2      = `${__dirname}/db${aTS}.json`

            fs.writeFileSync( aFile1, aJSON )
            fs.writeFileSync( aFile2,  JSON.stringify( pJSON ) )

            console.log(  "" )
            console.log(  `    Saved: ${aFile1}` )
            console.log(  `      and: ${aFile2}` )
            }
//--------  ----------  =  -------------------------------------------------------------

 async function  getJSON( aSQL ) {

//     var  aCmd        = `mysqlsh -u   nimdas   -pFormR\!1234 --host 127.0.0.1  --json=pretty --database=iodd       --sql --execute "${aSQL}"`
       var  aCmd        = `mysqlsh -u ${aDBUser} -p${aDBPass}  --host ${aDBHost} --json=pretty --database=${aDBName} --sql --execute "${aSQL}"`
            console.log(  `\n$ ${ aCmd.replace( /!/, "\\!" ) }` )

       var  aTable      =  aSQL.match( /FROM (\w+)( WHERE)*/i ); aTable = aTable != null ? aTable[1] : ''
            aTable      =  aTable.replace( /_view/i, "" )

//          ----------  =  --------------------------------
       try {
       var  pResult     =  await MyExec( aCmd, { silent: true } );
        if (pResult.stdout) {
                           parseJSON( pResult.stdout ); return }
        if (pResult.stdErr) {
            console.log(`* error:${ pResult.stdErr }`) }
       } catch( pErr ) {
            console.log(`* error:${ pErr }`)
            return
            } // err MyExec( aCmd )
//          ----------  =  --------------------------------
//     ---------------------------------------------------------------------

  function  parseJSON( aText ) {

       var  pMatch      =  aText.match( /.+"error":/ )
       if ( pMatch != null) {
            aText       =  "{ " + aText.substr( pMatch.index ) // }
       var  pJson       =  JSON.parse( aText )
       var  pError      =  pJson.error
            console.log(`  * Error: No rows were exported for ${aTable}: "${aSQL}"`)
            console.log(`           MySQL Error ${pError.code}: ${pError.message}` )
            return
            }
//          ----------  =  --------------------------------

       var  pMatch        =  aText.match( /.+"hasData":/ )
        if (pMatch == null ) {
            console.log(`  * Error: No rows were exported for ${aTable}: "${aSQL}"`)
            return
            }
//          ----------  =  --------------------------------

    try {   aText       = "{ " + aText.substr( pMatch.index ) // }
       var  pJson       =  JSON.parse( aText )

   } catch( pErr ) {
            console.log(`  * Error: ${pErr.message}`)
       var  nPos        =  pErr.message.replace( /.+position /, "" ) * 1
            aText       =  aText.replace( /\n/g, "|" )
            console.log(`${aText.substr( nPos - 20, 60 )}`)
            console.log(" ....................^........................................")
            console.log("                    | ")
            return
            } // err JSON.parse( aText )
//          ----------  =  --------------------------------

       var  mRows = pJson.rows ? pJson.rows : null
        if (mRows) {
            pJSON[ aTable ] = mRows
//          console.dir( aTable, mRows )
            console.log(`  Exported ${mRows.length} rows for ${aTable}: "${aSQL}"`)
        } else {
            console.log(`  * Error: Opps, no rows were exported for ${aTable}: "${aSQL}"`)
            }
         } // eof parseJSON
//     ---  ----------  =  ------------------------------------------
       } // eof getJSON
//--------  ----------  =  ------------------------------------------------------------------------------

  function  MyExec(     cmd, opts ) { return new Promise( resolve => {
            shell.exec( cmd, opts, (code, stdout, stderr) => { resolve( { code, stdout, stderr } ); } )
            } ) }
//     ---  ----------  =  ------------------------------------------

/*     var  MyExec  = ( cmd, opts ) => new Promise( resolve => {
            shell.exec( cmd, opts, (code, stdout, stderr) => { resolve( { code, stdout, stderr } ); } )
            } ) */
//     ---  ----------  =  ------------------------------------------

     async  function MyExec2( aCmd ) {
       var  pResult     =  await MyExec( aCmd );
        if (pResult.stdErr) { console.log(`error:${pResult.stdErr}`) }
        if (pResult.stdout) {              return  pResult.stdout.replace( /\n$/, "" ) }
            }
//     ---  ----------  =  ------------------------------------------

  function  inspect( pObj, nLv ) { return util.inspect( pObj, { depth: nLv ?nLv : 99 } ) }

//     ---  ----------  =  ------------------------------------------

  function  fmtTS( nFmt, nHrs, dDate ) {
         // nHrs -> Hrs from GMT time (default: 4 EST, not 3 EDT or 7 PDT)
         // nFmt -> 1:    ymmdd               90523
         //         2:   yymmdd              190523
         //         3: yyyymmdd            20190523
         //         4:    ymmdd-hhnn          90523-1031
         //         5:    ymmdd-hhnnss        90523-103107  (default)
         //         6:   yymmdd-hhnnss       190523-1031-078                                                    // .(90802.04.1 RAM was -07)
         //         7:   yymmdd-hhnnssiii    190523-1031-07899
         //         8: yyyymmdd-hhnnssiii  20190523-1031-07899
//     var  d = (new Date((new Date) - (nHrs ? nHrs : TheOffsetHrs) * 60 * 60 * 1000)).toISOString()            // .(90614.03.2 RAM Use Default TheOffsetHrs)
       var  d = (new Date( ( dDate ? dDate : new Date ) - (nHrs ? nHrs : 5 ) * 60 * 60 * 1000)).toISOString()   // .(90614.03.2 RAM Use Default TheOffsetHrs)
       var  i = ((nFmt || '') + '').match( /[12345678]/ ) ? nFmt - 1 : 5 - 1, j = [11, 11, 13][i-5]
       var  p = [3, 2, 0, 3, 3, 2, 2, 0][i]; var l = [7, 8, 10, 13, 16, 19, 21, 23][i]                          // .(90802.05.2 RAM was 18)
            d =  d.substr( p, l ).replace( /[-:.]/g, '').replace(/T/, '.')
    return  l > 16 ? d.substr(0, j) + '.' + d.substr(j) : d

       } // eof fmtTS
//     ---  ----------  =  ------------------------------------------

  function  backup( aFile, bUseOldDate ) {
//     var  aTS    =  await MyExec( "date '+%y%m%d.%H%M' " )
        if (fs.existsSync( aFile ) ) {
       var  pStats =  fs.statSync( aFile )
       var  aTS    =  fmtTS( 4, 0, bUseOldDate ? pStats.mtimeMs : 0 )
       var  aExt   =  aFile.replace( /^.+\./, "" )
            aExt   =  aFile.match( /\.json\.js/ ) ? 'json.js' : aExt
       var  aFile1 =  aFile.replace( new RegExp( `\\.${aExt}$` ), `_v${aTS}.${aExt}` )
            fs.renameSync( aFile, aFile1 )
            console.log( `  Renamed: ${aFile}\n       to: ${aFile1}`)
            }
       }  // eof backup
//     ---  ----------  =  ------------------------------------------
//--------  ----------  =  ------------------------------------------------------------------------------


