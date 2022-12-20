  import   shell           from 'shelljs' 
  import   util            from 'util' 
  import   fs              from 'fs' 
  import { fileURLToPath } from 'url'
  import { dirname       } from 'path'

   const __filename = fileURLToPath(import.meta.url)
   const __dirname  = dirname(__filename)

/* function  MyExec( cmd, opts ) { return new Promise( resolve => {
       shell.exec( cmd, opts, (code, stdout, stderr) => { resolve( { code, stdout, stderr } ); } )
       } ) } */

const  MyExec = ( cmd, opts ) => new Promise( resolve => {
            shell.exec( cmd, opts, (code, stdout, stderr) => { resolve( { code, stdout, stderr } ); } )
            } )

async  function MyExec2( aCmd )  { 
       var  pResult =  await MyExec( aCmd );
        if (pResult.stdErr) { console.log( "error:", pResult.stdErr ) }
        if (pResult.stdout) {                return  pResult.stdout.replace( /\n$/, "" ) }
            }       

//     var  aDir  =  await MyExec2( "ls -l" )
//     var  aDir  =  await MyExec2( "pwd" )
       var  aDir  =  __dirname

// -----------------------------------------------------------------------------------------

       var  pJSON = { } 

            await savEm( ) 

async function savEm() {  

            console.log( "" )

            await getJSON( 'SELECT * FROM members_view' )
            await getJSON( 'SELECT * FROM projects_view' )
            await getJSON( 'SELECT * FROM members_projects_view' )
//          await getJSON( 'SELECT FirstName, LastName FROM members WHERE MemberID < 3' )
//          await getJSON( 'SELECT * FROM members  WHERE MemberID  < 3' )
//          await getJSON( 'SELECT * FROM projects WHERE ProjectID < 3' )
//          await getJSON( 'SELECT * FROM members_projects WHERE ProjectID < 3' )
 
//          console.log( inspect( pJSON ) )

       var  aJSON  = "pJSON =\n" + inspect( pJSON ) + "\n\n   console.log( pJSON )"
 
//     var  aTS    =  await MyExec( "date '+%y%m%d.%H%M' " )
       var  aTS    =  fmtTS( 4 )   
       var  aFile1 = `${aDir}/db_v${aTS}.json.js` 
       var  aFile2 = `${aDir}/db_v${aTS}.json` 

            fs.writeFileSync( aFile1, aJSON )
            fs.writeFileSync( aFile2, JSON.stringify( pJSON ) )

            console.log( "" )
            console.log( `  Saved ${aFile1}` )
            console.log( `    and ${aFile2}` )
            }
// -----------------------------------------------------------------------------------------
     
 async function  getJSON( aSQL ) {

       var  aCmd      = `mysqlsh -u nimdas -pFormR\!1234 --host 45.32.219.12 --json=pretty --database=iodd --sql  --execute "${aSQL}"`
       var  aTable    =  aSQL.match( /FROM (\w+)( WHERE)*/i ); aTable = aTable[1] ? aTable[1] : ''
       console.log( aCmd )
       var  pResult   =  await MyExec( aCmd, { silent: true } );
        if (pResult.stdout) {             parseJSON( pResult.stdout ); return }
        if (pResult.stdErr) { console.log( "error:", pResult.stdErr ) }

//     ---------------------------------------------------------------------

  function  parseJSON( aText ) { 

       var  nPos      =  aText.match( /.+"hasData":/ ); nPos=nPos != null ? nPos.index : 0
  if (nPos == 0 ) {return} 
       aText     = "{ " + aText.substr( nPos )
    try { 
       var  pJson     =  JSON.parse( aText )
   } catch( pErr ) { 
            console.log( pErr.message )
       var  nPos      =  pErr.message.replace( /.+position /, "" ) * 1
            aText     =  aText.replace( /\n/g, "|" ) 
            console.log( aText.substr( nPos - 20, 60 ) )
            console.log( "....................^........................................")
            console.log( "                    | ")
            return 
            }   
       var  mRows = pJson.rows ? pJson.rows : null    
        if (mRows) { 
            pJSON[ aTable ] = mRows  
//          console.dir( aTable, mRows )
            console.log( `  Exported ${mRows.length} rows for "${aSQL}"` )
            }
       } // eof parseJSON 
//     ---------------------------------------------------------------------
   }
// -----------------------------------------------------------------------------------------

  function  inspect( pObj, nLv ) { return util.inspect( pObj, { depth: nLv ?nLv : 99 } ) }

  function  fmtTS( nFmt, nHrs ) {
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
       var  d = (new Date((new Date) - (nHrs ? nHrs :     5       ) * 60 * 60 * 1000)).toISOString()            // .(90614.03.2 RAM Use Default TheOffsetHrs)
       var  i = ((nFmt || '') + '').match( /[12345678]/ ) ? nFmt - 1 : 5 - 1, j = [11, 11, 13][i-5]
       var  p = [3, 2, 0, 3, 3, 2, 2, 0][i]; var l = [7, 8, 10, 13, 16, 19, 21, 23][i]                          // .(90802.05.2 RAM was 18)
            d =  d.substr( p, l ).replace( /[-:.]/g, '').replace(/T/, '.')
    return  l > 16 ? d.substr(0, j) + '.' + d.substr(j) : d
            }


