//---------------------------------------------------------------------------------------------------
 
       aTests='live in Browser'
//     aTests='test1 in NodeJS'

  if ( aTests.match( /test1/ ) ) { 

       var  pJSON    =  parseJSON( '../json/IODD-members_u2a.json.js' )
       var  aHTML    =  fmtMembers(  pJSON )

            console.log( aHTML )
            }

//---------------------------------------------------------------------------------------------------

function  fmtMembers( pJSON ) {     
       var  mMembers =  pJSON.results.items

       var  aHTML    =  mMembers.map( fmtMember ).join( "\n" )
//     var  mHTMLs=[];  mMembers.forEach( ( pMember, i ) => { fmtMember( pMember, i ) } ); aHTML = mHTMLs.join( "\n" ) 
    return  aHTML
            
//     ---  -------  =  -----------------------------------

  function  fmtMember( pMember, i ) {

       var  aMI      =     pMember.middlename;  aMI = ( aMI  > "" ) ?   `${ aMI.substr(0,1) }. ` : ""
       var  aName    = `${ pMember.firstname }${aMI} ${ pMember.lastname }`
       var  aPhone   =     pMember.phone1 + ( pMember.phone2 > ""   ? `, ${ pMember.phone2  }` : "" )
       var  aEmail   =     pMember.email

       var  aRow     = `<tr id="R${ `${ i + 1 }`.padStart( 3, "0" ) }">\n`
                     + `  <td><strong><a href="syschangepassword.js?username=${ aName }">${ aName }</a></strong></td>\n`
                     + `  <td><small ><a href="mailto:${ aEmail }">Email Address</a></small></td>\n`
                     + `  <td><small >${ aPhone }&nbsp;&nbsp;&nbsp;</small></td>\n`
                     + `</tr>\n`

//          mHTMLs.push( aRow )                  
     return aRow
            }   // eof  fmtMember
//     ---  -------  =  ----------------------------------
            }   // eof  fmtMembers
//--------  -------  =  -------------------------------------------------------

  function  parseJSON(  aFile ) {
       var  pFS      =  require( 'fs' )
       var  aDir     =  __dirname
       var  aText    =  pFS.readFileSync( `${aDir}/${aFile}`, "ASCII" )
            aText    =  aText.replace( //g, "'" )
            aText    =  aText.replace( //g, "'" )

        if (aFile.match( /\.json$/)) {   
       var  pJSON    =  JSON.parse( aText )
            }
        if (aFile.match( /\.js$/)) {   
            eval( aText )    
            }

    return  pJSON
            }   // eof  parseJSON
//--------  -------  =  -------------------------------------------------------

//---------------------------------------------------------------------------------------------------

