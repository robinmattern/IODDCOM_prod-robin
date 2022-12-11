       var  inspect  =  function( pObj ) { return require( 'util' ).inspect( pObj, { depth: 99 } ) }

//---------------------------------------------------------------------------------------------------

     const  shoJSON_Counts = function( aFile ) {
       var  pJSON    =  parseJSON( aFile )

            console.log( shoHeader( aFile, pJSON ) )
            }
//--------  -------  =  -------------------------------------------------------

  function  shoJSON_Object( aFile ) {
       var  pJSON    =  parseJSON( aFile )

            console.log( shoHeader( aFile, pJSON ) )
            console.log( inspect( pJSON ) )
            }
//--------  -------  =  -------------------------------------------------------

  function  shoJSON_HTML( aFile ) {
       var  pJSON    =  parseJSON( aFile )

       var  aHTML    =  fmtMembers( pJSON )
            console.log( aHTML )
            }
//--------  -------  =  -------------------------------------------------------

            shoJSON_Object( "IODD-members_u2c_export.json" )
            shoJSON_Counts( "IODD-members_u2c_export.json" )
            shoJSON_HTML(   "IODD-members_u2a_export.json" )
            process.exit()

            shoJSON_Counts( "IODD-users_u2a_export.json" )
            shoJSON_Counts( "IODD-user_roles_u2a_export.json" )
            shoJSON_Counts( "IODD-roles_u2a_export.json" )
            shoJSON_Counts( "IODD-roles_tables_u2a_export.json" )
            shoJSON_Counts( "IODD-tables_u2a_export.json" )

            shoJSON_Counts( "IODD-members_u2c_export.json" )
            shoJSON_Counts( "IODD-projects_u2a_export.json" )
            shoJSON_Counts( "IODD-members_projects_u2a_export.json" )

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

  function  shoHeader(  aFile, pJSON ) {
       var  nCols    =  pJSON.results.columns.length
       var  nRows    =  pJSON.results.items.length
  return "\n-------------------------------------------------------------\n"
         + `${aFile} (${nCols} Columns, ${nRows} Records)\n`
         + "-------------------------------------------------------------\n"
            }   // eof  shoHeader
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
