       var  pFS      =  require( 'fs' )
       var  inspect  =  function( pObj ) { return require( 'util' ).inspect( pObj, { depth: 99 } ) }

//---------------------------------------------------------------------------------------------------

  function  shoJSON_Counts( aFile ) {
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

            shoJSON_Object( "IODD-members_u2c_export.json" )
            shoJSON_Counts( "IODD-members_u2c_export.json" )
            shoJSON_HTML(   "IODD-members_u2c_export.json" )
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

  function  shoJSON_HTML( aFile ) {
       var  pJSON    =  parseJSON( aFile )
       var  pMembers =  pJSON.results.items

       var  aHTML    =  pMembers.map( fmtMember ).join( "\n" )

            console.log( aHTML )

//     ---  -------  =  -----------------------------------

  function  fmtMember( pMember ) {

       var  aMI      =     pMember.middlename;  aMI = ( aMI  > "" ) ?   `${ aMI.substr(0,1) }. ` : ""
       var  aName    = `${ pMember.firstname }${aMI} ${ pMember.lastname }`
       var  aPhone   =     pMember.phone1 + ( pMember.phone2 > ""   ? `, ${ pMember.phone2  }` : "" )
       var  aEmail   =     pMember.email

       var  aRow     = '<tr>\n'
                     + `  <td><strong><a href="syschangepassword.asp?username=${ aName }">${ aName }</a></strong></td>\n`
                     + `  <td><small ><a href="mailto:${ aEmail }">Email Address</a></small></td>\n`
                     + `  <td><small >${ aPhone }&nbsp;&nbsp;&nbsp;</small></td>\n`
                     + "</tr>\n"
     return aRow
            }   // eof  fmtMember
//     ---  -------  =  ----------------------------------
            }   // eof  shoJSON_HTML
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
       var  aText    =  pFS.readFileSync( aFile, "ASCII" )
            aText    =  aText.replace( //g, "'" )
            aText    =  aText.replace( //g, "'" )
       var  pJSON    =  JSON.parse( aText )
    return  pJSON
            }   // eof  parseJSON
//--------  -------  =  -------------------------------------------------------

//---------------------------------------------------------------------------------------------------
