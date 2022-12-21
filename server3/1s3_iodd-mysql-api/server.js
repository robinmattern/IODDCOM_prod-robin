                            require( 'dotenv').config();

//     var  mysql        =  require( 'mysql' );
       var  express      =  require( 'express' )
       var  bodyParser   =  require( 'body-parser' )

       var  app          =  express();

            app.use(        bodyParser.json() )
            app.use(        bodyParser.urlencoded( { extended: true } ) )

            app.route(    '/members'     ).get(    getMembers  )
            app.route(    '/members/:id' ).get(    getMembers  )
            app.route(    '/members/:id' ).post(   addMember   )
            app.route(    '/members/:id' ).delete( delMember   )
            app.route(    '/projects'    ).get(    getProjects )
            app.route(    '/projects/:id').get(    getProjects )

       var  port         =  process.env.APIPORT || 50000
            app.listen( port, ( ) => {    
            console.log( `  API Server located at http://localhost:${port}` )
            } )

       var{ connection } =  require( './connect.js' )

//--------  ------------ =  ------------------------------------------

  function  getMembers( request, response ) {

       var  id           =  request.params.id;
       var  sql          = `SELECT * FROM members ${ id ? `WHERE Id = ${id}` : `` }`
            console.log( `  sql: ${sql}`)
            connection.query( sql, onSelectMember )

  function  onSelectMember( error, results ) {

        if (error) { throw error; }
            response.status( 200 ).json( results );
            }
         } // eof getMembers
//--------  ------------ =  ------------------------------------------

  function  getProjects( request, response ) {

            connection.query( "SELECT * FROM projects", onSelect )

  function  onSelectMember( error, results ) {

        if (error) { throw error; }
            response.status( 200 ).json( results );
            }
         } // eof getMembers
//--------  ------------ =  ------------------------------------------

  function  addMember( request, response ) {

      var { FirstName, LastName, City } = request.body;
       var  sql          = "INSERT INTO Members( FirstName, LastName, City ) VALUES (?,?,?)"

            connection.query( sql, [ FirstName, LastName, City ], onInsertMember )

  function  onInsertMember( error, results ) {
        if (error) { throw error }
            response.status( 201 ).json( { "Member Added": results.affectedRows } );
            }
         } // eof addMember
//--------  ------------ =  ------------------------------------------

  function  delMember( request, response ) {

       var  id           =  request.params.id;
       var  sql          = "DELETE from Members where Id = ?"

            connection.query( sql, [ id ], onDelete )

  function  onDelete( error, results ) {
        if (error) { throw error }
            response.status( 201 ).json( { "Member Deleted": results.affectedRows } );
            }
         } // eof delMember
//--------  ------------ =  ------------------------------------------

            module.exports = app;
