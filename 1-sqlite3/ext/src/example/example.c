
/* Extencion de ejemplo de Sqlite3 	*/
/* 2021-07-26 By J. Elihu 			*/

#include "..\sqlite3ext.h"
SQLITE_EXTENSION_INIT1
#include <assert.h>

/* Insert your extension code here */
static void noopfunc(
  sqlite3_context *context,
  int argc,
  sqlite3_value **argv
){
  assert( argc==1 );
  sqlite3_result_value(context, argv[0]);
}



#ifdef _WIN32
__declspec(dllexport)
#endif
/* TODO: Change the entry point name so that "extension" is replaced by
** text derived from the shared library filename as follows:  Copy every
** ASCII alphabetic character from the filename after the last "/" through
** the next following ".", converting each character to lowercase, and
** discarding the first three characters if they are "lib".
*/
int sqlite3_example_init(
  sqlite3 *db, 
  char **pzErrMsg, 
  const sqlite3_api_routines *pApi
){
  int rc = SQLITE_OK;
  SQLITE_EXTENSION_INIT2(pApi);
  
  (void)pzErrMsg;  /* Unused parameter */
  rc = sqlite3_create_function(db, "myCustomFunc", 1, SQLITE_UTF8|SQLITE_INNOCUOUS|SQLITE_DETERMINISTIC, 0, noopfunc, 0, 0);
  return rc;
  
}