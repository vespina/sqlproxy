* SQLPROXY.PRG
* PROXY CLASS TO SEND DML OPERATIONS TO A REMOTE TABLE WITH A
* LOCAL TABLE EQUIVALENT.
*
* AUTOR: V. ESPINA
* VERSION: 1.0
* 
* EXAMPLES:
*
* A) CREATE AN INSTANCE
*
* oProxy = CREATE("sqlProxy","c:\data\company.dbc!customers")
* oProxy.remoteConnStr = "..."
* ?oProxy.localTable -> "customers"
* ?oProxy.remoteTable -> "customers"
* ?oProxy.dbcPath -> "c:\data\company.dbc"
*
* oProxy = CREATE("sqlProxy","customers","erp_customers","conn-str")
* oProxy.remoteConnStr = "conn-str"
* ?oProxy.localTable -> "customers"
* ?oProxy.remoteTable -> "erp_customers"
* ?oProxy.dbcPath -> ""
*
* nConn = SQLSTRINGCONNECT("...")
* oProxy = CREATE("sqlProxy","customers",,nConn)
* oProxy.remoteConnStr = ""
* ?oProxy.localTable -> "customers"
* ?oProxy.remoteTable -> "customers"
* ?oProxy.dbcPath -> ""
*
* B) ADD CURRENT ROW ON REMOTE TABLE
* SELECT customers
* IF NOT oProxy.insertOne()
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* C) ADD CUSTOM ROW ON REMOTE TABLE
* SELECT customers
* SCATTER NAME oRow BLANK
* oRow.<column> = <value>
* oRow.<column> = <value>
* ...
* IF NOT oProxy.insertOne(oRow)
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* D) ADD MANY ROWS ON REMOTE TABLE USING A DATA CURSOR
* SELECT * FROM customers WHERE status = 'ACTIVE' INTO CURSOR qactives
* IF NOT oProxy.insertMany("qactives")
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* E) ADD MANY ROWS ON REMOTE TABLE USING A COLLECTION
* SELECT customers
* oRows = CREATE("Collection")
* SCAN FOR status = "ACTIVE"
*   SCATTER NAME oRow
*   oRows.Add(oRow)
* ENDSCAN
* IF NOT oProxy.insertMany(oRows)
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* F) UPDATE ONE ROW FROM LOCAL TABLE
* SELECT customers
* LOCATE FOR id = "001"
* REPLACE status WITH "INACTIVE"
* IF NOT oProxy.updateOne()
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* G) UPDATE ONE ROW MANUALLY
* SELECT customers
* LOCATE FOR id = "001"
* SCATTER NAME oRow
* oRow.status = "INACTIVE"
* IF NOT oProxy.updateOne(oRow)
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* H) UPDATE MANY ROWS
* oSets = oProxy.Where("status,lastupd","INACTIVE",DATETIME())
* oWhere = oProxy.Where("zone","001")
* IF NOT oProxy.updateMany(oSets, oWhere)
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* I) DELETE ONE ROW FROM LOCAL TABLE
* SELECT customers
* LOCATE FOR id = "001"
* IF NOT oProxy.deleteOne()
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* J) UPDATE ONE ROW MANUALLY
* IF NOT oProxy.updateOne("001")
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* K) DELETE MANY ROWS
* oWhere = oProxy.Where("zone","001")
* IF NOT oProxy.deleteMany(oWhere)
*    MESSAGEBOX(oProxy.lastError)
* ENDIF
*
* L) GET DATA FROM ONE RECORD
* oRow = oProxy.findOne("001")
* IF ISNULL(oRow)
*    MESSAGEBOX("Not found!")
* ELSE
*    MESSAGEBOX("Name: " + oRow.fullName)
* ENDIF
*
* M) FIND MANY ROWS
* oWhere = oProxy.Where("zone,status")
* oWhere.Zone = "001"
* oWhere.Status = "ACTIVE"
* IF oProxy.findMany(oWhere, "qresult")
*    SELECT qresult
*    BROWSE
* ENDIF
*

DEFINE CLASS sqlProxy AS Custom

	version = ""               && VERSION ACTUAL
	dbcPath = ""               && RUTA AL DBC LOCAL (OPCIONAL)
	localTable = ""            && NOMBRE DE LA TABLA (O ALIAS) LOCAL
	remoteTable = ""           && NOMBRE DE LA TABLA REMOTA
	primaryKey = ""            && CLAVE PRIMARIA DE LA TABLA 
	remoteConnStr = ""         && CADENA DE CONEXION A LA BD REMOTA
	lastError = ""             && ULTIMO ERROR OCURRIDO
	sqlConnHnd = 0             && ACCESOR AL HANDLE ODBC
	fieldList = ""             && LISTA DE COLUMNAS
	lastSqlCmd = ""            && ULTIMO COMANDO SQL ENVIADO
	
	HIDDEN _sqlInsertCmd       && COMANDO PARA INSERT
	HIDDEN _sqlUpdateCmd       && COMANDO PARA UPDATE
	HIDDEN _sqlConnHnd         && HANDLE ODBC
	
	PROCEDURE sqlConnHnd_Access
		IF EMPTY(THIS._sqlConnHnd)
			IF !EMPTY(THIS.remoteConnStr)
				THIS._sqlConnHnd = SQLSTRINGCONNECT(THIS.remoteConnStr)
			ELSE
				THIS.lastError = "No se ha indicado la cadena de conexion"
				THIS._sqlConnHnd = 0
			ENDIF
		ENDIF
		RETURN THIS._sqlConnHnd
		
	PROCEDURE sqlConnHnd_Assign(vNewVal)
		RETURN
		
	PROCEDURE version_access
		RETURN "1.0"
		
	PROCEDURE version_assign(vNewVal)
		RETURN
	
	
	* CONSTRUCTOR
	PROCEDURE Init(pcLocal, pcRemote, puConn)
		IF PCOUNT() = 1
			pcRemote = pcLocal
		ENDIF
		IF "!" $ pcLocal
			THIS.dbcPath = GETWORDNUM(pcLocal,1,"!")
			pcLocal = GETWORDNUM(pcLocal,2,"!")
		ENDIF
		THIS.localTable = pcLocal
		THIS.remoteTable = pcRemote
		THIS._sqlConnHnd = 0
		DO CASE
		   CASE VARTYPE(puConn) = "C"
		        THIS.remoteConnStr = puConn
		        
		   CASE VARTYPE(puConn) = "N"
		   	    THIS._sqlConnHnd = puConn
		ENDCASE
		RETURN
		
		
	* CAMBIAR LA SESION DE DATOS ASOCIADA
	PROCEDURE setDataSession(pnDSID)
		SET DATASESSION TO (pnDSID)
		RETURN	
		
		
	* INSERTAR UN REGISTRO EN LA TABLA REMOTA
	PROCEDURE insertOne
		PARAMETERS poData
		THIS.lastError = ""

		* PREPARAMOS EL COMANDO		
		LOCAL cCmd
		cCmd = THIS.getInsertCmd("poData")
		IF EMPTY(cCmd)
			RETURN .F.
		ENDIF
		
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ENVIAMOS EL COMANDO
		LOCAL lResult
		lResult = THIS._Execute(nConn, cCmd)
		RETURN lResult
		
		
		
	* INSERTAR VARIOS REGISTROS EN LA TABLA REMOTA
	PROCEDURE insertMany
		PARAMETERS puData
		THIS.lastError = ""

		* PREPARAMOS EL COMANDO		
		LOCAL cCmd
		cCmd = THIS.getInsertCmd("oRow")
		IF EMPTY(cCmd)
			RETURN .F.
		ENDIF
		
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ENVIAMOS EL COMANDO
		LOCAL lResult,oRow
		lResult = .F.
		DO CASE
		   CASE VARTYPE(puData) = "C" AND USED(puData)      && ALIAS
		   		SELECT (puData)
		   		GO TOP
		   		lResult = .T.
		   		SCAN WHILE lResult
		   		    SCATTER NAME oRow MEMO
					lResult = THIS._Execute(nConn, cCmd)
					SELECT (puData)
				ENDSCAN
				
			CASE VARTYPE(puData) = "O" AND LOWER(puData.baseClass) == "collection"    && COLECCION
			     FOR EACH oRow IN puData	
			     	lResult = THIS._Execute(nConn, cCmd)
			     	IF NOT lResult
			     		EXIT
			     	ENDIF
			     ENDFOR
			     
			OTHERWISE
				THIS.lastError = "El parametro indicado no corresponde ni a una coleccion ni a un alias de datos"
		ENDCASE
		RETURN lResult	
		
		
		
	
	* ACTUALIZAR UN REGISTRO EN LA TABLA REMOTA
	PROCEDURE updateOne
		PARAMETERS poData
		THIS.lastError = ""

		* PREPARAMOS EL COMANDO		
		LOCAL cCmd
		cCmd = THIS.getUpdateCmd()
		IF EMPTY(cCmd)
			RETURN .F.
		ENDIF
		
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ENVIAMOS EL COMANDO
		LOCAL lResult
		lResult = THIS._Execute(nConn, cSQL)
		RETURN lResult	
		


	* ACTUALIZAR VARIOS REGISTROS DESDE LA TABLA REMOTA
	PROCEDURE updateMany
		PARAMETERS poSets, poWhere			
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ARMAMOS EL COMANDO
		LOCAL cSQL,lResult,oData,nCount,cWhere,i,cSetList
		LOCAL ARRAY aCols[1]
		nCount = AMEMBERS(aCols, poSets, 0)
		cSetList = ""
		FOR i = 1 TO nCount
			cSetList = cSetList + IIF(i > 1,",","") + aCols[i] + " = ?poSets." + aCols[i]
		ENDFOR		
		nCount = AMEMBERS(aCols, poWhere, 0)
		cWhere = ""
		FOR i = 1 TO nCount
			cWhere = cWhere + IIF(i > 1,",","") + aCols[i] + " = ?poWhere." + aCols[i]
		ENDFOR
		cSQL = [UPDATE ] + THIS.remoteTable + [ SET ] + cSetList + [ WHERE ] + cWhere
		lResult = THIS._Execute(nCnn, cSQL)
		RETURN lResult
		


	* ELIMNINAR UN REGISTRO ESPECIFICO DESDE LA TABLA REMOTA
	PROCEDURE deleteOne
		PARAMETERS puPKExpr
		* OBTENEMOS LOS COMANDOS DE INSERT Y UPDATE
		* PARA ASEGURARNOS DE TENER TODOS LOS DATOS
		* NECESARIOS PARA ARMAR EL COMANDO
		IF EMPTY(THIS.getUpdateCmd())
			RETURN NULL
		ENDIF
		
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ARMAMOS EL COMANDO
		LOCAL cSQL,lResult
		cSQL = [DELETE FROM ] + THIS.remoteTable + [ WHERE ] + THIS.primaryKey + [ = ?puPKExpr]	
		lResult = THIS._Execute(nConn, cSQL, "qresult")
		RETURN lResult	




	* ELIMINAR VARIOS REGISTROS DESDE LA TABLA REMOTA
	PROCEDURE deleteMany
		PARAMETERS poWhere, pcOrderBy, pcCursor 	
		* OBTENEMOS LOS COMANDOS DE INSERT Y UPDATE
		* PARA ASEGURARNOS DE TENER TODOS LOS DATOS
		* NECESARIOS PARA ARMAR EL COMANDO
		IF EMPTY(THIS.getInsertCmd())
			RETURN NULL
		ENDIF
		
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ARMAMOS EL COMANDO
		LOCAL cSQL,lResult,oData,nCount,cWhere,i
		LOCAL ARRAY aCols[1]
		nCount = AMEMBERS(aCols, poWhere, 0)
		cWhere = ""
		FOR i = 1 TO nCount
			cWhere = cWhere + IIF(i > 1,",","") + aCols[i] + " = ?poWhere." + aCols[i]
		ENDFOR
		cSQL = [DELETE FROM ] + THIS.remoteTable + [ WHERE ] + cWhere
		lResult = THIS._Execute(nConn, cSQL)
		RETURN lResult


	* DEVOLVER UN REGISTRO ESPECIFICO DESDE LA TABLA REMOTA
	PROCEDURE findOne
		PARAMETERS puPKExpr, plReturnEmpty
	 	LOCAL nWkArea
	 	nWkArea = SELECT()
	 	
		* OBTENEMOS LOS COMANDOS DE INSERT Y UPDATE
		* PARA ASEGURARNOS DE TENER TODOS LOS DATOS
		* NECESARIOS PARA ARMAR EL COMANDO
		IF EMPTY(THIS.getInsertCmd()) ;
		OR EMPTY(THIS.getUpdateCmd())
			RETURN NULL
		ENDIF
		
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ARMAMOS EL COMANDO
		LOCAL cSQL,lResult,oData
		oData = NULL
		cSQL = [SELECT ] + THIS.fieldList + [ FROM ] + THIS.remoteTable + [ WHERE ] + THIS.primaryKey + [ = ?puPKExpr]	
		SELECT 0
		lResult = THIS._Execute(nConn, cSQL, "qresult")
		IF lResult
			SELECT qresult
			IF !EOF() OR plReturnEmpty
				SCATTER NAME oData MEMO
			ENDIF
			USE
		ENDIF
		SELECT (nWkArea)
		RETURN oData		



	* DEVOLVER VARIOS REGISTROS DESDE LA TABLA REMOTA
	PROCEDURE findMany
		PARAMETERS poWhere, pcOrderBy, pcCursor
			
		* OBTENEMOS LOS COMANDOS DE INSERT Y UPDATE
		* PARA ASEGURARNOS DE TENER TODOS LOS DATOS
		* NECESARIOS PARA ARMAR EL COMANDO
		IF EMPTY(THIS.getInsertCmd())
			RETURN NULL
		ENDIF
		
		* OBTENEMOS EL HANDLE DE CONEXION
		LOCAL nConn
		nConn = THIS.sqlConnHnd
		IF nConn = 0 
			RETURN .F.
		ENDIF
		
		* ARMAMOS EL COMANDO
		LOCAL cSQL,lResult,oData,nCount,cWhere,i
		LOCAL ARRAY aCols[1]
		cWhere = ""
		IF VARTYPE(poWhere) = "O"
			nCount = AMEMBERS(aCols, poWhere, 0)		
			FOR i = 1 TO nCount
				cWhere = cWhere + IIF(i > 1,",","") + aCols[i] + " = ?poWhere." + aCols[i]
			ENDFOR
		ENDIF
		cSQL = [SELECT ] + THIS.fieldList + [ FROM ] + THIS.remoteTable
		IF !EMPTY(cWhere)
			cSQL = cSQL + [ WHERE ] + cWhere
		ENDIF
		IF !EMPTY(pcORderBy)
			cSQL = cSQL + " ORDER BY " + THIS.pcOrderBy
		ENDIF
		pcCursor = EVL(pcCursor, "Q" + SYS(2015))
		SELECT 0
		lResult = THIS._Execute(nConn, cSQL, pcCursor)
		IF NOT lResult
			RETURN -1
		ENDIF
		RETURN RECCOUNT(pcCursor)




	* GENERAR UN OBJECT WHERE
	PROCEDURE Where(pcColList,pu0,pu1,pu2,pu3,pu4,pu5,pu6,pu7,pu8,pu9)
		LOCAL oWhere,i,puX, cColName
		oWhere = CREATEOBJECT("EMPTY")
		FOR i = 1 TO MIN(10,GETWORDCOUNT(pcColList))
			cColName = GETWORDNUM(pcColList,i)
			puX = "pu" + ALLTRIM(STR(i - 1))
			ADDPROPERTY(oWhere, cColName, EVALUATE(puX))
		ENDFOR
		RETURN oWhere
		

	* GENERAR EL CMD PARA INSERT	
	PROCEDURE getInsertCmd(pcSource)
	    LOCAL nWkArea
	    nWkArea = SELECT()
		pcSource = EVL(pcSource, "m")
		IF !EMPTY(THIS._sqlInsertCmd)   && YA HABIAMOS GENERADO EL COMANDO ?
			RETURN THIS._sqlInsertCmd   && ... LO DEVOLVEMOS DE UNA VEZ
		ENDIF 
		LOCAL cAlias                    && DE LO CONTRARIO, LO GENERAMOS    
		cAlias = THIS._openLocal()
		IF EMPTY(cAlias)                
			RETURN ""
		ENDIF
		THIS.lastError = "" 
		THIS._sqlInsertCmd = ""
		THIS.fieldList = ""
		LOCAL cColList,cInsList,i,cFieldName
		STORE "" TO cColList,cInsList,cFieldName
		SELECT (cAlias)
		FOR i = 1 TO FCOUNT()
		    cFieldNAme = LOWER(FIELD(i))
			cColList = cColList + IIF(i > 1,",","") + cFieldName
			cInsList = cInsList + IIF(i > 1,",","") + "?" + pcSource + "." + cFieldName
		ENDFOR
		USE IN (cAlias)
		THIS.fieldList = cColList
		THIS._sqlInsertCmd = "INSERT INTO " + THIS.remoteTable + " (" + cColList + ") VALUES (" + cInsList + ")"
		SELECT (nWkArea)
		RETURN THIS._sqlInsertCmd
			
		
	* GENERAR EL CMD PARA INSERT	
	PROCEDURE getUpdateCmd
	    LOCAL nWkArea
	    nWkArea = SELECT()
		IF !EMPTY(THIS._sqlUpdateCmd)   && YA HABIAMOS GENERADO EL COMANDO ?
			RETURN THIS._sqlUpdateCmd   && ... LO DEVOLVEMOS DE UNA VEZ
		ENDIF
		LOCAL cAlias                    && DE LO CONTRARIO, LO GENERAMOS    
		cAlias = THIS._openLocal()
		IF EMPTY(cAlias)                
			RETURN ""
		ENDIF
		THIS.lastError = "" 
		THIS._sqlUpdateCmd = ""
		LOCAL cSetList,i,cFieldName
		STORE "" TO cSetList,cFieldName
		SELECT (cAlias)
		FOR i = 1 TO FCOUNT()
		    cFieldNAme = LOWER(FIELD(i))
			cSetList = cSetList + IIF(i > 1,",","") + cFieldName + " = ?poData." + cFieldName
			IF EMPTY(THIS.primaryKey) AND (PRIMARY(i) OR CANDIDATE(i))
				THIS.primaryKey = KEY(i)
			ENDIF
		ENDFOR
		USE IN (cAlias)
		SELECT (nWkArea)
		IF EMPTY(THIS.primaryKey)
			THIS.lastError = "La tabla local no tiene un indice primario/candidato"
			RETURN ""
		ENDIF
		LOCAL cWhere
		cWhere = THIS.primaryKey + " = ?poData." + THIS.primaryKey
		THIS._sqlUpdateCmd = "UPDATE " + THIS.remoteTable + " SET " + cSetList + " WHERE " + cWhere
		RETURN THIS._sqlUpdateCmd		
		
		
		
	* EJECUTAMOS UN COMANDO SOBRE LA TABLA REMOTA
	PROCEDURE _Execute(pnConn, pcCmd, pcCursor)
		THIS.lastError = ""
		LOCAL nResult
		THIS.lastSqlCmd = pcCmd
		IF VARTYPE(pcCursor) = "C"
			nResult = SQLEXEC(pnConn, pcCmd, pcCursor)
		ELSE
			nResult = SQLEXEC(pnConn, pcCmd)
		ENDIF
		IF nResult <= 0
			THIS.lastError = THIS._getLastError()
		ENDIF
		RETURN (nResult > 0)
		
		
	
	* OBTENEMOS EL ULTIMO ERROR OCURRRIDO
	HIDDEN PROCEDURE _getLastError
		LOCAL ARRAY aERRInfo[1]
		AERROR(aERRInfo)
		RETURN aERRInfo[1,2]
	
	
	* ABRIR LA TABLA LOCAL (SI NO ESTA ABIERTA YA)
	HIDDEN PROCEDURE _openLocal
	    THIS.lastError = ""
		IF EMPTY(THIS.localTable) 
			THIS.lastError = "No se ha indicado la tabla local"
			RETURN ""
		ENDIF
		LOCAL cAlias
		cAlias = THIS.localTable + "_sqlproxy"
		IF !USED(cAlias)
			LOCAL cTabla
			cTabla = IIF(!EMPTY(THIS.dbcPath),THIS.dbcPath + "!","") + THIS.localTable
			TRY
				SELECT 0
				USE (cTabla) ALIAS (cAlias) AGAIN 
				
			CATCH TO ex
				THIS.lastError = ex.Message
				cAlias = ""
			ENDTRY
		ENDIF
		RETURN cAlias
		
		
		
ENDDEFINE






