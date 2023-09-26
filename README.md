# sqlproxy
Proxy class for a VFP local table with a remote copy

AUTOR: V. ESPINA
VERSION: 1.0

EXAMPLES:

A) CREATE AN INSTANCE
    oProxy = CREATE("sqlProxy","c:\data\company.dbc!customers")
    oProxy.remoteConnStr = "..."
    ?oProxy.localTable -> "customers"
    ?oProxy.remoteTable -> "customers"
    ?oProxy.dbcPath -> "c:\data\company.dbc"

    oProxy = CREATE("sqlProxy","customers","erp_customers","conn-str")
    oProxy.remoteConnStr = "conn-str"
    ?oProxy.localTable -> "customers"
    ?oProxy.remoteTable -> "erp_customers"
    ?oProxy.dbcPath -> ""

    nConn = SQLSTRINGCONNECT("...")
    oProxy = CREATE("sqlProxy","customers",,nConn)
    oProxy.remoteConnStr = ""
    ?oProxy.localTable -> "customers"
    ?oProxy.remoteTable -> "customers"
    ?oProxy.dbcPath -> ""

B) ADD CURRENT ROW ON REMOTE TABLE
    SELECT customers
    IF NOT oProxy.insertOne()
        MESSAGEBOX(oProxy.lastError)
    ENDIF

C) ADD CUSTOM ROW ON REMOTE TABLE
    SELECT customers
    SCATTER NAME oRow BLANK
    oRow.<column> = <value>
    oRow.<column> = <value>
    ...
    IF NOT oProxy.insertOne(oRow)
       MESSAGEBOX(oProxy.lastError)
    ENDIF

D) ADD MANY ROWS ON REMOTE TABLE USING A DATA CURSOR
    SELECT FROM customers WHERE status = 'ACTIVE' INTO CURSOR qactives
    IF NOT oProxy.insertMany("qactives")
       MESSAGEBOX(oProxy.lastError)
    ENDIF

E) ADD MANY ROWS ON REMOTE TABLE USING A COLLECTION
    SELECT customers
    oRows = CREATE("Collection")
    SCAN FOR status = "ACTIVE"
      SCATTER NAME oRow
      oRows.Add(oRow)
    ENDSCAN
    IF NOT oProxy.insertMany(oRows)
       MESSAGEBOX(oProxy.lastError)
    ENDIF

F) UPDATE ONE ROW FROM LOCAL TABLE
    SELECT customers
    LOCATE FOR id = "001"
    REPLACE status WITH "INACTIVE"
    IF NOT oProxy.updateOne()
       MESSAGEBOX(oProxy.lastError)
    ENDIF

G) UPDATE ONE ROW MANUALLY
    SELECT customers
    LOCATE FOR id = "001"
    SCATTER NAME oRow
    oRow.status = "INACTIVE"
    IF NOT oProxy.updateOne(oRow)
       MESSAGEBOX(oProxy.lastError)
    ENDIF

H) UPDATE MANY ROWS
    oSets = oProxy.Where("status,lastupd","INACTIVE",DATETIME())
    oWhere = oProxy.Where("zone","001")
    IF NOT oProxy.updateMany(oSets, oWhere)
       MESSAGEBOX(oProxy.lastError)
    ENDIF

I) DELETE ONE ROW FROM LOCAL TABLE
    SELECT customers
    LOCATE FOR id = "001"
    IF NOT oProxy.deleteOne()
       MESSAGEBOX(oProxy.lastError)
    ENDIF

J) UPDATE ONE ROW MANUALLY
    IF NOT oProxy.updateOne("001")
       MESSAGEBOX(oProxy.lastError)
    ENDIF

K) DELETE MANY ROWS
    oWhere = oProxy.Where("zone","001")
    IF NOT oProxy.deleteMany(oWhere)
       MESSAGEBOX(oProxy.lastError)
    ENDIF

L) GET DATA FROM ONE RECORD
    oRow = oProxy.findOne("001")
    IF ISNULL(oRow)
       MESSAGEBOX("Not found!")
    ELSE
       MESSAGEBOX("Name: " + oRow.fullName)
    ENDIF

M) FIND MANY ROWS
    oWhere = oProxy.Where("zone,status")
    oWhere.Zone = "001"
    oWhere.Status = "ACTIVE"
    IF oProxy.findMany(oWhere, "qresult")
       SELECT qresult
       BROWSE
    ENDIF
