

function readDataToSheet() {
  /* 
  Connect to RDS and query data
  */

  // connection parameters
  var strConnectionName = 'connection name';
  var strUsername = 'username';
  var strPassword = 'password';
  var strDb = 'database';
  // a MySQL connection url must be of the form "jdbc:msql://[subprotocol]/[db name]"
  var strUrl = 'jdbc:mysql://url/database';
  
  // read query to result set
  var conn = Jdbc.getConnection(strUrl, strUsername, strPassword);
  var strQuery = "SELECT id \
	,external_item_id \
	,sku \
	,external_order_id \
	,marketplace \
	,account_id \
    ,DATE(CONVERT_TZ(created_at, 'UTC', 'America/New_York')) AS date_created_est \
	,CONVERT_TZ(created_at, 'UTC', 'America/New_York') as created_at_EST \
	,CONVERT_TZ(updated_at, 'UTC', 'America/New_York') AS updated_at_EST \
	,failure_reason \
	,description \
	,fixed \
    FROM unfulfillable_marketplace_orders \
    WHERE created_at >= CURRENT_DATE - INTERVAL (19 + DAYOFWEEK(CURRENT_DATE)) DAY \
    AND sku NOT LIKE '%FBA%' \
    ORDER BY created_at"
  var connStmt = conn.createStatement();
  var connResults = connStmt.executeQuery(strQuery)
  
  // set range variables
  var shtData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var cell = shtData.getRange("A1");
  shtData.clear(); // clear range
  
  // print result set to sheet
  var numCols = connResults.getMetaData().getColumnCount();
  var countRow = 0;

  for (var col = 1; col <= numCols; col++){
    cell.offset(0, col - 1).setValue(connResults.getMetaData().getColumnName(col)); // add column names
  }
  while (connResults.next()) {
    for (var col = 0; col < numCols; col++){
      cell.offset(countRow+1, col).setValue(connResults.getString(col+1)); // offset row by 1 for column names
    }
    countRow++;
  }
  connResults.close();
  connStmt.close();
  
  // rebuild pivot table
  buildPivotTable();
} 


function buildPivotTable(){
  /*
	This function rebuilds the pivot table after the data source is rebuilt.
	In order to update a pivot table through GAS, you have to rebuild the 
	table. Note: this function saves date filters before rebuilding the table.
  */

  var arrDateFilters = getDateFilters(); // save date filters already set

  // set data and pivot ranges
  var sht = SpreadsheetApp.getActiveSpreadsheet();
  var shtPvt = sht.getSheetByName("Pivot Table 1");
  var rngPvt = shtPvt.getRange("A1");
  var shtData = sht.getSheetByName("Data");
  var rngData = shtData.getRange("A1").getDataRegion();

  // get column indices
  var columnNames = shtData.getDataRange().getValues()[0]
  var colId = columnNames.indexOf("id")+1;
  var colDate = columnNames.indexOf("date_created_est")+1;
  var colAccountId = columnNames.indexOf("account_id")+1;
  var colMarketplace = columnNames.indexOf("marketplace")+1;
  var colSku = columnNames.indexOf("sku")+1;
  var colExternalOrderId = columnNames.indexOf("external_order_id")+1;
  var colFixed = columnNames.indexOf("fixed")+1;

  // create pivot table
  var pvtTable = rngPvt.createPivotTable(rngData);
  // set aggregation
  var pvtValue = pvtTable.addPivotValue(colId, 
  	SpreadsheetApp.PivotTableSummarizeFunction.COUNT); 
  // set grouping
  var pvtGroup = pvtTable.addRowGroup(colDate); 
  pivotGroup = pvtTable.addRowGroup(colAccountId);
  pivotGroup = pvtTable.addRowGroup(colMarketplace);
  pivotGroup.showTotals(false);
  pivotGroup = pvtTable.addRowGroup(colSku);
  pivotGroup.showTotals(false);
  pivotGroup = pvtTable.addRowGroup(colExternalOrderId);
  pivotGroup.showTotals(false);
  // set filters
  var pvtFilters = SpreadsheetApp.newFilterCriteria() // date filter
  .setVisibleValues(arrDateFilters)
  .build();
  pvtTable.addFilter(colDate, pvtFilters); 
  pvtFilters = SpreadsheetApp.newFilterCriteria() // fixed = {0,1} filter
  .setVisibleValues(['0'])
  .build();
  pvtTable.addFilter(colFixed, pvtFilters);
  pvtFilters = SpreadsheetApp.newFilterCriteria() // account id filter
  .setVisibleValues(['1', '2','3','4','5','6','7','8','9','10',
  	'11','12','13','14','15','16','17','18','19','20'])
  .build();
  pvtTable.addFilter(colAccountId, pvtFilters); 
}

function getDateFilters(){
  /*
  Get potential filter dates from the pivot table and paste to a dummy sheet
	*/
  var shtDateFilters = SpreadsheetApp.getActiveSpreadsheet()
	.getSheetByName("filters");
	shtDateFilters.getRange("A1").setFormula("UNIQUE('Pivot Table 1'!A:A)")
	var rngDateValues = shtDateFilters.getRange("A1:A").getValues();

	// find text of the form yyyy-mm-dd
	var re = new RegExp("[0-9]{4}-[0-9]{2}-[0-9]{2}");

	// initialize array
	var arrDateFilters = [];

	// loop through date column of the pivot table to extract dates
	for (var i = 0; i < rngDateValues.length; i++){
		var match = re.exec(rngDateValues[i]);
		if (match != null){
			arrDateFilters.push(String(match));
		}
	}
	return arrDateFilters;
}



function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu('Update Sheet')
  .addItem('Refresh Data', 'readDataToSheet')
  .addToUi();
}
