//Return the last row id of the list
//I use this value as a requirement for other functions
//Reusable with other lists
function sp_getLastRowID(str_theListName) {
    return new Promise(function (resolve, reject) {
        try {

            //internal
            let resultListItems;

            //output
            let int_id;

            //specific query to get the last row ID
            async function getLastRowID() {
                let clientContext = SP.ClientContext.get_current();
                let sourceList = clientContext.get_web().get_lists().getByTitle(str_theListName);
                let camlQuery = new SP.CamlQuery();
                let str_queryCamlXML = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy></Query><RowLimit>1</RowLimit></View>";
                camlQuery.set_viewXml(str_queryCamlXML);
                resultListItems = sourceList.getItems(camlQuery);
                clientContext.load(resultListItems);
                clientContext.executeQueryAsync(
                    Function.createDelegate(this, onSuccess),
                    Function.createDelegate(this, onFail)
                )
            }

            function onSuccess() {
                var listEnumerator = resultListItems.getEnumerator();
                while (listEnumerator.moveNext()) {
                    let item = listEnumerator.get_current();
                    int_id = item.get_item('ID');
                }
                resolve(int_id);
            }

            function onFail(sender, args) {
                console.log(args);
                reject(args);
            }

            getLastRowID();

        } catch (error) {
            console.log(error);
            reject(error);
        }
    })
}


//Search a list and return an array of ID's that match your criteria
//Part of the function is a CAML query builder
//Reusable with other lists
function sp_listSearch(
    str_theListName, //The name of the sharepoint list
    arr_theKeywords, //An array containing keywords. it is assumed that this array will contain lowercase characters only
    arr_theTextColumnsToMatchKeywords, //Columns of the list where the keywords will be searched
    arr_theCustomCamlFilters, //Expects a well formatted caml e.g <Eq><FieldRef Name='ID'/><Value Type='Number'>123</Value></Eq>. The elements will be added as nested AND conditions to the other filters.
    bool_isAscending, //Return results in ascending ID order (equivalent to ascending creation date order).
    bool_isRowLimitUsed, //If row limit is used, the list will stop if the either the row limit for results is reached or if the end of the list is reached.
    int_theRowLimit, //This defines how many rows this search will return. This function is designed not to break at >5000 results. Further details below.
    int_theLastRowID //The last row ID tells this function that the last row of the list is reached(in ascending mode), or it tells the list which row to start from (descending mode)
) {
    return new Promise(function (resolve, reject) {
        try {

            //Internal Variables
            //The min ID and max ID limits the number of records to be searched each time due to the 5000-row limit for query results.
            let int_variableMinID;
            let int_variableMaxID;
            let int_defaultSpan = 5000; //By limiting the searched rows, we guarantee that the query won't fail due to the 5000-row threshold limit.
            let int_variableSpan = int_defaultSpan; //This variable will be increased in order to increase the number of searched rows. This is to try and lessen the recursive search and speed up the function

            //Output Variables
            let arr_resultIDs = [];//This will contain the output record IDs that matches the search criteria

            //Recursive search
            //The sharepoint CAML query will fail if the result set goes beyond 5000 rows.
            //This function will run repeatedly, searching a small part of the list each time until the conditions are satisfied.
            //If it succeeds, but is allowed to search another segment of the list with a broader range.
            //If it fails due to threshold limit, then it will re run at the last starting point but reduce its range to 5000, to guarantee success.
            function recursiveSearch(int_theMinID, int_theMaxID) {

                //Sharepoint query initialization
                let clientContext = SP.ClientContext.get_current(); //This gets the client context from the sharepoint site where this script is running.
                let sourceList = clientContext.get_web().get_lists().getByTitle(str_theListName); //The actual list is found by calling it by name from the client context.
                let camlQuery = new SP.CamlQuery(); //This initializes the CAML query.

                //Custom filter
                let str_customFilters = ""; //This will contain a hard coded or generated caml query string via below tasks
                let arr_andConditionFilters = []; //This will contain all conditions. By looping through this array, we can create a nested AND condition that is well-formatted for the query.

                //Min max ID custom filter
                //All list rows have an ID. Since they are incremental and unique, then they can be used sort the list and search it part by part.
                let str_minIDCaml = "<Gt><FieldRef Name='ID'/><Value Type='Number'>" + int_theMinID + "</Value></Gt>"; //Construct the min ID filter
                arr_andConditionFilters.push(str_minIDCaml); //Add to Array
                let str_maxIDCaml = "<Leq><FieldRef Name='ID'/><Value Type='Number'>" + int_theMaxID + "</Value></Leq>"; //Construct the max ID filter
                arr_andConditionFilters.push(str_maxIDCaml); //Add to Array

                //Keyword filter. Common among lists
                //If keywords and the columns to match keywords are present, then we can create a filter for each of them.
                if (arr_theKeywords.length > 0 && arr_theTextColumnsToMatchKeywords.length > 0) {

                    let str_andConditionsForKeywords = ""; //This will be a mixture of ANDs and ORs depending on the number of columns and keywords to search.
                    //Each keyword must be found (AND)
                    //Each keyword must be found in at least one column (OR)

                    //Loop through the keywords
                    for (let i = 0; i < arr_theKeywords.length; i++) {

                        let str_orConditionsForColumns = ""; //This will contain the possible OR type matches, e.g. if keywordX is in columnX OR columnY

                        //Loop through the columns
                        for (let j = 0; j < arr_theTextColumnsToMatchKeywords.length; j++) {

                            //If there is only one column...
                            if (arr_theTextColumnsToMatchKeywords.length == 1) {
                                str_orConditionsForColumns = "<Contains><FieldRef Name='" + arr_theTextColumnsToMatchKeywords[j] + "'/><Value Type='Text'>" + arr_theKeywords[i] + "</Value></Contains>";
                            } else {
                                //Wrap the first two conditions in one OR
                                if (j == 0) {
                                    str_orConditionsForColumns += "<Or>";
                                    str_orConditionsForColumns += "<Contains><FieldRef Name='" + arr_theTextColumnsToMatchKeywords[j] + "'/><Value Type='Text'>" + arr_theKeywords[i] + "</Value></Contains>";
                                } else if (j == 1) {
                                    str_orConditionsForColumns += "<Contains><FieldRef Name='" + arr_theTextColumnsToMatchKeywords[j] + "'/><Value Type='Text'>" + arr_theKeywords[i] + "</Value></Contains>";
                                    str_orConditionsForColumns += "</Or>";
                                    //Then nest the others
                                } else {
                                    str_orConditionsForColumns = "<Or>" + str_orConditionsForColumns + "<Contains><FieldRef Name='" + arr_theTextColumnsToMatchKeywords[j] + "'/><Value Type='Text'>" + arr_theKeywords[i] + "</Value></Contains></Or>";
                                }
                            }
                        }
                        console.log(str_orConditionsForColumns);

                        //If there is only one keyword
                        if (arr_theKeywords.length == 1) {
                            str_andConditionsForKeywords = str_orConditionsForColumns;

                            //if there are many keywords
                        } else {

                            //Enclose the first Two keywords in an AND filter
                            if (i == 0) {
                                str_andConditionsForKeywords += "<And>";
                                str_andConditionsForKeywords += str_orConditionsForColumns;
                            } else if (i == 1) {
                                str_andConditionsForKeywords += str_orConditionsForColumns;
                                str_andConditionsForKeywords += "</And>";
                                //Then nest the others
                            } else {
                                str_andConditionsForKeywords = "<And>" + str_andConditionsForKeywords + str_orConditionsForColumns + "</And>";
                            }
                        }

                    }
                    console.log(str_andConditionsForKeywords);

                    arr_andConditionFilters.push(str_andConditionsForKeywords); //Add to array
                }

                //Add externally contsructed CAML filters to the filter array
                if (arr_theCustomCamlFilters.length > 0) {
                    arr_andConditionFilters = arr_andConditionFilters.concat(arr_theCustomCamlFilters);
                }

                //Convert filter array into nested AND conditions for the query
                for (let i = 0; i < arr_andConditionFilters.length; i++) {

                    //if there's only one filter, no AND wrapper is needed
                    if (arr_andConditionFilters.length == 1) {
                        str_customFilters += arr_andConditionFilters[i];
                    } else {
                        //if there's more than one filter
                        //stack the first two filters in one AND wrapper...
                        if (i == 0) {
                            str_customFilters += "<And>";
                            str_customFilters += arr_andConditionFilters[i];
                        } else if (i == 1) {
                            str_customFilters += arr_andConditionFilters[i];
                            str_customFilters += "</And>";
                        } else {
                            //then nest the others
                            str_customFilters = "<And>" + str_customFilters + arr_andConditionFilters[i] + "</And>";
                        }
                    }
                }

                //ascending or not
                let str_isSearchAscending = "False";
                if (bool_isAscending) {
                    str_isSearchAscending = "True";
                }

                //caml query builder
                let str_query = "";
                str_query += "<View>";
                str_query += "<Query>";
                str_query += "<Where>";
                str_query += str_customFilters;
                str_query += "</Where>";
                str_query += "<OrderBy>";
                str_query += "<FieldRef Name='ID' Ascending='" + str_isSearchAscending + "' />";
                str_query += "</OrderBy>";
                str_query += "</Query>";

                //Apply row limit if used
                if (bool_isRowLimitUsed) {
                    str_query += "<RowLimit>" + int_theRowLimit;
                    str_query += "</RowLimit>";
                }

                str_query += "</View>";

                //Print query (for debugging only)
                console.log(str_query);

                //Execute caml query
                camlQuery.set_viewXml(str_query);
                let resultListItems = sourceList.getItems(camlQuery);
                clientContext.load(resultListItems, 'Include(ID)'); //Reduce the size of the result set by including the ID column only
                clientContext.executeQueryAsync(
                    Function.createDelegate(
                        this,
                        function onSuccess() {
                            //If the query ends without error then...
                            //append results to array
                            let listEnumerator = resultListItems.getEnumerator();
                            while (listEnumerator.moveNext()) {

                                //Get the ID from the resultListItems and add it to the result set
                                let item = listEnumerator.get_current();
                                let int_id = item.get_item('ID');
                                arr_resultIDs.push(int_id);
                            }

                            //Check if the last row is reached
                            let bool_isLastRowReached = false;
                            if (bool_isAscending) {
                                if (int_theMaxID >= int_theLastRowID) {
                                    bool_isLastRowReached = true;
                                }
                            } else {
                                if (int_theMinID <= 0) {
                                    bool_isLastRowReached = true;
                                }
                            }

                            //Check if the row limit is reached
                            let bool_isRowLimitReached = true;
                            if (bool_isRowLimitUsed) {
                                bool_isRowLimitReached = false;
                                if (arr_resultIDs.length >= int_theRowLimit) {
                                    bool_isRowLimitReached = true;
                                }
                            }

                            //Resolve if last row is reached or if the row limit is reached
                            if (bool_isRowLimitReached || bool_isLastRowReached) {
                                arr_resultIDs.slice(int_theRowLimit); //In case the final result set exceeds the page limit, trim it.
                                resolve(arr_resultIDs);
                            } else {
                                //Try to search the next part of the list with bigger span (search more rows at once). It could fail, but we are trying to reduce the number of times the search has to run in order to search the whole list if it has to.
                                int_variableSpan = int_variableSpan * 2;
                                if (bool_isAscending) {
                                    int_variableMinID = int_variableMaxID;
                                    int_variableMaxID = int_variableMaxID + int_variableSpan;

                                } else {
                                    int_variableMaxID = int_variableMinID;
                                    int_variableMinID = int_variableMinID - int_variableSpan;
                                }
                                recursiveSearch(int_variableMinID, int_variableMaxID);
                            }
                        }
                    ),
                    Function.createDelegate(
                        this,
                        function onFail(sender, args) {

                            //Check if the stackTrace can be taken. It's not always possible.
                            let str_error = args.get_stackTrace();
                            if (str_error != null) {
                                str_error = str_error.toString().toLowerCase();

                                //Check if the error is related to threshold (>5k results) overshoot. Else reject.
                                if (str_error.includes("exceeds the list view threshold")) {

                                    //Search using the same starting point (minID for ascending or MaxID for descending searches) but reset the span to 5000 to ensure that this search attempt doesn't fail.
                                    int_variableSpan = int_defaultSpan;
                                    if (bool_isAscending) {
                                        int_variableMinID = int_variableMinID;
                                        int_variableMaxID = int_variableMinID + int_variableSpan;
                                    } else {
                                        int_variableMaxID = int_variableMaxID;
                                        int_variableMinID = int_variableMaxID - int_variableSpan;
                                    }
                                    recursiveSearch(int_variableMinID, int_variableMaxID);
                                } else {
                                    console.log(str_error);
                                    reject(str_error);
                                }
                            } else {
                                console.log(args);
                                reject(args);
                            }
                        }
                    )
                )
            }

            //First search
            if (bool_isAscending) {
                int_variableMinID = 0;
                int_variableMaxID = int_variableSpan;
            } else {
                int_variableMaxID = int_theLastRowID;
                int_variableMinID = int_theLastRowID - int_variableSpan;
            }
            recursiveSearch(int_variableMinID, int_variableMaxID);

        } catch (error) {
            console.log(error);
            reject(error);
        }
    })
}

//Get list items filtered by their record IDs//
//Reusable with other lists
function sp_getRecordsFromList(str_theListName, arr_theIDs, bool_isAscending) {
    return new Promise(function (resolve, reject) {
        try {

            //Output variables
            let arr_resultRecords = []; //This will contain whole records

            //Recursive search.
            function recursiveSearch() {

                //Sharepoint query initialization
                let clientContext = SP.ClientContext.get_current(); //This gets the client context from the sharepoint site where this script is running.
                let sourceList = clientContext.get_web().get_lists().getByTitle(str_theListName); //The actual list is found by calling it by name from the client context.
                let camlQuery = new SP.CamlQuery(); //This initializes the CAML query.

                //Custom filter
                let str_customFilters = ""; //This will contain a hard coded or generated caml query string via below tasks

                //Convert id array into nested OR conditions for the query
                for (let i = 0; i < arr_theIDs.length; i++) {

                    //If there's only one ID, no OR wrapper is needed
                    if (arr_theIDs.length == 1) {
                        str_customFilters += "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>";
                    } else {
                        //if there's more than one ID
                        //stack the first two ID in one OR wrapper...
                        if (i == 0) {
                            str_customFilters += "<Or>";
                            str_customFilters += "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>";
                        } else if (i == 1) {
                            str_customFilters += "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>";
                            str_customFilters += "</Or>";
                        } else {
                            //Then nest the others
                            str_customFilters = "<Or>" + str_customFilters + "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>" + "</Or>";
                        }
                    }
                }

                //ascending or not
                let str_isSearchAscending = "False";
                if (bool_isAscending) {
                    str_isSearchAscending = "True";
                }

                //caml query builder
                let str_query = "";
                str_query += "<View>";
                str_query += "<Query>";
                str_query += "<Where>";
                str_query += str_customFilters;
                str_query += "</Where>";
                str_query += "<OrderBy>";
                str_query += "<FieldRef Name='ID' Ascending='" + str_isSearchAscending + "' />";
                str_query += "</OrderBy>";
                str_query += "</Query>";
                str_query += "</View>";

                //Print the query for debugging.
                console.log(str_query);

                //Execute caml query
                camlQuery.set_viewXml(str_query);
                let resultListItems = sourceList.getItems(camlQuery);
                clientContext.load(resultListItems);
                clientContext.executeQueryAsync(
                    Function.createDelegate(
                        this,
                        function onSuccess() {

                            //Append results to array
                            let listEnumerator = resultListItems.getEnumerator();
                            while (listEnumerator.moveNext()) {
                                
                                //Get the list item from the list enumerator
                                let item = listEnumerator.get_current();

                                //Get only the record data from each item
                                let obj_item = item.get_fieldValues();
                                arr_resultRecords.push(obj_item);
                            }
                            resolve(arr_resultRecords);
                        }
                    ),
                    Function.createDelegate(
                        this,
                        function onFail(sender, args) {
                            console.log(args);
                            //check if the error is related to threshold (>5k results) overshoot
                            let str_error = args.get_stackTrace();
                            str_error = str_error.toString().toLowerCase();
                            if (str_error.includes("exceeds the list view threshold")) {
                                //reset span and search again
                                int_variableSpan = int_span;
                                if (bool_isAscending) {
                                    int_minID = int_minID;
                                    int_maxID = int_minID + int_variableSpan;
                                    recursiveSearch(int_minID, int_maxID);
                                } else {
                                    int_maxID = int_maxID;
                                    int_minID = int_maxID - int_variableSpan;
                                    recursiveSearch(int_minID, int_maxID);
                                }
                            } else {
                                console.log(str_error);
                                reject(str_error);
                            }
                        })
                )
            }
            recursiveSearch();
        } catch (error) {
            console.log(error);
            reject(error);
        }
    })
}

//Get list items filtered by their record IDs
//Reusable with other lists
function sp_getRecordsFromList_bak(str_theListName, arr_theIDs, bool_isAscending) {
    return new Promise(function (resolve, reject) {
        try {

            let arr_resultRecords = [];
            let resultListItems;

            //Recursive search.
            function recursiveSearch() {

                //Sharepoint query initialization.
                let clientContext = SP.ClientContext.get_current();
                let sourceList = clientContext.get_web().get_lists().getByTitle(str_theListName);
                let camlQuery = new SP.CamlQuery();
                let str_customFilters = "";

                //Convert filter array into nested OR conditions for the query
                for (let i = 0; i < arr_theIDs.length; i++) {

                    //If there's only one filter, no AND wrapper is needed
                    if (arr_theIDs.length == 1) {
                        str_customFilters += "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>";
                    } else {
                        //if there's more than one filter
                        //stack the first two filters in one AND wrapper...
                        if (i == 0) {
                            str_customFilters += "<Or>";
                            str_customFilters += "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>";
                        } else if (i == 1) {
                            str_customFilters += "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>";
                            str_customFilters += "</Or>";
                        } else {
                            //then nest the others
                            str_customFilters = "<Or>" + str_customFilters + "<Eq><FieldRef Name='ID' /><Value Type='Number'>" + arr_theIDs[i] + "</Value></Eq>" + "</Or>";
                        }
                    }
                }

                //ascending or not
                let str_isSearchAscending = "False";
                if (bool_isAscending) {
                    str_isSearchAscending = "True";
                }

                //caml query builder
                let str_query = "";
                str_query += "<View>";
                str_query += "   <Query>";
                str_query += "      <Where>";
                str_query += str_customFilters;
                str_query += "      </Where>";
                str_query += "      <OrderBy>";
                str_query += "         <FieldRef Name='ID' Ascending='" + str_isSearchAscending + "' />";
                str_query += "      </OrderBy>";
                str_query += "   </Query>";
                str_query += "</View>";

                //remove 3-space indents
                str_query = str_query.replace("   ", "");
                console.log(str_query);

                //execute caml query
                camlQuery.set_viewXml(str_query);
                resultListItems = sourceList.getItems(camlQuery);
                clientContext.load(resultListItems);
                clientContext.executeQueryAsync(
                    Function.createDelegate(
                        this,
                        function onSuccess() {

                            //append results to array
                            let listEnumerator = resultListItems.getEnumerator();
                            while (listEnumerator.moveNext()) {
                                //columns to retrieve
                                let item = listEnumerator.get_current();
                                let int_id = item.get_item('ID');
                                let str_title = item.get_item("tbl000_title");
                                let str_details = item.get_item("tbl000_details");
                                let date_creationDate = item.get_item("tbl000_creation_date");
                                let str_author = item.get_item("tbl000_author");
                                let str_requester = item.get_item("tbl000_requester");

                                arr_resultRecords.push({
                                    int_id,
                                    str_title,
                                    str_details,
                                    date_creationDate,
                                    str_author,
                                    str_requester
                                });
                            }

                            resolve(arr_resultRecords);
                        }
                    ),
                    Function.createDelegate(
                        this,
                        function onFail(sender, args) {
                            console.log(args);
                            //check if the error is related to threshold (>5k results) overshoot
                            let str_error = args.get_stackTrace();
                            str_error = str_error.toString().toLowerCase();
                            if (str_error.includes("exceeds the list view threshold")) {
                                //reset span and search again
                                int_variableSpan = int_span;
                                if (bool_isAscending) {
                                    int_minID = int_minID;
                                    int_maxID = int_minID + int_variableSpan;
                                    recursiveSearch(int_minID, int_maxID);
                                } else {
                                    int_maxID = int_maxID;
                                    int_minID = int_maxID - int_variableSpan;
                                    recursiveSearch(int_minID, int_maxID);
                                }
                            } else {
                                console.log(str_error);
                                reject(str_error);
                            }
                        })
                )
            }
            recursiveSearch();
        } catch (error) {
            console.log(error);
            reject(error);
        }
    })
}