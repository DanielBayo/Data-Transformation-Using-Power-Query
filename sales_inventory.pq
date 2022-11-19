let
  // Loaded #sales query from the Excel file
  sales = Table.TransformColumnTypes(
    Table.SelectColumns(
      Excel.CurrentWorkbook(){[Name = "sales"]}[Content], //To load sales table from the excel file
      {"datePlaced", "Sellerid"}
    ), // Select #datePlaced and #Sellerid columns from the query
    {{"datePlaced", type date}, {"Sellerid", type text}}
  ), // change the data types of the columns

  // Loaded #stock_count query from the Excel file
  stock_count = Table.TransformColumnTypes(
    Table.AddColumn(
      Excel.CurrentWorkbook(){[Name = "stock_count"]}[Content],//To load stock_count table from the excel file
      "Sellerid",
      each "XLR123"
    ), // Add a column #Sellerid to the query
    {
      {"date", type date},
      {"company_name", type text},
      {"country", type text},
      {"Sellerid", type text}
    }
  ), //change the data types of the columns

  // Join the #sales query with the #stock_count query using a left outer join
  Source = Table.NestedJoin(
    sales,
    {"Sellerid", "datePlaced"},
    stock_count,
    {"Sellerid", "date"},
    "stock_count",
    JoinKind.LeftOuter
  ),
  #"Expanded stock_count" = Table.ExpandTableColumn(Source, "stock_count", {"date"}, {"date"}), //Expand the stock_count query joined and select only the date
  #"Renamed Columns" = Table.RenameColumns(#"Expanded stock_count", {{"date", "Stock_countdate"}}),//Rename the date column from the expanded query to Stock_countdate
  #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([Stock_countdate] = null)),//Select null Rows on the Stock_countdate column as this are the dayes stocks were not taken
  #"Grouped Rows" = Table.Group(
    #"Filtered Rows",
    {"datePlaced", "Sellerid"},
    {{"Stock_Count?", each "No", Text.Type}}
  ),//Group Selected rows by the sellerid and the date.
  #"Rename datePlaced to date" = Table.RenameColumns(#"Grouped Rows", {{"datePlaced", "date"}}),//Rename the datePlaced column to date
  #"Combine stock_count and no stockcount together" = Table.Combine(
    {
       #"Rename datePlaced to date",
      Table.TransformColumnTypes(
        Table.AddColumn(
          Table.SelectColumns(
            Table.TransformColumnTypes(
              Table.AddColumn(
                Excel.CurrentWorkbook(){[Name = "stock_count"]}[Content],
                "Sellerid",
                each "XLR123"
              ),
              {
                {"date", type date},
                {"company_name", type text},
                {"country", type text},
                {"Sellerid", type text}
              }
            ),
            {"date", "Sellerid"}
          ),
          "Stock_Count?",
          each "Yes"
        ),
        {{"Stock_Count?", type text}}
      )
    }
  )
in
    #"Combine stock_count and no stockcount together"
