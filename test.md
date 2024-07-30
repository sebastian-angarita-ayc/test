To implement incremental refresh with hourly iterations using the provided code, we'll need to adjust the plan to include the hourly data retrieval while maintaining the incremental refresh capabilities. Here is a step-by-step plan:

### Step-by-Step Plan

1. **Set Up Parameters**:
   - Define parameters for `RangeStart` and `RangeEnd` in Power BI.

2. **Create a Function to Fetch Data by Hour**:
   - Modify the existing function to retrieve data for each hour within a given date-time range using the provided code structure.

3. **Generate Hourly Date-Time Intervals**:
   - Create a list of hourly intervals within the refresh period.

4. **Iterate Over Hourly Intervals and Fetch Data**:
   - Loop through the hourly intervals to make API calls and combine the results.

5. **Configure Incremental Refresh**:
   - Use Power BI’s incremental refresh features to specify the refresh period and the historical data period.

### Implementation Details

#### Step 1: Set Up Parameters

Define parameters for start and end date-time, and the refresh period.

1. Go to the `Manage Parameters` section in Power BI.
2. Create parameters:
   - `RangeStart`: DateTime type, default value to some past date-time.
   - `RangeEnd`: DateTime type, default value to the current date-time.

#### Step 2: Create a Function to Fetch Data by Hour

Define a function in Power Query to make an API call for a specific hourly date-time range using the provided code structure.

```m
let
    GetSendGridDataByHour = (start_date as datetime, end_date as datetime) as table =>
    let
        startDateNati = DateTime.ToText(start_date, [Format="yyyy-MM-dd'T'HH:mm:ss'Z'"]),
        endDateNati = DateTime.ToText(end_date, [Format="yyyy-MM-dd'T'HH:mm:ss'Z'"]),
        url = "https://api.sendgrid.com/v3/messages?limit=1000&query=last_event_time%20BETWEEN%20TIMESTAMP%20%22" & startDateNati & "%22%20AND%20TIMESTAMP%20%22" & endDateNati & "%22",
        headers = [Authorization="Bearer SG.Natacha"],
        Source = Json.Document(Web.Contents(url, [Headers=headers])),
        messages = Source[messages],
        #"Converted to Table" = Table.FromList(messages, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"from_email", "msg_id", "subject", "to_email", "status", "opens_count", "clicks_count", "last_event_time"}, {"from_email", "msg_id", "subject", "to_email", "status", "opens_count", "clicks_count", "last_event_time"}),
        #"Duplicated Column" = Table.DuplicateColumn(#"Expanded Column1", "last_event_time", "last_event_time - Copy"),
        #"Split Column by Position" = Table.SplitColumn(#"Duplicated Column", "last_event_time - Copy", Splitter.SplitTextByPositions({0, 10}, false), {"last_event_time - Copy.1", "last_event_time - Copy.2"}),
        #"Changed Type" = Table.TransformColumnTypes(#"Split Column by Position",{{"last_event_time - Copy.1", type date}, {"last_event_time - Copy.2", type text}}),
        #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"last_event_time - Copy.2"})
    in
        #"Removed Columns"
in
    GetSendGridDataByHour
```

#### Step 3: Generate Hourly Date-Time Intervals

Create a function to generate a list of hourly intervals within the given date-time range.

```m
let
    GenerateHourlyIntervals = (start_datetime as datetime, end_datetime as datetime) as list =>
    let
        hourList = List.Generate(
            () => start_datetime,
            each _ <= end_datetime,
            each DateTime.AddHours(_, 1)
        )
    in
        hourList
in
    GenerateHourlyIntervals
```

#### Step 4: Iterate Over Hourly Intervals and Fetch Data

Create a query that uses the hourly intervals to fetch data iteratively and combine the results.

```m
let
    // Parameters
    startDateTime = RangeStart,
    endDateTime = RangeEnd,
    hourlyIntervals = GenerateHourlyIntervals(startDateTime, endDateTime),

    // Retrieve data for each hourly interval and combine
    allData = List.Combine(
        List.Transform(
            hourlyIntervals,
            each let
                currentStartDateTime = _,
                currentEndDateTime = DateTime.AddHours(_, 1),
                data = GetSendGridDataByHour(currentStartDateTime, currentEndDateTime)
            in
                data
        )
    ),

    // Convert to table
    finalTable = Table.Combine(allData)
in
    finalTable
```

#### Step 5: Configure Incremental Refresh

1. **Import Data Using Parameters**:
   - In Power Query Editor, filter your data source using `RangeStart` and `RangeEnd`.

```m
let
    Source = GenerateHourlyIntervals(RangeStart, RangeEnd),
    FilteredRows = Table.SelectRows(Source, each [last_event_time] >= RangeStart and [last_event_time] < RangeEnd)
in
    FilteredRows
```

2. **Enable Incremental Refresh**:
   - In Power BI Desktop, right-click your table in the Fields pane, and select `Incremental refresh`.

3. **Set Incremental Refresh Policy**:
   - Define the period for which historical data is kept.
   - Specify the period for which incremental data refreshes.

For example:
   - `Store rows in the last`: 5 years.
   - `Refresh rows in the last`: 1 day.

### Summary

- **Set Parameters**: Define `RangeStart` and `RangeEnd` parameters.
- **Fetch Data by Hour**: Create a function to fetch data for each hourly interval using the provided structure.
- **Generate Hourly Intervals**: Create a function to generate hourly intervals within the date-time range.
- **Iterate and Fetch**: Loop through hourly intervals, fetch data, and combine results.
- **Configure Incremental Refresh**: Set up and enable incremental refresh in Power BI.

This approach ensures efficient data retrieval by handling large datasets through hourly iterations and maintaining incremental refresh for optimal performance.
