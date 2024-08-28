# ------------------------------------------------------------------------------
#
# This script will extract an html table located within a rich text field that
# can then be used to select a specific row based on the table's "Artifact Id"
# column. When executed, the script will output the value of the row as CSV.
#
# ------------------------------------------------------------------------------
# Manual execution (Matter Object)
#
# ------------------------------------------------------------------------------
# In order for this script to work properly, there must be a rich text field 
# that contains a table with at least an "Artifact Id" header colum in the 
# first row.
# 
# Example:
# +-------------+---------------------+------------------+
# | Artifact Id | Workspace Name      | Workspace Status |
# +------------ +---------------------+------------------+
# | 100001      | Rice vs Corn        | Active           |
# +------------ +---------------------+------------------+
# | 100002      | Apples vs Oranges   | Active           |
# +------------ +---------------------+------------------+
# | 100003      | Salt vs Pepper      | Archived         |
# +------------ +---------------------+------------------+

# ------------------------------------------------------------------------------


# ------------------------------------------------------------------------------
# USER-DEFINED VARIABLES (Remove/comment these if using script parameters)

# The matter to retrieve
$matterId = 1

# The name of the rich text field on the matter that contains a table
$label = "Relativity Workspaces"

# The Relativity artifact id in the table to retrieve
$artifactId = "100001"


# ------------------------------------------------------------------------------
# FUNCTIONS

# Function to parse HTML content that contains a table from a rich text field
function ExtractTableFromHtml {
    param (
        [string]$htmlContent
    )
    
    if ([string]::IsNullOrWhiteSpace($htmlContent)) {
        throw "No content provided"
    }

    # Replace common HTML entities that are not XML-compliant
    # Note: This covers most basic html-encoded characters, but you may need to do 
    #       more depending on how complex your html is stored
    $htmlContent = $htmlContent -replace '&nbsp;', ' '
    $htmlContent = $htmlContent -replace '&(?!amp;|lt;|gt;|quot;|apos;)', '&amp;'

    # Convert and store the HTML content into an xml object
    $xmlContent = [xml]("<root>$htmlContent</root>")

    # Locate the first table amongst the HTML content
    $table = $xmlContent.root.table

    # If there is no table that exists in the content, throw an error
    if (-not $table) {
        throw "No table found in the provided HTML content."
    }

    # This parser is assuming that the provided html is generally well-formed 
    # where there is a <tbody> parent tag that contains <tr> table rows. We'll 
    # access the tbody section of the table here to locate the row content
    $tbody = $table.tbody

    # Throw an error if the tbody doesn't exist
    if (-not $tbody) {
        throw "No tbody found in the table."
    }

    # We're far enough ahead here to now initialize an array that will hold our 
    # table data
    $tableData = @()

    # The first row will contain the table headers for this parser
    $headerRow = $tbody.tr[0]

    if ($headerRow -eq $null) {
        throw "No header row found in the table."
    }

    # Extract headers into a variable. The headers are contains within cells
    $headers = $headerRow.SelectNodes("td") | ForEach-Object { $_.InnerText.Trim() }

    # Throw an error if we're unable to locate any headers
    if ($headers.Count -eq 0) {
        throw "No headers found in the table."
    }

    # Process each row in the table body after the header
    foreach ($row in $tbody.tr[1..$($tbody.tr.Count - 1)]) {
        $rowData = @{}
        $cells = $row.SelectNodes("td")

        # Move to the next row if no cells exist in the current row
        if (-not $cells) {
            continue
        }

        # Loop through each cell and map it to the corresponding header
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $header = $headers[$i]
            $rowData[$header] = $cells[$i].InnerText.Trim()
        }
        
        # Add the row data to the table array
        $tableData += $rowData
    }

    return $tableData
}

# Function to select data that was extracted by the ExtractTableFromHtml 
# function by Artifact Id
function SelectDataByArtifactId {
    param (
        [array]$tableData,
        [string]$artifactId
    )

    # Filter the table data for the given artifact id
    $filteredRow = $tableData | Where-Object {
        $_['Artifact Id'] -eq $artifactId
    } | Select-Object -First 1

    if ($filteredRow) {
        return $filteredRow
    } else {
        Write-Output "No matching row found for Artifact Id: $artifactId"
        
        return $null
    }
}



# ------------------------------------------------------------------------------
# SCRIPT

$matter = Get-Matter -Id $matterId
$richTextFieldWithTable = $matter.Fields | Where-Object { $_.Label -eq $label } | Select-Object -First 1

try {
    # Extract the html data into an array of hashtable objects
    $tableData = ExtractTableFromHtml `
        -htmlContent $richTextFieldWithTable.Value.ValueAsString
    
    # Select a row from the table based on the artifact id
    $selectedRow = SelectDataByArtifactId `
        -tableData $tableData `
        -artifactId $artifactId

    # Check if the selected row is not null before converting to CSV
    if ($selectedRow -ne $null) {
        # Output the selected row as CSV
        $selectedRow | ConvertTo-CSV -NoTypeInformation
    } else {
        Write-Output "No data found for the specified Artifact Id"
    }
}
catch {
    Write-Error "An error occurred: $_"
}
