function Get-LatestExchangeVersion {
    param(
        [ValidateSet($null,'2013','2016','2019')]
        $Version
    )
    # https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
    $response = Invoke-WebRequest -Uri 'https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates'
    if ($response.StatusCode -eq 200) {
        # grab the tables
        $tables = [regex]::Matches($response.Content,'(?<=Exchange Server \d{4}\.<\/p>\n)<table>.*?<\/table>','SingleLine')
        $versions = foreach($table in $tables.Value) {
            # grab the rows for each table
            $rows = [regex]::Matches($table,'<(tr)>.*?<\/\1>','SingleLine')
            # grab the data from the 2nd row containing the latest version data
            $data = [regex]::Matches($rows[1].Value, '<(th|td).*?>(?<data>.*?)<\/\1>','SingleLine')
            # groups 2 contain the data
            $hrefLink = $data[0].groups['data'].Value
            $date = $data[1].groups['data'].Value
            $shortVersion = $data[2].groups['data'].Value
            $data = [regex]::Matches($hrefLink,'(?<=a href=")(?<link>.*?)".*?>(?<description>.*)<\/a>')
            [PSCustomObject]@{
                Description = $data[0].groups['description'].Value
                Version = $shortVersion
                Date = $date
                Link = $data[0].groups['link']
            }
        }
    }
    if ($null -eq $Version) {
        return $versions
    } else {
        return $versions | Where-Object Description -like "*$Version*"
    }
}

function Get-LatestSharePointVersion {
    param(
        [ValidateSet($null,'2013','2016','2019','Subscription')]
        $Version
    )
    # https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates
    $response = Invoke-WebRequest -Uri 'https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates'
    if ($response.StatusCode -eq 200) {
        # grab the tables
        $tables = [regex]::Matches($response.Content,'<table>(.*?)<\/table','SingleLine')
        $versions = foreach($table in $tables.Value) {
            # grab the rows for each table
            $rows = [regex]::Matches($table,'<(tr)>.*?<\/\1>','SingleLine')
            # grab the data from the 2nd row containing the latest version data
            $data = [regex]::Matches($rows[1].Value, '<(th|td).*?>(?<data>.*?)<\/\1>','SingleLine')
            # groups 2 contain the data
            $description = $data[0].groups['data'].Value.replace(' <br>','')
            $hrefLink = $data[1].groups['data'].Value
            $shortVersion = $data[2].groups['data'].Value
            $date = $data[3].groups['data'].Value
            $data = [regex]::Matches($hrefLink,'(?<=a href=")(?<link>.*?)".*?>(?<description>.*)<\/a>')
            [PSCustomObject]@{
                Description = $description
                Version = $shortVersion
                Date = $date
                Link = $data[0].groups['link']
            }
        }
    }
    if ($null -eq $Version) {
        return $versions
    } else {
        return $versions | Where-Object Description -like "*$Version*"
    }
}
