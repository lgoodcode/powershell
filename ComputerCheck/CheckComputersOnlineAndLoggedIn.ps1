Write-Host 'This script can verify both whether a computer is online and if there is a user logged in.'
Write-Host "You can specify a room or a specific computer from any campus.`n"

Write-Host "Initializing AD`n"
Import-Module -Name ActiveDirectory

$WRITE = 0
$CONTINUE = 0
$CHOICES = '&Yes', '&No'
$DESKTOP = "$env:USERPROFILE\desktop\results.csv"
# $ROOM_AND_COMP_REGEX = '^(id\d{5}|\d{5}[a-fA-F]?|[PpSs]\d{4})(-\d\d-\d{5})?$'

while ($CONTINUE -eq 0) {
    $DOMAIN = 'instruction'
    $query = Read-Host -Prompt 'Enter the computer name (e.g., BBRRR[A]-CC-TTTTT or (P|S)BBRR-CC-TTTTT)'
    # Will contain the results of the testing so it can be used to output to a file if desired
    $queries = $null

    if (-not ($query -imatch '^(\d\d|id\d{5}|\d{5}[a-fA-F]?|[PpSs](\d|\d{4}))(-\d\d-\d{5})?$')) {
        Write-Host 'Must enter a valid room or computer name'
        continue
    } else {
        # If querying an entire building
        if ($query -match '^(\d{2}|[ps]\d)$') {
            Write-Host "`nWARNING: This can take a long time (1-2 hours) if there a lot of offline computers e.g., building 04`n" -ForegroundColor Red

            # Prompt the user to confirm if they want to continue with the query
            if ($host.UI.PromptForChoice($null, 'Contiue query?', $CHOICES, 1)) {
                continue
            }
        }

        # Edge case: single testing center room and office computers are on office domain
        if ($query -imatch '^04222[a]|^id\d{5}$') {
            $DOMAIN = 'office'
        }

        # If specifying a string less than a full computer name, then it is a room
        if ($query.Length -lt 7) {
            $query = $query + '*'
        }

        Write-Host 'Retrieving computer(s) from AD...'
        $SERVER = "$DOMAIN.oc.ctc.edu"
        $SEARCH_BASE = "DC=$DOMAIN,DC=oc,DC=ctc,DC=edu"
        $computers = Get-ADComputer -LDAPFilter "(name=$query)" -Server $SERVER -SearchBase $SEARCH_BASE | ForEach-Object{ $_.Name }

        # If query cannot find any results, it will return null
        if ($null -eq $computers) {
            Write-Host "Computer(s) not found for: $query"
            $CONTINUE = $host.UI.PromptForChoice('CONTINUE', 'Query another computer or room?', $CHOICES, 1)
            break
        }

        # If querying a single computer, it will return the name as a string and not an array.
        # If a string, we place it in an array.
        if ($computers.GetType().Name -eq 'String') {
            $computers = @($computers)
        }

        Write-Host 'Testing computer(s) network status...'

        # For each computer, start a job testing the connection. Once each job is then, it will return 
        # the object we will use containing the computer name, boolean value indicating if it is 
        # online or not, and then the username, which defaults to null for no user logged in.
        $online =  $computers | ForEach-Object { Test-Connection $_ -Count 1 -AsJob } | Wait-Job | Receive-Job | ForEach-Object {[PSCustomObject]@{
            Computer = $_.Address
            Online = $(if ($_.StatusCode -eq 0) { $true } else { $false })
            User = $null
        }}

        Write-Host 'Querying computer(s)...'

        # With the results of the connection testing, if the computer is online, we will perform the
        # query using `qwinsta` rather than PS `query user`, which for some reason would return buggy
        # results and sometimes takes longer than `qwinsta`. 
        $queries = $online | ForEach-Object {
            if ($_.Online) {
                # Run the command to check active session and sends the error, if any, to null
                # since it will be manually handled to include the computer name that failed.
                $result = qwinsta /SERVER:$($_.Computer) 2>$null

                # If query failed
                if (-not $?) {
                    Write-Host "Failed to connect to $($_.Computer)"
                # Otherwise, perform query
                } else {
                    # qwinsta returns a string so we need to split it and parse each line
                    foreach ($line in $result.split([Environment]::NewLine)) {
                        if (($line -match 'Active') -eq $true) {
                            # We want to find the username in the second column word but, console 
                            # and rdp will also match so we explicitly omit it from capture.
                            if (-not ($line -match '(?<= {1,})(?!console|rdp)[a-z\d]+')) {
                                $_.User = 'Error'
                            } else {
                                $_.User = $matches[0]
                            }
                            # We found the active user line, stop looping the result
                            break
                        }
                    }
                }
            }

            return $_
        }
        
        Write-Output $queries | Format-Table -AutoSize
    }

    $WRITE = $host.UI.PromptForChoice(' ', 'Write output to file?', $CHOICES, 1)

    if ($WRITE -eq 0) {
        $path = Read-Host -Prompt "Enter directory to save file [desktop]"

        while ($true) {
            if ($path -eq '') {
                $path = $DESKTOP
                break
            }

            if (Test-Path $path -IsValid) {
                if (-not $path.endsWith('\')) {
                    $path = $path + '\'
                }
                $path = $path + "results.csv"
                break
            }

            Write-Host 'Invalid path. Try again.'
        }

        # Convert the result object to CSV and replace the first line which contains
        # PS information about the object converted, which we don't want, to specify
        # the delimiter for excel to use.
        $csv = $queries | ConvertTo-Csv
        $csv[0] = 'sep=,'

        write-output "path $path"
        $csv | Out-File -FilePath $path
    }

    $CONTINUE = $host.UI.PromptForChoice(' ', 'Query another computer or room?', $CHOICES, 1)
}