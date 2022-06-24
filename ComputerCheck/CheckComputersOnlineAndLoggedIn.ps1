Write-Host 'This script can verify both whether a computer is online and if there is a user logged in.'
Write-Host "You can specify a room or a specific computer from any campus.`n"

Write-Host "Initializing AD`n"
Import-Module -Name ActiveDirectory

$CONTINUE = 0
$CHOICES = '&Yes', '&No'
# $ROOM_AND_COMP_REGEX = '^(id\d{5}|\d{5}[a-fA-F]?|[PpSs]\d{4})(-\d\d-\d{5})?$'

while ($CONTINUE -eq 0) {
    $DOMAIN = 'instruction'
    $query = Read-Host -Prompt 'Enter the computer name (e.g., BBRRR[A]-CC-TTTTT or (P|S)BBRR-CC-TTTTT)'

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

        Write-Host 'Determining computer(s) network status...'
        # Create a fixed length array to store custom objects of results
        $online = @(0..$($computers.count-1))

        for ($i = 0; $i -lt $computers.count; $i++) {
            $percent = [int](($i / $computers.count) * 100)
            Write-Progress -Activity 'Testing Network Connection(s)' -Status "$percent complete | Pinging computer $($computers[$i])"
            $online[$i] = Test-Connection $($computers[$i]) -Count 1 -Quiet
        }
        Write-Progress -Activity 'Testing Network Connection(s)' -Status 'Done' -Completed

        # Create a fixed length array to store custom objects of results
        $results = @(0..$($computers.count-1))

        Write-Host 'Querying computer(s)...'
        for ($i = 0; $i -lt $computers.count; $i++) {
            $percent = [int](($i / $computers.count) * 100)

            # Set custom object to default to offline result
            $results[$i] = [PSCustomObject]@{
                ComputerName = $computers[$i]
                Online = $online[$i]
                LoggedIn = $false
                Username = ''
            }         

            Write-Progress -Activity 'Querying Computer(s)' -Status "$percent complete | Querying computer $($computers[$i])"

            if ($online[$i] -eq $true) {
                # Run the command to check active session and sends the error, if any, to null
                # since it will be manually handled to include the computer name that failed.
                $result = qwinsta /SERVER:$($computers[$i]) 2>$null

                # If query failed
                if (-not $?) {
                    Write-Host "Failed to connect to $($computers[$i])"
                # Otherwise, perform query
                } else {
                    # qwinsta returns a string so we need to split it and parse each line
                    foreach ($line in $result.split([Environment]::NewLine)) {
                        if (($line -match 'Active') -eq $true) {
                            $results[$i].LoggedIn = $true

                            # We want to find the username in the second column word but, console 
                            # and rdp will also match so we explicitly omit it from capture.
                            if (-not ($line -match '(?<= {1,})(?!console|rdp)[a-z\d]+')) {
                                $results[$i].Username = 'Error'
                            } else {
                                $results[$i].Username = $matches[0]
                            }
                            break
                        }
                    }
                }
            }

        }
        Write-Progress -Activity 'Querying Computer(s)' -Status 'Done' -Completed

        Write-Output $results | Format-Table -AutoSize
    }

    $CONTINUE = $host.UI.PromptForChoice(' ', 'Query another computer or room?', $CHOICES, 1)
}