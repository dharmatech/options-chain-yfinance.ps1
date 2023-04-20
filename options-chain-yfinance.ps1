




$result = Invoke-RestMethod 'https://query1.finance.yahoo.com/v7/finance/options/SPY?date=1710460800'


([DateTimeOffset] (Get-Date '2024-03-15')).ToUnixTimeSeconds()

$symbol = 'SPY'
$date = '2024-03-15'

function get-options-chain ($symbol, $date)
{
    # echo ('https://query1.finance.yahoo.com/v7/finance/options/{0}?date={1}' -f $symbol, ([DateTimeOffset] (Get-Date $date)).ToUnixTimeSeconds())
    # Invoke-RestMethod ('https://query1.finance.yahoo.com/v7/finance/options/{0}?date={1}' -f $symbol, ([DateTimeOffset] (Get-Date $date)).ToUnixTimeSeconds())

    # Get-Date 2024-03-15 -UFormat %s

    $chain = Invoke-RestMethod ('https://query1.finance.yahoo.com/v7/finance/options/{0}?date={1}' -f $symbol, (Get-Date $date -UFormat %s))
    # ----------------------------------------------------------------------
    # mid price    
    # ----------------------------------------------------------------------
    foreach ($row in $chain.optionChain.result[0].options[0].puts)
    {        
        $val = ($row.ask + $row.bid) / 2
        
        $row | Add-Member -MemberType NoteProperty -Name mid_price -Value $val
    }
    # ----------------------------------------------------------------------
    $prev = $chain.optionChain.result[0].options[0].puts[0]

    foreach ($row in $chain.optionChain.result[0].options[0].puts | Select-Object -Skip 1)
    {
        $val = [math]::Round(( (percent $prev.mid_price $row.mid_price) / ($row.strike - $prev.strike) ), 5)          # change up : mid price

        # $val = [math]::Round(( (percent $prev.lastPrice $row.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change up

        # $val = [math]::Round(( (percent $row.lastPrice $prev.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change down

        # $val = [math]::Round((percent $prev.lastPrice $row.lastPrice), 5)

        $row | Add-Member -MemberType NoteProperty -Name pct_chg_per_dol -Value $val -Force

        $prev = $row
    }
    # ----------------------------------------------------------------------

    foreach ($row in $chain.optionChain.result[0].options[0].puts)
    {
        $vol_max = ($result.optionChain.result[0].options[0].puts | Measure-Object -Maximum volume).Maximum

        $vol_pct = [math]::Round($row.volume / $vol_max, 6)

        $row | Add-Member -MemberType NoteProperty -Name vol_pct -Value $vol_pct

        # ----------------------------------------------------------------------
        $oi_max = ($result.optionChain.result[0].options[0].puts | Measure-Object -Maximum openInterest).Maximum

        $oi_pct = [math]::Round($row.openInterest / $oi_max, 6)

        $row | Add-Member -MemberType NoteProperty -Name oi_pct -Value $oi_pct

    }

    $chain    
}

$result = get-options-chain SPY 2023-04-20

$result = get-options-chain SPY 2024-03-15




$result_0 = get-options-chain SPY 2023-05-19
$result_1 = get-options-chain SPY 2023-06-16
$result_2 = get-options-chain SPY 2024-03-15








1710460800

1710486000

([DateTimeOffset] (Get-Date $date)).ToUnixTimeSeconds()


1710460800

[System.DateTimeOffset]::FromUnixTimeSeconds(1710460800).DateTime.ToString('s')


[System.DateTimeOffset]::FromUnixTimeSeconds(1710460800).DateTime.ToString('yyyy-MM-dd')


[System.DateTimeOffset]



[System.DateTimeOffset](Get-Date 2024-03-15)

[System.DateTimeOffset](Get-Date 1970-01-01)

([System.DateTimeOffset](Get-Date 1970-01-01)).ToUnixTimeSeconds()


[System.DateTimeOffset]::new(1970, 1, 1, 0, 0, 0, [timespan]::Zero)



[System.DateTimeOffset]::new(1970, 1, 1, 0, 0, 0, [timespan]::Zero).ToUnixTimeSeconds()


[System.DateTimeOffset]::new(1970, 1, 1, 0, 0, 0, [timespan]::Zero)

[System.DateTimeOffset]((Get-Date 1970-01-01).ToUniversalTime())


Get-Date 1970-01-01 -UFormat %s

Get-Date 2024-03-15 -UFormat %s




# ----------------------------------------------------------------------

$table = $result.optionChain.result[0].options[0].puts

$json = @{
    chart = @{
        type = 'line'
        data = @{
            
            labels = $table | ForEach-Object strike

            datasets = @(

                @{ label = 'SPY';       data = $table | ForEach-Object lastPrice; pointRadius = 2; }

            )
        }
        options = @{
            
            title = @{ display = $true; text = 'Options Chain' }

            scales = @{ }
        }
    }
} | ConvertTo-Json -Depth 100

$result = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'

# Start-Process $result.url

$id = ([System.Uri] $result.url).Segments[-1]

Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)
# ----------------------------------------------------------------------

$table = $result.optionChain.result[0].options[0].puts



$table | ForEach-Object { @{ x = $_.strike; y = $_.lastPrice } }


$json = @{
    chart = @{
        type = 'line'
        data = @{
            
            # labels = $table | ForEach-Object strike

            datasets = @(

                # @{ label = 'SPY';       data = $table | ForEach-Object lastPrice; pointRadius = 2; }

                @{ 
                    label = 'SPY'
                    
                    # data = $table | ForEach-Object lastPrice

                    data = $table | ForEach-Object { @{ x = $_.strike; y = $_.lastPrice } }
                    
                    pointRadius = 2
                }

            )
        }
        options = @{
            
            title = @{ display = $true; text = 'Options Chain' }

            scales = @{ 
                xAxes = @(@{ type = 'linear'})
                yAxes = @(@{ type = 'linear'})
            }
        }
    }
} | ConvertTo-Json -Depth 100

$result = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'

# Start-Process $result.url

$id = ([System.Uri] $result.url).Segments[-1]

Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)







# ----------------------------------------------------------------------
# linear axis
# ----------------------------------------------------------------------

function create-dataset-puts ($result)
{

    $symbol = $result.optionChain.result[0].underlyingSymbol
    
    $seconds = $result.optionChain.result[0].options[0].expirationDate

    $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')

    $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)   

    @{
        # label = $date

        label = "$date ${dte}d"

        data = $result.optionChain.result[0].options[0].puts | ForEach-Object { @{ x = $_.strike; y = $_.lastPrice } }

        # pointRadius = 2

        fill = $false
    }
}
# ----------------------------------------------------------------------



$json = @{
    chart = @{
        type = 'line'
        data = @{            
            datasets = @(
                create-dataset-puts $result_0
                create-dataset-puts $result_1
                create-dataset-puts $result_2
            )
        }
        options = @{
            
            title = @{ display = $true; text = 'Options Chain' }

            scales = @{ 
                xAxes = @(@{ type = 'linear'})
                yAxes = @(@{ type = 'linear'})
            }

            annotation = @{

                annotations = @(

                    @{
                        type = 'line'
                        mode = 'vertical'
                        value = $result_0.optionChain.result[0].quote.regularMarketPrice
                        scaleID = 'X1'
                        borderColor = 'red'
                        borderWidth = 1
                        label = @{ }
                    }
                )
            }

            plugins = @{ datalabels = @{ display = $false } }

        }
    }
} | ConvertTo-Json -Depth 100

$result = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'

# Start-Process $result.url

$id = ([System.Uri] $result.url).Segments[-1]

Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)
# ----------------------------------------------------------------------
function chart-puts ($symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    $datasets = foreach ($chain in $chains)
    {
        create-dataset-puts $chain
    }
  
    $json = @{
        chart = @{
            type = 'line'
            data = @{ datasets = $datasets }
            options = @{
                
                title = @{ display = $true; text = 'Options Chain : puts : {0}' -f $symbol }
    
                scales = @{ 
                    xAxes = @(@{ type = 'linear'; scaleLabel = @{ display = $true; labelString = "Strike Price" } })
                    yAxes = @(@{ type = 'linear'; scaleLabel = @{ display = $true; labelString = "Last Price"   } })
                }
    
                annotation = @{
    
                    annotations = @(
    
                        @{
                            type = 'line'
                            mode = 'vertical'
                            # value = $result_0.optionChain.result[0].quote.regularMarketPrice
                            value = $chains[0].optionChain.result[0].quote.regularMarketPrice                            
                            scaleID = 'X1'
                            borderColor = 'red'
                            borderWidth = 1
                            label = @{ }
                        }
                    )
                }
    
                plugins = @{ datalabels = @{ display = $false } }
    
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    # Start-Process $result.url
    
    $id = ([System.Uri] $result.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)       
}



exit

# ----------------------------------------------------------------------

(Get-Date 2024-12-20) - (Get-Date)



$chains = chart-puts SPY 2023-05-19, 2023-06-16, 2023-07-21, 2024-03-15, 2024-12-20


$chains = chart-puts TSLA 2023-05-19, 2023-06-16, 2024-03-15

$chains = chart-puts QQQ 2023-05-19, 2023-06-16, 2024-03-15


$symbol = 'SPY'
$expirations = '2023-05-19','2023-06-16','2024-03-15'

$chains[0].optionChain.result[0].quote.regularMarketPrice


# ----------------------------------------------------------------------
$chain = get-options-chain SPY 2024-03-15

$chain.optionChain.result[0].options[0].puts | ft *



# $row = $chain.optionChain.result[0].options[0].puts[1]


function percent ($a, $b) { ($b - $a) / $a }


$prev = $chain.optionChain.result[0].options[0].puts[0]

foreach ($row in $chain.optionChain.result[0].options[0].puts | Select-Object -Skip 1)
{
    $row | Add-Member -MemberType NoteProperty -Name pct_chg_per_dol `
        -Value ([math]::Round(( (percent $prev.lastPrice $row.lastPrice) / ($row.strike - $prev.strike) ), 5)) `
        -Force
}




# ----------------------------------------------------------------------
# plot percent change in price
# ----------------------------------------------------------------------
function create-dataset-puts-chg ($result)
{

    $symbol = $result.optionChain.result[0].underlyingSymbol
    
    $seconds = $result.optionChain.result[0].options[0].expirationDate

    $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')

    $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)   

    # $vol_max = ($result.optionChain.result[0].options[0].puts | Measure-Object -Maximum volume).Maximum
    
    # $result.optionChain.result[0].options[0].puts | ft *

    # $result.optionChain.result[0].options[0].puts | ForEach-Object { $_.vol_pct * 50 }

    @{
        label = "$date ${dte}d"

        data = $result.optionChain.result[0].options[0].puts | ForEach-Object { @{ x = $_.strike; y = $_.pct_chg_per_dol } }

        # pointRadius = 2

        fill = $false

        # pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.vol_pct * 1000, 10) }

        pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.oi_pct * 100, 10) }
    }
}



function chart-puts-chg ($symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    $datasets = foreach ($chain in $chains)
    {
        create-dataset-puts-chg $chain
    }
  
    $json = @{
        chart = @{
            type = 'line'
            data = @{ datasets = $datasets }
            options = @{
                
                title = @{ display = $true; text = 'Options Chain : puts : {0}' -f $symbol }
    
                scales = @{ 
                    xAxes = @(@{ type = 'linear'; scaleLabel = @{ display = $true; labelString = "Strike Price" } })
                    yAxes = @(@{ type = 'linear'; scaleLabel = @{ display = $true; labelString = "Percent change of price per dollar"   } })
                }
    
                annotation = @{
    
                    annotations = @(
    
                        @{
                            type = 'line'
                            mode = 'vertical'
                            # value = $result_0.optionChain.result[0].quote.regularMarketPrice
                            value = $chains[0].optionChain.result[0].quote.regularMarketPrice                            
                            scaleID = 'X1'
                            borderColor = 'red'
                            borderWidth = 1
                            label = @{ }
                        }
                    )
                }
    
                plugins = @{ datalabels = @{ display = $false } }
    
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    # Start-Process $result.url
    
    $id = ([System.Uri] $result.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)       
}




$chains = chart-puts     SPY 2023-05-19, 2023-06-16, 2023-07-21, 2024-03-15, 2024-12-20
$chains = chart-puts-chg SPY 2023-05-19, 2023-06-16, 2023-07-21, 2024-03-15, 2024-12-20

# $chains = chart-puts-chg SPY 2023-05-19, 2023-06-16