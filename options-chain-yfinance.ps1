
function percent ($a, $b) { ($b - $a) / $a }
# ----------------------------------------------------------------------
function get-options-chain ($symbol, $date)
{
    $chain = Invoke-RestMethod ('https://query1.finance.yahoo.com/v7/finance/options/{0}?date={1}' -f $symbol, (Get-Date $date -UFormat %s))
    # ----------------------------------------------------------------------
    # mid price    
    # ----------------------------------------------------------------------
    foreach ($row in $chain.optionChain.result[0].options[0].calls)
    {        
        $val = ($row.ask + $row.bid) / 2
        
        $row | Add-Member -MemberType NoteProperty -Name mid_price -Value $val -Force
    }

    foreach ($row in $chain.optionChain.result[0].options[0].puts)
    {        
        $val = ($row.ask + $row.bid) / 2
        
        $row | Add-Member -MemberType NoteProperty -Name mid_price -Value $val -Force
    }
    # ----------------------------------------------------------------------
    # pct_chg_per_dol : calls
    # ----------------------------------------------------------------------
    $prev = $chain.optionChain.result[0].options[0].calls[0]

    foreach ($row in $chain.optionChain.result[0].options[0].calls | Select-Object -Skip 1)
    {
        $val = [math]::Round(( (percent $prev.mid_price $row.mid_price) / ($row.strike - $prev.strike) ), 5)          # change up : mid price

        # $val = [math]::Round(( (percent $prev.lastPrice $row.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change up

        # $val = [math]::Round(( (percent $row.lastPrice $prev.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change down

        # $val = [math]::Round((percent $prev.lastPrice $row.lastPrice), 5)

        $row | Add-Member -MemberType NoteProperty -Name pct_chg_per_dol -Value $val -Force

        $prev = $row
    }
    # ----------------------------------------------------------------------
    # pct_chg_per_dol : puts
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
    $vol_max = ($chain.optionChain.result[0].options[0].calls | Measure-Object -Maximum volume).Maximum

    $oi_max = ($chain.optionChain.result[0].options[0].calls | Measure-Object -Maximum openInterest).Maximum
    foreach ($row in $chain.optionChain.result[0].options[0].calls)
    {
        $vol_pct = [math]::Round($row.volume / $vol_max, 6)

        $row | Add-Member -MemberType NoteProperty -Name vol_pct -Value $vol_pct
        # ----------------------------------------------------------------------
        $oi_pct = [math]::Round($row.openInterest / $oi_max, 6)

        $row | Add-Member -MemberType NoteProperty -Name oi_pct -Value $oi_pct

    }
    # ----------------------------------------------------------------------
    $vol_max = ($chain.optionChain.result[0].options[0].puts | Measure-Object -Maximum volume).Maximum

    $oi_max = ($chain.optionChain.result[0].options[0].puts | Measure-Object -Maximum openInterest).Maximum
    foreach ($row in $chain.optionChain.result[0].options[0].puts)
    {
        $vol_pct = [math]::Round($row.volume / $vol_max, 6)

        $row | Add-Member -MemberType NoteProperty -Name vol_pct -Value $vol_pct
        # ----------------------------------------------------------------------
        $oi_pct = [math]::Round($row.openInterest / $oi_max, 6)

        $row | Add-Member -MemberType NoteProperty -Name oi_pct -Value $oi_pct

    }

    $chain    
}
# ----------------------------------------------------------------------
function create-dataset ($result, [ValidateSet("calls", "puts")]$type)
{
    # $symbol = $result.optionChain.result[0].underlyingSymbol
    
    $seconds = $result.optionChain.result[0].options[0].expirationDate

    $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')

    $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)   

    @{
        label = "$date ${dte}d"

        data = $result.optionChain.result[0].options[0].$type | ForEach-Object { @{ x = $_.strike; y = $_.lastPrice } }

        # pointRadius = 2

        fill = $false

        # pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.oi_pct * 100, 10) }

        pointRadius = $result.optionChain.result[0].options[0].$type | ForEach-Object { [math]::Min($_.oi_pct * 100, 5) }
    }
}

function chart ([ValidateSet("calls", "puts")]$type, $symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    $datasets = foreach ($chain in $chains)
    {
        create-dataset $chain $type
    }
  
    $json = @{
        chart = @{
            type = 'line'
            data = @{ datasets = $datasets }
            options = @{
                
                title = @{ display = $true; text = ('Options Chain : {1} : {0}' -f $symbol, $type) }
    
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

function create-dataset-chg ($result, [ValidateSet("calls", "puts")]$type)
{

    $symbol = $result.optionChain.result[0].underlyingSymbol
    
    $seconds = $result.optionChain.result[0].options[0].expirationDate

    $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')

    $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)   

    # $vol_max = ($result.optionChain.result[0].options[0].$type | Measure-Object -Maximum volume).Maximum
    
    # $result.optionChain.result[0].options[0].$type | ft *

    # $result.optionChain.result[0].options[0].$type | ForEach-Object { $_.vol_pct * 50 }

    @{
        label = "$date ${dte}d"

        data = $result.optionChain.result[0].options[0].$type | ForEach-Object { @{ x = $_.strike; y = $_.pct_chg_per_dol } }

        # pointRadius = 2

        fill = $false

        # pointRadius = $result.optionChain.result[0].options[0].$type | ForEach-Object { [math]::Min($_.vol_pct * 1000, 10) }

        # pointRadius = $result.optionChain.result[0].options[0].$type | ForEach-Object { [math]::Min($_.oi_pct * 100, 10) }

        pointRadius = $result.optionChain.result[0].options[0].$type | ForEach-Object { [math]::Min($_.oi_pct * 100, 5) }
    }
}

function chart-chg ([ValidateSet("calls", "puts")]$type, $symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    $datasets = foreach ($chain in $chains)
    {
        create-dataset-chg $chain $type
    }
  
    $json = @{
        chart = @{
            type = 'line'
            data = @{ datasets = $datasets }
            options = @{
                
                title = @{ display = $true; text = ('Options Chain : {1} : {0}' -f $symbol, $type) }
    
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




# ----------------------------------------------------------------------
function create-dataset-puts ($result)
{

    # $symbol = $result.optionChain.result[0].underlyingSymbol
    
    $seconds = $result.optionChain.result[0].options[0].expirationDate

    $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')

    $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)   

    @{
        label = "$date ${dte}d"

        data = $result.optionChain.result[0].options[0].puts | ForEach-Object { @{ x = $_.strike; y = $_.lastPrice } }

        # pointRadius = 2

        fill = $false

        # pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.oi_pct * 100, 10) }

        pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.oi_pct * 100, 5) }
    }
}
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

        # pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.oi_pct * 100, 10) }

        pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.oi_pct * 100, 5) }
    }
}
# ----------------------------------------------------------------------
# $chains = chart-puts-chg UNG 2025-01-17

# $symbol = 'UNG'
# $expirations = 

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
# ----------------------------------------------------------------------
exit
# ----------------------------------------------------------------------


$chains = chart-puts     SPY 2023-05-19, 2023-06-16, 2023-07-21, 2024-03-15, 2024-12-20
$chains = chart-puts-chg SPY 2023-05-19, 2023-06-16, 2023-07-21, 2024-03-15, 2024-12-20


$chains = chart-puts-chg COIN 2023-05-19, 2023-06-16,            2024-01-19, 2024-06-21


$chains = chart-puts     UNG 2024-01-19, 2025-01-17
$chains = chart-puts-chg UNG 2024-01-19, 2025-01-17

$chains = chart     'calls'  UNG 2024-01-19, 2025-01-17
$chains = chart-chg 'calls'  UNG 2024-01-19, 2025-01-17
# ----------------------------------------------------------------------
$chains = chart     'puts'   COIN 2024-01-19, 2025-01-17
$chains = chart-chg 'puts'   COIN 2024-01-19, 2025-01-17


$result = chart 'calls' SPY '2023-06-22', '2023-07-21'


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





$prev = $chain.optionChain.result[0].options[0].puts[0]

foreach ($row in $chain.optionChain.result[0].options[0].puts | Select-Object -Skip 1)
{
    $row | Add-Member -MemberType NoteProperty -Name pct_chg_per_dol `
        -Value ([math]::Round(( (percent $prev.lastPrice $row.lastPrice) / ($row.strike - $prev.strike) ), 5)) `
        -Force
}









$expirations = '2023-05-19 2023-06-16 2023-07-21 2023-08-18 2023-09-15 2023-10-20 2023-12-15 2024-01-19 2024-03-15 2024-12-20' -split ' '

$chains = chart-puts     SPY $expirations
$chains = chart-puts-chg SPY $expirations





$chains = chart-puts     SPY 2023-05-19, 2023-06-16, 2023-07-21, 2023-08-18, 2023-09-15, 2024-03-15, 2024-12-20    
$chains = chart-puts-chg SPY 2023-05-19, 2023-06-16, 2023-07-21, 2023-08-18, 2023-09-15, 2024-03-15, 2024-12-20

# $chains = chart-puts-chg SPY 2023-05-19, 2023-06-16
# ----------------------------------------------------------------------
$chain = get-options-chain SPY 2023-08-18

$chain.optionChain.result[0]
# ----------------------------------------------------------------------
# open interest chart

$chain = get-options-chain SPY 2023-08-18



$json = @{
    chart = @{
        type = 'bar'
        data = @{
            labels = $chain.optionChain.result[0].options[0].calls | % strike
            datasets = @(
                @{ label = 'calls open interest';      data = $chain.optionChain.result[0].options[0].calls | % openInterest }
            )
        }
        options = @{
            title = @{ display = $true; text = '' }
            scales = @{ }
        }
    }
} | ConvertTo-Json -Depth 100

$result = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'

$id = ([System.Uri] $result.url).Segments[-1]

Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)



$type = 'calls'
$symbol = 'SPY'
$expirations = '2023-07-21', '2023-08-18', '2023-09-15', '2023-10-20'


$result = $chains[0]



function create-dataset-open-interest ($result, [ValidateSet("calls", "puts")]$type)
{
    # $symbol = $result.optionChain.result[0].underlyingSymbol
    
    $seconds = $result.optionChain.result[0].options[0].expirationDate

    $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')

    $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)   

    @{
        label = "$date ${dte}d"
       
        data = $result.optionChain.result[0].options[0].$type | ForEach-Object { @{ x = $_.strike; y = $_.openInterest } }
        
        fill = $false
        
        # pointRadius = $result.optionChain.result[0].options[0].puts | ForEach-Object { [math]::Min($_.oi_pct * 100, 10) }

        # pointRadius = $result.optionChain.result[0].options[0].$type | ForEach-Object { [math]::Min($_.oi_pct * 100, 5) }
    }
}

function chart-open-interest ([ValidateSet("calls", "puts")]$type, $symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    $datasets = foreach ($chain in $chains)
    {
        create-dataset-open-interest $chain $type
    }
  
    $json = @{
        chart = @{
            type = 'bar'
            data = @{ datasets = $datasets }
            options = @{
                
                title = @{ display = $true; text = ('Options Chain : {1} : {0}' -f $symbol, $type) }
    
                scales = @{ 
                    xAxes = @(@{ type = 'linear'; scaleLabel = @{ display = $true; labelString = "Strike Price" } })
                    yAxes = @(@{ type = 'linear'; scaleLabel = @{ display = $true; labelString = "Open Interest"   } })
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

$result = chart-open-interest calls SPY '2023-07-21', '2023-08-18', '2023-09-15', '2023-10-20'


$result[0]


$chains[0]


$chains = $result



$strikes_union = $chains | ForEach-Object { $_.optionChain.result[0].options[0].calls | % strike } | Sort-Object -Unique



$chain = $chains[0]


$call = $chain.optionChain.result[0].options[0].calls[0]

$strike = $strikes_union[0]

$data = @()

$data = $data + @(1,2,3)





# ----------------------------------------------------------------------

$chains = $result

function chart-open-interest ($symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    # ----------------------------------------------------------------------
    $strikes_union = $chains | ForEach-Object { $_.optionChain.result[0].options[0].calls | % strike } | Sort-Object -Unique
    # ----------------------------------------------------------------------
    $datasets_calls = foreach ($chain in $chains)
    {
        $data = $strikes_union | ForEach-Object { 
            $strike = $_
    
            $call = $chain.optionChain.result[0].options[0].calls | Where-Object strike -EQ $strike
    
            if ($call -eq $null)
            {
                0
            }
            else
            {
                $call.openInterest
            }
        }
        
        $seconds = $chain.optionChain.result[0].options[0].expirationDate
    
        $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')
    
        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)
           
        @{
        
            label = "C $date ${dte}d"
    
            data = $data
        }
    }

    $datasets_puts = foreach ($chain in $chains)
    {
        $data = $strikes_union | ForEach-Object { 
            $strike = $_
    
            $option = $chain.optionChain.result[0].options[0].puts | Where-Object strike -EQ $strike
    
            if ($option -eq $null) { 0 } else { -$option.openInterest }
        }
        
        $seconds = $chain.optionChain.result[0].options[0].expirationDate
    
        $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')
    
        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)
           
        @{
        
            label = "P $date ${dte}d"
    
            data = $data
        }
    }
    # ----------------------------------------------------------------------
    $json = @{
        chart = @{
            type = 'bar'
            data = @{
                labels = $strikes_union
    
                datasets = $datasets_calls + $datasets_puts
            }
            options = @{
                title = @{ display = $true; text = ('{0} Open Interest' -f $symbol) }
                scales = @{ 
                    xAxes = @(@{ id = 'X1' })
                    yAxes = @(
                        @{
                            stacked = $true
                        }
                    )
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
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------
}

$result = chart-open-interest SPY '2023-06-22', '2023-07-21'

$result | ConvertTo-Json -Depth 100 > c:\temp\out.json


$result_spy = chart-open-interest SPY '2023-06-22', '2023-07-21', '2023-08-18', '2023-09-15', '2023-10-20'

$result_TSLA = chart-open-interest TSLA '2023-07-21', '2023-08-18', '2023-09-15', '2023-10-20'
# ----------------------------------------------------------------------
function chart-volume ($symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    # ----------------------------------------------------------------------
    $strikes_union = $chains | ForEach-Object { $_.optionChain.result[0].options[0].calls | % strike } | Sort-Object -Unique
    # ----------------------------------------------------------------------
    $datasets_calls = foreach ($chain in $chains)
    {
        $data = $strikes_union | ForEach-Object { 
            $strike = $_
    
            $call = $chain.optionChain.result[0].options[0].calls | Where-Object strike -EQ $strike
    
            if ($call -eq $null)
            {
                0
            }
            else
            {
                $call.volume
            }
        }
        
        $seconds = $chain.optionChain.result[0].options[0].expirationDate
    
        $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')
    
        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)
           
        @{
        
            label = "C $date ${dte}d"
    
            data = $data
        }
    }

    $datasets_puts = foreach ($chain in $chains)
    {
        $data = $strikes_union | ForEach-Object { 
            $strike = $_
    
            $option = $chain.optionChain.result[0].options[0].puts | Where-Object strike -EQ $strike
    
            if ($option -eq $null) { 0 } else { -$option.volume }
        }
        
        $seconds = $chain.optionChain.result[0].options[0].expirationDate
    
        $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')
    
        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)
           
        @{
        
            label = "P $date ${dte}d"
    
            data = $data
        }
    }
    # ----------------------------------------------------------------------
    $json = @{
        chart = @{
            type = 'bar'
            data = @{
                labels = $strikes_union
    
                datasets = $datasets_calls + $datasets_puts
            }
            options = @{
                title = @{ display = $true; text = ('{0} Volume' -f $symbol) }
                scales = @{ 
                    xAxes = @(@{ id = 'X1' })
                    yAxes = @(
                        @{
                            stacked = $true
                        }
                    )
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
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------
}

$result_spy = chart-volume SPY '2023-06-22', '2023-07-21', '2023-08-18', '2023-09-15', '2023-10-20'

$result_spy | ConvertTo-Json -Depth 100 > spy.json



# ----------------------------------------------------------------------

$chain = Invoke-RestMethod ('https://query1.finance.yahoo.com/v7/finance/options/{0}?date={1}' -f 'SPY', (Get-Date '2023-08-18' -UFormat %s))
# ----------------------------------------------------------------------
# term structure
# ----------------------------------------------------------------------

$symbol = 'TLT'

$expirations = '2023-06-23', '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18'

$chains = foreach ($date in $expirations)
{
    get-options-chain $symbol $date
}

$chain = $chains[0]




function expirationDate-to-date ()
{
    param ([Parameter(Mandatory = $true, ValueFromPipeline = $true)]$val)
        
    [System.DateTimeOffset]::FromUnixTimeSeconds($val).DateTime.ToString('yyyy-MM-dd')
}


# $chain.optionChain.result[0].options[0].expirationDate | expirationDate-to-date


function chart-term-structure ($symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }
    
    $chains = foreach ($chain in $chains)
    {
        if ($chain.optionChain.result[0].options[0].calls.Count -gt 0)
        {
            $chain
        }
    }

    $chains

    $regularMarketPrice = $chains[0].optionChain.result[0].quote.regularMarketPrice

    # ----------------------------------------------------------------------
    # $strikes_union = $chains | ForEach-Object { $_.optionChain.result[0].options[0].calls | % strike } | Sort-Object -Unique
    # ----------------------------------------------------------------------
    $datasets_calls = foreach ($chain in $chains)
    {
        $data = $strikes_union | ForEach-Object { 
            $strike = $_
    
            $call = $chain.optionChain.result[0].options[0].calls | Where-Object strike -EQ $strike
    
            if ($call -eq $null)
            {
                0
            }
            else
            {
                $call.volume
            }
        }
        
        $seconds = $chain.optionChain.result[0].options[0].expirationDate
    
        $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')
    
        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)
           
        @{
        
            label = "C $date ${dte}d"
    
            data = $data
        }
    }

    $datasets_puts = foreach ($chain in $chains)
    {
        $data = $strikes_union | ForEach-Object { 
            $strike = $_
    
            $option = $chain.optionChain.result[0].options[0].puts | Where-Object strike -EQ $strike
    
            if ($option -eq $null) { 0 } else { -$option.volume }
        }
        
        $seconds = $chain.optionChain.result[0].options[0].expirationDate
    
        $date = [System.DateTimeOffset]::FromUnixTimeSeconds($seconds).DateTime.ToString('yyyy-MM-dd')
    
        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)
           
        @{
        
            label = "P $date ${dte}d"
    
            data = $data
        }
    }

    

    $atm_calls = foreach ($chain in $chains)
    {
        $chain[0].optionChain.result[0].options[0].calls | ? strike -GE $regularMarketPrice | Select-Object -First 1
    }

    $atm_puts = foreach ($chain in $chains)
    {
        $chain[0].optionChain.result[0].options[0].puts | ? strike -GE $regularMarketPrice | Select-Object -First 1
    }    

    $otm_calls = foreach ($chain in $chains)
    {
        $chain[0].optionChain.result[0].options[0].calls[-1]
    }

    $otm_puts = foreach ($chain in $chains)
    {
        $chain[0].optionChain.result[0].options[0].puts[0]
    }

    



    # ----------------------------------------------------------------------
    $json = @{
        chart = @{
            type = 'line'
            data = @{
                                
                labels = $chains | % { 
                    $chain = $_; 

                    $chain.optionChain.result[0].options[0].expirationDate | expirationDate-to-date
                }

                datasets = @(
                    @{ label = 'ATM Call'; data = $atm_calls | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = 'ATM Put' ; data = $atm_puts  | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = 'OTM Call'; data = $otm_calls | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = 'OTM Put' ; data = $otm_puts  | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                )
            }
            options = @{
                title = @{ display = $true; text = ('{0} IV term structure' -f $symbol) }
                scales = @{ 
                    # xAxes = @(@{ id = 'X1' })
                    # yAxes = @(
                    #     @{
                    #         stacked = $true
                    #     }
                    # )
                }

                # annotation = @{
    
                #     annotations = @(
    
                #         @{
                #             type = 'line'
                #             mode = 'vertical'
                #             # value = $result_0.optionChain.result[0].quote.regularMarketPrice
                #             value = $chains[0].optionChain.result[0].quote.regularMarketPrice                            
                #             scaleID = 'X1'
                #             borderColor = 'red'
                #             borderWidth = 1
                #             label = @{ }
                #         }
                #     )
                # }
    
                plugins = @{ datalabels = @{ display = $true } }                
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------
}

$result_spy = chart-term-structure SPY '2023-06-23', '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18'