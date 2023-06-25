
Param([string]$symbol, [string[]]$expirations)

function percent ($a, $b) { ($b - $a) / $a }

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
        if ($prev.mid_price -eq 0)
        {
            $val = $null
        }
        else
        {
            $val = [math]::Round(( (percent $prev.mid_price $row.mid_price) / ($row.strike - $prev.strike) ), 5)          # change up : mid price

            # $val = [math]::Round(( (percent $prev.lastPrice $row.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change up
    
            # $val = [math]::Round(( (percent $row.lastPrice $prev.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change down
    
            # $val = [math]::Round((percent $prev.lastPrice $row.lastPrice), 5)
        }
       
        $row | Add-Member -MemberType NoteProperty -Name pct_chg_per_dol -Value $val -Force

        $prev = $row
    }
    # ----------------------------------------------------------------------
    # pct_chg_per_dol : puts
    # ----------------------------------------------------------------------
    $prev = $chain.optionChain.result[0].options[0].puts[0]

    foreach ($row in $chain.optionChain.result[0].options[0].puts | Select-Object -Skip 1)
    {
        if ($prev.mid_price -eq 0)
        {
            $val = $null
        }
        else
        {
            $val = [math]::Round(( (percent $prev.mid_price $row.mid_price) / ($row.strike - $prev.strike) ), 5)          # change up : mid price

            # $val = [math]::Round(( (percent $prev.lastPrice $row.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change up
    
            # $val = [math]::Round(( (percent $row.lastPrice $prev.lastPrice) / ($row.strike - $prev.strike) ), 5)          # change down
    
            # $val = [math]::Round((percent $prev.lastPrice $row.lastPrice), 5)
        }

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

function expirationDate-to-date ()
{
    param ([Parameter(Mandatory = $true, ValueFromPipeline = $true)]$val)
        
    [System.DateTimeOffset]::FromUnixTimeSeconds($val).DateTime.ToString('yyyy-MM-dd')
}


function chart-term-structure ($symbol, $expirations, $chains)
{
    if ($chains -eq $null)
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
    }
    else
    {
        $symbol = $chains[0].optionChain.result[0].underlyingSymbol
    }
    
    $chains

    $regularMarketPrice = $chains[0].optionChain.result[0].quote.regularMarketPrice
    
    $atm_strike = $chain[0].optionChain.result[0].options[0].calls | Select-Object strike, @{ Label = 'dist'; Expression = { [math]::Abs($_.strike - $regularMarketPrice) } } | Sort-Object dist | Select-Object -First 1 | % strike
        
    # ----------------------------------------------------------------------
    $atm_calls = foreach ($chain in $chains)
    {
        # $chain[0].optionChain.result[0].options[0].calls | ? strike -GE $regularMarketPrice | Select-Object -First 1

        $chain[0].optionChain.result[0].options[0].calls | ? strike -EQ $atm_strike | Select-Object -First 1
    }

    # $atm_calls | Select-Object contractSymbol, strike, impliedVolatility

    $atm_puts = foreach ($chain in $chains)
    {
        $chain[0].optionChain.result[0].options[0].puts | ? strike -EQ $atm_strike | Select-Object -First 1
    }
    # ----------------------------------------------------------------------
    # $otm_calls = foreach ($chain in $chains)
    # {
    #     $chain[0].optionChain.result[0].options[0].calls[-1]
    # }

    # $otm_puts = foreach ($chain in $chains)
    # {
    #     $chain[0].optionChain.result[0].options[0].puts[0]
    # }


    $otm_call_strike = $chains | % { $chain = $_; $chain[0].optionChain.result[0].options[0].calls[-1].strike } | Measure-Object -Minimum | % Minimum
   
    # $chains | % { $chain = $_; $chain[0].optionChain.result[0].options[0].calls | ? strike -EQ $otm_call_strike | Select-Object contractSymbol, impliedVolatility }

    $otm_calls = foreach ($chain in $chains)
    {
        $chain[0].optionChain.result[0].options[0].calls | ? strike -EQ $otm_call_strike
    }

    # $otm_calls | Select-Object contractSymbol, strike, impliedVolatility

    $otm_put_strike = $chains | % { $chain = $_; $chain[0].optionChain.result[0].options[0].puts[0].strike } | Measure-Object -Maximum | % Maximum

    $otm_puts = foreach ($chain in $chains)
    {
        $chain[0].optionChain.result[0].options[0].puts | ? strike -EQ $otm_put_strike
    }    
    # ----------------------------------------------------------------------

    $partial_otm_call_strike = [math]::Round(($otm_call_strike - $atm_strike)/5 + $atm_strike)

    $partial_otm_calls = $chains | % {
        $chain = $_
        
        $chain[0].optionChain.result[0].options[0].calls | ? strike -EQ $partial_otm_call_strike
    }

    $partial_otm_put_strike = [math]::Round($atm_strike - ($atm_strike - $otm_put_strike)/5)

    $partial_otm_puts = $chains | % {
        $chain = $_
        
        $chain[0].optionChain.result[0].options[0].puts | ? strike -EQ $partial_otm_put_strike
    }    

    # ----------------------------------------------------------------------

    # $partial_otm_call = $chains | % {
    #     $chain = $_

    #     $otm = $chain[0].optionChain.result[0].options[0].calls[-1]

    #     $strike = [math]::Round(($otm.strike - $atm_strike) / 5 + $atm_strike)

    #     $chain[0].optionChain.result[0].options[0].calls | ? strike -EQ $strike
    # }

    # $partial_otm_put = $chains | % {
    #     $chain = $_

    #     $otm = $chain[0].optionChain.result[0].options[0].puts[0]

    #     $strike = [math]::Round($atm_strike - ($atm_strike - $otm.strike) / 5)

    #     $chain[0].optionChain.result[0].options[0].puts | ? strike -EQ $strike
    # }    

    # ----------------------------------------------------------------------




    

    # $chains | % { 
    #     $chain = $_
        
    #     $otm_strike = $chain[0].optionChain.result[0].options[0].calls[-1].strike


    # }


    # $chains | % { $chain = $_; $chain[0].optionChain.result[0].calls }

    # $chains[0].optionChain.result[0].options[0].calls | Sort-Object strike -Descending | ft *

    # $chains[0].optionChain.result[0].options[0].calls[0]

    


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
                    @{ label = 'ATM Call';                           data = $atm_calls | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = 'ATM Put' ;                           data = $atm_puts  | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = ('OTM Call {0}' -f $otm_call_strike); data = $otm_calls | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = ('OTM Put {0}'  -f $otm_put_strike ); data = $otm_puts  | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = ('OTM Call {0}' -f $partial_otm_call_strike); data = $partial_otm_calls | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    @{ label = ('OTM Put {0}'  -f $partial_otm_put_strike) ; data = $partial_otm_puts  | % { $_.impliedVolatility.ToString('N4') }; fill = $false }                    
                )
            }
            options = @{
                title = @{ display = $true; text = ('{0} IV term structure' -f $symbol) }
                scales = @{ }   
                plugins = @{ datalabels = @{ display = $true } }                
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------
}

chart-term-structure $symbol $expirations
# ----------------------------------------------------------------------
exit
# ----------------------------------------------------------------------


$result = get-options-chain SPY '2023-08-04'
$result = get-options-chain SPY '2023-08-18'



$chains_spy = .\term-structure-yfinance.ps1 SPY '2023-06-26', '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18'



$chains_spy = chart-term-structure SPY '2023-06-23', '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18'

$chains = $chains_spy

chart-term-structure -chains $chains_spy


$chains_spy