Function Get-SortBy {
<#
.SYNOPSIS
    A helper function which sorts the search results by specific criteria.
.DESCRIPTION
    It's not supposed to be run as stand-alone
#>
    Param (
        [Parameter(Mandatory=$true)]$DOMObject,
        [Parameter(Mandatory=$true)]$Criteria,
        [Parameter(Mandatory=$true)]$Timeout
    )
    Begin {
        Write-Verbose "< Sorting results by: $Criteria"
        $maxcount = $Timeout * 2
        $count = 0
        $element = $DOMObject.Document.IHTMLDocument3_getElementById('sortBy')
        $defaultCriteria = $element[$element.selectedIndex].text
    }
    Process {
        $itemsListArray = @()
        $itemsList = @{}
        $i = 0
        $element | %{
            $itemsListArray += $_.text.Trim()
        }
        $itemsListArray | Sort-Object | %{
            $key = $i++
            $itemsList[$key] = $_
        }

        if ($itemsList.ContainsValue($Criteria)) {
            $element | %{
                if ($_.text.Trim() -eq $Criteria) {
                    $element.value = $_.value
                    $element.FireEvent('onchange') | Out-Null
                }
            }
            do {
                Start-Sleep -Milliseconds 500
                $count++
            } until (($DOMObject.Document.readyState -eq 'complete') -or ($count -ge $maxcount))
            if ($DOMObject.Document.readyState -eq 'complete') {
                Write-Verbose "> Sorting results finished successfully"
            }
            else {
                Write-Verbose "> Request timeout. Please check website message."
                $DOMObject.Visible = $true
                break
            }
        }
        else {
            Write-Verbose "> Criteria `"$Criteria`" could not found"
            $itemsList.Keys | Sort-Object | %{ Write-Host $_":" $itemsList[$_] }
            Write-Host -NoNewline -ForegroundColor green "Please select sort criteria`: "
            $itemIndex = Read-Host

            if ($itemIndex) {
                $item = $itemsList[[int]$itemIndex]
                Write-Verbose "< Sorting results by: `"$item`".."
                $element | %{
                    if ($_.text.Trim() -eq $item) {
                        $element.value = $_.value
                        $element.FireEvent('onchange') | Out-Null
                    }
                }
                do {
                    Start-Sleep -Milliseconds 500
                    $count++
                } until (($DOMObject.Document.readyState -eq 'complete') -or ($count -ge $maxcount))
                if ($DOMObject.Document.readyState -eq 'complete') {
                    Write-Verbose "> Sorting results finished successfully"
                }
                else {
                    Write-Verbose "> Request timeout. Please check website message."
                    $DOMObject.Visible = $true
                    break
                }
            }
            else {
                Write-Verbose "Empty or invalid choice. Select default criteria: $defaultCriteria"
            }
        }
    }
}

Function Get-NarrowFocus {
<#
.SYNOPSIS
    A helper function which narrows focus of the search results
.DESCRIPTION
    It's not supposed to be run as stand-alone
#>
    Param (
        [Parameter(Mandatory=$true)]$DOMObject,
        [Parameter(Mandatory=$true)]$focus,
        [Parameter(Mandatory=$true)]$focusItem,
        [Parameter(Mandatory=$true)]$Timeout
    )
    Begin {
        Write-Verbose "< Selecting $focus`: `"$focusItem`".."
        $maxcount = $Timeout * 2
        $count = 0
        $idList = $DOMObject.Document.IHTMLDocument3_getElementsByName('idList')
        $table = $idList[0].parentElement

        if ($table) {
            $narrowFocusTable = $table.getElementsByClassName('GS_bgcolor')
            Switch ($focus) {
                "Language" { $narrowFocusItems = $narrowFocusTable[2].getElementsByTagName('A'); break }
                "Category" { $narrowFocusItems = $narrowFocusTable[1].getElementsByTagName('A'); break }
                "Product" { $narrowFocusItems = $narrowFocusTable[0].getElementsByTagName('A'); break }
            }
        }
        else {
            Write-Verbose "> Could not narrow focus by $focus."
        }
    }
    Process {
        # Build focus list
        $itemsListArray = @()
        $itemsList = @{}
        $i = 0

        if ($narrowFocusItems) {
            $narrowFocusItems | %{
                $itemsListArray += $_.innerTEXT.Trim()
            }
            $itemsListArray | Sort-Object | %{
                $key = $i++
                $itemsList[$key] = $_
            }

            if ($itemsList.ContainsValue($focusItem)) {
                $narrowFocusItems | %{
                    if ($_.innerTEXT.Trim() -eq $focusItem) {
                        $_.click()
                    }
                }
                do {
                    Start-Sleep -Milliseconds 500
                    $count++
                } until (($DOMObject.Document.readyState -eq 'complete') -or ($count -ge $maxcount))
                if ($DOMObject.Document.readyState -eq 'complete') {
                    Write-Verbose "> Selecting $focus finished successfully"
                }
                else {
                    Write-Verbose "> Request timeout. Please check website message."
                    $DOMObject.Visible = $true
                    break
                }
            }
            else {
                Write-Verbose "> $focus `"$focusItem`" could not found"
                $itemsList.Keys | Sort-Object | %{ Write-Host $_":" $itemsList[$_] }
                Write-Host -NoNewline -ForegroundColor green "Please select $focus`: "
                $itemIndex = Read-Host
                $item = $itemsList[[int]$itemIndex]

                if ($itemIndex -and $item) {
                    Write-Verbose "< Selecting $focus`: `"$item`".."
                    $narrowFocusItems | %{
                        if ($_.innerTEXT.Trim() -eq $item) {
                            $_.click()
                        }
                    }
                    do {
                        Start-Sleep -Milliseconds 500
                        $count++
                    } until (($DOMObject.Document.readyState -eq 'complete') -or ($count -ge $maxcount))
                    if ($DOMObject.Document.readyState -eq 'complete') {
                        Write-Verbose "> Selecting $focus finished successfully"
                    }
                    else {
                        Write-Verbose "> Request timeout. Please check website message."
                        $DOMObject.Visible = $true
                        break
                    }
                }
                else {
                    Write-Verbose "Empty or invalid choice. Select all $focus"
                }
            }
        }
    }
}

Function Search-VMwareKB {
<#
.SYNOPSIS
    A PowerShell Module for searching VMware KB articles on the command line.

.DESCRIPTION
    A PowerShell Module for searching VMware KB articles on the command line.
    It uses Internet Explorer COM Object to interact with the VMware KB site.

    You can search for a keyword, sort the results and narrow focus down to
    specific language, category, product.
    * Just as VMware KB site, narrow focus conditions are dynamic.
      (For example, if there's no article written in a language, you cannot
      narrow focus to it)

    Search results are returned as a PowerShell Array which contains following
    properties:
        . Title
        . URL - https://kb.vmware.com/kb/[Article #]
        . Description
        . Rating - # of stars (if there)
        . Published / Created Date / Last Modified Date
            - DateTime.ToShortDateString()

.PARAMETER Keyword
    A search keyword
.PARAMETER SortBy
    Sort the search results by specific criteria such as "Most Relevant",
    "Publication Date", etc.
.PARAMETER Language
    Narrow focus down to specific language such as "English", "日本語", etc.
.PARAMETER Category
    Narrow focus down to specific category such as "Troubleshooting", etc.
.PARAMETER Product
    Narrow focus down to specific product such as "VMware ESXi 6.5.x", etc.
.PARAMETER Timeout
    Set timeout value of fetching HTML document, DOM element, etc.

.EXAMPLE
    Search-VMwareKB PSOD
    Search VMware KB site using the keyword 'PSOD'

    === Sample Output ===
    Title            : "PF Exception 14 in world 32868:helper11-0 IP 0x418008f10260" PSOD in ESXi 5.x or 6.0.x host (2114745)
    URL              : https://kb.vmware.com/kb/2114745
    Description      : change the Latency sensitivity of the virtual machine to normal to prevent any further occurrence of the PSOD. Note: The host failing with PSOD has the virtual machine configured for High Latency sensitivity. To change...
    Rating           : 5
    Published        : 2017-01-24
    CreatedDate      : 2015-04-21
    LastModifiedDate : 2017-01-24
.EXAMPLE
    Search-VMwareKB -Keyword 'no workaround' -SortBy 'Publication Date'
    Search VMware KB site for most recently published article using the keyword 'no workaround' 
.EXAMPLE
    Search-VMwareKB -Keyword 'no workaround' -SortBy * -Language * -Category * -Product *
    Use '*' as a parameter value if you don't know
.EXAMPLE
    Search-VMwareKB -Keyword 'fails' | %{ start $_.URL }
    Open all KB articles at once in search results using default web browser

.NOTES
    Author                      : Han Ho-Sung
    Author email                : funksoul@insdata.co.kr
    Version                     : 1.0
    Dependencies                : 
    ===Tested Against Environment====
    ESXi Version                : 
    PowerCLI Version            : 
    PowerShell Version          : 5.1.14393.693
#>

    Param (
        [Parameter(Mandatory=$true, Position=0)][String]$Keyword,
        [Parameter(Mandatory=$false)]$SortBy,
        [Parameter(Mandatory=$false)]$Language,
        [Parameter(Mandatory=$false)]$Category,
        [Parameter(Mandatory=$false)]$Product,
        [Parameter(Mandatory=$false)]$Timeout = 60
    )

    Begin {
        $ie = New-Object -ComObject 'InternetExplorer.Application'
        $url = 'https://kb.vmware.com/selfservice/microsites/microsite.do'
        $count = 0
        $maxcount = $Timeout * 2
    }

    Process {
        Write-Verbose "< Opening URL: $url"
        $ie.Navigate($url)
        do {
            Start-Sleep -Milliseconds 500
            $count++
        } until (($ie.Document.readyState -eq 'complete') -or ($count -ge $maxcount))
        if ($ie.Document.readyState -eq 'complete') {
            Write-Verbose "> URL opened successfully"
        }
        else {
            Write-Host "> Request timeout. Please check website message."
            $ie.Visible = $true
            break
        }

        # Search for a keyword
        Write-Verbose "< Searching for the keyword: `"$Keyword`""
        Try {
            $searchForm = $ie.Document.IHTMLDocument3_getElementById('id_searchForm')
        }
        Catch {
            Write-Host "> $_"
            Write-Host "Unknown error occurred. Please check website message."
            $ie.Visible = $true
            break
        }
        $searchString = $searchForm | Where-Object { $_.name -eq 'searchString' }
        $btnSearchAll = $searchForm | Where-Object { $_.name -eq 'btnSearchAll' }
        $searchString.value = $Keyword
        $btnSearchAll.click()
        do {
            Start-Sleep -Milliseconds 1000
        } until ($ie.Document.readyState -eq 'complete')
        Write-Verbose "> Keyworld search finished successfully"

        # Sort results
        if ($SortBy) {
            Get-SortBy -DOMObject $ie -Criteria $SortBy -Timeout $Timeout
        }

        # Narrow Focus (Language)
        if ($Language) {
            Get-NarrowFocus -DOMObject $ie -focus 'Language' -focusItem $Language -Timeout $Timeout
        }

        # Narrow Focus (Category)
        if ($Category) {
            Get-NarrowFocus -DOMObject $ie -focus 'Category' -focusItem $Category -Timeout $Timeout
        }

        # Narrow Focus (Products)
        if ($Product) {
            Get-NarrowFocus -DOMObject $ie -focus 'Product' -focusItem $Product -Timeout $Timeout
        }

        # Create array of sort results
        $result = @()
        $searchRes = $ie.Document.IHTMLDocument3_getElementById('searchres')
        $searchRes.getElementsByClassName('vmdoc') | %{
            $row = "" | Select-Object "Title", "URL", "Description", "Rating", "Published", "CreatedDate", "LastModifiedDate"
            $row.Title = $_.getElementsByClassName('doctitle')[0].innerText.Trim()
            $row.URL = 'https://kb.vmware.com/kb/' + ($_.getElementsByClassName('doctitle')[0].getElementsByTagName('A')[0].href -split 'externalId=')[1] -replace '&.*',''
            $row.Description = $_.getElementsByTagName('synopsis')[0].innerText.Trim()
            $metadata = $_.getElementsByClassName('metadata')[0]
            $Rating = 0
            if ($metadata.innerText -like "*Rating:*") {
                $row.Rating = $metadata.getElementsByTagName('img') | %{
                    if ($_.src -like '*icon_rating_star.gif') {
                        $Rating++
                    }
                }
                $row.Rating = $Rating
                $row.Published = ([datetime](($metadata.innerText -split '\|')[1] -split ':')[1].Trim()).ToShortDateString()
                $row.CreatedDate = ([datetime](($metadata.innerText -split '\|')[2] -split ':')[1].Trim()).ToShortDateString()
                $row.LastModifiedDate = ([datetime](($metadata.innerText -split '\|')[3] -split ':')[1].Trim()).ToShortDateString()
            }
            else {
                $row.Published = ([datetime](($metadata.innerText -split '\|')[0] -split ':')[1].Trim()).ToShortDateString()
                $row.CreatedDate = ([datetime](($metadata.innerText -split '\|')[1] -split ':')[1].Trim()).ToShortDateString()
                $row.LastModifiedDate = ([datetime](($metadata.innerText -split '\|')[2] -split ':')[1].Trim()).ToShortDateString()
            }

            $result += $row
        }
    }

    End {
        # Quit (hidden) IE instance at exit
        if ($ie.Visible -ne $true) {
            $ie.Quit()
        }
        return $result
    }
}
