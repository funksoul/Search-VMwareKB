Function Wait-Document {
<#
.SYNOPSIS
    A helper function which waits for document loading.
.DESCRIPTION
    It's not supposed to be run as stand-alone
#>
    Param (
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)]$Timeout
    )

    Begin {
        $maxcount = $Timeout * 2
        $count = 0

        # Determine browser characteristics
        if (Get-Member -InputObject $Object.Document -MemberType Property -Name "readyState" -ErrorAction SilentlyContinue) {
            $readyStateExpression = '($Object.Document.readyState -eq "complete")'
        }
        else {
            $readyStateExpression = '(($Object.ReadyState -eq 4) -and ($Object.Busy -eq $false))'
        }
    }

    Process {
        # Wait for document loading
        do {
            Start-Sleep -Milliseconds 500
            $count++
        } until ((Invoke-Expression $readyStateExpression) -or ($count -ge $maxcount))

        if ((Invoke-Expression $readyStateExpression)) {
            return $true
        }
        else {
            return $false
        }
    }
}

Function Get-SortBy {
<#
.SYNOPSIS
    A helper function which sorts the search results by specific criteria.
.DESCRIPTION
    It's not supposed to be run as stand-alone
#>
    Param (
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)]$Criteria,
        [Parameter(Mandatory=$true)]$Timeout
    )
    Begin {
        Write-Verbose "< Sorting results by: $Criteria"

        # Read default sort criteria
        if (Get-Member -InputObject $Object.Document -MemberType Method -Name IHTMLDocument3_getElementById) {
            $element = $Object.Document.IHTMLDocument3_getElementById('sortBy')
        }
        else {
            $element = $Object.Document.getElementById('sortBy')
        }
        if (((Wait-Document -Object $Object -Timeout $Timeout) -eq $true) -and ($element -ne $null)) {
            $defaultCriteria = $element[$element.selectedIndex].text
        }
        else {
            Write-Verbose "> Request timeout. Please check website message."
            $Object.Visible = $true
            break
        }
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

        # Select sort criteria
        if ($itemsList.ContainsValue($Criteria)) {
            $element | %{
                if ($_.text.Trim() -eq $Criteria) {
                    $element.value = $_.value
                    $element.FireEvent('onchange') | Out-Null
                }
            }
            if ((Wait-Document -Object $Object -Timeout $Timeout) -eq $true) {
                Write-Verbose "> Sorting results finished successfully"
            }
            else {
                Write-Verbose "> Request timeout. Please check website message."
                $Object.Visible = $true
                break
            }
        }

        # Change to interactive mode if needed
        else {
            Write-Verbose "> Criteria `"$Criteria`" could not be found"
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
                if ((Wait-Document -Object $Object -Timeout $Timeout) -eq $true) {
                    Write-Verbose "> Sorting results finished successfully"
                }
                else {
                    Write-Verbose "> Request timeout. Please check website message."
                    $Object.Visible = $true
                    break
                }
            }
            else {
                Write-Host -ForegroundColor yellow "Empty or invalid choice. Select default criteria: $defaultCriteria"
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
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)]$focus,
        [Parameter(Mandatory=$true)]$focusItem,
        [Parameter(Mandatory=$true)]$Timeout
    )
    Begin {
        Write-Verbose "< Selecting $focus`: `"$focusItem`".."

        if (Get-Member -InputObject $Object.Document -MemberType Method -Name IHTMLDocument3_getElementsByName) {
            $idList = $Object.Document.IHTMLDocument3_getElementsByName('idList')
        }
        else {
            $idList = $Object.Document.getElementsByName('idList')
        }
        $table = ($idList | Select-Object -First 1).parentElement

        if ($table) {
            $narrowFocusTable = @()
            $table.getElementsByClassName('GS_bgcolor') | %{
                $narrowFocusTable += $_
            }
            Switch ($focus) {
                "Language" { $narrowFocusItems = $narrowFocusTable[2].getElementsByTagName('A'); break }
                "Category" { $narrowFocusItems = $narrowFocusTable[1].getElementsByTagName('A'); break }
                "Product" { $narrowFocusItems = $narrowFocusTable[0].getElementsByTagName('A'); break }
            }
        }
        else {
            Write-Verbose "> Cannot narrow focus by $focus."
        }
    }
    Process {
        if ($narrowFocusItems) {

            # Build focus list
            $itemsListArray = @()
            $itemsList = @{}
            $i = 0

            $narrowFocusItems | %{
                $itemsListArray += $_.innerTEXT.Trim()
            }
            $itemsListArray | Sort-Object | %{
                $key = $i++
                $itemsList[$key] = $_
            }

            # Select focus item
            if ($itemsList.ContainsValue($focusItem)) {
                $narrowFocusItems | %{
                    if ($_.innerTEXT) {
                        if ($_.innerTEXT.Trim() -eq $focusItem) {
                            $_.click()
                        }
                    }
                }
                if ((Wait-Document -Object $Object -Timeout $Timeout) -eq $true) {
                    Write-Verbose "> Selecting $focus finished successfully"
                }
                else {
                    Write-Verbose "> Request timeout. Please check website message."
                    $Object.Visible = $true
                    break
                }
            }

            # Change to interactive mode if needed
            else {
                Write-Verbose "> $focus `"$focusItem`" could not found"
                $itemsList.Keys | Sort-Object | %{ Write-Host $_":" $itemsList[$_] }
                Write-Host -NoNewline -ForegroundColor green "Please select $focus`: "
                $itemIndex = Read-Host
                $item = $itemsList[[int]$itemIndex]

                if ($itemIndex -and $item) {
                    Write-Verbose "< Selecting $focus`: `"$item`".."
                    $narrowFocusItems | %{
                        if ($_.innerTEXT) {
                            if ($_.innerTEXT.Trim() -eq $item) {
                                $_.click()
                            }
                        }
                    }
                    if ((Wait-Document -Object $Object -Timeout $Timeout) -eq $true) {
                        Write-Verbose "> Selecting $focus finished successfully"
                    }
                    else {
                        Write-Verbose "> Request timeout. Please check website message."
                        $Object.Visible = $true
                        break
                    }
                }
                else {
                    Write-Host -ForegroundColor yellow "Empty or invalid choice. Select all $focus"
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
.PARAMETER Interactive
    Run Cmdlet in interactive mode
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
    Version                     : 1.1
    Dependencies                : 
    ===Tested Against Environment====
    ESXi Version                : 
    PowerCLI Version            : 
    PowerShell Version          : 5.1.14393.693
#>

    Param (
        [Parameter(Mandatory=$true, Position=0)][String]$Keyword,
        [Parameter(Mandatory=$false)][switch]$Interactive = $false,
        [Parameter(Mandatory=$false)]$SortBy,
        [Parameter(Mandatory=$false)]$Language,
        [Parameter(Mandatory=$false)]$Category,
        [Parameter(Mandatory=$false)]$Product,
        [Parameter(Mandatory=$false)]$Timeout = 60
    )

    Begin {
        if ($Interactive) {
            $SortBy = $Language = $Category = $Product = '*'
        }

        $ie = New-Object -ComObject 'InternetExplorer.Application'
        $url = 'https://kb.vmware.com/selfservice/microsites/microsite.do'

        # Add KB Article Type
        Try {
            Add-Type @"
namespace InsData {
    public struct KBArticleSearchResult {
        public string Title;
        public string URL;
        public string Description;
        public string Rating;
        public string Published;
        public string CreatedDate;
        public string LastModifiedDate;
    }
}
"@
        }
        Catch {
            Write-Error $_
        }
    }

    Process {
        Write-Verbose "< Opening URL: $url"
        $ie.Navigate($url)

        if ((Wait-Document -Object $ie -Timeout $Timeout) -eq $true) {
            Write-Verbose "> URL opened successfully"
        }
        else {
            Write-Verbose "> Request timeout. Please check website message."
            $ie.Visible = $true
            break
        }

        # Search for a keyword
        Write-Verbose "< Searching for the keyword: `"$Keyword`""
        Try {
            if (Get-Member -InputObject $ie.Document -MemberType Method -Name IHTMLDocument3_getElementById) {
                $searchForm = $ie.Document.IHTMLDocument3_getElementById('id_searchForm')
            }
            else {
                $searchForm = $ie.Document.getElementById('id_searchForm')
            }
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

        if ((Wait-Document -Object $ie -Timeout $Timeout) -eq $true) {
            Write-Verbose "> Keyword search finished successfully"
        }
        else {
            Write-Verbose "> Request timeout. Please check website message."
            $ie.Visible = $true
            break
        }

        # Sort results
        if ($SortBy) {
            Get-SortBy -Object $ie -Criteria $SortBy -Timeout $Timeout
        }

        # Narrow Focus (Language)
        if ($Language) {
            Get-NarrowFocus -Object $ie -focus 'Language' -focusItem $Language -Timeout $Timeout
        }

        # Narrow Focus (Category)
        if ($Category) {
            Get-NarrowFocus -Object $ie -focus 'Category' -focusItem $Category -Timeout $Timeout
        }

        # Narrow Focus (Products)
        if ($Product) {
            Get-NarrowFocus -Object $ie -focus 'Product' -focusItem $Product -Timeout $Timeout
        }

        # Create array of sort results
        $result = @()
        if (Get-Member -InputObject $ie.Document -MemberType Method -Name IHTMLDocument3_getElementById) {
            $searchRes = $ie.Document.IHTMLDocument3_getElementById('searchres')
        }
        else {
            $searchRes = $ie.Document.getElementById('searchres')
        }
        $searchRes.getElementsByClassName('vmdoc') | %{
            $doctitleClass = $_.getElementsByClassName('doctitle') | Select-Object -First 1
            $doctitleClassATag = $doctitleClass.getElementsByTagName('A') | Select-Object -First 1
            $synopsisTag = $_.getElementsByTagName('synopsis') | Select-Object -First 1

            $row = New-Object -TypeName InsData.KBArticleSearchResult

            $row.Title = $doctitleClass.innerText.Trim()
            $row.URL = 'https://kb.vmware.com/kb/' + ($doctitleClassATag.href -split 'externalId=')[1] -replace '&.*',''
            $row.Description = $synopsisTag.innerText.Trim()
            $metadata = $_.getElementsByClassName('metadata') | Select-Object -First 1
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
        Write-Verbose "> Exiting normally.."
        if ($ie.Visible -ne $true) {
            $ie.Quit()
        }
        return $result
    }
}

Function Get-LinkItems {
    Param (
        [Parameter(Mandatory=$true)]$Element
    )

    Begin {
        $linkItems = @()
    }

    Process {
        $Element.getElementsByTagName('A') | %{
            $linkItem = "" | Select-Object "Title", "URL"
            $linkItem.Title = $_.innerText
            if ($_.href -like "javascript:openConsole*") {
                $linkItem.URL = "https://kb.vmware.com/selfservice/viewAttachment.do?attachID=" + ($_.href -split "`'")[3] + "&documentID=" + ($_.href -split "`'")[1]
            }
            else {
                $linkItem.URL = $_.href
            }
            $linkItems += $linkItem
        }
    }

    End {
        return $linkItems
    }
}

Function Get-KBArticle {
<#
.SYNOPSIS
    A PowerShell Cmdlet for fetching a VMware KB article on the command line.

.DESCRIPTION
    A PowerShell Cmdlet for fetching a VMware KB article on the command line.
    It uses Internet Explorer COM Object to interact with the VMware KB site
    and to extract DOM elements from the HTML Document.

    A fetched article is processed and converted to a PowerShell custom object
    which contains following properties:
        . Title
        . Symptoms, Purpose, Cause, Details, Resolution, Solution, Impact/Risks,
          Update History, Additional Information, Tags
        . See Also, Attachments
        . Rating Value, Rating Count, RSS URL
        . Updated, Categories, Languages, Product(s), Product Version(s), Language Editions
        . Links, RunAt

    You can save the object to a file to track changes of an article in detail.

.PARAMETER ArticleNumber
    A VMware KB article number
.PARAMETER Timeout
    Set timeout value of fetching HTML document, DOM element, etc.

.EXAMPLE
    Get-KBArticle 2144934
    Fetch VMware KB article 2144934

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
        [Parameter(Mandatory=$true, Position=0)]$ArticleNumber,
        [Parameter(Mandatory=$false)]$Timeout = 60
    )

    Begin {
        $ie = New-Object -ComObject 'InternetExplorer.Application'
        $url = "https://kb.vmware.com/kb/$ArticleNumber"

        # Add KB Article Type
        Try {
            Add-Type @"
namespace InsData {
    public struct KBArticle {
        public System.DateTime RunAt;
    }
}
"@
        }
        Catch {
            Write-Error $_
        }
    }

    Process {
        Write-Verbose "< Opening URL: $url"
        $ie.Navigate($url)

        if ((Wait-Document -Object $ie -Timeout $Timeout) -eq $true) {
            Write-Verbose "> URL opened successfully"
        }
        else {
            Write-Verbose "> Request timeout. Please check website message."
            $ie.Visible = $true
            break
        }

        # Get document elements
        Write-Verbose "< Getting document elements.."
        Try {
            if (Get-Member -InputObject $ie.Document -MemberType Method -Name IHTMLDocument3_getElementsByTagName) {
                $divs = $ie.Document.IHTMLDocument3_getElementsByTagName('DIV')
                $h4s = $ie.Document.IHTMLDocument3_getElementsByTagName('H4')
                $spans = $ie.Document.IHTMLDocument3_getElementsByTagName('SPAN')
                $metas = $ie.Document.IHTMLDocument3_getElementsByTagName('META')
            }
            else {
                $divs = $ie.Document.getElementsByTagName('DIV')
            }
            Write-Verbose "< Getting document elements finished successfully"
        }
        Catch {
            Write-Host "> $_"
            Write-Host "Unknown error occurred. Please check website message."
            $ie.Visible = $true
            break
        }

        $row = New-Object -TypeName InsData.KBArticle
        $Links = @{}

        # RunAt
        $row.RunAt = Get-Date

        # Title
        $title = $spans | Where-Object { $_.getAttribute('itemprop') -eq 'name' } | Select-Object -First 1 | %{ $_.parentNode.innerText.Trim() }
        $row | Add-Member -MemberType NoteProperty -Name "Title" -Value $title

        # Symptoms, Purpose, Cause, Details, Resolution, Solution, Impact/Risks, Update History, Additional Information, Tags
        $divs | ForEach-Object {
            $className = $_.className
            $div = $_
            $linkItems = $null

            Switch ($className) {
                "doccontent cc_Symptoms" {
                    $symptoms = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Symptoms" -Value $symptoms
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Symptoms"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Purpose" {
                    $purpose = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Purpose" -Value $purpose
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Purpose"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Cause" {
                    $cause = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Cause" -Value $cause
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Cause"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Details" {
                    $details = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Details" -Value $details
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Details"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Resolution" {
                    $resolution = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Resolution" -Value $resolution
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Resolution"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Solution" {
                    $solution = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Solution" -Value $solution
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Solution"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Impact/Risks" {
                    $impact_risks = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Impact/Risks" -Value $impact_risks
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Impact/Risks"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Update_History" {
                    $update_history = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Update History" -Value $update_history
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Update History"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Additional_Information" {
                    $additional_information = $div.innerText.Trim() -replace "`r`n`r`n","`r`n"
                    $row | Add-Member -MemberType NoteProperty -Name "Additional Information" -Value $additional_information
                    if ($linkItems = Get-LinkItems -Element $div) {
                        $Links["Additional Information"] = $linkItems
                    }
                    break
                }
                "doccontent cc_Tags" {
                    $tags = $div.innerText.Trim()
                    $row | Add-Member -MemberType NoteProperty -Name "Tags" -Value $tags
                    break
                }
            }
        }

        # See Also
        $h4 = $h4s | Where-Object { $_.innerText.Trim() -eq 'See Also' }
        if ($h4) {
            $see_also = $h4.nextSibling.innerText.Trim() -replace "`r`n`r`n","`r`n"
            $row | Add-Member -MemberType NoteProperty -Name "See Also" -Value $see_also
            if ($linkItems = Get-LinkItems -Element $h4.nextSibling) {
                $Links["See Also"] = $linkItems
            }
        }

        # Attachments
        $h4 = $h4s | Where-Object { $_.innerText.Trim() -eq 'Attachments' }
        if ($h4) {
            $attachments = $h4.nextSibling.innerText.Trim() -replace "`r`n`r`n","`r`n"
            $row | Add-Member -MemberType NoteProperty -Name "Attachments" -Value $attachments
            if ($linkItems = Get-LinkItems -Element $h4.nextSibling) {
                $Links["Attachments"] = $linkItems
            }
        }

        # Rating Value
        $row | Add-Member -MemberType NoteProperty -Name "Rating Value" -Value $null
        $row."Rating Value" = $metas | Where-Object { $_.getAttribute('itemprop') -eq 'ratingValue' } | Select-Object -First 1 | %{ $_.content }

        # Rating Count
        $row | Add-Member -MemberType NoteProperty -Name "Rating Count" -Value $null
        $row."Rating Count" = $spans | Where-Object { $_.getAttribute('itemprop') -eq 'ratingCount' } | Select-Object -First 1 | %{ $_.innerText.Trim() }

        # RSS URL
        $row | Add-Member -MemberType NoteProperty -Name "RSS URL" -Value $null
        $row."RSS URL" = ($ie.Document.IHTMLDocument3_getElementById('rssfeed-link')).href

        # Updated, Categories, Languages, Product(s), Product Version(s), Language Editions
        $spans | ForEach-Object {
            $id = $_.id
            $span = $_

            if ($span.innerText) {
                Switch ($id) {
                    "kbarticledate" {
                        $updated = $span.innerText.Trim()
                        $row | Add-Member -MemberType NoteProperty -Name "Updated" -Value $updated
                        break
                    }
                    "kbcategory" {
                        $categories = $span.innerText.Trim()
                        $row | Add-Member -MemberType NoteProperty -Name "Categories" -Value $categories
                        break
                    }
                    "kblanguages" {
                        $languages = $span.innerText.Trim()
                        $row | Add-Member -MemberType NoteProperty -Name "Languages" -Value $languages
                        break
                    }
                    "kbarticleproducts" {
                        $products = $span.innerText.Trim()
                        $row | Add-Member -MemberType NoteProperty -Name "Product(s)" -Value $products
                        break
                    }
                    "productversions" {
                        $product_versions = $span.innerText.Trim()
                        $row | Add-Member -MemberType NoteProperty -Name "Product Version(s)" -Value $product_versions
                        break
                    }
                    "langedition" {
                        $language_editions = $span.innerText.Trim()
                        $row | Add-Member -MemberType NoteProperty -Name "Language Editions" -Value $language_editions
                        break
                    }
                }
            }
        }

        # Links
        $row | Add-Member -MemberType NoteProperty -Name "Links" -Value $null
        $row.Links = $Links

        $row
    }

    End {
        # Quit (hidden) IE instance at exit
        Write-Verbose "> Exiting normally.."
        if ($ie.Visible -ne $true) {
            $ie.Quit()
        }
    }
}
