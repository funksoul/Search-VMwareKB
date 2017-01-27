# Search-VMwareKB

A PowerShell Module for searching VMware KB articles on the command line.
It uses Internet Explorer COM Object to interact with the VMware KB site.


You can search for a keyword, sort the results and narrow focus to specific language, category, product.
Just as VMware KB site, narrow focus conditions are dynamic.
(For example, if there's no article written in a language, you cannot narrow focus to it)

Search results are returned as a PowerShell Array which contains following properties:

- Title
- URL - https://kb.vmware.com/kb/[Article #]
- Description
- Rating - # of stars (if there)
- Published / Created Date / Last Modified Date - DateTime.ToShortDateString()




### Requirements

- Desktop edition of Windows PowerShell
- Internet Explorer / Edge




### Installation

1. Download repo as .zip file and extract it to your preferred PowerShell module path.
2. Remove extra suffix from the folder name. (-master, -devel, ..etc)
3. Check if the PowerShell recognizes the module properly.
4. Load the module (yay)

```powershell
PS C:\> $env:PSModulePath -split ';'
	C:\Users\Foo\Documents\WindowsPowerShell\Modules\Search-VMwareKB
PS C:\> Get-Module -ListAvailable Search-VMwareKB
PS C:\> Import-Module Search-VMwareKB
```



### Usage

- Search VMware KB site using the keyword 'PSOD'

  ```powershell
  PS C:\> Search-VMwareKB PSOD

  === Sample Output ===
  Title            : "PF Exception 14 in world 32868:helper11-0 IP 0x418008f10260" PSOD in ESXi 5.x or 6.0.x host (2114745)
  URL              : https://kb.vmware.com/kb/2114745
  Description      : change the Latency sensitivity of the virtual machine to normal to prevent any further occurrence of the PSOD. Note: The host failing with PSOD has the virtual machine configured for High Latency sensitivity. To change...
  Rating           : 5
  Published        : 2017-01-24
  CreatedDate      : 2015-04-21
  LastModifiedDate : 2017-01-24
  ```


- Search VMware KB site for most recently published article using the keyword 'no workaround' 

  ```powershell
  PS C:\> Search-VMwareKB -Keyword 'no workaround' -SortBy 'Publication Date'
  ```


- Use '*' as a parameter value if you don't know

  ```powershell
  PS C:\> Search-VMwareKB -Keyword 'no workaround' -SortBy * -Language * -Category * -Product *
  ```

- Open all KB articles at once in search results using default web browser

  ```powershell
  PS C:\> Search-VMwareKB -Keyword 'fails' | %{ start $_.URL }
  ```




### Etc

- There's no page navigation so the maximum number of search results will be 25.
- When something goes wrong, a browser window pops up to help you identify the problem.
  (ex. HTTP communication error, Capcha is required, ..etc)