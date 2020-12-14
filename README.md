JobParser
======

How to use
-------
### Unblock script
Open PowerShell and type

    > Unblock-File -Path C:\Path\to\script\jobparser_v3.ps1
    
This will remove a hidden parameter that indicates that the script was downloaded from the internet

### Run it from command line

1. If the files on a network drive

CMD does not support UNC paths (path to a folder or file on a network) as current directories. 
We can get around this by mapping a drive to the UNC path and referencing that network drive

    > net use Batches: \\path\on\network\drive\to\batches
    > Batches:

It will change our working directory to batches folder 
    Batches:\>

Now we can run the script

    Batches:\> ./jobparser_v3.ps1

or

    Batches:\> Batches:\jobparser_v3.ps1

2. If the files on local machine

    > cd C:\Path\to\files

    > ./jobparser_v3.ps1
or
    > C:\Path\to\files\jobparser_v3.ps1


### Supported commands

    > Batches:\jobparser_v3.ps1 split

Outputs 1 .csv file for every source .txt file it finds in the batch directory


    > Batches:\jobparser_v3.ps1 ondate YYYY-MM-DD

for example

    > Batches:\jobparser_v3.ps1 ondate 2020-12-10

Will output only jobs that Started on 2020-12-10.
Possible date formats DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD, YYYY/MM/DD

