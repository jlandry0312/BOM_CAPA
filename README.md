# BOM_CAPA

Date: 6/25/2019
To-Do:
    - Consolidate BOMs into 1 BOM. Save as .xlms
    - Write VBA code to hide colums per dept
    - Write VBA code to Export to .csv / Find trigger
    - Investigate VDS powershell issue
    - Investivate Re-Indexing Vault Server
    - Email D3

    After research, I've made the decision to utilize a file dump folder on the network that Excel will export to, and SQL server will loop through to update data tables. Need to build an application to query SQL Server for information needed. Stay outside of querying within Excel. 