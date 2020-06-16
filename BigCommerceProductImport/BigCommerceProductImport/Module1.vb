Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Collections
Imports System.Text.RegularExpressions
Module Module1
    Dim FileName As String
    Public sqlArray(30) As String
    Dim OutputFile As String
    'Public stPath As String = System.AppDomain.CurrentDomain.BaseDirectory()
    Public stPath As String = "e:\core\heltontools\test\"
    Dim FileExt As String
    Public stChannelAccount As String
    Public stFBAChannelAccount As String
    Public stCompanyList(50, 2) As String
    Public stFileType As String
    Dim bolImportError As Boolean = False
    Public stIsUseCurrentInStock As String = "False"
    Public stIsUseRetailPrice As String = "False"
    Public stUseProductNoAsSKU As String = "False"
    Public bolUpdateExistingProd As Boolean = False
    Public bolIsError As Boolean = False
    Public stODBCString As String
    Public bolIsClientServer As Boolean
    Public bolAddNewProducts As Boolean
    Sub Main(ByVal stfilename As String, ByVal stpath As String)
        Dim FileContents As String
        Dim stSQL As String
        Dim stErrorFile As String = stpath + "\ErrorLog.csv"



        If Right(stpath, 1) = "\" Then
            stErrorFile = stpath + "ErrorLog.csv"
        Else
            stErrorFile = stpath + "\ErrorLog.csv"
        End If

        'Reset the boolean error flag
        bolIsError = False


        'Make sure the dialog box is cleared.
        UpdateDialog("NULL")

        'Check if file is in DOS format
        If CheckIfDOS(stfilename) = False Then
            'First, convert the file to DOS format
            ConvUnix2Dos(stfilename)
        Else
            'Read source text file
            FileContents = FileGet(stfilename)
            FileExt = Right(stfilename, 4)
            OutputFile = Left(stfilename, Len(stfilename) - 4) + "_MW" + FileExt
            FileWrite(OutputFile, FileContents)
        End If


        BigCommerceImport(OutputFile, stODBCString)

        'Show completion message
        UpdateDialog("Done!")
    End Sub
    Sub BigCommerceImport(ByRef stFileName As String, ByRef stODBCString As String)
        Dim stSQL As String
        Dim stBCAccountNo As String
        Dim stExportfile As String = stPath + "TempBCProducts.csv"
        Dim Exportfile As StreamWriter
        Dim iRowCount As Integer = 1
        Dim iRowID As Integer = 1
        Dim iMaxRowID As Integer
        Dim stMasterProduct(30) As String
        Dim stNewProductFile As String = stPath + stChannelAccount + ".csv"



        If My.Computer.FileSystem.FileExists(stFileName) = True Then
            'Delete the export file if it exits
            If My.Computer.FileSystem.FileExists(stExportfile) = True Then
                My.Computer.FileSystem.DeleteFile(stExportfile)
            End If

            Exportfile = My.Computer.FileSystem.OpenTextFileWriter(stExportfile, True)
        End If


        'Get the BigCommerce account information
        stSQL = "SELECT ChannelAccountNo FROM ChannelAccounts WHERE ChannelNo=23 AND Description='" & stChannelAccount & "';"
        SQLTextQuery("S", stSQL, stODBCString, 1)
        stBCAccountNo = sqlArray(0)

        'Drop table if exists
        stSQL = "DROP TABLE IF EXISTS TempBCProducts;"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        'Create Temp Table
        'stSQL = "CREATE TABLE TempBCProducts( " +
        'Chr(34) + "ItemType" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductID" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductName" + Chr(34) + " VARCHAR(255), " +
        'Chr(34) + "ProductType" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductSKU" + Chr(34) + " VARCHAR(100), " +
        'Chr(34) + "BinPickingNumber" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "BrandName" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "OptionSet" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "OptionSetAlign" + Chr(34) + " VARCHAR(50), " +
        ' Chr(34) + "ProductDescription" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "Price" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "CostPrice" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "RetailPrice" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "SalePrice" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "FixedShippingCost" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "FreeShipping" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductWarranty" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductWeight" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductWidth" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductHeight" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductDepth" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "AllowPurchases?" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductVisible?" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductAvailability" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "TrackInventory" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "CurrentStockLevel" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "LowStockLevel" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "Category" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID1" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile1" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription1" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail1" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort1" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID2" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile2" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription2" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail2" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort2" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID3" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile3" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription3" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail3" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort3" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID4" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile4" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription4" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail4" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort4" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID5" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile5" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription5" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail5" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort5" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID6" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile6" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription6" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail6" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort6" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID7" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile7" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription7" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail7" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort7" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID8" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile8" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription8" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail8" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort8" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID9" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile9" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription9" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail9" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort9" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID10" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile10" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription10" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail10" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort10" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID11" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile11" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription11" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail11" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort11" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID12" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile12" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription12" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail12" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort12" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID13" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile13" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription13" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail13" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort13" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID14" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile14" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription14" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail14" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort14" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID15" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile15" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription15" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail15" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort15" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID16" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile16" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription16" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail16" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort16" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID17" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile17" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription17" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail17" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort17" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID18" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile18" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription18" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail18" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort18" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID19" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile19" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription19" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail19" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort19" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID20" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile20" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription20" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail20" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort20" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID21" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile21" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription21" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail21" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort21" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID22" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile22" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription22" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail22" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort22" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID23" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile23" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription23" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail23" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort23" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID24" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile24" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription24" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail24" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort24" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID25" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile25" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription25" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail25" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort25" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID26" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile26" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription26" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail26" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort26" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID27" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile27" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription27" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail27" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort27" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID28" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile28" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription28" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail28" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort28" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageID29" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageFile29" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageDescription29" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageIsThumbnail29" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductImageSort29" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "SearchKeywords" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "PageTitle" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "MetaKeywords" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "MetaDescription" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "MYOBAssetAcct" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "MYOBIncomeAcct" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "MYOBExpenseAcct" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductCondition" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "EventDateRequired?" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "EventDateName" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "EventDateIsLimited?" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "EventDateStartDate" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "EventDateEndDate" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "SortOrder" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductTaxClass" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductUPC/EAN" + Chr(34) + " VARCHAR(50)); " +
        'Chr(34) + "StopProcessingRules" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductURL" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "RedirectOldURL?" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSGlobalTradeItemNumber" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSManufacturerPartNumber" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSGender" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSAgeGroup" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSColor" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSSize" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSMaterial" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSPattern" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSItemGroupID" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSCategory" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "GPSEnabled" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "AvalaraProductTaxCode" + Chr(34) + " VARCHAR(50), " +
        'Chr(34) + "ProductCustomFields" + Chr(34) + " VARCHAR(50));"
        stSQL = "CREATE TABLE TempBCProducts( " +
        Chr(34) + "ItemType" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductID" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductName" + Chr(34) + " VARCHAR(255), " +
        Chr(34) + "ProductType" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductSKU" + Chr(34) + " VARCHAR(100), " +
        Chr(34) + "BinPickingNumber" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "BrandName" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "OptionSet" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "OptionSetAlign" + Chr(34) + " VARCHAR(50), " +
         Chr(34) + "Price" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "CostPrice" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "RetailPrice" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "SalePrice" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "FixedShippingCost" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "FreeShipping" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductWeight" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductWidth" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductHeight" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductDepth" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "AllowPurchases?" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductVisible?" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductAvailability" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "TrackInventory" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "CurrentStockLevel" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "LowStockLevel" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "Category" + Chr(34) + " VARCHAR(50), " +
        Chr(34) + "ProductUPC/EAN" + Chr(34) + " VARCHAR(50)); "

        SQLTextQuery("I", stSQL, stODBCString, 0)

        'Read in the file clean up the records and then export it to the TempBCProducts.csv

        UpdateDialog("Cleaning up Big Commerce text file.")
        For Each stLine In File.ReadLines(stFileName)
            If iRowCount = 1 Then
                Exportfile.WriteLine(stLine)
            Else
                'stRow = stLine.Split(",")
                ''Create the values string for inserting into the temp table
                'For count = 0 To stRow.Length - 1
                '    stISQL = stISQL + CleanString(stRow(count)) + "','"
                'Next

                ''Add the closing );
                'stISQL = stISQL + "');"
                'SQLTextQuery("I", stSQL + stISQL, stODBCString, 0)

                Exportfile.WriteLine(CleanString(stLine))


            End If
            iRowCount = iRowCount + 1
        Next
        Exportfile.Close()

        'Import the cleaned up file into TempBCProducts
        UpdateDialog("Importing BigCommerce products into a temporary table.")
        stSQL = "IMPORT TABLE TempBCProducts EXCLUSIVE FROM " + Chr(34) + stExportfile + Chr(34) + " ;"
        SQLTextQuery("I", stSQL, stODBCString, 0)

        'Delete the Header row
        stSQL = "DELETE FROM TempBCProducts WHERE ProductID='Product ID';"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        UpdateDialog("Adding Channergy data fields.")

        'Add a RowID field
        stSQL = "ALTER TABLE TempBCProducts ADD COLUMN RowID AUTOINC AT 1,ADD COLUMN ProductNo VARCHAR(100) AT 2,ADD COLUMN MasterCode VARCHAR(20) AT 3,ADD COLUMN Desc VARCHAR(255) AT 4,ADD COLUMN ChannelNo FLOAT DEFAULT 23, ADD COLUMN IsNewProduct BOOLEAN DEFAULT True,ADD COLUMN ChannelAccountNo FLOAT DEFAULT " + stBCAccountNo + ";"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        'Index fields
        stSQL = "CREATE INDEX IF NOT EXISTS idxSKU ON TempBCProducts(ProductSKU);"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        stSQL = "CREATE INDEX IF NOT EXISTS idxChannelAccount ON TempBCProducts(ChannelAccountNo);"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        stSQL = "CREATE INDEX IF NOT EXISTS idxProductNo ON TempBCProducts(ProductNo);"
        SQLTextQuery("D", stSQL, stODBCString, 0)


        'Update the ProductNo with the ProductNo from the related ChannelListings
        UpdateDialog("Updating ProductNo's for existing products.")
        stSQL = "UPDATE TempBCProducts T SET ProductNo=C.ProductNo, IsNewProduct=False FROM TempBCProducts T JOIN ChannelListings C ON T.ProductSKU=C.SKU AND T.ChannelAccountNo=C.ChannelAccountNo JOIN Products P ON C.ProductNo=P.ProductNo;"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        'Update the ProductNo with the ProductNo from the Products table.
        stSQL = "UPDATE TempBCProducts T SET ProductNo=P.ProductNo, IsNewProduct=False FROM TempBCProducts T JOIN Products P ON T.ProductSKU=P.ProductNo WHERE T.ProductNo='';"
        SQLTextQuery("D", stSQL, stODBCString, 0)



        'Get the max number of rows
        stSQL = "SELECT MAX(RowID) FROM TempBCProducts;"
        SQLTextQuery("S", stSQL, stODBCString, 1)

        If sqlArray(0) <> "0" Then
            UpdateDialog("Processing sub products.")
            iMaxRowID = CInt(sqlArray(0))
            'Set up progress bar
            Form1.ProgressBar1.Minimum = 1
            Form1.ProgressBar1.Maximum = iMaxRowID
            Form1.ProgressBar1.Visible = True
        Else
            iMaxRowID = 0
        End If

        'Loop through the table and update the fields
        Do While iRowID <= iMaxRowID
            'Get the core data
            stSQL = "SELECT ProductType, ProductName,BinPickingNumber,BrandName,Price,CostPrice,RetailPrice,SalePrice,FixedShippingCost,ProductWeight,ProductWidth,ProductHeight,ProductDepth,CurrentStockLevel,ProductImageFile1,ProductImageFile2 INTO Temp FROM TempBCProducts WHERE RowID=" + CStr(iRowID) + ";"
            SQLTextQuery("S", stSQL, stODBCString, 16)

            'If the product type is a P, then store it in the product array
            If sqlArray(0) = "P" Then
                Array.Copy(sqlArray, stMasterProduct, stMasterProduct.Length)
                stSQL = "UPDATE TempBCProducts SET ProductNo=IF(ProductNo='',UPPER(ProductSKU),ProductNo),"
                stSQL = stSQL + "Desc ='" + stMasterProduct(1) + "' "
                stSQL = stSQL + "WHERE RowID=" + CStr(iRowID) + ";"
                SQLTextQuery("U", stSQL, stODBCString, 0)
            Else 'Update the sub product fields
                UpdateDialog("Updating sub products for " + stMasterProduct(1))
                stSQL = "UPDATE TempBCProducts SET "
                'Update for BH Moto
                'stSQL = stSQL + "MasterCode=IF(MasterCode='',UPPER(LEFT(ProductSKU,6)),MasterCode),"
                stSQL = stSQL + "MasterCode=IF(MasterCode='',UPPER(LEFT(ProductSKU,POS('_' IN ProductSKU)-1)),MasterCode),"
                stSQL = stSQL + "ProductNo=IF(ProductNo='',UPPER(ProductSKU),ProductNo),"
                stSQL = stSQL + "Desc='" + stMasterProduct(1) + "',"
                stSQL = stSQL + "BinPickingNumber='" + stMasterProduct(2) + "',"
                stSQL = stSQL + "BrandName='" + stMasterProduct(3) + "',"
                'stSQL = stSQL + "ProductDescription='" + stMasterProduct(4) + "',"
                stSQL = stSQL + "Price='" + stMasterProduct(4) + "',"
                stSQL = stSQL + "CostPrice='" + stMasterProduct(5) + "',"
                stSQL = stSQL + "RetailPrice='" + stMasterProduct(6) + "',"
                stSQL = stSQL + "SalePrice='" + stMasterProduct(7) + "',"
                stSQL = stSQL + "FixedShippingCost='" + stMasterProduct(8) + "',"
                stSQL = stSQL + "ProductWeight='" + stMasterProduct(9) + "',"
                stSQL = stSQL + "ProductWidth='" + stMasterProduct(10) + "',"
                stSQL = stSQL + "ProductHeight='" + stMasterProduct(11) + "',"
                stSQL = stSQL + "ProductDepth='" + stMasterProduct(12) + "',"
                stSQL = stSQL + "ProductImageFile1='" + stMasterProduct(16) + "',"
                stSQL = stSQL + "ProductImageFile2='" + stMasterProduct(17) + "' WHERE RowID=" + CStr(iRowID) + ";"
                SQLTextQuery("U", stSQL, stODBCString)

            End If
            Form1.ProgressBar1.Value = iRowID
            iRowID = iRowID + 1
        Loop

        'If the user opted to update existing products in Channergy then do so here
        If bolUpdateExistingProd = True Then
            UpdateDialog("Updating existing product information.")
            stSQL = "UPDATE Products P SET "
            stSQL = stSQL + "Description=T.Desc,"
            stSQL = stSQL + "ManufacturerID=T.BrandName,"
            'stSQL = stSQL + "LongDesc=CAST(T.ProductDescription AS MEMO),"
            stSQL = stSQL + "Price=CAST(T.Price AS FLOAT),"
            stSQL = stSQL + "Cost=CAST(T.CostPrice AS FLOAT),"
            stSQL = stSQL + "Price3=CAST(T.RetailPrice AS FLOAT),"
            stSQL = stSQL + "ShipCost=CAST(T.FixedShippingCost AS FLOAT),"
            stSQL = stSQL + "Weight=CAST(T.ProductWeight AS FLOAT),"
            stSQL = stSQL + "Width=CAST(T.ProductWidth AS FLOAT),"
            stSQL = stSQL + "Height=CAST(T.ProductHeight AS FLOAT),"
            stSQL = stSQL + Chr(34) + "Length" + Chr(34) + "=CAST(T.ProductDepth AS FLOAT),"
            stSQL = stSQL + "ImageURLLarge=T.ProductImageFile1,"
            stSQL = stSQL + "ImageURLSmall=T.ProductImageFile2, "
            stSQL = stSQL + "Serial=T.MasterCode, "
            stSQL = stSQL + "FROM Products P JOIN TempBCProducts T ON P.ProductNo=T.ProductNo;"
            SQLTextQuery("U", stSQL, stODBCString, 0)

            UpdateDialog("Updating existing channel listings.")
            stSQL = "UPDATE ChannelListings C SET Price=CAST(T.Price AS FLOAT),InStock=CAST(T.CurrentStockLevel AS FLOAT),MasterSKU=T.MasterCode FROM ChannelListings C JOIN TempBCProducts T ON C.SKU=T.ProductSKU AND C.ChannelAccountNo=T.ChannelAccountNo;"
            SQLTextQuery("U", stSQL, stODBCString, 0)

        End If

        'Update the basic products
        If bolAddNewProducts = False Then
            stSQL = "UPDATE TempBCProducts T SET ProductNo=P.ProductNo,IsNewProduct=False FROM TempBCProducts T JOIN Products P ON T.ProductSKU=P.ProductNo WHERE T.ProductNo='';"
            SQLTextQuery("U", stSQL, stODBCString)
        Else
            stSQL = "UPDATE TempBCProducts SET ProductNo=ProductSKU,IsNewProduct=True  WHERE ProductNo='';"
            SQLTextQuery("U", stSQL, stODBCString)
        End If
  

        'If the user does not want to add new Products to the product table then export the data out
        If bolAddNewProducts = False Then
            'Check to see if there are any new products
            stSQL = "SELECT COUNT(ProductNo) FROM TempBCProducts WHERE IsNewProduct=True;"
            SQLTextQuery("S", stSQL, stODBCString, 1)

            If sqlArray(0) <> "0" And sqlArray(0) <> "" Then
                stSQL = "SELECT * INTO NewBCProducts FROM TempBCProducts WHERE IsNewProduct=True;"
                SQLTextQuery("I", stSQL, stODBCString, 0)

                'Drop the Desc and Master code from NewBCProducts.
                stSQL = "ALTER TABLE NewBCProducts DROP COLUMN MasterCode,DROP COLUMN Desc;"
                SQLTextQuery("D", stSQL, stODBCString, 0)

                'Update the ProductNo in the NewBCProducts
                stSQL = "UPDATE NewBCProducts SET ProductNo='';"
                SQLTextQuery("I", stSQL, stODBCString, 0)

                'Export file
                stSQL = "EXPORT TABLE NewBCProducts TO " + Chr(34) + stNewProductFile + Chr(34) + " WITH HEADERS;"
                SQLTextQuery("U", stSQL, stODBCString, 0)
                UpdateDialog("Unmatched products exported to " + stNewProductFile)

            End If
        Else

            UpdateDialog("Adding new products to the products table.")

            'Add the new products
            stSQL = "INSERT INTO Products(ProductNo) SELECT DISTINCT ProductNo FROM TempBCProducts T LEFT OUTER JOIN Products P ON T.ProductNo=P.ProductNo WHERE IsNewProduct=True AND P.ProductNo='' AND T.Desc<>'';"
            SQLTextQuery("U", stSQL, stODBCString, 0)

            stSQL = "UPDATE Products P SET "
            stSQL = stSQL + "Description=T.Desc,"
            stSQL = stSQL + "ManufacturerID=T.BrandName,"
            stSQL = stSQL + "LongDesc=CAST(T.ProductDescription AS MEMO),"
            stSQL = stSQL + "Price=CAST(T.Price AS FLOAT),"
            stSQL = stSQL + "Cost=CAST(T.CostPrice AS FLOAT),"
            stSQL = stSQL + "Price3=CAST(T.RetailPrice AS FLOAT),"
            stSQL = stSQL + "ShipCost=CAST(T.FixedShippingCost AS FLOAT),"
            stSQL = stSQL + "Weight=CAST(T.ProductWeight AS FLOAT),"
            stSQL = stSQL + "Width=CAST(T.ProductWidth AS FLOAT),"
            stSQL = stSQL + "Height=CAST(T.ProductHeight AS FLOAT),"
            stSQL = stSQL + Chr(34) + "Length" + Chr(34) + "=CAST(T.ProductDepth AS FLOAT),"
            stSQL = stSQL + "ImageURLLarge=T.ProductImageFile1,"
            stSQL = stSQL + "ImageURLSmall=T.ProductImageFile2 "
            stSQL = stSQL + "FROM Products P JOIN TempBCProducts T ON P.ProductNo=T.ProductNo;"
            SQLTextQuery("U", stSQL, stODBCString, 0)

            'ParseBCSubProd()

        End If


        'Delete the existing Channellistings for matching products
        UpdateDialog("Deleting " + stChannelAccount + " listings already in Channergy from the temporary table.")

        stSQL = "DELETE FROM TempBCProducts T JOIN ChannelListings C ON T.ProductSKU=C.SKU AND T.ChannelAccountNo=C.ChannelAccountNo;"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Add the new Channel Listings
        UpdateDialog("Adding new " + stChannelAccount + " listings.")

        If bolAddNewProducts = False Then
            stSQL = "INSERT INTO ChannelListings("
            stSQL = stSQL + "ChannelNo,"
            stSQL = stSQL + "ChannelAccountNo,"
            stSQL = stSQL + "ProductNo,"
            stSQL = stSQL + "SKU,"
            stSQL = stSQL + "Price,"
            stSQL = stSQL + "IsActive,"
            stSQL = stSQL + "IsUseProductNoAsSKU,"
            stSQL = stSQL + "InStock,"
            stSQL = stSQL + "IsUseCurrentInStock,"
            stSQL = stSQL + "IsUseCurrentRetailPrice,"
            stSQL = stSQL + "MasterSKU) "
            stSQL = stSQL + "SELECT "
            stSQL = stSQL + "ChannelNo,"
            stSQL = stSQL + "ChannelAccountNo,"
            stSQL = stSQL + "ProductNo,"
            stSQL = stSQL + "ProductSKU,"
            stSQL = stSQL + "CAST(Price AS MONEY),"
            stSQL = stSQL + "True,"
            stSQL = stSQL + stUseProductNoAsSKU + ","
            stSQL = stSQL + "CAST(CurrentStockLevel AS FLOAT),"
            stSQL = stSQL + stIsUseCurrentInStock + ","
            stSQL = stSQL + stIsUseRetailPrice + ","
            'Code Added For Blanc Hills Moto
            'stSQL = stSQL + "LEFT(ProductSKU,8)"
            stSQL = stSQL + "MasterCode"
            stSQL = stSQL + " FROM TempBCProducts WHERE ProductNo<>'' AND IsNewProduct=False;"
            SQLTextQuery("I", stSQL, stODBCString, 0)
        Else
            stSQL = "INSERT INTO ChannelListings("
            stSQL = stSQL + "ChannelNo,"
            stSQL = stSQL + "ChannelAccountNo,"
            stSQL = stSQL + "ProductNo,"
            stSQL = stSQL + "SKU,"
            stSQL = stSQL + "Price,"
            stSQL = stSQL + "IsActive,"
            stSQL = stSQL + "IsUseProductNoAsSKU,"
            stSQL = stSQL + "InStock,"
            stSQL = stSQL + "IsUseCurrentInStock,"
            stSQL = stSQL + "IsUseCurrentRetailPrice,"
            stSQL = stSQL + "MasterSKU) "
            stSQL = stSQL + "SELECT "
            stSQL = stSQL + "ChannelNo,"
            stSQL = stSQL + "ChannelAccountNo,"
            stSQL = stSQL + "ProductNo,"
            stSQL = stSQL + "ProductSKU,"
            stSQL = stSQL + "CAST(Price AS MONEY),"
            stSQL = stSQL + "True,"
            stSQL = stSQL + stUseProductNoAsSKU + ","
            stSQL = stSQL + "CAST(CurrentStockLevel AS FLOAT),"
            stSQL = stSQL + stIsUseCurrentInStock + ","
            stSQL = stSQL + stIsUseRetailPrice + ","
            'Code Added For Blanc Hills Moto
            'stSQL = stSQL + "LEFT(ProductSKU,8)"
            stSQL = stSQL + "MasterCode"
            stSQL = stSQL + " FROM TempBCProducts WHERE ProductNo<>'' AND IsNewProduct=True;"
            SQLTextQuery("I", stSQL, stODBCString, 0)
        End If
 
        UpdateDialog("BigCommerce product channel listings for " + stChannelAccount + " created.")


    End Sub
    Sub ParseBCSubProd()
        Dim stSQL As String
        Dim stProductName() As String
        Dim stAttributes() As String
        Dim iCounter As Integer = 1
        Dim iMaxCount As Integer
        Dim stFind As String = "\:.*$"
        Dim stReplace As String = ""
        Dim stResult As String
        Dim stProductNo As String
        Dim stMasterCode As String


        'Grab all of the products that have MasterCodes
        stSQL = "SELECT ProductNo,MasterCode,ProductName INTO TempBCSubProd FROM TempBCProducts WHERE MasterCode<>'';"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Create the temp table for dumping the parsed attributes pairs
        stSQL = "DROP TABLE IF EXISTS TempBCSubProducts;"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        stSQL = "CREATE TABLE TempBCSubProducts(ProductNo VARCHAR(100),MasterCode VARCHAR(20),Option VARCHAR(30),Value VARCHAR(40));"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Add the fields to the table
        stSQL = "ALTER TABLE TempBCSubProd ADD COLUMN RowID AUTOINC AT 1,ADD COLUMN CustomDesc VARCHAR(255);"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Create Indexes
        stSQL = "CREATE INDEX IF NOT EXISTS idxProductOptionValue ON TempBCSubProducts(ProductNo,Option,Value);"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        stSQL = "CREATE INDEX IF NOT EXISTS idxProducts ON TempBCSubProducts(ProductNo);"
        SQLTextQuery("U", stSQL, stODBCString, 0)


        'Get the iMaxCount
        stSQL = "SELECT MAX(RowID) FROM TempBCSubProd;"
        SQLTextQuery("S", stSQL, stODBCString, 1)

        If sqlArray(0) <> "0" And sqlArray(0) <> "" Then
            iMaxCount = CInt(sqlArray(0))
        Else
            iMaxCount = 0
        End If

        Do While iCounter <= iMaxCount
            'Get the ProductName
            stSQL = "SELECT ProductName,ProductNo,MasterCode FROM TempBCSubProd WHERE RowID=" + CStr(iCounter) + ";"
            SQLTextQuery("S", stSQL, stODBCString, 3)
            stProductNo = sqlArray(1)
            stMasterCode = sqlArray(2)
            UpdateDialog("Creating Subproducts for " + stProductNo)

            'Strip the trailing color information
            Dim rgx As New Regex(stFind)
            If rgx.IsMatch(sqlArray(0)) = True Then
                stResult = rgx.Replace(sqlArray(0), stReplace)
            Else
                stResult = sqlArray(0)
            End If

            'Split the options
            stProductName = stResult.Split(",")

            For Each stAttribute In stProductName
                stAttributes = stAttribute.Split("=")
                stSQL = "INSERT INTO TempBCSubProducts(ProductNo,MasterCode,Option,Value) "
                stSQL = stSQL + "VALUES('" + stProductNo + "','" + stMasterCode + "',UPPER('" + stAttributes(0) + "'),UPPER('" + stAttributes(1) + "'));"
                SQLTextQuery("I", stSQL, stODBCString, 0)

                'Update the CustomDesc in TempBCSubProd
                stSQL = "UPDATE TempBCSubProd SET CustomDesc=IF(CustomDesc='',UPPER('" + stAttributes(0) + "')+':'+UPPER('" + stAttributes(1) + "'),CustomDesc+ #13+#10 +UPPER('" + stAttributes(0) + "')+':'+UPPER('" + stAttributes(1) + "')) WHERE ProductNo='" + stProductNo + "';"
                SQLTextQuery("I", stSQL, stODBCString, 0)
            Next
            iCounter = iCounter + 1
        Loop

        'Delete products that are already in the Subproducts table
        UpdateDialog("Deleting exising subproducts from the temp table.")
        stSQL = "DELETE FROM TempBCSubProducts T JOIN SubProducts S ON T.ProductNo=S.ProductNo AND T.Option=S.Option AND T.Value=S.Value;"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        'Add the new Subproducts
        UpdateDialog("Adding new subproducts to Channergy.")
        stSQL = "INSERT INTO SubProducts(ProductNo,MasterCode,Option,Value) "
        stSQL = stSQL + "SELECT ProductNo,MasterCode,Option,Value FROM TempBCSubProducts;"
        SQLTextQuery("U", stSQL, stODBCString, 0)


        'Update the Serial number in the products table
        UpdateDialog("Updating the Master Code and Custom Description in Channergy.")
        stSQL = "UPDATE Products P SET Serial=T.MasterCode,CustomDesc=CAST(T.CustomDesc AS MEMO) FROM Products P JOIN TempBCSubProd T ON P.ProductNo=T.ProductNo;"
        SQLTextQuery("U", stSQL, stODBCString, 0)

    End Sub
    Sub ConvUnix2Dos(ByVal stFileName As String)
        Dim strFind As String
        Dim strReplace As String
        Dim FileContents As String
        Dim dFileContents As String
        Dim FileExt As String

        FileExt = Right(stFileName, 4)

        OutputFile = Left(stFileName, Len(stFileName) - 4) + "_MW" + FileExt

        strFind = Chr(10)
        strReplace = Chr(13) + Chr(10)
        'Read source text file
        FileContents = FileGet(stFileName)

        'replace all string In the source file
        dFileContents = Replace(FileContents, strFind, strReplace, 1, -1, 1)

        'Compare source And result
        If dFileContents <> FileContents Then
            'write result If different
            FileWrite(OutputFile, dFileContents)
            UpdateDialog("New file saved as " + OutputFile)
            'MsgBox("New file saved as " + OutputFile)
        End If

    End Sub
    Sub LoadForm()
        Dim Count As Integer = 1
        Dim stSQL As String
        Dim Max As Integer
        Dim ChannelAccount As String

        'Clear out existing combo boxes
        Form1.ComboBox1.Items.Clear()
        Form1.ComboBox1.Refresh()


        stSQL = "SELECT Description FROM ChannelAccounts WHERE ChannelNo=23;"

        LoadCombo2(stSQL)

        'Get Number of rows 
        stSQL = "SELECT COUNT(Temp) FROM Temp;"
        SQLTextQuery("S", stSQL, stODBCString, 1)
        If sqlArray(0) = "" Then
            Max = 0
        Else
            Max = CInt(sqlArray(0))
        End If

        If Max > 0 Then


            Do While Count <= Max
                'Get item to add
                stSQL = "SELECT Temp FROM Temp WHERE Mark=False TOP 1;"
                SQLTextQuery("S", stSQL, stODBCString, 1)
                ChannelAccount = sqlArray(0)
                Form1.ComboBox1.Items.Add(ChannelAccount)

                'Update Mark flag
                stSQL = "UPDATE Temp SET Mark=True WHERE Temp='" & ChannelAccount & "';"
                SQLTextQuery("U", stSQL, stODBCString, 0)

                'Update counter
                Count = Count + 1
            Loop

        Else
            MessageBox.Show("No channel account for Big Commerce has been set up in Channergy. Please set up your Big Commerce account in Tools->Preferences before continuing.", "No Big Commerce Account", MessageBoxButtons.OK)
        End If
    End Sub
    Sub LoadCombo2(ByVal SelectSQL As String)
        Dim stSQL As String

        'Drop temp table if it exists
        stSQL = "DROP TABLE IF EXISTS Temp;"
        SQLTextQuery("I", stSQL, stODBCString, 0)

        'Create temp table 
        stSQL = "CREATE TABLE Temp (Temp VARCHAR(255),Mark BOOLEAN DEFAULT False);"
        SQLTextQuery("I", stSQL, stODBCString, 0)

        'Empty Table
        stSQL = "DELETE FROM Temp;"
        SQLTextQuery("D", stSQL, stODBCString, 0)

        'Append to table
        stSQL = "INSERT INTO Temp(Temp) " & SelectSQL
        SQLTextQuery("I", stSQL, stODBCString, 0)



    End Sub
    Function CleanString(ByRef stString) As String
        Dim stFind As String = Chr(34)
        Dim stReplace As String = ""
        Dim stResult As String
        Dim stResult1 As String
        Dim stResult2 As String
        Dim stResult3 As String


        'Replace html tags
        stFind = "<[^>]*>"
        Dim rgx As New Regex(stFind)
        If rgx.IsMatch(stString) = True Then
            stResult = rgx.Replace(stString, stReplace)
        Else
            stResult = stString
        End If

        'Remove []
        stFind = "\[[^\]]*\]"
        Dim rgx1 As New Regex(stFind)
        If rgx1.IsMatch(stResult) = True Then
            stResult1 = rgx1.Replace(stResult, stReplace)
        Else
            stResult1 = stResult
        End If

        'stResult1 = stResult

        'Replace single quotes
        stFind = "'"
        Dim rgx2 As New Regex(stFind)
        If rgx2.IsMatch(stResult1) = True Then
            stResult2 = rgx2.Replace(stResult1, stReplace)
        Else
            stResult2 = stResult1
        End If


        'Replace double quotes
        stFind = Chr(34)

        Dim rgx3 As New Regex(stFind)
        If rgx3.IsMatch(stResult2) = True Then
            stResult3 = rgx3.Replace(stResult2, stReplace)
        Else
            stResult3 = stResult2
        End If
        Return stResult3

    End Function
    Function CheckIfDOS(ByVal stFileName As String) As Boolean
        Dim strFind As String = Chr(13) + Chr(10)
        Dim FileContents As String

        'Read source text file
        FileContents = FileGet(stFileName)

        'Find /r
        Dim FirstCharacter As Integer = FileContents.IndexOf(strFind)

        'If the <CR> character is found we can assume that the file has been converted to DOS format
        If FirstCharacter > 0 Then
            Return True
        Else
            Return False
        End If


    End Function
    'Read text file

    Function FileGet(ByVal FileName)
        If Len(FileName) > 0 Then
            Dim FS, FileStream
            FS = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            FileStream = FS.OpenTextFile(FileName)
            FileGet = FileStream.ReadAll
        End If
    End Function

    'Write string As a text file.
    Function FileWrite(ByVal FileName, ByVal Contents)
        Dim OutStream, FS, Dummy

        Dummy = Left(FileName, Len(FileName) - 4) & ".bak"
        On Error Resume Next
        FS = CreateObject("Scripting.FileSystemObject")
        OutStream = FS.OpenTextFile(Dummy, 2, True)
        OutStream.Write(Contents)
        FS.Close()
        FileCopy(Dummy, FileName)


    End Function
    Function GetDataPath(ByRef stExePath As String) As String
        Dim stIniFilePath As String
        Dim stIniFile = New IniFile()
        Dim stDatapath As String
        Dim virtualFilePath As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VirtualStore\Program Files (x86)\")
        Dim stVDataPath As String
        Dim stProgram As String

        'Get the version path and append to the virtualFilePath
        If Right(stExePath, 1) = "\" Then
            stProgram = Left(stExePath, Len(stExePath) - 1)
        Else
            stProgram = stExePath
        End If
        stProgram = Right(stProgram, Len(stProgram) - InStrRev(stProgram, "\"))
        virtualFilePath = virtualFilePath + stProgram


        'Check to see if the AppData folder contains the ini files
        If My.Computer.FileSystem.DirectoryExists(virtualFilePath) = True Then
            stVDataPath = virtualFilePath
        Else
            stVDataPath = stExePath
        End If


        'Fix path if it does not have a trailing \
        If Right(stVDataPath, 1) <> "\" Then
            stVDataPath = stVDataPath + "\"
        End If


        'Check to see if they are using Channergy or Mailware
        If My.Computer.FileSystem.FileExists(stVDataPath + "MAILWARE.INI") Then
            stIniFilePath = stVDataPath + "MAILWARE.INI"
        ElseIf My.Computer.FileSystem.FileExists(stVDataPath + "CHANNERGY.INI") Then
            stIniFilePath = stVDataPath + "CHANNERGY.INI"
        Else
            stIniFilePath = ""
        End If

        If stIniFilePath <> "" Then
            stIniFile.Load(stIniFilePath)
            stDatapath = stIniFile.GetKeyValue("Network", "NetworkDataDirectory")
            'Fix path if it does not have a trailing \
            If Right(stDatapath, 1) <> "\" Then
                stDatapath = stDatapath + "\"
            End If
        Else
            stDatapath = ""
        End If
        Return stDatapath

    End Function
    Function GetDSN(ByRef stDataPath As String) As String
        Dim stDSN As String
        Dim stIniFile = New IniFile()
        Dim stIniFilePath As String = stDataPath + "clientserver.ini"
        Dim stIpAddress As String
        Dim stCatalog As String
        Dim builder As New OdbcConnectionStringBuilder()

        builder.Driver = "DBISAM 4 ODBC Driver"

        'Check to see if the clientserver.ini file is in the passed path
        If My.Computer.FileSystem.FileExists(stIniFilePath) = True Then
            stIniFile.Load(stIniFilePath)
            stIpAddress = stIniFile.GetKeyValue("Settings", "IPAddress")
            stCatalog = stIniFile.GetKeyValue("Settings", "RemoteDatabase")
            builder.Add("UID", "Admin")
            builder.Add("PWD", "DBAdmin")
            builder.Add("ConnectionType", "Remote")
            builder.Add("RemoteIPAddress", stIpAddress)
            builder.Add("CatalogName", stCatalog)
            stDSN = builder.ConnectionString
            bolIsClientServer = True
        ElseIf My.Computer.FileSystem.FileExists(stDataPath + "Version.dat") = True Then
            builder.Add("ConnectionType", "Local")
            builder.Add("CatalogName", stDataPath)
            stDSN = builder.ConnectionString
        Else
            stDSN = ""
        End If

        Return stDSN
    End Function
    Sub SQLTextQuery(ByVal QueryType As String, ByVal CommandText As String, ByVal stODBC As String, Optional ByVal Columns As Integer = 0)
        'Dim DBCString As String = "Driver={C:\dbisam\odbc\std\ver4\lib\dbodbc\dbodbc.dll};connectiontype=Local;remoteipaddress=127.0.0.1;RemotePort=12005;remotereadahead=50;catalogname=" + stODBCString + ";readonly=False;lockretrycount=15;lockwaittime=100;forcebufferflush=False;strictchangedetection=False;"

        'Dim DBC As New System.Data.Odbc.OdbcConnection
        'DBC.ConnectionString = DBCString

        Dim DBC As New OdbcConnection(stODBC)
        'Dim DBCString As String = "Dsn=" & stODBC & ";"

        'Dim DBC As New OdbcConnection(DBCString)

        If QueryType = "S" And Columns > 0 Then
            Try
                Dim SQL1 As New OdbcCommand
                SQL1.Connection = DBC
                SQL1.CommandType = CommandType.Text
                SQL1.CommandText = CommandText
                DBC.Open()

                Dim DataRow As OdbcDataReader
                DataRow = SQL1.ExecuteReader()
                DataRow.Read()
                If DataRow.HasRows Then
                    Dim Counter As Integer = 0
                    While Counter < Columns
                        sqlArray(Counter) = DataRow(Counter).ToString
                        Counter = Counter + 1
                    End While
                Else
                    sqlArray(0) = "NoData"
                End If
                DataRow.Close()
                DBC.Close()
                SQL1.Dispose()

            Catch ex As Exception
                MessageBox.Show("SQL2 Exception Message: " & ex.Message)
                UpdateDialog("SQL1 Exception Message: " & ex.Message)
                bolIsError = True
            End Try
        End If
        If QueryType = "U" Or QueryType = "I" Or QueryType = "D" Then
            Try
                DBC.Open()
                Dim SQL2 As New OdbcCommand
                SQL2.Connection = DBC
                SQL2.CommandType = CommandType.Text
                SQL2.CommandTimeout = 60
                SQL2.CommandText = CommandText
                SQL2.ExecuteScalar()
                DBC.Close()
                SQL2.Dispose()

            Catch ex As Exception
                MessageBox.Show("SQL2 Exception Message: " & ex.Message)
                UpdateDialog("SQL1 Exception Message: " & ex.Message)
                bolIsError = True
            End Try
        End If
    End Sub
    Sub UpdateDialog(ByVal stMessage)
        Dim CRLF As String = Chr(13) + Chr(10)
        If stMessage = "NULL" Then
            Form1.TextBox2.Clear()
        Else
            Form1.TextBox2.AppendText(TimeString + ": " + stMessage + CRLF)
        End If
    End Sub
    ' IniFile class used to read and write ini files by loading the file into memory
    Public Class IniFile
        ' List of IniSection objects keeps track of all the sections in the INI file
        Private m_sections As Hashtable

        ' Public constructor
        Public Sub New()
            m_sections = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
        End Sub

        ' Loads the Reads the data in the ini file into the IniFile object
        Public Sub Load(ByVal sFileName As String, Optional ByVal bMerge As Boolean = False)
            If Not bMerge Then
                RemoveAllSections()
            End If
            '  Clear the object... 
            Dim tempsection As IniSection = Nothing
            Dim oReader As New StreamReader(sFileName)
            Dim regexcomment As New Regex("^([\s]*#.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            ' Broken but left for history
            'Dim regexsection As New Regex("\[[\s]*([^\[\s].*[^\s\]])[\s]*\]", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim regexsection As New Regex("^[\s]*\[[\s]*([^\[\s].*[^\s\]])[\s]*\][\s]*$", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim regexkey As New Regex("^\s*([^=\s]*)[^=]*=(.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            While Not oReader.EndOfStream
                Dim line As String = oReader.ReadLine()
                If line <> String.Empty Then
                    Dim m As Match = Nothing
                    If regexcomment.Match(line).Success Then
                        m = regexcomment.Match(line)
                        Trace.WriteLine(String.Format("Skipping Comment: {0}", m.Groups(0).Value))
                    ElseIf regexsection.Match(line).Success Then
                        m = regexsection.Match(line)
                        Trace.WriteLine(String.Format("Adding section [{0}]", m.Groups(1).Value))
                        tempsection = AddSection(m.Groups(1).Value)
                    ElseIf regexkey.Match(line).Success AndAlso tempsection IsNot Nothing Then
                        m = regexkey.Match(line)
                        Trace.WriteLine(String.Format("Adding Key [{0}]=[{1}]", m.Groups(1).Value, m.Groups(2).Value))
                        tempsection.AddKey(m.Groups(1).Value).Value = m.Groups(2).Value
                    ElseIf tempsection IsNot Nothing Then
                        '  Handle Key without value
                        Trace.WriteLine(String.Format("Adding Key [{0}]", line))
                        tempsection.AddKey(line)
                    Else
                        '  This should not occur unless the tempsection is not created yet...
                        Trace.WriteLine(String.Format("Skipping unknown type of data: {0}", line))
                    End If
                End If
            End While
            oReader.Close()
        End Sub
        ' Used to save the data back to the file or your choice
        Public Sub ListSections(ByVal sSection As String)
            Dim Counter As Integer = 0

            For Each s As IniSection In Sections
                If s.Name = sSection Then
                    For Each k As IniSection.IniKey In s.Keys
                        If k.Value <> String.Empty Then
                            stCompanyList(Counter, 0) = k.Name
                            stCompanyList(Counter, 1) = k.Value
                            stCompanyList(Counter, 2) = CStr(Counter + 1)
                            Counter = Counter + 1
                        End If
                    Next
                End If
            Next

        End Sub
        ' Used to save the data back to the file or your choice
        Public Sub Save(ByVal sFileName As String)
            Dim oWriter As New StreamWriter(sFileName, False)
            For Each s As IniSection In Sections
                Trace.WriteLine(String.Format("Writing Section: [{0}]", s.Name))
                oWriter.WriteLine(String.Format("[{0}]", s.Name))
                For Each k As IniSection.IniKey In s.Keys
                    If k.Value <> String.Empty Then
                        Trace.WriteLine(String.Format("Writing Key: {0}={1}", k.Name, k.Value))
                        oWriter.WriteLine(String.Format("{0}={1}", k.Name, k.Value))
                    Else
                        Trace.WriteLine(String.Format("Writing Key: {0}", k.Name))
                        oWriter.WriteLine(String.Format("{0}", k.Name))
                    End If
                Next
            Next
            oWriter.Close()
        End Sub

        ' Gets all the sections
        Public ReadOnly Property Sections() As System.Collections.ICollection
            Get
                Return m_sections.Values
            End Get
        End Property

        ' Adds a section to the IniFile object, returns a IniSection object to the new or existing object
        Public Function AddSection(ByVal sSection As String) As IniSection
            Dim s As IniSection = Nothing
            sSection = sSection.Trim()
            ' Trim spaces
            If m_sections.ContainsKey(sSection) Then
                s = DirectCast(m_sections(sSection), IniSection)
            Else
                s = New IniSection(Me, sSection)
                m_sections(sSection) = s
            End If
            Return s
        End Function

        ' Removes a section by its name sSection, returns trus on success
        Public Function RemoveSection(ByVal sSection As String) As Boolean
            sSection = sSection.Trim()
            Return RemoveSection(GetSection(sSection))
        End Function

        ' Removes section by object, returns trus on success
        Public Function RemoveSection(ByVal Section As IniSection) As Boolean
            If Section IsNot Nothing Then
                Try
                    m_sections.Remove(Section.Name)
                    Return True
                Catch ex As Exception
                    Trace.WriteLine(ex.Message)
                End Try
            End If
            Return False
        End Function

        '  Removes all existing sections, returns trus on success
        Public Function RemoveAllSections() As Boolean
            m_sections.Clear()
            Return (m_sections.Count = 0)
        End Function

        ' Returns an IniSection to the section by name, NULL if it was not found
        Public Function GetSection(ByVal sSection As String) As IniSection
            sSection = sSection.Trim()
            ' Trim spaces
            If m_sections.ContainsKey(sSection) Then
                Return DirectCast(m_sections(sSection), IniSection)
            End If
            Return Nothing
        End Function

        '  Returns a KeyValue in a certain section
        Public Function GetKeyValue(ByVal sSection As String, ByVal sKey As String) As String
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.Value
                End If
            End If
            Return String.Empty
        End Function

        ' Sets a KeyValuePair in a certain section
        Public Function SetKeyValue(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
            Dim s As IniSection = AddSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.AddKey(sKey)
                If k IsNot Nothing Then
                    k.Value = sValue
                    Return True
                End If
            End If
            Return False
        End Function

        ' Renames an existing section returns true on success, false if the section didn't exist or there was another section with the same sNewSection
        Public Function RenameSection(ByVal sSection As String, ByVal sNewSection As String) As Boolean
            '  Note string trims are done in lower calls.
            Dim bRval As Boolean = False
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                bRval = s.SetName(sNewSection)
            End If
            Return bRval
        End Function

        ' Renames an existing key returns true on success, false if the key didn't exist or there was another section with the same sNewKey
        Public Function RenameKey(ByVal sSection As String, ByVal sKey As String, ByVal sNewKey As String) As Boolean
            '  Note string trims are done in lower calls.
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.SetName(sNewKey)
                End If
            End If
            Return False
        End Function

        ' Remove a key by section name and key name
        Public Function RemoveKey(ByVal sSection As String, ByVal sKey As String) As Boolean
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Return s.RemoveKey(sKey)
            End If
            Return False
        End Function

        ' IniSection class 
        Public Class IniSection
            '  IniFile IniFile object instance
            Private m_pIniFile As IniFile
            '  Name of the section
            Private m_sSection As String
            '  List of IniKeys in the section
            Private m_keys As Hashtable

            ' Constuctor so objects are internally managed
            Protected Friend Sub New(ByVal parent As IniFile, ByVal sSection As String)
                m_pIniFile = parent
                m_sSection = sSection
                m_keys = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
            End Sub

            ' Returns all the keys in a section
            Public ReadOnly Property Keys() As System.Collections.ICollection
                Get
                    Return m_keys.Values
                End Get
            End Property

            ' Returns the section name
            Public ReadOnly Property Name() As String
                Get
                    Return m_sSection
                End Get
            End Property

            ' Adds a key to the IniSection object, returns a IniKey object to the new or existing object
            Public Function AddKey(ByVal sKey As String) As IniKey
                sKey = sKey.Trim()
                Dim k As IniSection.IniKey = Nothing
                If sKey.Length <> 0 Then
                    If m_keys.ContainsKey(sKey) Then
                        k = DirectCast(m_keys(sKey), IniKey)
                    Else
                        k = New IniSection.IniKey(Me, sKey)
                        m_keys(sKey) = k
                    End If
                End If
                Return k
            End Function

            ' Removes a single key by string
            Public Function RemoveKey(ByVal sKey As String) As Boolean
                Return RemoveKey(GetKey(sKey))
            End Function

            ' Removes a single key by IniKey object
            Public Function RemoveKey(ByVal Key As IniKey) As Boolean
                If Key IsNot Nothing Then
                    Try
                        m_keys.Remove(Key.Name)
                        Return True
                    Catch ex As Exception
                        Trace.WriteLine(ex.Message)
                    End Try
                End If
                Return False
            End Function

            ' Removes all the keys in the section
            Public Function RemoveAllKeys() As Boolean
                m_keys.Clear()
                Return (m_keys.Count = 0)
            End Function

            ' Returns a IniKey object to the key by name, NULL if it was not found
            Public Function GetKey(ByVal sKey As String) As IniKey
                sKey = sKey.Trim()
                If m_keys.ContainsKey(sKey) Then
                    Return DirectCast(m_keys(sKey), IniKey)
                End If
                Return Nothing
            End Function

            ' Sets the section name, returns true on success, fails if the section
            ' name sSection already exists
            Public Function SetName(ByVal sSection As String) As Boolean
                sSection = sSection.Trim()
                If sSection.Length <> 0 Then
                    ' Get existing section if it even exists...
                    Dim s As IniSection = m_pIniFile.GetSection(sSection)
                    If s IsNot Me AndAlso s IsNot Nothing Then
                        Return False
                    End If
                    Try
                        ' Remove the current section
                        m_pIniFile.m_sections.Remove(m_sSection)
                        ' Set the new section name to this object
                        m_pIniFile.m_sections(sSection) = Me
                        ' Set the new section name
                        m_sSection = sSection
                        Return True
                    Catch ex As Exception
                        Trace.WriteLine(ex.Message)
                    End Try
                End If
                Return False
            End Function

            ' Returns the section name
            Public Function GetName() As String
                Return m_sSection
            End Function

            ' IniKey class
            Public Class IniKey
                '  Name of the Key
                Private m_sKey As String
                '  Value associated
                Private m_sValue As String
                '  Pointer to the parent CIniSection
                Private m_section As IniSection

                ' Constuctor so objects are internally managed
                Protected Friend Sub New(ByVal parent As IniSection, ByVal sKey As String)
                    m_section = parent
                    m_sKey = sKey
                End Sub

                ' Returns the name of the Key
                Public ReadOnly Property Name() As String
                    Get
                        Return m_sKey
                    End Get
                End Property

                ' Sets or Gets the value of the key
                Public Property Value() As String
                    Get
                        Return m_sValue
                    End Get
                    Set(ByVal value As String)
                        m_sValue = value
                    End Set
                End Property

                ' Sets the value of the key
                Public Sub SetValue(ByVal sValue As String)
                    m_sValue = sValue
                End Sub
                ' Returns the value of the Key
                Public Function GetValue() As String
                    Return m_sValue
                End Function

                ' Sets the key name
                ' Returns true on success, fails if the section name sKey already exists
                Public Function SetName(ByVal sKey As String) As Boolean
                    sKey = sKey.Trim()
                    If sKey.Length <> 0 Then
                        Dim k As IniKey = m_section.GetKey(sKey)
                        If k IsNot Me AndAlso k IsNot Nothing Then
                            Return False
                        End If
                        Try
                            ' Remove the current key
                            m_section.m_keys.Remove(m_sKey)
                            ' Set the new key name to this object
                            m_section.m_keys(sKey) = Me
                            ' Set the new key name
                            m_sKey = sKey
                            Return True
                        Catch ex As Exception
                            Trace.WriteLine(ex.Message)
                        End Try
                    End If
                    Return False
                End Function

                ' Returns the name of the Key
                Public Function GetName() As String
                    Return m_sKey
                End Function
            End Class
            ' End of IniKey class
        End Class
        ' End of IniSection class

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
    ' End of IniFile class
End Module
