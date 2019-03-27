Option Explicit

Public Sub CallbackFind(control As IRibbonControl)

    FindSetsMain    'ToFindSets module

End Sub

Public Sub CallbackToSixbit(control As IRibbonControl)

    FormattedToSixbit   'ToSixbit module

End Sub

Public Sub Callbackfixfitments(control As IRibbonControl)

    Formatted = False   'fitments not formatted yet
    
    FixFitments 'FormatFitments Module

End Sub

Public Sub CallbackInheritance(control As IRibbonControl)

    InheritanceMain 'MyFitment module

End Sub

Public Sub CallbackOOS(control As IRibbonControl)

    OOSMain 'OOS module

End Sub

Public Sub CallbackManageInv(control As IRibbonControl)

    ManageInventoryMain 'ManageInventory module

End Sub

Public Sub CallbackDropship(control As IRibbonControl)

    DropshipMain    'Dropship module

End Sub

Public Sub CallbackImportHerko(control As IRibbonControl)

    FindHerkoReport

End Sub

Public Sub CallbackImportShipstation(control As IRibbonControl)

    FindShipstationReport

End Sub

Public Sub CallbackSoldReport(control As IRibbonControl)

    FormatReportMain

End Sub

Public Sub CallbackRemoveKeep(control As IRibbonControl)

    RemoveKeepMain

End Sub

Public Sub CallbackConfirmInv(control As IRibbonControl)

    ConfirmInvMain

End Sub

Public Sub CallbackRemoveInactive(control As IRibbonControl)

    RemoveInactiveMain

End Sub

Public Sub CallbackFormatBoM(control As IRibbonControl)

    BoMMain     'BillofMaterials module

End Sub

Public Sub CallbackFinaleProducts(contorl As IRibbonControl)

    FinaleProductsMain

End Sub

Public Sub CallbackFinaleStockTake(control As IRibbonControl)

    FinaleStockTakeMain

End Sub

Public Sub CallbackFinaleBillofMaterials(control As IRibbonControl)

    FinaleBoMMain

End Sub

Public Sub CallbackFinaleLookups(control As IRibbonControl)

    FinaleLookupsMain

End Sub

Public Sub CallbackAmazontoFinale(control As IRibbonControl)

    

End Sub

Public Sub CallbackShipstationFields(control As IRibbonControl)

    ShipstationFieldsMain   'Shipstation module

End Sub

Public Sub CallbackAmazon(control As IRibbonControl)

    AmazonMain 'Amazon Module

End Sub

Public Sub CallbackUPC(control As IRibbonControl)

    UPC

End Sub

Public Sub CallbackEbayVolumePricing(control As IRibbonControl)

    VolumePricing

End Sub

Public Sub CallbackExportxlsx(control As IRibbonControl)

    XLSX    'Export Module

End Sub

Public Sub CallbackExportcsv(control As IRibbonControl)

    CSV     'Export Module

End Sub

Public Sub CallbackExporttxt(control As IRibbonControl)

    TXT     'Export Module

End Sub

Public Sub CallbackEmail(control As IRibbonControl)

    EmailMain   'Export Module

End Sub

Public Sub CallbackDbConn(control As IRibbonControl)

    DBConns.Show

End Sub

Public Sub CallbackVariables(control As IRibbonControl)

    VariablesForm.Show

End Sub

Public Sub CallbackVersion(control As IRibbonControl)

    OfficeVersion   'About module

End Sub

Public Sub AboutADP(control As IRibbonControl)

    ADPVersion  'About module

End Sub
