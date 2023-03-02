import { IDropdownOption } from "office-ui-fabric-react";
import { IFilterOptions, IListItem, IListItemBounce, ISelectDisp } from "../Models/IListItem";
import {  SummaryArray } from "./WorkFlowReport";

export interface IWorkFlowReportstate{
    pageSize: number;
    items?: IListItem[];
    paginatedItems?: IListItem[];
    itemsAfterDropdown?: IListItem[];
    searchFilterItems?: IListItem[];
    
    allColumnItems?: any[];
    viewFields: any[];
    viewFieldsBounce: any[];
    viewFieldsSummary: any[];
    isCurrentUserAdmin: boolean;
    paginationCurrentPage: number;
    isFilterPanelOpen: boolean;
    isAdvancedSearchPanelOpen: boolean;
    isAdminPanelOpen: boolean;
    filterOptions: IFilterOptions;
    selectDisp: ISelectDisp;
    showExportToExcelPopup: boolean;
    showWorkflowPopup: boolean;
    showDeleteItemPopup: boolean;
    showSendRequestConfirmPopup: boolean;
    deleteItemID: any;
    exportToExcelColumnName: any[];
    exportToExcelSelectedColumnName: any[];
    exportToExcelFinalValues: any[];
    workflowColumnName: any[];
    workflowSelectedColumnName: any[];
    workflowComments: string;
    selectedItems?: IListItem[];
    advanceSearch?: {};
    advanceSearchDisp: ISelectDisp;
    sortingItem?: {};

    managementTitles?: any[];
    managementOptions: any[];
    managementtitlesKEY:any;
    managementtitles:any;
    Temp: string;
    searchTextbox?: string;
    startDate: any;  
    endDate: any;  
    dateFilter:string;
    WorkflowItemsArray?: IListItem[];
    InsideArray?: any[];
    TestArray?: SummaryArray[];
    count?: number;
    DateArray? : any[];
    Item_SummaryReport?: boolean;
    Item_WorkFlowReport?: boolean;

    Item_Bounce?: boolean;
    itemsBounce?: IListItemBounce[];
    paginatedItemsBounce?: IListItemBounce[];
    searchFilterItemsBounce?: IListItemBounce[];
    allColumnItemsBounce?: any[];
    exportToExcelColumnNameBounce: any[];
    exportToExcelSelectedColumnNameBounce: any[];
    exportToExcelFinalValuesBounce: any[];
    selectedItemsBounce?: IListItemBounce[];
    SummaryColor?: boolean;
    BounceColor?: boolean;
    AllColor?: boolean;

}