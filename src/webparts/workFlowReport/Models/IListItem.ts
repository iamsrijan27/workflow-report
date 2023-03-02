export interface IListItem {
    ItemID: any;
    WorkflowName: any;
    SendDate: any;
    IsCompleted: any;
    CompletedDate: any;
    IsBounceBack: any;
    ContactName: any;
    userEmail: any;
}
export interface IListItemBounce {
    ItemID: any;
    WorkflowName: any;
    SendDate: any;
    userEmail: any;
}


export interface IFilterOptions {
    WorkflowName?: string[];
    SendDate?: string[];
    IsCompleted?: string[];
    CompletedDate?: string[];
    IsBounceBack?: string[];
    ContactName?: string[];
    userEmail?: string[];
}

export interface ISelectDisp {
    WorkflowName?: any[];
    SendDate?: any[];
    IsCompleted?: any[];
    CompletedDate?: any[];
    IsBounceBack?: any[];
    ContactName?: any[];
    userEmail?: any[];
}