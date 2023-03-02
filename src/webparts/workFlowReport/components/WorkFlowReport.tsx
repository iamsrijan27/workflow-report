import * as React from "react";
import styles from "./WorkFlowReport.module.scss";
import { IWorkFlowReportProps } from "./IWorkFlowReportProps";
import { CSVLink } from "react-csv";
import { IViewField, ListView, Pagination } from "@pnp/spfx-controls-react";
import {
  CheckboxVisibility,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  IDetailsRowStyles,
  DetailsRow,
} from "office-ui-fabric-react/lib/DetailsList";
import DualListBox from "react-dual-listbox";
import "react-dual-listbox/lib/react-dual-listbox.css";
import {
  getTheme,
  mergeStyleSets,
  PrimaryButton,
  mergeStyles,
  IDropdownOption,
  IStackTokens,
  IDropdownStyles,
  DialogType,
  Stack,
  Dropdown,
  CommandButton,
  Dialog,
  FontIcon,
  DialogFooter,
} from "office-ui-fabric-react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import {
  DateTimePicker,
  DateConvention,
} from "@pnp/spfx-controls-react/lib/dateTimePicker";
import { colors, withStyles } from "@material-ui/core";
import Tooltip from "@material-ui/core/Tooltip";
import * as moment from "moment";
import { IListItem, IListItemBounce } from "../Models/IListItem";
import { IWorkFlowReportstate } from "./IWorkFlowReportstate";
import * as jquery from "jquery";

export interface SummaryArray {
  WorkflowName: any;
  StartDate: any;
  ReportDate: any;
  Total: any;
  Completed: any;
  Percentage: any;
}
const liTheme = getTheme();
const selectCustomStyles = {
  control: (state: any) => ({
    ...state,
    // color: "rgb(33, 33, 33)",
    borderColor: "rgb(153, 153, 153)",
    "&:hover": { borderColor: "rgb(33, 33, 33)" },
    "&:active,:focus,:focus-within": {
      borderColor: "rgb(219, 0, 7)",
      boxShadow: "0 0 0 1px #db0007",
    },
    "&.disabled": { borderColor: "rgb(166, 166, 166);" },
    minHeight: "32px",
    //height: "32px",
  }),
  placeholder: (provided: any, state: any) => {
    return {
      ...provided,
      //color: "#000",
    };
  },
  menu: (provided: any, state: any) => {
    return {
      ...provided,
      color: "#000",
      padding: 0,
      boxShadow: "rgba(0, 0, 0, 0.2) 0px 0px 2px 0px",
      outline: "transparent",
      borderColor: "rgb(234, 234, 234)",
      //fontSize: "14px",
      margin: 0,
      fontWeight: 400,
    };
  },
  menuList: (provided: any, state: any) => {
    return {
      ...provided,
      color: "#000",
    };
  },
  valueContainer: (provided: any, state: any) => ({
    ...provided,
    //height: "32px",
    padding: "0 6px",
    overflowY: "auto",
  }),
  indicatorsContainer: (provided: any, state: any) => ({
    ...provided,
    //height: "32px",
    "> div": {
      padding: "0 8px",
    },
  }),
  multiValue: (provided: any, state: any) => ({
    ...provided,
    //fontSize: "16px",
  }),
};
export default class WorkFlowReport extends React.Component<
  IWorkFlowReportProps,
  IWorkFlowReportstate
> {
  constructor(props: IWorkFlowReportProps, state: IWorkFlowReportstate) {
    super(props);
    const LightTooltip = withStyles((theme) => ({
      tooltip: {
        backgroundColor: theme.palette.common.white,
        color: "rgba(0, 0, 0, 0.87)",
        boxShadow: theme.shadows[1],
        fontSize: 11,
      },
    }))(Tooltip);
    const iconClass = mergeStyles({
      fontSize: 17,
      height: 20,
      width: 20,
      margin: "0 5px",
      cursor: "pointer",
    });
    const classNames = mergeStyleSets({
      deepSkyBlue: [{ color: "#db0007" }, iconClass],
    });
    const _trimLength: any = 20;
    // const summary {
    //   color:'#db0007';
    // }
    var _viewFields: IViewField[] = [
      {
        name: "WorkflowName",
        displayName: "Work Flow Name",
        minWidth: 100,
        maxWidth: 250,
        sorting: true,
        isResizable: true,
        render: (rowitem: any) => {
          const _WorkflowName = rowitem["WorkflowName"];
          if (_WorkflowName) {
            let _WorkflowNameTrimmed: string = _WorkflowName;
            if (_WorkflowName.length > 20) {
              _WorkflowNameTrimmed = _WorkflowName.substring(0, 30) + '...';
            }
            return <LightTooltip title={_WorkflowName} arrow>
              <span className="ms-fontColor-black">{_WorkflowName}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
      },
      {
        name: "SendDate",
        displayName: "Send Date",
        minWidth: 100,
        maxWidth: 400,
        sorting: true,
        isResizable: true,
        render: (rowitem: any) => {
          const _SendDate = rowitem["SendDate"];
          if (_SendDate) {
            let _SendDateTrimmed: string = _SendDate;
            if (_SendDate.length > 20) {
              _SendDateTrimmed = _SendDate.substring(0, 30) + '...';
            }
            return <LightTooltip title={_SendDate} arrow>
              <span className="ms-fontColor-black">{_SendDate}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
        // render: (item: any, index, column: any) => {
        //   return moment(item.SendDate).format("MM/DD/YYYY");
        // },
      },
      {
        name: "IsCompleted",
        displayName: "Is Completed",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        sorting: true,
        render: (rowitem: any) => {
          const _IsCompleted = rowitem["IsCompleted"];
          if (_IsCompleted) {
            let _IsCompletedTrimmed: string = _IsCompleted;
            if (_IsCompleted.length > 20) {
              _IsCompletedTrimmed = _IsCompleted.substring(0, 30) + '...';
            }
            return <LightTooltip title={_IsCompleted} arrow>
              <span className="ms-fontColor-black">{_IsCompleted}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
        // render: (item: any, index, column: any) => {
        //   if (item.IsCompleted == true) {
        //     return "Completed";
        //   } else {
        //     return "Pending";
        //   }
        // },
      },
      {
        name: "ContactName",
        displayName: "Contact Name",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        sorting: true,
        render: (rowitem: any) => {
          const _ContactName = rowitem["ContactName"];
          if (_ContactName) {
            let _ContactNameTrimmed: string = _ContactName;
            if (_ContactName.length > 20) {
              _ContactNameTrimmed = _ContactName.substring(0, 30) + '...';
            }
            return <LightTooltip title={_ContactName} arrow>
              <span className="ms-fontColor-black">{_ContactName}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
      },
      {
        name: "userEmail",
        displayName: "User Email",
        minWidth: 100,
        maxWidth: 400,
        sorting: true,
        isResizable: true,
        render: (rowitem: any) => {
          const _userEmail = rowitem["userEmail"];
          if (_userEmail) {
            let _userEmailTrimmed: string = _userEmail;
            if (_userEmail.length > 20) {
              _userEmailTrimmed = _userEmail.substring(0, 30) + '...';
            }
            return <LightTooltip title={_userEmail} arrow>
              <span className="ms-fontColor-black">{_userEmail}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
      },
      {
        name: "IsBounceBack",
        displayName: "Is Bounce Back",
        minWidth: 100,
        maxWidth: 400,
        sorting: true,
        isResizable: true,
        // render: (rowitem: any) => {
        //   const _IsBounceBack = rowitem["IsBounceBack"];
        //   if (_IsBounceBack) {
        //     let _IsBounceBackTrimmed: string = _IsBounceBack;
        //     if (_IsBounceBack.length > 20) {
        //       _IsBounceBackTrimmed = _IsBounceBack.substring(0, 30) + '...';
        //     }
        //     return <LightTooltip title={_IsBounceBack} arrow>
        //       <span className="ms-fontColor-black">{_IsBounceBack}</span>
        //     </LightTooltip>;
        //   }
        //   else {
        //     // return <span></span>;
        //   }
        // },
      },
      {
        name: "CompletedDate",
        displayName: "Completed Date",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        sorting: true,
        render: (rowitem: any) => {
          const _CompletedDate = rowitem["CompletedDate"];
          if (_CompletedDate) {
            let _CompletedDateTrimmed: string = _CompletedDate;
            if (_CompletedDate.length > 20) {
              _CompletedDateTrimmed = _CompletedDate.substring(0, 30) + '...';
            }
            return <LightTooltip title={_CompletedDate} arrow>
              <span className="ms-fontColor-black">{_CompletedDate}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
        // render: (item: any, index, column: any) => {
        //   if (
        //     moment(item.CompletedDate).format("MM/DD/YYYY") == "Invalid date"
        //   ) {
        //     return "NA";
        //   } else {
        //     return moment(item.CompletedDate).format("MM/DD/YYYY");
        //   }
        // },
      },
    ];
    var _viewFieldsBounce: IViewField[] = [
      {
        name: "WorkflowName",
        displayName: "Work Flow Name",
        minWidth: 100,
        maxWidth: 250,
        sorting: true,
        isResizable: true,
        render: (rowitem: any) => {
          const _WorkflowName = rowitem["WorkflowName"];
          if (_WorkflowName) {
            let _WorkflowNameTrimmed: string = _WorkflowName;
            if (_WorkflowName.length > 20) {
              _WorkflowNameTrimmed = _WorkflowName.substring(0, 30) + '...';
            }
            return <LightTooltip title={_WorkflowName} arrow>
              <span className="ms-fontColor-black">{_WorkflowName}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
      },
      {
        name: "SendDate",
        displayName: "Send Date",
        minWidth: 100,
        maxWidth: 210,
        sorting: true,
        isResizable: true,
        render: (rowitem: any) => {
          const _SendDate = rowitem["SendDate"];
          if (_SendDate) {
            let _SendDateTrimmed: string = _SendDate;
            if (_SendDate.length > 20) {
              _SendDateTrimmed = _SendDate.substring(0, 30) + '...';
            }
            return <LightTooltip title={_SendDate} arrow>
              <span className="ms-fontColor-black">{_SendDate}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
      },
      {
        name: "userEmail",
        displayName: "User Email",
        minWidth: 100,
        maxWidth: 250,
        isResizable: true,
        sorting: true,
        render: (rowitem: any) => {
          const _userEmail = rowitem["userEmail"];
          if (_userEmail) {
            let _userEmailTrimmed: string = _userEmail;
            if (_userEmail.length > 20) {
              _userEmailTrimmed = _userEmail.substring(0, 30) + '...';
            }
            return <LightTooltip title={_userEmail} arrow>
              <span className="ms-fontColor-black">{_userEmail}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },
      },
    ];

    ///////////////////////////////
    var _viewFieldsSummary: IViewField[] = [
      {
        name: "WorkflowName",
        displayName: "Work Flow Name",
        minWidth: 100,
        maxWidth: 250,
        isResizable: true,
        render: (rowitem: any) => {
          const _WorkflowName = rowitem["WorkflowName"];
          if (_WorkflowName) {
            let _WorkflowNameTrimmed: string = _WorkflowName;
            if (_WorkflowName.length > 20) {
              _WorkflowNameTrimmed = _WorkflowName.substring(0, 30) + '...';
            }
            return <LightTooltip title={_WorkflowName} arrow>
              <span className="ms-fontColor-black">{_WorkflowName}</span>
            </LightTooltip>;
          }
          else {
            // return <span></span>;
          }
        },        
      },
      {
        name: "StartDate",
        displayName: "Start Date",
        minWidth: 100,
        maxWidth: 400,
        // sorting: true,
        isResizable: true,
      },
      {
        // key: "column3",
        name: "ReportDate",
        displayName: "Report Date",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        // sorting: true,
      },
      {
        // key: "column4",
        name: "Total",
        displayName: "Total",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        // sorting: true,
      },
      {
        // key: "column5",
        name: "Completed",
        displayName: "Completed",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        //sorting: true,
      },
      {
        // key: "column1",
        name: "Percentage",
        displayName: "Percentage",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        // sorting: true,
      },
    ];
    ////////////////////////
    this.state = {
      pageSize: this.props.pageSize,
      items: [],
      paginatedItems: [],
      itemsAfterDropdown: [],
      searchFilterItems: [],
      allColumnItems: [],
      viewFields: _viewFields,
      viewFieldsBounce: _viewFieldsBounce,
      viewFieldsSummary: _viewFieldsSummary,
      isCurrentUserAdmin: false,
      paginationCurrentPage: 1,
      selectedItems: [],
      isFilterPanelOpen: false,
      isAdvancedSearchPanelOpen: false,
      isAdminPanelOpen: false,
      filterOptions: {},
      selectDisp: {},
      showExportToExcelPopup: false,
      showWorkflowPopup: false,
      showDeleteItemPopup: false,
      showSendRequestConfirmPopup: false,
      deleteItemID: "",
      exportToExcelColumnName: [],
      exportToExcelSelectedColumnName: [],
      exportToExcelFinalValues: [],
      workflowColumnName: [],
      workflowSelectedColumnName: [],
      workflowComments: "",
      advanceSearch: {},
      advanceSearchDisp: {},
      sortingItem: {},

      managementTitles: [],
      managementOptions: [""],
      managementtitlesKEY: "All",
      managementtitles: "All",
      Temp: "",
      searchTextbox: "",
      startDate: null,
      endDate: null,
      dateFilter: "",
      WorkflowItemsArray: [],
      InsideArray: [],
      TestArray: [],
      count: 0,
      DateArray: [],
      Item_SummaryReport: true,
      Item_WorkFlowReport: false,
      Item_Bounce: false,

      itemsBounce: [],
      paginatedItemsBounce: [],
      exportToExcelColumnNameBounce: [],
      exportToExcelSelectedColumnNameBounce: [],
      exportToExcelFinalValuesBounce: [],
      selectedItemsBounce: [],
      SummaryColor:true,
      BounceColor: false,
      AllColor:false
    };
  }
  public componentDidMount() {
    this._getItems();
    this.getMultipleChoiceFieldNameValue();
    this._getItemsBounce();
  }
  private _getItemsBounce = (): void => {
    this.props.helperService
      .getAllItemsBounce(this.props.BouncelistName)
      .then((resultSetsBounce) => {
        var listItemsBounce: Array<IListItemBounce> =
          new Array<IListItemBounce>();
        resultSetsBounce.map((item, index) => {
          listItemsBounce.push({
            ItemID: item.ID,
            WorkflowName: item.Title,
            SendDate: moment(item.SendDate).format("MM/DD/YYYY"),
            userEmail: item.UserEmail,
          });
        });
        var allColumnValueBounce: Array<any> = new Array<any>();
        resultSetsBounce.map((listItem, index) => {
          listItem.SendDate = moment(listItem.SendDate).format("MM/DD/YYYY") == "Invalid date" ? "NA" : moment(listItem.SendDate).format("MM/DD/YYYY");
          allColumnValueBounce.push({ listItem });
        });
        this.setState({
          itemsBounce: listItemsBounce,
          paginatedItemsBounce: listItemsBounce.slice(0, this.state.pageSize),
          paginationCurrentPage: 1,
          searchFilterItemsBounce: listItemsBounce,
          allColumnItemsBounce: allColumnValueBounce.map((i) => i.listItem),
        });
      });
  };
  private _getItems = (): void => {
    this.props.helperService
      .getAllItems(this.props.listName)
      .then((resultSets) => {
        var listItems: Array<IListItem> = new Array<IListItem>();
        resultSets.map((item, index) => {
          listItems.push({
            ItemID: item.ID,
            WorkflowName: item.WorkflowName,
            SendDate: moment(item.SendDate).format("MM/DD/YYYY")== "Invalid date" ? "NA" : moment(item.SendDate).format("MM/DD/YYYY"),
            IsCompleted: item.IsCompleted == true ? "Completed" : "Pending",
            CompletedDate: moment(item.CompletedDate).format("MM/DD/YYYY")== "Invalid date" ? "NA" :moment(item.CompletedDate).format("MM/DD/YYYY"),
            IsBounceBack: item.IsBounceBack,
            ContactName: item.ContactName,
            userEmail: item.UserEmail,
          });
          //to get all the management titles from the list.
          this.state.managementTitles.push(item.WorkflowName);
        });

        //To get unique management value for the dropdown
        var managementOptionArray: any[] = [];
        var uniqueManagementTitles: any[] = this.state.managementTitles.filter(
          (v, i, a) => a.indexOf(v) === i
        );
        uniqueManagementTitles.forEach((option) => {
          let managementOption = { option: option, focused: false };
          // managementOptionArray.push({ option });
          managementOptionArray.push(managementOption);
        });
        this.setState({ managementOptions: managementOptionArray });

        // array inside of array for getting buttons values as dynamic
        this.state.managementOptions.forEach((options) => {
          listItems.forEach((e) => {
            if (e.WorkflowName == options.option) {
              this.state.WorkflowItemsArray.push(e);
            }
          });

          this.state.InsideArray.push(this.state.WorkflowItemsArray);
          this.setState({
            WorkflowItemsArray: [],
          });
        });
        let TempCount = 0;
        for (let i = 0; i < this.state.InsideArray.length; i++) {
          for (let j = 0; j < this.state.InsideArray[i].length; j++) {
            // if (this.state.InsideArray[i][j].IsCompleted == true) {
            if (this.state.InsideArray[i][j].IsCompleted == "Completed") {
              TempCount += 1;
            }
            this.state.DateArray.push(this.state.InsideArray[i][j].SendDate);
          }
          this.state.TestArray.push({
            WorkflowName: this.state.InsideArray[i][0].WorkflowName,
            StartDate: this.state.DateArray[0],
            ReportDate: this.state.DateArray[this.state.DateArray.length - 1],
            Total: this.state.InsideArray[i].length,
            Completed: TempCount,
            Percentage: ((TempCount / this.state.InsideArray[i].length) * 100).toFixed(2),
            // Percentage: Math.floor((TempCount / this.state.InsideArray[i].length) * 100),
          });
          TempCount = 0;
          this.setState({
            DateArray: [],
          });
        }

        var allColumnValue: Array<any> = new Array<any>();
        resultSets.map((listItem, index) => {
          listItem.SendDate = moment(listItem.SendDate).format("MM/DD/YYYY") == "Invalid date" ? "NA" : moment(listItem.SendDate).format("MM/DD/YYYY");
          listItem.CompletedDate = moment(listItem.CompletedDate).format("MM/DD/YYYY") == "Invalid date" ? "NA" : moment(listItem.CompletedDate).format("MM/DD/YYYY");
          allColumnValue.push({ listItem });
        });
        this.setState({
          items: listItems,
          paginatedItems: listItems.slice(0, this.state.pageSize),
          searchFilterItems: listItems,
          allColumnItems: allColumnValue.map((i) => i.listItem),
          paginationCurrentPage: 1,
          sortingItem: {},
          itemsAfterDropdown: listItems,
          managementtitles: this.state.managementtitles.text,
          managementtitlesKEY: this.state.managementtitles.key,
          dateFilter: this.state.managementtitles,
        });
      });
  };
  private getMultipleChoiceFieldNameValue = (): void => {
    //this.props.helperService.getFieldDropdownValues("Category Contacts", "").then(listFields => {
    this.props.helperService
      .getFieldDropdownValues(this.props.listName, "")
      .then((listFields) => {
        let fieldName: any[] = [];
        fieldName = [...listFields];
        //fieldName.unshift({ key: "All Fields", text: "All Fields", value: "All Fields", label: "All Fields" });
        this.setState({
          exportToExcelColumnName: fieldName,
          workflowColumnName: listFields,
        });
      });
    this.props.helperService
      .getFieldDropdownValues(this.props.BouncelistName, "")
      .then((listFields) => {
        let fieldName: any[] = [];
        fieldName = [...listFields];
        //fieldName.unshift({ key: "All Fields", text: "All Fields", value: "All Fields", label: "All Fields" });
        this.setState({ exportToExcelColumnNameBounce: fieldName });
      });
  };
  private OnExportToExcelPopupOpen = () => {
    var listViewSelectedItems: IListItem[] = [];
    if (this.state.selectedItems.length == 0) {
      listViewSelectedItems = this.state.searchFilterItems;
    } else {
      listViewSelectedItems = this.state.selectedItems;
    }

    if (listViewSelectedItems != null && listViewSelectedItems.length > 0) {
      var allColumnItem = this.state.allColumnItems;
      var exportValueTemp: any[] = [];
      var exportValueFinal: any[] = [];
      exportValueTemp = allColumnItem.filter((val) =>
        listViewSelectedItems.some((i) => Number(val.ID) == Number(i.ItemID))
      );
      if (exportValueTemp != null && exportValueTemp.length > 0) {
        exportValueTemp.map((exp) => {
          let expLocal = {};
          this.state.exportToExcelColumnName.map((col) => {
            if (col.key != "All Fields") {
              expLocal[col.label] = exp[col.key];
            }
          });
          exportValueFinal.push(expLocal);
        });
      }
    }
    this.setState({
      exportToExcelFinalValues: exportValueFinal,
      showExportToExcelPopup: true,
      showWorkflowPopup: false,
    });
  };
  private onDualListExcelChange = (selected) => {
    let selectedValues: any[] = [];
    selectedValues = this.state.exportToExcelColumnName.filter((selVal) =>
      selected.includes(selVal.key)
    );
    var listViewSelectedItems: IListItem[] = [];
    if (this.state.selectedItems.length == 0) {
      // listViewSelectedItems = this.state.items;
      listViewSelectedItems = this.state.searchFilterItems;
    } else {
      listViewSelectedItems = this.state.selectedItems;
    }

    if (listViewSelectedItems != null && listViewSelectedItems.length > 0) {
      var allColumnItem = this.state.allColumnItems;
      var exportValueTemp: any[] = [];
      var exportValueFinal: any[] = [];
      exportValueTemp = allColumnItem.filter((val) =>
        listViewSelectedItems.some((i) => Number(val.ID) == Number(i.ItemID))
      );
      if (exportValueTemp != null && exportValueTemp.length > 0) {
        exportValueTemp.map((exp) => {
          let expLocal = {};
          // if (selectedValues != null && selectedValues != [] && selectedValues.length > 0 && selectedValues.some(sel => sel.key == "All Fields")) {
          if (
            selectedValues != null &&
            selectedValues.length > 0 &&
            selectedValues.some((sel) => sel.key == "All Fields")
          ) {
            this.state.exportToExcelColumnName.map((col) => {
              if (col.key != "All Fields") {
                expLocal[col.label] = exp[col.key];
              }
            });
          }
          // else if (selectedValues == [] && selectedValues.length == 0) {
          else if (selectedValues.length == 0) {
            this.state.exportToExcelColumnName.map((col) => {
              if (col.key != "All Fields") {
                expLocal[col.label] = exp[col.key];
              }
            });
          } else {
            selectedValues.map((col) => {
              if (col.key != "All Fields") {
                expLocal[col.label] = exp[col.key];
              }
            });
          }
          exportValueFinal.push(expLocal);
        });
      }
    }
    this.setState({
      exportToExcelSelectedColumnName: selected,
      exportToExcelFinalValues: exportValueFinal,
    });
  };
  private OnExportToExcelBouncePopupOpen = () => {
    var listViewSelectedItems: IListItemBounce[] = [];
    if (this.state.selectedItemsBounce.length == 0) {
      listViewSelectedItems = this.state.searchFilterItemsBounce;
    } else {
      listViewSelectedItems = this.state.selectedItemsBounce;
    }

    if (listViewSelectedItems != null && listViewSelectedItems.length > 0) {
      var allColumnItem = this.state.allColumnItemsBounce;
      var exportValueTemp: any[] = [];
      var exportValueFinal: any[] = [];
      exportValueTemp = allColumnItem.filter((val) =>
        listViewSelectedItems.some((i) => Number(val.ID) == Number(i.ItemID))
      );
      if (exportValueTemp != null && exportValueTemp.length > 0) {
        exportValueTemp.map((exp) => {
          let expLocal = {};
          this.state.exportToExcelColumnNameBounce.map((col) => {
            if (col.key != "All Fields") {
              expLocal[col.label] = exp[col.key];
            }
          });
          exportValueFinal.push(expLocal);
        });
      }
    }
    this.setState({
      exportToExcelFinalValues: exportValueFinal,
      showExportToExcelPopup: true,
      showWorkflowPopup: false,
    });
  };
  private onDualListExcelChangeBounce = (selected) => {
    let selectedValues: any[] = [];
    selectedValues = this.state.exportToExcelColumnNameBounce.filter((selVal) =>
      selected.includes(selVal.key)
    );
    var listViewSelectedItems: IListItemBounce[] = [];
    if (this.state.selectedItemsBounce.length == 0) {
      // listViewSelectedItems = this.state.items;
      listViewSelectedItems = this.state.searchFilterItemsBounce;
    } else {
      listViewSelectedItems = this.state.selectedItemsBounce;
    }

    if (listViewSelectedItems != null && listViewSelectedItems.length > 0) {
      var allColumnItem = this.state.allColumnItemsBounce;
      var exportValueTemp: any[] = [];
      var exportValueFinal: any[] = [];
      exportValueTemp = allColumnItem.filter((val) =>
        listViewSelectedItems.some((i) => Number(val.ID) == Number(i.ItemID))
      );
      if (exportValueTemp != null && exportValueTemp.length > 0) {
        exportValueTemp.map((exp) => {
          let expLocal = {};
          // if (selectedValues != null && selectedValues != [] && selectedValues.length > 0 && selectedValues.some(sel => sel.key == "All Fields")) {
          if (
            selectedValues != null &&
            selectedValues.length > 0 &&
            selectedValues.some((sel) => sel.key == "All Fields")
          ) {
            this.state.exportToExcelColumnNameBounce.map((col) => {
              if (col.key != "All Fields") {
                expLocal[col.label] = exp[col.key];
              }
            });
          }
          // else if (selectedValues == [] && selectedValues.length == 0) {
          else if (selectedValues.length == 0) {
            this.state.exportToExcelColumnNameBounce.map((col) => {
              if (col.key != "All Fields") {
                expLocal[col.label] = exp[col.key];
              }
            });
          } else {
            selectedValues.map((col) => {
              if (col.key != "All Fields") {
                expLocal[col.label] = exp[col.key];
              }
            });
          }
          exportValueFinal.push(expLocal);
        });
      }
    }
    this.setState({
      exportToExcelSelectedColumnNameBounce: selected,
      exportToExcelFinalValues: exportValueFinal,
    });
  };
  private _onListViewRenderRow(props: any) {
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (props) {
      if (props.itemIndex % 2 === 0) {
        // Every other row renders with a different background color
        customStyles.root = {
          backgroundColor: liTheme.palette.themeLighterAlt,
        };
      }
      return <DetailsRow {...props} styles={customStyles} />;
    }
    return null;
  }
  private _getSelection = (items: any[]): void => {
    if (items.length > 0) {
      let _selectedValues: IListItem[] = [];
      let _selectedValuesBounce: IListItemBounce[] = [];
      items.map((item) => {
        _selectedValues.push(item);
        _selectedValuesBounce.push(item);
      });
      this.setState({
        selectedItems: _selectedValues,
        selectedItemsBounce: _selectedValuesBounce,
      });
    } else {
      this.setState({ selectedItems: [], selectedItemsBounce: [] });
    }
  };
  public itemsSortingListView = (
    items: any[],
    columnKey: string,
    descending: boolean
  ): any[] => {
    const key = columnKey as keyof any;
    let newState: any = this.state;
    if (descending) {
      newState.sortingItem[key] = descending;
    } else if (
      newState.sortingItem[key] != null &&
      newState.sortingItem[key] != undefined
    ) {
      newState.sortingItem[key] = !newState.sortingItem[key];
    } else {
      newState.sortingItem[key] = false;
    }

    let sortValue = this.state.searchFilterItems
      .slice(0)
      .sort((a: any, b: any) =>
        (
          newState.sortingItem[key]
            ? String(a[key]).toLowerCase() < String(b[key]).toLowerCase()
            : String(a[key]).toLowerCase() > String(b[key]).toLowerCase()
        )
          ? 1
          : -1
      );
    newState.searchFilterItems = sortValue;
    newState.paginatedItems = sortValue.slice(0, this.state.pageSize);
    newState.paginationCurrentPage = 1;
    this.setState(newState);

    return sortValue.slice(0, this.state.pageSize);
  };
  private _getPage(page: number) {
    const roundupPage = Math.ceil(page - 1);
    this.setState({
      paginatedItemsBounce: this.state.searchFilterItemsBounce.slice(
        roundupPage * this.state.pageSize,
        roundupPage * this.state.pageSize + this.state.pageSize
      ),
      paginatedItems: this.state.searchFilterItems.slice(
        roundupPage * this.state.pageSize,
        roundupPage * this.state.pageSize + this.state.pageSize
      ),
      paginationCurrentPage: page,
    });
  }
  private onNumberItemChange = (event: any, item: IDropdownOption): void => {
    this.setState({
      paginatedItemsBounce: this.state.searchFilterItemsBounce.slice(
        0,
        Number(item.key)
      ),
      paginatedItems: this.state.searchFilterItems.slice(0, Number(item.key)),
      paginationCurrentPage: 1,
      pageSize: Number(item.key),
    });
  };
  private _choicegroupChanged = (option: any): void => {
    if (option == "All") {
      this.setState({
        paginatedItems: this.state.items.slice(0, this.state.pageSize),
        searchFilterItems: this.state.items,
        paginationCurrentPage: 1,
      });
    }
  };
  private _highlightButton = (option: string): void =>{
    let managementOptions = this.state.managementOptions;
      managementOptions.forEach((i) => {
        if (i.option == option) {
          i.focused = true;
        }
        else{
          i.focused = false;
        }
      });
      this.setState({
        managementOptions: managementOptions
      });
  }
  private _getitemsDropdown = (option: any): void => {
    if (option == "All") {
      this._highlightButton("All")
      this.setState({
        paginatedItems: this.state.items.slice(0, this.state.pageSize),
        searchFilterItems: this.state.items,
        paginationCurrentPage: 1,
        dateFilter: option, //9/26/22
        Item_SummaryReport: false, // 28/9/22
        Item_WorkFlowReport: true, //28/9/22
        Item_Bounce: false,
        SummaryColor: false,
        BounceColor: false,
        endDate: null,
        startDate: null,
        AllColor: true
      });
    } else {
      let filterDropdownManaged: boolean =
        this.state.items != undefined && this.state.items.length > 0;
      var filterDropdownManagedResult: IListItem[] = [];
      if (filterDropdownManaged) {
        filterDropdownManagedResult = this.state.items.filter(
          (val) =>
            val.WorkflowName.split(" ").join("") ==
            option.option.split(" ").join("")
        );
      }
      this._highlightButton(option.option);
      this.setState({
        // managementOptions: managementOptions,
        itemsAfterDropdown: filterDropdownManagedResult,
        paginatedItems: filterDropdownManagedResult.slice(0,this.state.pageSize),
        searchFilterItems: filterDropdownManagedResult,
        paginationCurrentPage: 1,
        managementtitles: option.option,
        managementtitlesKEY: option.option,
        advanceSearch: {},
        advanceSearchDisp: {},
        filterOptions: {},
        selectDisp: {},
        endDate: null,
        startDate: null,
        dateFilter: option.option,
        Item_SummaryReport: false, // 28/9/22
        Item_WorkFlowReport: true, //28/9/22
        Item_Bounce: false,
        SummaryColor: false,
        BounceColor: false,
        AllColor:false
      });
    }
  };
  private OnDialogPopupClose = () => {
    this.setState({
      exportToExcelFinalValues: [],
      exportToExcelSelectedColumnName: [],
      exportToExcelSelectedColumnNameBounce: [],
      workflowComments: "",
      workflowSelectedColumnName: [],
      showExportToExcelPopup: false,
      showWorkflowPopup: false,
    });
  };
  /*Start Date*/
  private handleStartDate = (date: any): void => {
    if (this.state.endDate != null && this.state.endDate > date) {
      this.setState({ startDate: date });
    } else if (this.state.endDate != null && this.state.endDate < date) {
      alert("Start date should be less than end date");
      this.setState({ endDate: null, startDate: date });
    } else {
      this.setState({ startDate: date });
    }
  };
  /*End Date */
  private handleEndDate = (date: any): void => {
    if (this.state.startDate != null && this.state.startDate <= date) {
      this.setState({ endDate: date });
    } else {
      alert("Start date should be less than end date");
      this.setState({ endDate: null });
    }
  };
  private _resetForm = () => {
    this.setState({
      endDate: null,
      startDate: null,
      paginatedItems: this.state.itemsAfterDropdown.slice(
        0,
        this.state.pageSize
      ),
      searchFilterItems: this.state.itemsAfterDropdown,
      paginatedItemsBounce: this.state.itemsBounce.slice(
        0,
        this.state.pageSize
      ),
      searchFilterItemsBounce: this.state.itemsBounce,
      paginationCurrentPage: 1,
    });
  };
  private _searchfilter = (id: any): void => {
    let filterarray: any[] = [];
    let filterarrayfinal: any[] = [];
    let startdate = moment(this.state.startDate);
    let enddate = moment(this.state.endDate);
    // if(this.state.startDate == null && this.state.endDate != null){
    //   alert("Please Select  Start Date")
    // }
    // if(this.state.startDate != null && this.state.endDate == null){
    //   alert("Please Select  End Date")
    // }
    if (
      (this.state.startDate == null && this.state.endDate == null) ||
      (this.state.startDate == null && this.state.endDate != null) ||
      (this.state.startDate != null && this.state.endDate == null)
    ) {
      alert("Please Select  Both Start Date and End Date");
    } else if (
      this.state.startDate != null &&
      this.state.startDate != undefined &&
      this.state.endDate != null &&
      this.state.endDate != undefined
    ) {
      let filterDropdownManaged: boolean =
        this.state.items != undefined && this.state.items.length > 0;
      ////////////////////////////////
      if (filterDropdownManaged && id == "All") {
        this.state.items.forEach((e) => {
          let senddate = moment(new Date(e.SendDate));
          if (senddate >= startdate && senddate <= enddate) {
            filterarrayfinal.push(e);
          }
        });
      } else if (filterDropdownManaged && id == "BounceBack") {
        this.state.itemsBounce.forEach((e) => {
          let senddate = moment(new Date(e.SendDate));
          if (senddate >= startdate && senddate <= enddate) {
            filterarrayfinal.push(e);
          }
        });

        //filterarray = this.state.items.filter(val => (val.IsBounceBack == true));
      } else if (filterDropdownManaged) {
        filterarray = this.state.items.filter(
          (val) =>
            val.WorkflowName.split(" ").join("") == id.split(" ").join("")
        );
      }

      filterarray.forEach((e) => {
        let senddate = moment(new Date(e.SendDate));
        if (senddate >= startdate && senddate <= enddate) {
          filterarrayfinal.push(e);
        }
      });
      this.setState({
        paginatedItemsBounce: filterarrayfinal.slice(0, this.state.pageSize),
        searchFilterItemsBounce: filterarrayfinal,
        paginatedItems: filterarrayfinal.slice(0, this.state.pageSize),
        searchFilterItems: filterarrayfinal,
        paginationCurrentPage: 1,
      });
    } else {
      this.setState({
        endDate: null,
        startDate: null,
        paginatedItems: this.state.itemsAfterDropdown.slice(
          0,
          this.state.pageSize
        ),
        searchFilterItems: this.state.itemsAfterDropdown,
        paginationCurrentPage: 1,
      });
    }
  };
  //Avoid enter key press in List View Search
  public onFilterKeyPress(event: any): void {
    if (event.keyCode == 13) {
      event.preventDefault();
    }
  }
  private SummaryReport(e: string): void {
    if (e == "Summary") {
      this._highlightButton("Summary")
      this.setState({
        paginationCurrentPage: 1,
        Item_SummaryReport: true,
        Item_WorkFlowReport: false,
        Item_Bounce: false,
        SummaryColor: true,
        BounceColor: false,
        AllColor:false
      });
    }
  }
  private BounceBack(e: string): void {
    if (e == "BounceBack") {
      this._highlightButton("BounceBack")
      this.setState({
        paginatedItemsBounce: this.state.itemsBounce.slice(0, this.state.pageSize),
        paginationCurrentPage: 1,
        endDate: null,
        startDate: null,
        // searchFilterItemsBounce: this.state.itemsBounce,
        // allColumnItemsBounce: this.state.allColumnItemsBounce.map((i) => i.listItem),
        Item_SummaryReport: false,
        Item_WorkFlowReport: false,
        Item_Bounce: true,
        SummaryColor: false,
        BounceColor: true,
        AllColor:false
      });
    }
  }
  private contentStyles = {
    content: {
      backgroundColor: "Lavender",
      // marginBottom: '5%'
    },
  };
  public render(): React.ReactElement<IWorkFlowReportProps> {
    const LightTooltip = withStyles((theme) => ({
      tooltip: {
        backgroundColor: theme.palette.common.white,
        color: "rgba(0, 0, 0, 0.87)",
        boxShadow: theme.shadows[1],
        fontSize: 11,
      },
    }))(Tooltip);
    const _viewFieldsSummary: IColumn[] = [
      {
        key: "column1",
        name: "Work Flow Name",
        fieldName: "WorkflowName",
        minWidth: 100,
        maxWidth: 250,
        isResizable: true,
        // onRender: (rowitem: any) => {
        //   const _WorkflowName = rowitem["column1"];
        //   if (_WorkflowName) {
        //     let _WorkflowNameTrimmed: string = _WorkflowName;
        //     if (_WorkflowName.length > 20) {
        //       _WorkflowNameTrimmed = _WorkflowName.substring(0, 30) + '...';
        //     }
        //     return <LightTooltip title={_WorkflowName} arrow>
        //       <span className="ms-fontColor-black">{_WorkflowName}</span>
        //     </LightTooltip>;
        //   }
        //   else {
        //     // return <span></span>;
        //   }
        // },        
      },
      {
        key: "column2",
        name: "Start Date",
        fieldName: "StartDate",
        minWidth: 100,
        maxWidth: 400,
        // sorting: true,
        isResizable: true,
      },
      {
        key: "column3",
        name: "Report Date",
        fieldName: "ReportDate",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        // sorting: true,
      },
      {
        key: "column4",
        name: "Total",
        fieldName: "Total",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        // sorting: true,
      },
      {
        key: "column5",
        name: "Completed",
        fieldName: "Completed",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        //sorting: true,
      },
      {
        key: "column1",
        name: "Percentage",
        fieldName: "Percentage",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        // sorting: true,
      },
    ];
    const _viewFieldsBounce: IColumn[] = [
      {
        key: "column1",
        name: "Work Flow Name",
        fieldName: "WorkflowName",
        minWidth: 100,
        maxWidth: 250,
        isResizable: true,
      },
      {
        key: "column2",
        name: "Send Date",
        fieldName: "SendDate",
        minWidth: 100,
        maxWidth: 400,
        // sorting: true,
        isResizable: true,
      },
      {
        key: "column1",
        name: "user Email",
        fieldName: "userEmail",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        // sorting: true,
      },
    ];
    const { userDisplayName } = this.props;
    const stackTokens: IStackTokens = { childrenGap: 15 };
    const buttonStyles = { root: { marginRight: 8 } };

    const dropdownStylescss: Partial<IDropdownStyles> = {
      dropdown: { width: 142, paddingLeft: "20px" },
    };
    const dropdown: Partial<IDropdownStyles> = {
      dropdown: {
        width: 253,
        paddingLeft: "20px",
      },
    };
    const iconClass = mergeStyles({
      fontSize: 20,
      height: 15,
      width: 15,
      margin: "6px 0px",
    });
    const classNames = mergeStyleSets({
      deepSkyBlue: [{ color: "deepskyblue" }, iconClass],
    });
    const dialogExcelContentProps = {
      type: DialogType.normal,
      title: "Export to Excel",
      closeButtonAriaLabel: "Close",
      subText:
        "Please select the fields to export to excel. For all fields, please select All Fields or leave it empty.",
    };
    const Content = mergeStyleSets({
      Summary: {
        backgroundColor: "black",
      },
    });
    // const theme = getTheme();
    // const customStyles: Partial<IDetailsRowStyles> = {};
    // customStyles.root = { backgroundColor: theme.palette.themeLighterAlt };
    SPComponentLoader.loadCss(
      "https://collaborate.mcd.com/sites/ContactDatabaseDev/Style%20Library/WorkFlowCss/bootstrap.min.css"
    );
    SPComponentLoader.loadCss(
      "https://collaborate.mcd.com/sites/ContactDatabaseDev/Style%20Library/WorkFlowCss/main.css"
    );
    

    let _ItemSummaryReport = <div style={{ display: "none" }}></div>;
    if (this.state.Item_SummaryReport) {
      _ItemSummaryReport = (
        <div>
          <DetailsList
            items={this.state.TestArray}
            columns={_viewFieldsSummary}
            isHeaderVisible={true}
            className={styles.listViewStyle}
            setKey="Id"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.multiple}
            onRenderRow={this._onListViewRenderRow}
            checkboxVisibility={CheckboxVisibility.onHover}
          />
        </div>
      );
    } else {
      _ItemSummaryReport = <div style={{ display: "none" }}></div>;
    }

    let _ItemWorkFlowReport = <div style={{ display: "none" }}></div>;
    if (this.state.Item_WorkFlowReport) {
      _ItemWorkFlowReport = (
        <div>
          <Stack horizontal tokens={stackTokens} className={styles.topMenuItem}>
            <DateTimePicker
              dateConvention={DateConvention.Date}
              formatDate={(date: Date) => date.toLocaleDateString()}
              showGoToToday={false}
              allowTextInput={false}
              showLabels={false}
              value={this.state.startDate}
              onChange={this.handleStartDate}
              placeholder="Start Date"
            />
            <DateTimePicker
              dateConvention={DateConvention.Date}
              formatDate={(date: Date) => date.toLocaleDateString()}
              showGoToToday={false}
              allowTextInput={false}
              showLabels={false}
              value={this.state.endDate}
              onChange={this.handleEndDate}
              placeholder="End Date"
            />
            <PrimaryButton
              text="Search"
              className="LeftButtonCss"
              onClick={() => this._searchfilter(this.state.dateFilter)}
            />
            <PrimaryButton
              text="Reset"
              className="LeftButtonCss"
              onClick={() => this._resetForm()}
            />
            <CommandButton
              className={styles.Excel}
              iconProps={{ iconName: "ExcelLogoInverse" }}
              text="Export to Excel"
              onClick={this.OnExportToExcelPopupOpen}
            />
          </Stack>
          <ListView
            items={this.state.paginatedItems}
            viewFields={this.state.viewFields}
            compact={true}
            selectionMode={SelectionMode.multiple}
            selection={this._getSelection}
            listClassName={styles.listViewStyle}
            onRenderRow={this._onListViewRenderRow}
            sortItems={this.itemsSortingListView}
          />
          <Pagination
            currentPage={this.state.paginationCurrentPage}
            totalPages={Math.ceil(
              this.state.searchFilterItems.length / this.state.pageSize
            )}
            onChange={(page) => this._getPage(page)}
            limiter={3}
          />
          <div className={styles.showDropdown}>
            <Stack horizontal tokens={stackTokens}>
              <label htmlFor="Show">Show</label>
              <Dropdown
                //label="Show"
                options={[
                  { key: 10, text: "10" },
                  { key: 25, text: "25" },
                  { key: 50, text: "50" },
                ]}
                styles={{ dropdown: { width: 55 } }}
                selectedKey={this.state.pageSize}
                onChange={this.onNumberItemChange}
              />
            </Stack>
          </div>
          <Dialog
            minWidth={500}
            hidden={!this.state.showExportToExcelPopup}
            onDismiss={this.OnDialogPopupClose}
            dialogContentProps={dialogExcelContentProps}
            modalProps={{ isBlocking: true }}
          >
            <DualListBox
              options={this.state.exportToExcelColumnName}
              selected={this.state.exportToExcelSelectedColumnName}
              onChange={this.onDualListExcelChange}
              icons={{
                moveLeft: <FontIcon iconName="ChevronLeftSmall" />,
                moveAllLeft: <FontIcon iconName="DoubleChevronLeft8" />,
                moveRight: <FontIcon iconName="ChevronRightSmall" />,
                moveAllRight: <FontIcon iconName="DoubleChevronRight8" />,
              }}
            />
            <DialogFooter className={styles.dialogFooter}>
              <CSVLink
                data={this.state.exportToExcelFinalValues}
                filename={"Work FLow Report.csv"}
              >
                <PrimaryButton text="Export to Excel" />
              </CSVLink>
              <PrimaryButton onClick={this.OnDialogPopupClose} text="Close" />
            </DialogFooter>
          </Dialog>
        </div>
      );
    } else {
      _ItemWorkFlowReport = <div style={{ display: "none" }}></div>;
    }

    let _ItemBounce = <div style={{ display: "none" }}></div>;
    if (this.state.Item_Bounce) {
      _ItemBounce = (
        <div>
          <Stack horizontal tokens={stackTokens} className={styles.topMenuItem}>
            <DateTimePicker
              dateConvention={DateConvention.Date}
              formatDate={(date: Date) => date.toLocaleDateString()}
              showGoToToday={false}
              allowTextInput={false}
              showLabels={false}
              value={this.state.startDate}
              onChange={this.handleStartDate}
              placeholder="Start Date"
            />
            <DateTimePicker
              dateConvention={DateConvention.Date}
              formatDate={(date: Date) => date.toLocaleDateString()}
              showGoToToday={false}
              allowTextInput={false}
              showLabels={false}
              value={this.state.endDate}
              onChange={this.handleEndDate}
              placeholder="End Date"
            />
            <PrimaryButton
              text="Search"
              className="LeftButtonCss"
              onClick={() => this._searchfilter("BounceBack")}
            />
            <PrimaryButton
              text="Reset"
              className="LeftButtonCss"
              onClick={() => this._resetForm()}
            />
            <CommandButton
              className={styles.Excel}
              iconProps={{ iconName: "ExcelLogoInverse" }}
              text="Export to Excel"
              onClick={this.OnExportToExcelBouncePopupOpen}
            />
          </Stack>
          <ListView
            items={this.state.paginatedItemsBounce}
            viewFields={this.state.viewFieldsBounce}
            compact={true}
            selectionMode={SelectionMode.multiple}
            selection={this._getSelection}
            listClassName={styles.listViewStyle}
            onRenderRow={this._onListViewRenderRow}
            // sortItems={this.itemsSortingListView}
          />
          <Pagination
            currentPage={this.state.paginationCurrentPage}
            totalPages={Math.ceil(
              this.state.searchFilterItemsBounce.length / this.state.pageSize
            )}
            onChange={(page) => this._getPage(page)}
            limiter={3}
          />
          <div className={styles.showDropdown}>
            <Stack horizontal tokens={stackTokens}>
              <label htmlFor="Show">Show</label>
              <Dropdown
                //label="Show"
                options={[
                  { key: 10, text: "10" },
                  { key: 25, text: "25" },
                  { key: 50, text: "50" },
                ]}
                styles={{ dropdown: { width: 55 } }}
                selectedKey={this.state.pageSize}
                onChange={this.onNumberItemChange}
              />
            </Stack>
          </div>
          <Dialog
            minWidth={500}
            hidden={!this.state.showExportToExcelPopup}
            onDismiss={this.OnDialogPopupClose}
            dialogContentProps={dialogExcelContentProps}
            modalProps={{ isBlocking: true }}
          >
            <DualListBox
              options={this.state.exportToExcelColumnNameBounce}
              selected={this.state.exportToExcelSelectedColumnNameBounce}
              onChange={this.onDualListExcelChangeBounce}
              icons={{
                moveLeft: <FontIcon iconName="ChevronLeftSmall" />,
                moveAllLeft: <FontIcon iconName="DoubleChevronLeft8" />,
                moveRight: <FontIcon iconName="ChevronRightSmall" />,
                moveAllRight: <FontIcon iconName="DoubleChevronRight8" />,
              }}
            />
            <DialogFooter className={styles.dialogFooter}>
              <CSVLink
                data={this.state.exportToExcelFinalValues}
                filename={"Work FLow Report.csv"}
              >
                <PrimaryButton text="Export to Excel" />
              </CSVLink>
              <PrimaryButton onClick={this.OnDialogPopupClose} text="Close" />
            </DialogFooter>
          </Dialog>
        </div>
      );
    } else {
      _ItemBounce = <div style={{ display: "none" }}></div>;
    }
    
    return (
      <div className={styles.workFlowReport}>
        <div>
          <h3 className={styles.bottom}>{this.props.webPartTitle}</h3>
        </div>
        <div className="row">
          <div className="form-group col-md-3">
            {/* <label>Work Flow Name</label> */}
            <PrimaryButton
              className={this.state.SummaryColor == true ? styles.Summaryoptionscss: styles.optionscss}
              text="Summary"
              id="Summary"
              onClick={() => this.SummaryReport("Summary")}
            />

            <PrimaryButton
              className={this.state.AllColor == true ? styles.Summaryoptionscss: styles.optionscss}
              text="All"
              onClick={() => this._getitemsDropdown("All")}
            />

            <PrimaryButton
              className={this.state.BounceColor == true ? styles.Summaryoptionscss: styles.optionscss}
              text="Bounce Back"
              onClick={() => this.BounceBack("BounceBack")}
            />
            
            {this.state.managementOptions &&
              this.state.managementOptions.map((options, i) => {
                return [
                  <PrimaryButton
                    id={options.option}
                    className={options.focused ? styles.Summaryoptionscss: styles.optionscss}
                    text={options.option}
                    // onClick={() => this._getitemsDropdown(options.option)}
                    onClick={() => this._getitemsDropdown(options)}
                  />,
                ];
              })}
          </div>
          <div className="form-group col-md-9">
            {_ItemSummaryReport}
            {_ItemWorkFlowReport}
            {_ItemBounce}
          </div>
        </div>
      </div>
    );
  }
}
