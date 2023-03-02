import { WebPartContext } from "@microsoft/sp-webpart-base";
import IService from "../../../common/Services/IServices";

export interface IWorkFlowReportProps {
  userDisplayName: string;
  userEmail: string;
  context: WebPartContext;
  webPartTitle: string;
  listName: string;
  listURL: string;
  BouncelistName: string;
  helperService: IService;
  pageSize: number;
  webPartPageURL: string;
}
