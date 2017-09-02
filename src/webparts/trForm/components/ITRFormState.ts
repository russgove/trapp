import { TRDocument, TR } from "../dataModel";
import * as md from "./MessageDisplay";
export interface ITRFormState {
  tr: TR;
  childTRs: Array<TR>;
  errorMessages: Array<md.Message>;
  isDirty: boolean;
  showTRSearch: boolean;
  documents: Array<TRDocument>;
  documentCalloutVisible: boolean;
  documentCalloutTarget: HTMLElement;
  documentCalloutIframeUrl: string;
  customer;

}